import express from 'express';
import path from 'path';
import dotenv from 'dotenv';
import { neon } from '@neondatabase/serverless';
import { createServer as createViteServer } from 'vite';

dotenv.config();

const app = express();
const PORT = 3000;

app.use(express.json({ limit: '50mb' }));

// Helper function to get Neon SQL client
function getSqlClient() {
  const connectionString = process.env.DATABASE_URL || process.env.VITE_SUPABASE_URL;
  if (!connectionString) {
    throw new Error('DATABASE_URL or VITE_SUPABASE_URL environment variable is missing.');
  }
  return neon(connectionString);
}

// RESTful API route to handle direct database operations from our custom client
app.post('/api/db', async (req, res) => {
  try {
    const { action, table, columns, limit, offset, filters, data, functionName, params, count, head, orderColumn, orderAscending } = req.body;
    const sql = getSqlClient();

    // 1. SELECT Action
    if (action === 'select') {
      const values: any[] = [];
      let whereClause = '';
      
      if (filters && Array.isArray(filters) && filters.length > 0) {
        whereClause += ' WHERE ';
        const filterClauses = filters.map((f) => {
          const placeholder = `$${values.length + 1}`;
          values.push(f.value);
          if (f.operator === 'in') {
            return `"${f.column}" = ANY(${placeholder})`;
          }
          const op = f.operator === 'eq' ? '=' : f.operator === 'neq' ? '!=' : '=';
          return `"${f.column}" ${op} ${placeholder}`;
        });
        whereClause += filterClauses.join(' AND ');
      }

      let totalCount = null;
      if (count === 'exact') {
        const countQueryText = `SELECT COUNT(*)::int as count FROM public."${table}"${whereClause}`;
        console.log(`[SQL COUNT] Executing: ${countQueryText} with values:`, values);
        try {
          const countRes = await sql.query(countQueryText, values);
          totalCount = countRes[0]?.count || 0;
        } catch (cntErr: any) {
          console.warn('[SQL COUNT Warning] Failed to fetch exact count:', cntErr.message || cntErr);
        }
      }

      if (head === true) {
        return res.json({ data: [], count: totalCount, error: null });
      }

      let queryText = 'SELECT ';
      // Determine columns
      if (typeof columns === 'string') {
        queryText += columns;
      } else if (Array.isArray(columns)) {
        queryText += columns.map(c => `"${c}"`).join(', ');
      } else {
        queryText += '*';
      }

      queryText += ` FROM public."${table}"${whereClause}`;

      if (orderColumn) {
        const direction = orderAscending !== false ? 'ASC' : 'DESC';
        queryText += ` ORDER BY "${orderColumn}" ${direction}`;
      }

      if (limit !== undefined) {
        queryText += ` LIMIT ${parseInt(limit)}`;
      }
      if (offset !== undefined) {
        queryText += ` OFFSET ${parseInt(offset)}`;
      }

      console.log(`[SQL SELECT] Executing: ${queryText} with values:`, values);
      const result = await sql.query(queryText, values);
      return res.json({ data: result, count: totalCount, error: null });
    }

    // 2. INSERT Action
    if (action === 'insert') {
      if (!data) {
        return res.status(400).json({ data: null, error: 'No data provided for insert' });
      }

      const rowsToInsert = Array.isArray(data) ? data : [data];
      if (rowsToInsert.length === 0) {
        return res.json({ data: [], error: null });
      }

      // Collect all distinct keys from the rows to form the columns list
      const allKeys = new Set<string>();
      rowsToInsert.forEach(row => {
        Object.keys(row).forEach(k => {
          if (row[k] !== undefined) {
            allKeys.add(k);
          }
        });
      });

      const columnsList = Array.from(allKeys);
      if (columnsList.length === 0) {
        return res.status(400).json({ data: null, error: 'Data objects contain no valid keys' });
      }

      // Map keys to double quoted columns
      const colsSql = columnsList.map(c => `"${c}"`).join(', ');

      const values: any[] = [];
      const rowPlaceholders: string[] = [];

      rowsToInsert.forEach(row => {
        const placeholders = columnsList.map(col => {
          const val = row[col];
          const actualVal = val === undefined ? null : val;
          values.push(actualVal);
          return `$${values.length}`;
        });
        rowPlaceholders.push(`(${placeholders.join(', ')})`);
      });

      const queryText = `INSERT INTO public."${table}" (${colsSql}) VALUES ${rowPlaceholders.join(', ')} RETURNING *`;
      console.log(`[SQL INSERT] Executing insert to ${table} for ${rowsToInsert.length} rows`);
      
      const result = await sql.query(queryText, values);
      return res.json({ data: result, error: null });
    }

    // 3. DELETE Action
    if (action === 'delete') {
      let queryText = `DELETE FROM public."${table}"`;
      const values: any[] = [];

      if (filters && Array.isArray(filters) && filters.length > 0) {
        queryText += ' WHERE ';
        const filterClauses = filters.map((f, idx) => {
          const placeholder = `$${values.length + 1}`;
          values.push(f.value);
          const op = f.operator === 'eq' ? '=' : f.operator === 'neq' ? '!=' : '=';
          return `"${f.column}" ${op} ${placeholder}`;
        });
        queryText += filterClauses.join(' AND ');
      }

      queryText += ' RETURNING *';

      console.log(`[SQL DELETE] Executing: ${queryText} with values:`, values);
      const result = await sql.query(queryText, values);
      return res.json({ data: result, error: null });
    }

    // 4. RPC (Function call) Action
    if (action === 'rpc') {
      if (functionName === 'reset_sequence_to_one' || functionName === 'sync_sequence_to_max') {
        const pTable = params?.p_table;
        if (!pTable) {
          return res.status(400).json({ data: null, error: 'Missing p_table parameter' });
        }
        
        let queryText = '';
        if (functionName === 'reset_sequence_to_one') {
          queryText = `SELECT setval(pg_get_serial_sequence('public."${pTable}"', 'id'), 1, false)`;
        } else {
          queryText = `SELECT setval(pg_get_serial_sequence('public."${pTable}"', 'id'), COALESCE((SELECT MAX(id) FROM public."${pTable}"), 1), true)`;
        }

        try {
          console.log(`[SQL RPC] Running sequence alignment for table: ${pTable}`);
          const result = await sql.query(queryText);
          return res.json({ data: result, error: null });
        } catch (rpcErr: any) {
          console.warn('[SQL RPC Warning] Failed to reset sequence (might not have serial id/sequence):', rpcErr.message || rpcErr);
          return res.json({ data: null, error: null });
        }
      }

      return res.status(400).json({ data: null, error: `RPC function '${functionName}' is not supported` });
    }

    return res.status(400).json({ data: null, error: `Action '${action}' is not supported` });
  } catch (err: any) {
    console.error('[SQL Error] General failure:', err.message || err);
    return res.json({ data: null, error: { message: err.message || String(err) } });
  }
});

async function startServer() {
  // Auto-cleanup and maintain only the tables requested by the user
  try {
    const sql = getSqlClient();
    console.log('[Neon Cleanup] Dropping unwanted tables as requested...');
    
    // Actively drop the unwanted tables
    await sql.query('DROP TABLE IF EXISTS public.don_giao CASCADE;');
    await sql.query('DROP TABLE IF EXISTS public.don_nhan CASCADE;');
    await sql.query('DROP TABLE IF EXISTS public.don_kho CASCADE;');
    await sql.query('DROP TABLE IF EXISTS public.du_an CASCADE;');
    
    console.log('[Neon Cleanup] Cleaned up unwanted tables successfully!');
  } catch (err: any) {
    console.error('[Neon Cleanup Error] Failed to drop tables:', err.message || err);
  }

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), 'dist');
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server is running on http://0.0.0.0:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
});
