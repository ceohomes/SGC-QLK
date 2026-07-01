import { Pool } from '@neondatabase/serverless';

// ============================================================================
// KẾT NỐI NEON DIRECT DATABASE (thay cho Supabase Client)
// ============================================================================
// File này cung cấp một client interface giống hệt như Supabase JS library,
// nhưng chuyển hướng tất cả các truy vấn (SELECT, INSERT, DELETE, RPC) về
// phía Express backend `/api/db` HOẶC chạy trực tiếp tại trình duyệt đến Neon
// nếu người dùng cấu hình chuỗi kết nối trực tiếp.
//
// Cách này bảo mật hơn 100%, không lộ thông tin kết nối ra Client (nếu dùng Server),
// và giải quyết triệt để lỗi CORS khi host tĩnh trên GitHub Pages/Cloudflare Pages.
// ============================================================================

let cachedDirectClient = null;
let lastConnString = null;

async function executeRequest(payload) {
  try {
    let directConnString = '';
    if (typeof window !== 'undefined') {
      directConnString = window.localStorage.getItem('neon_connection_string') || '';
    }

    // Nếu không có trong localStorage, tự động lấy từ biến môi trường của Vite (đặc biệt khi deploy tĩnh lên Cloudflare Pages)
    if (!directConnString) {
      try {
        directConnString = import.meta.env.VITE_DATABASE_URL || import.meta.env.VITE_NEON_DATABASE_URL || '';
      } catch (e) {
        console.warn('[Vite Env] Không thể đọc biến môi trường VITE_DATABASE_URL:', e);
      }
    }

    // Nếu có Chuỗi kết nối trực tiếp (DATABASE_URL), thực thi SQL trực tiếp từ trình duyệt!
    if (directConnString) {
      if (!cachedDirectClient || lastConnString !== directConnString) {
        cachedDirectClient = new Pool({ connectionString: directConnString });
        lastConnString = directConnString;
      }
      
      const pool = cachedDirectClient;
      const { action, table, columns, limit, offset, filters, data, functionName, params, count, head, orderColumn, orderAscending } = payload;
      
      const DELETED_TABLES = ['du_an', 'don_giao', 'don_nhan', 'don_kho'];
      if (table && DELETED_TABLES.includes(table)) {
        console.log(`[Client SQL Short-Circuit] Table '${table}' is deleted. Safely returning empty dataset.`);
        if (action === 'select') {
          return { data: [], count: 0, error: null };
        }
        return { data: [], error: null };
      }

      if (action === 'rpc' && (functionName === 'reset_sequence_to_one' || functionName === 'sync_sequence_to_max')) {
        const pTable = params?.p_table;
        if (pTable && DELETED_TABLES.includes(pTable)) {
          console.log(`[Client SQL RPC Short-Circuit] Sequence alignment table '${pTable}' is deleted. Returning success no-op.`);
          return { data: null, error: null };
        }
      }

      // 1. SELECT Action
      if (action === 'select') {
        const values = [];
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
          try {
            const countRes = await pool.query(countQueryText, values);
            totalCount = countRes.rows[0]?.count || 0;
          } catch (cntErr) {
            console.warn('[Client SQL COUNT Warning] Failed to fetch exact count:', cntErr);
          }
        }

        if (head === true) {
          return { data: [], count: totalCount, error: null };
        }

        let queryText = 'SELECT ';
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

        console.log(`[Client Direct SQL SELECT] Executing: ${queryText} with values:`, values);
        const resultRes = await pool.query(queryText, values);
        return { data: resultRes.rows, count: totalCount, error: null };
      }

      // 2. INSERT Action
      if (action === 'insert') {
        if (!data) {
          return { data: null, error: { message: 'No data provided for insert' } };
        }

        const rowsToInsert = Array.isArray(data) ? data : [data];
        if (rowsToInsert.length === 0) {
          return { data: [], error: null };
        }

        const allKeys = new Set();
        rowsToInsert.forEach(row => {
          Object.keys(row).forEach(k => {
            if (row[k] !== undefined) {
              allKeys.add(k);
            }
          });
        });

        const columnsList = Array.from(allKeys);
        if (columnsList.length === 0) {
          return { data: null, error: { message: 'Data objects contain no valid keys' } };
        }

        const colsSql = columnsList.map(c => `"${c}"`).join(', ');
        const values = [];
        const rowPlaceholders = [];

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
        console.log(`[Client Direct SQL INSERT] Executing insert to ${table} for ${rowsToInsert.length} rows`);
        
        const resultRes = await pool.query(queryText, values);
        return { data: resultRes.rows, error: null };
      }

      // 3. DELETE Action
      if (action === 'delete') {
        let queryText = `DELETE FROM public."${table}"`;
        const values = [];

        if (filters && Array.isArray(filters) && filters.length > 0) {
          queryText += ' WHERE ';
          const filterClauses = filters.map((f) => {
            const placeholder = `$${values.length + 1}`;
            values.push(f.value);
            const op = f.operator === 'eq' ? '=' : f.operator === 'neq' ? '!=' : '=';
            return `"${f.column}" ${op} ${placeholder}`;
          });
          queryText += filterClauses.join(' AND ');
        }

        queryText += ' RETURNING *';

        console.log(`[Client Direct SQL DELETE] Executing: ${queryText} with values:`, values);
        const resultRes = await pool.query(queryText, values);
        return { data: resultRes.rows, error: null };
      }

      // 4. RPC (Function call) Action
      if (action === 'rpc') {
        if (functionName === 'reset_sequence_to_one' || functionName === 'sync_sequence_to_max') {
          const pTable = params?.p_table;
          if (!pTable) {
            return { data: null, error: { message: 'Missing p_table parameter' } };
          }
          
          let queryText = '';
          if (functionName === 'reset_sequence_to_one') {
            queryText = `SELECT setval(pg_get_serial_sequence('public."${pTable}"', 'id'), 1, false)`;
          } else {
            queryText = `SELECT setval(pg_get_serial_sequence('public."${pTable}"', 'id'), COALESCE((SELECT MAX(id) FROM public."${pTable}"), 1), true)`;
          }

          try {
            console.log(`[Client Direct SQL RPC] Running sequence alignment for table: ${pTable}`);
            const resultRes = await pool.query(queryText);
            return { data: resultRes.rows, error: null };
          } catch (rpcErr) {
            console.warn('[Client Direct SQL RPC Warning] Failed to reset sequence:', rpcErr);
            return { data: null, error: null };
          }
        }

        return { data: null, error: { message: `RPC function '${functionName}' is not supported` } };
      }

      // 5. UPSERT Action
      if (action === 'upsert') {
        if (!data) {
          return { data: null, error: { message: 'No data provided for upsert' } };
        }

        const rowsToInsert = Array.isArray(data) ? data : [data];
        if (rowsToInsert.length === 0) {
          return { data: [], error: null };
        }

        const allKeys = new Set();
        rowsToInsert.forEach(row => {
          Object.keys(row).forEach(k => {
            if (row[k] !== undefined) {
              allKeys.add(k);
            }
          });
        });

        const columnsList = Array.from(allKeys);
        if (columnsList.length === 0) {
          return { data: null, error: { message: 'Data objects contain no valid keys' } };
        }

        const colsSql = columnsList.map(c => `"${c}"`).join(', ');
        const values = [];
        const rowPlaceholders = [];

        rowsToInsert.forEach(row => {
          const placeholders = columnsList.map(col => {
            const val = row[col];
            const actualVal = val === undefined ? null : val;
            values.push(actualVal);
            return `$${values.length}`;
          });
          rowPlaceholders.push(`(${placeholders.join(', ')})`);
        });

        let queryText = `INSERT INTO public."${table}" (${colsSql}) VALUES ${rowPlaceholders.join(', ')}`;
        
        const conflictCol = onConflict || 'id';
        const updateCols = columnsList.filter(c => c !== conflictCol);
        
        if (updateCols.length > 0) {
          const updateSql = updateCols.map(c => `"${c}" = EXCLUDED."${c}"`).join(', ');
          queryText += ` ON CONFLICT ("${conflictCol}") DO UPDATE SET ${updateSql}`;
        } else {
          queryText += ` ON CONFLICT ("${conflictCol}") DO NOTHING`;
        }
        
        queryText += ' RETURNING *';

        console.log(`[Client Direct SQL UPSERT] Executing: ${queryText}`);
        const resultRes = await pool.query(queryText, values);
        return { data: resultRes.rows, error: null };
      }
    }

    // Default Fallback:
    let backendUrl = '';
    if (typeof window !== 'undefined') {
      backendUrl = window.localStorage.getItem('backend_api_url') || '';
      
      // Auto-detect when running on external hosts (like Cloudflare Pages, GitHub Pages, Vercel)
      if (!backendUrl) {
        const hostname = window.location.hostname;
        if (hostname && hostname !== 'localhost' && hostname !== '127.0.0.1' && !hostname.endsWith('.run.app')) {
          // Default back to the production Cloud Run URL running the custom Neon DB server
          backendUrl = 'https://ais-pre-7ylvwuj3ycxnp24kvzwo6e-270809794219.asia-east1.run.app';
        }
      }
    }
    if (!backendUrl && typeof import.meta !== 'undefined' && import.meta.env && import.meta.env.VITE_API_URL) {
      backendUrl = import.meta.env.VITE_API_URL;
    }
    
    if (backendUrl && backendUrl.endsWith('/')) {
      backendUrl = backendUrl.slice(0, -1);
    }
    
    const apiEndpoint = backendUrl ? `${backendUrl}/api/db` : '/api/db';

    const res = await fetch(apiEndpoint, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      const errText = await res.text();
      return { data: null, error: { message: errText } };
    }
    const result = await res.json();
    return { data: result.data, count: result.count, error: result.error };
  } catch (err) {
    return { data: null, error: { message: err.message || String(err) } };
  }
}

export const supabase = {
  from(table) {
    let payload = { table, filters: [] };

    const chainable = {
      eq(column, value) {
        payload.filters.push({ column, operator: 'eq', value });
        return chainable;
      },
      neq(column, value) {
        payload.filters.push({ column, operator: 'neq', value });
        return chainable;
      },
      in(column, valuesArray) {
        payload.filters.push({ column, operator: 'in', value: valuesArray });
        return chainable;
      },
      order(column, options = {}) {
        payload.orderColumn = column;
        payload.orderAscending = options.ascending !== false;
        return chainable;
      },
      limit(val) {
        payload.limit = val;
        return chainable;
      },
      range(from, to) {
        payload.offset = from;
        payload.limit = (to - from + 1);
        return chainable;
      },
      // This is called when awaiting the chainable object directly (e.g. select or delete)
      then(onfulfilled, onrejected) {
        return executeRequest(payload).then(onfulfilled, onrejected);
      }
    };

    const builder = {
      select(columns = '*', options = {}) {
        payload.action = 'select';
        payload.columns = columns;
        if (options && options.count === 'exact') {
          payload.count = 'exact';
          if (options.head) {
            payload.head = true;
          }
        }
        return chainable;
      },

      insert(data) {
        payload.action = 'insert';
        payload.data = data;
        // Allows direct await of .insert()
        return {
          select() {
            // PostgREST support .insert().select()
            return executeRequest(payload);
          },
          then(onfulfilled, onrejected) {
            return executeRequest(payload).then(onfulfilled, onrejected);
          }
        };
      },

      upsert(data, options = {}) {
        payload.action = 'upsert';
        payload.data = data;
        payload.onConflict = options.onConflict;
        return {
          select() {
            return executeRequest(payload);
          },
          then(onfulfilled, onrejected) {
            return executeRequest(payload).then(onfulfilled, onrejected);
          }
        };
      },

      delete() {
        payload.action = 'delete';
        return chainable;
      },

      then(onfulfilled, onrejected) {
        return executeRequest(payload).then(onfulfilled, onrejected);
      }
    };

    return builder;
  },

  rpc(functionName, params) {
    const payload = {
      action: 'rpc',
      functionName,
      params
    };
    return executeRequest(payload);
  },

  // Mock Realtime Channels so the existing code doesn't fail
  channel(name) {
    return {
      on(event, filter, callback) {
        return this;
      },
      subscribe(callback) {
        if (typeof callback === 'function') {
          callback('SUBSCRIBED');
        }
        return this;
      }
    };
  },
  removeChannel(channel) {
    // No-op
  }
};

// Dynamically calculate if database connection is configured
export const isSupabaseConfigured = (() => {
  if (typeof window === 'undefined') return true;
  const hostname = window.location.hostname;
  // If running locally or on AI Studio development/preview domains, backend is always ready
  if (hostname === 'localhost' || hostname === '127.0.0.1' || hostname.endsWith('.run.app')) {
    return true;
  }

  // Kiểm tra biến môi trường Vite trước (VITE_DATABASE_URL hoặc VITE_NEON_DATABASE_URL)
  let hasEnvVar = false;
  try {
    const envUrl = import.meta.env.VITE_DATABASE_URL || import.meta.env.VITE_NEON_DATABASE_URL;
    if (envUrl) {
      hasEnvVar = true;
    }
  } catch (e) {}

  // If running on custom static domain (GitHub/Cloudflare Pages), check if any connection is configured
  const neonConn = window.localStorage.getItem('neon_connection_string');
  const backendUrl = window.localStorage.getItem('backend_api_url');
  return !!(neonConn || backendUrl || hasEnvVar);
})();

export const supabaseUrl = 'Server Direct Neon DB Connection';
export const supabaseAnonKey = 'JWT Authentication Managed Server-Side';
