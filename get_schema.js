import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://ojknvcprfwcrnydkspqp.supabase.co'
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9qa252Y3ByZndjcm55ZGtzcHFwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM2NDM5MDIsImV4cCI6MjA5OTIxOTkwMn0.UpuOPw6WyXQQv-2zYzRdExjg1MzzQvYRdRcTM-wHLUY'

async function tryOptions() {
  try {
    let total = 0
    let lastId = 0
    let hasMore = true
    const start = Date.now()
    
    console.log('Fetching all IDs from don_chung using pagination...')
    while (hasMore) {
      const res = await fetch(`${supabaseUrl}/rest/v1/don_chung?select=id&id=gt.${lastId}&order=id.asc&limit=1000`, {
        method: 'GET',
        headers: {
          'apikey': supabaseAnonKey,
          'Authorization': `Bearer ${supabaseAnonKey}`
        }
      })
      if (!res.ok) {
        console.error('Fetch error status:', res.status, await res.text())
        break
      }
      const data = await res.json()
      if (data.length === 0) {
        hasMore = false
      } else {
        total += data.length
        lastId = data[data.length - 1].id
        console.log(`Fetched ${data.length} rows, current total: ${total}, last ID: ${lastId}`)
        if (data.length < 1000) {
          hasMore = false
        }
      }
    }
    console.log(`Total rows in don_chung: ${total} in ${Date.now() - start} ms`)
  } catch (e) {
    console.error('Error:', e)
  }
}

tryOptions()

