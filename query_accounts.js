import { createClient } from '@supabase/supabase-js'

const url = 'https://ojknvcprfwcrnydkspqp.supabase.co'
const key = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9qa252Y3ByZndjcm55ZGtzcHFwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM2NDM5MDIsImV4cCI6MjA5OTIxOTkwMn0.UpuOPw6WyXQQv-2zYzRdExjg1MzzQvYRdRcTM-wHLUY'

const supabase = createClient(url, key)

async function run() {
  console.log('--- 1. Testing checkSupabaseTableExists logic ---')
  const countRes = await supabase.from('quan_ly_tai_khoan').select('id', { count: 'exact', head: true })
  console.log('Count Result:', JSON.stringify(countRes))

  console.log('--- 2. Testing select * ---')
  const selectRes = await supabase.from('quan_ly_tai_khoan').select('*')
  console.log('Select * Result:', JSON.stringify(selectRes))
}

run()
