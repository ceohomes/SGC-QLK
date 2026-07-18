import { createClient } from '@supabase/supabase-js'

const url = 'https://ojknvcprfwcrnydkspqp.supabase.co'
const key = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9qa252Y3ByZndjcm55ZGtzcHFwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM2NDM5MDIsImV4cCI6MjA5OTIxOTkwMn0.UpuOPw6WyXQQv-2zYzRdExjg1MzzQvYRdRcTM-wHLUY'

const supabase = createClient(url, key)

async function run() {
  console.log('--- Trying to insert a record into quan_ly_tai_khoan ---')
  const res = await supabase.from('quan_ly_tai_khoan').insert([{
    ho_ten: 'Test Account',
    ten_dang_nhap: 'test_insert',
    mat_khau: '123456',
    quyen: 'User'
  }]).select()
  
  console.log('Insert Result:', JSON.stringify(res))
}

run()
