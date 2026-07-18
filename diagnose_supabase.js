import { createClient } from '@supabase/supabase-js'

const url = 'https://ojknvcprfwcrnydkspqp.supabase.co'
const key = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9qa252Y3ByZndjcm55ZGtzcHFwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM2NDM5MDIsImV4cCI6MjA5OTIxOTkwMn0.UpuOPw6WyXQQv-2zYzRdExjg1MzzQvYRdRcTM-wHLUY'

const supabase = createClient(url, key)

async function run() {
  console.log('Querying cau_hinh_khau_hao...')
  const { data: configData, error: configError } = await supabase.from('cau_hinh_khau_hao').select('*')
  if (configError) {
    console.error('Config Error:', configError)
  } else {
    console.log(`Fetched ${configData.length} configs:`)
    configData.forEach(c => {
      console.log(`id: ${c.id}, months: ${c.months}, is_approved: ${c.is_approved}, ma_sap: ${c.ma_sap}`)
    })
  }

  console.log('\nQuerying don_gia_vat_tu...')
  const { data: priceData, error: priceError } = await supabase.from('don_gia_vat_tu').select('*')
  if (priceError) {
    console.error('Price Error:', priceError)
  } else {
    console.log(`Fetched ${priceData.length} price records:`)
    priceData.forEach(p => {
      console.log(p)
    })
  }
}

run()
