import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://ojknvcprfwcrnydkspqp.supabase.co'
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9qa252Y3ByZndjcm55ZGtzcHFwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM2NDM5MDIsImV4cCI6MjA5OTIxOTkwMn0.UpuOPw6WyXQQv-2zYzRdExjg1MzzQvYRdRcTM-wHLUY'

const supabase = createClient(supabaseUrl, supabaseAnonKey)

async function run() {
  console.log('--- INSERTING AND QUERYING COLS FOR TYPE DETECTIONS ---')
  
  const sampleData = {
    ngay_xuat_nhap: '2026-05-29',
    ma_vat_tu: 'VT002',
    ma_sap: 'SAP002',
    ten_vat_tu: 'Test types lookup',
    khoi_luong_nhap: 12.34,
    khoi_luong_xuat: 56.78,
    du_an: 'BCH Cọc Khoan Nhồi - Star City Thanh Hóa',
    // Let's add extra string values
    thong_so_ky_thuat: 'specs',
    dvt: 'Cái',
    loai_don: 'Đơn giao',
    trang_thai: 'Đã phê duyệt'
  }
  
  // Insert
  const { data: insData, error: insErr } = await supabase
    .from('don_giao')
    .insert([sampleData])
    .select('*')
    
  if (insErr) {
    console.error('Insert error on don_giao:', insErr)
    return
  }
  
  if (insData && insData.length > 0) {
    const row = insData[0]
    console.log('Successfully inserted row! Columns and Types:')
    Object.keys(row).forEach(key => {
      console.log(`- ${key}: val=${JSON.stringify(row[key])} typ=${typeof row[key]}`)
    })
    
    // Clean up
    await supabase.from('don_giao').delete().eq('ma_vat_tu', 'VT002')
  }
}

run()
