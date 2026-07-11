import { createClient } from '@supabase/supabase-js'

const supabaseUrl = 'https://ojknvcprfwcrnydkspqp.supabase.co'
const supabaseAnonKey = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im9qa252Y3ByZndjcm55ZGtzcHFwIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODM2NDM5MDIsImV4cCI6MjA5OTIxOTkwMn0.UpuOPw6WyXQQv-2zYzRdExjg1MzzQvYRdRcTM-wHLUY'

const supabase = createClient(supabaseUrl, supabaseAnonKey)

async function testInsert() {
  const row = {
    ngayXuatNhap: '2026-05-29',
    maVatTu: 'VT001',
    maSAP: 'SAP001',
    thongSoKyThuat: 'test specs',
    tenVatTu: 'Vật tư test',
    dvt: 'Cái',
    loaiDon: 'Đơn giao',
    maDonNhapKho: 'NHAP01',
    maDonXuatKho: 'XUAT01',
    khoiLuongNhap: 10,
    maDonViGiao: 'DV01',
    donViGiao: 'Đơn vị giao test',
    nguoiGiao: 'Phạm Anh Tuấn',
    khoiLuongXuat: 5,
    maDonViNhan: 'DV02',
    donViNhan: 'Đơn vị nhận test',
    nguoiPheDuyet: 'Phạm Anh Tuấn',
    tenNguon: 'Công ty máy móc',
    maNguon: 'NGUON01',
    lo: 'Lô 1',
    hangMuc: 'Hạng mục 1',
    soHopDong: 'HD01',
    thuKho: 'Phạm Anh Tuấn',
    bienSoXe: '29C-1234',
    phanKhu: 'Phân khu A',
    duAn: 'BCH Cọc Khoan Nhồi - Star City Thanh Hóa',
    tinhTrang: 'USED',
    nguoiNhan: 'Phạm Anh Tuấn',
    maDonLienQuan: 'LQ01',
    nhaCungCap: 'NCC 01',
    maDonChuyenTiepLC: 'LC01',
    maDonChuyenTiepNB: 'NB01',
    ghiChu: 'Ghi chú test',
    ghiChuVatTu: 'Ghi chú vật tư test',
    trangThai: 'Đã phê duyệt',
    nhanHieu: 'Airman'
  }

  function toSnakeCase(str) {
    if (str === 'maSAP') return 'ma_sap'
    if (str === 'maDonChuyenTiepLC') return 'ma_don_chuyen_tiep_lc'
    if (str === 'maDonChuyenTiepNB') return 'ma_don_chuyen_tiep_nb'
    return str.replace(/[A-Z]/g, letter => `_${letter.toLowerCase()}`)
  }

  function convertToCasingStyle(row, style) {
    const newRow = {}
    Object.keys(row).forEach(k => {
      if (style === 'snake') {
        newRow[toSnakeCase(k)] = row[k]
      } else if (style === 'lower') {
        newRow[k.toLowerCase()] = row[k]
      } else {
        newRow[k] = row[k]
      }
    })
    return newRow
  }

  console.log('--- Testing inserts ---')
  for (const style of ['snake', 'lower', 'camel']) {
    console.log(`Trying ${style} style on don_giao:`)
    const styled = convertToCasingStyle(row, style)
    const res = await supabase.from('don_giao').insert([styled]).select()
    console.log(`Result ${style}:`, res.error ? `Error: ${res.error.message} (code: ${res.error.code})` : `Success! ID: ${res.data?.[0]?.id}`)
  }
}

testInsert()
