// Columns for Đơn Giao & Đơn Nhận (same structure, different label for upload)
export const COLS_GIAO_NHAN = [
  { key: 'ngayXuatNhap', label: 'Ngày xuất nhập', width: 110 },
  { key: 'maVatTu', label: 'Mã vật tư', width: 100 },
  { key: 'maSAP', label: 'Mã SAP', width: 110 },
  { key: 'thongSoKyThuat', label: 'Thông số KT', width: 130 },
  { key: 'tenVatTu', label: 'Tên vật tư', width: 220 },
  { key: 'dvt', label: 'ĐVT', width: 60 },
  { key: 'loaiDon', label: 'Loại đơn', width: 130 },
  { key: 'maDonNhapKho', label: 'Mã đơn nhập kho', width: 280 },
  { key: 'maDonXuatKho', label: 'Mã đơn xuất kho', width: 280 },
  // Nhập kho group
  { key: 'khoiLuongNhap', label: 'KL nhập', width: 80, group: 'Nhập kho' },
  { key: 'maDonViGiao', label: 'Mã đơn vị giao', width: 130, group: 'Nhập kho' },
  { key: 'donViGiao', label: 'Đơn vị giao', width: 200, group: 'Nhập kho' },
  { key: 'nguoiGiao', label: 'Người giao', width: 150, group: 'Nhập kho' },
  // Xuất kho group
  { key: 'khoiLuongXuat', label: 'KL xuất', width: 80, group: 'Xuất kho' },
  { key: 'maDonViNhan', label: 'Mã đơn vị nhận', width: 130, group: 'Xuất kho' },
  { key: 'donViNhan', label: 'Đơn vị nhận', width: 200, group: 'Xuất kho' },
  { key: 'nguoiPheDuyet', label: 'Người phê duyệt', width: 150, group: 'Xuất kho' },
  // Thông tin chung group
  { key: 'tenNguon', label: 'Tên nguồn xuất/nhập', width: 300, group: 'Thông tin chung' },
  { key: 'maNguon', label: 'Mã nguồn', width: 200, group: 'Thông tin chung' },
  { key: 'lo', label: 'Lô', width: 80, group: 'Thông tin chung' },
  { key: 'hangMuc', label: 'Hạng mục', width: 100, group: 'Thông tin chung' },
  { key: 'soHopDong', label: 'Số HĐ tổ đội', width: 150, group: 'Thông tin chung' },
  { key: 'thuKho', label: 'Thủ kho BCH', width: 200, group: 'Thông tin chung' },
  { key: 'bienSoXe', label: 'Biển số xe', width: 100, group: 'Thông tin chung' },
  { key: 'phanKhu', label: 'Phân khu', width: 130, group: 'Thông tin chung' },
  { key: 'duAn', label: 'Dự án', width: 200, group: 'Thông tin chung' },
  { key: 'tinhTrang', label: 'Tình trạng VT', width: 120, group: 'Thông tin chung' },
  { key: 'nguoiNhan', label: 'Người nhận', width: 200, group: 'Thông tin chung' },
  { key: 'maDonLienQuan', label: 'Mã đơn liên quan', width: 280, group: 'Thông tin chung' },
  { key: 'nhaCungCap', label: 'Nhà cung cấp', width: 280, group: 'Thông tin chung' },
  { key: 'maDonChuyenTiepLC', label: 'Mã ĐCT liên công ty', width: 200, group: 'Thông tin chung' },
  { key: 'maDonChuyenTiepNB', label: 'Mã ĐCT nội bộ', width: 200, group: 'Thông tin chung' },
  { key: 'ghiChu', label: 'Ghi chú', width: 300, group: 'Thông tin chung' },
  { key: 'ghiChuVatTu', label: 'Ghi chú VT', width: 200, group: 'Thông tin chung' },
  { key: 'trangThai', label: 'Trạng thái', width: 130, group: 'Thông tin chung' },
  { key: 'nhanHieu', label: 'Nhãn hiệu', width: 120, group: 'Thông tin chung' },
]

export function parseXlsxToRows(data) {
  // data is array of arrays (from XLSX.utils.sheet_to_json with header:1)
  // Row 0 = group headers, Row 1 = sub-headers, Row 2+ = data
  if (!data || data.length < 3) return []
  const rows = []
  for (let i = 2; i < data.length; i++) {
    const r = data[i]
    if (!r || r.every(v => v === null || v === undefined || v === '')) continue
    rows.push({
      id: i,
      ngayXuatNhap: r[0] ?? '',
      maVatTu: r[1] ?? '',
      maSAP: r[2] ?? '',
      thongSoKyThuat: r[3] ?? '',
      tenVatTu: r[4] ?? '',
      dvt: r[5] ?? '',
      loaiDon: r[6] ?? '',
      maDonNhapKho: r[7] ?? '',
      maDonXuatKho: r[8] ?? '',
      khoiLuongNhap: r[9] ?? '',
      maDonViGiao: r[10] ?? '',
      donViGiao: r[11] ?? '',
      nguoiGiao: r[12] ?? '',
      khoiLuongXuat: r[13] ?? '',
      maDonViNhan: r[14] ?? '',
      donViNhan: r[15] ?? '',
      nguoiPheDuyet: r[16] ?? '',
      tenNguon: r[17] ?? '',
      maNguon: r[18] ?? '',
      lo: r[19] ?? '',
      hangMuc: r[20] ?? '',
      soHopDong: r[21] ?? '',
      thuKho: r[22] ?? '',
      bienSoXe: r[23] ?? '',
      phanKhu: r[24] ?? '',
      duAn: r[25] ?? '',
      tinhTrang: r[26] ?? '',
      nguoiNhan: r[27] ?? '',
      maDonLienQuan: r[28] ?? '',
      nhaCungCap: r[29] ?? '',
      maDonChuyenTiepLC: r[30] ?? '',
      maDonChuyenTiepNB: r[31] ?? '',
      ghiChu: r[32] ?? '',
      ghiChuVatTu: r[33] ?? '',
      trangThai: r[34] ?? '',
      nhanHieu: r[35] ?? '',
    })
  }
  return rows
}

export function formatVal(val) {
  if (val === null || val === undefined) return ''
  if (typeof val === 'number') {
    // Check if it's an Excel date serial (between 1 and 50000)
    if (val > 1 && val < 50000 && Number.isInteger(val)) {
      // Could be date, return as-is for now  
      return val
    }
    return val.toLocaleString('vi-VN')
  }
  return String(val)
}

export function getTrangThaiColor(val) {
  if (!val) return 'badge-gray'
  const v = String(val).toLowerCase()
  if (v.includes('phê duyệt') || v.includes('hoàn thành') || v.includes('đã')) return 'badge-green'
  if (v.includes('từ chối') || v.includes('hủy')) return 'badge-red'
  if (v.includes('chờ') || v.includes('chưa')) return 'badge-yellow'
  return 'badge-blue'
}
