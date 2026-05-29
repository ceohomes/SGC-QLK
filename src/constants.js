// Columns for Đơn Giao & Đơn Nhận (same structure, different label for upload)
export const COLS_GIAO_NHAN = [
  { key: 'ngayXuatNhap', label: 'Ngày xuất nhập', width: 85 },
  { key: 'maVatTu', label: 'Mã vật tư', width: 80 },
  { key: 'maSAP', label: 'Mã SAP', width: 80 },
  { key: 'thongSoKyThuat', label: 'Thông số kỹ thuật', width: 160 },
  { key: 'tenVatTu', label: 'Tên vật tư', width: 160 },
  { key: 'dvt', label: 'ĐVT', width: 60 },
  { key: 'loaiDon', label: 'Loại đơn (Xuất, nhập kho)', width: 80 },
  { key: 'maDonNhapKho', label: 'Mã đơn nhập kho', width: 200 },
  { key: 'maDonXuatKho', label: 'Mã đơn xuất kho', width: 200 },
  { key: 'khoiLuongNhap', label: 'Khối lượng nhập', width: 80 },
  { key: 'maDonViGiao', label: 'Mã đơn vị giao', width: 120 },
  { key: 'donViGiao', label: 'Đơn vị giao', width: 120 },
  { key: 'nguoiGiao', label: 'Người giao', width: 120 },
  { key: 'khoiLuongXuat', label: 'Khối lượng xuất', width: 80 },
  { key: 'maDonViNhan', label: 'Mã đơn vị nhận', width: 120 },
  { key: 'donViNhan', label: 'Đơn vị nhận', width: 120 },
  { key: 'nguoiPheDuyet', label: 'Người phê duyệt 1', width: 120 },
  { key: 'tenNguon', label: 'Tên nguồn xuất/ nhập kho', width: 220 },
  { key: 'maNguon', label: 'Mã nguồn nhập/xuất kho', width: 120 },
  { key: 'lo', label: 'Lô', width: 80 },
  { key: 'hangMuc', label: 'Hạng mục', width: 100 },
  { key: 'soHopDong', label: 'Số Hợp đồng tổ đội', width: 120 },
  { key: 'thuKho', label: 'Thủ kho BCH Giao/Nhận hàng', width: 120 },
  { key: 'bienSoXe', label: 'Biển số xe', width: 100 },
  { key: 'phanKhu', label: 'Phân khu', width: 130 },
  { key: 'duAn', label: 'Dự án', width: 200 },
  { key: 'tinhTrang', label: 'Tình trạng vật tư', width: 80 },
  { key: 'nguoiNhan', label: 'Người nhận', width: 120 },
  { key: 'maDonLienQuan', label: 'Mã đơn liên quan', width: 120 },
  { key: 'nhaCungCap', label: 'Nhà cung cấp', width: 120 },
  { key: 'maDonChuyenTiepLC', label: 'Mã đơn chuyển tiếp liên công ty', width: 120 },
  { key: 'maDonChuyenTiepNB', label: 'Mã đơn chuyển tiếp nội bộ', width: 120 },
  { key: 'ghiChu', label: 'Ghi chú', width: 220 },
  { key: 'ghiChuVatTu', label: 'Ghi chú vật tư', width: 180 },
  { key: 'trangThai', label: 'Trạng Thái', width: 130 },
  { key: 'nhanHieu', label: 'Nhãn hiệu', width: 120 },
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

export function formatVal(val, colKey) {
  if (val === null || val === undefined) return ''
  if (typeof val === 'number') {
    // Các cột khối lượng/số lượng: luôn format với dấu phân cách hàng nghìn
    const isQuantityCol = colKey && (
      colKey.toLowerCase().includes('khoiluong') ||
      colKey.toLowerCase().includes('soluong') ||
      colKey.toLowerCase().includes('quantity') ||
      colKey === 'khoiLuongNhap' ||
      colKey === 'khoiLuongXuat'
    )
    if (isQuantityCol) {
      // Format số với dấu phân cách hàng nghìn
      if (Number.isInteger(val)) return val.toLocaleString('vi-VN')
      return val.toLocaleString('vi-VN', { maximumFractionDigits: 3 })
    }
    // Check if it's an Excel date serial (between 1 and 50000, integer)
    if (val > 1 && val < 50000 && Number.isInteger(val)) {
      // Could be date, return as-is for now  
      return val
    }
    return val.toLocaleString('vi-VN')
  }
  // Nếu là chuỗi số (string) trong cột khối lượng, thử parse và format
  if (typeof val === 'string' && colKey && (
    colKey.toLowerCase().includes('khoiluong') ||
    colKey === 'khoiLuongNhap' ||
    colKey === 'khoiLuongXuat'
  )) {
    const cleaned = val.replace(/[^\d.,]/g, '').replace(',', '.')
    const num = parseFloat(cleaned)
    if (!isNaN(num)) {
      if (Number.isInteger(num)) return num.toLocaleString('vi-VN')
      return num.toLocaleString('vi-VN', { maximumFractionDigits: 3 })
    }
  }
  return String(val)
}

export function isApprovedStatus(statusStr) {
  if (!statusStr) return false
  const s = String(statusStr).toLowerCase()
  // Exclude explicit pending or negative status
  if (s.includes('chờ') || s.includes('chưa') || s.includes('pending')) return false
  if (s.includes('từ chối') || s.includes('hủy') || s.includes('reject') || s.includes('không')) return false
  // Include valid approved indications
  return s.includes('đã') || s.includes('duyệt') || s.includes('approved') || s.includes('hoàn thành')
}

export function isPendingStatus(statusStr) {
  if (!statusStr) return false
  const s = String(statusStr).toLowerCase()
  return s.includes('chờ') || s.includes('chưa') || s.includes('pending')
}

export function isRejectedStatus(statusStr) {
  if (!statusStr) return false
  const s = String(statusStr).toLowerCase()
  return s.includes('từ chối') || s.includes('hủy') || s.includes('reject')
}

export function getTrangThaiColor(val) {
  if (!val) return 'badge-gray'
  if (isRejectedStatus(val)) return 'badge-red'
  if (isPendingStatus(val)) return 'badge-yellow'
  if (isApprovedStatus(val)) return 'badge-green'
  return 'badge-blue'
}
