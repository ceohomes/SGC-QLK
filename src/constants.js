// Columns for Đơn Giao & Đơn Nhận (same structure, different label for upload)
export const COLS_GIAO_NHAN = [
  { key: 'ngayXuatNhap', label: 'Ngày xuất nhập', width: 85 },
  { key: 'maVatTu', label: 'Mã vật tư', width: 80 },
  { key: 'maSAP', label: 'Mã SAP', width: 80 },
  { key: 'thongSoKyThuat', label: 'Thông số kỹ thuật', width: 60 },
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
  // Row 1 (index 0) = empty/title
  // Row 2 (index 1) = report title
  // Row 3 (index 2) = empty
  // Row 4 (index 3) = group headers
  // Row 5 (index 4) = sub-headers
  // Row 6 (index 5) onwards = actual data
  if (!data || data.length === 0) return []

  // Always start strictly from index 5 (Row 6) as per request: "lấy dữ liệu từ dòng số 6 trở đi"
  const startIndex = 5

  // Helper to check if a value looks like a date or serial date
  function isDateValue(val) {
    if (val === null || val === undefined) return false
    const s = String(val).trim()
    if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4}$/.test(s)) return true
    if (/^\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}$/.test(s)) return true
    const num = Number(val)
    if (!isNaN(num) && num > 30000 && num < 60000) return true
    return false
  }

  let offset = 0
  // Method 1: Check header row 4 (index 3) and row 5 (index 4)
  for (const hRowIdx of [3, 4]) {
    const hRow = data[hRowIdx]
    if (hRow && Array.isArray(hRow)) {
      const idx = hRow.findIndex(cell => {
        if (!cell) return false
        const s = String(cell).toLowerCase().trim()
        return s.includes('ngày') || s.includes('ngay') || s.includes('date')
      })
      if (idx !== -1) {
        offset = idx
        break
      }
    }
  }

  // Method 2 (Fallback): Check first few data rows
  if (offset === 0) {
    for (let i = startIndex; i < Math.min(data.length, startIndex + 5); i++) {
      const r = data[i]
      if (r && Array.isArray(r)) {
        if (isDateValue(r[1]) && !isDateValue(r[0])) {
          offset = 1
          break
        }
      }
    }
  }

  const rows = []
  for (let i = startIndex; i < data.length; i++) {
    const r = data[i]
    if (!r) continue

    // Only process row if Column for ngayXuatNhap is not empty
    const cellDate = r[offset]
    if (cellDate === null || cellDate === undefined || String(cellDate).trim() === '') {
      continue
    }

    rows.push({
      id: i,
      ngayXuatNhap: r[offset + 0] ?? '',
      maVatTu: r[offset + 1] ?? '',
      maSAP: r[offset + 2] ?? '',
      thongSoKyThuat: r[offset + 3] ?? '',
      tenVatTu: r[offset + 4] ?? '',
      dvt: r[offset + 5] ?? '',
      loaiDon: r[offset + 6] ?? '',
      maDonNhapKho: r[offset + 7] ?? '',
      maDonXuatKho: r[offset + 8] ?? '',
      khoiLuongNhap: r[offset + 9] ?? '',
      maDonViGiao: r[offset + 10] ?? '',
      donViGiao: r[offset + 11] ?? '',
      nguoiGiao: r[offset + 12] ?? '',
      khoiLuongXuat: r[offset + 13] ?? '',
      maDonViNhan: r[offset + 14] ?? '',
      donViNhan: r[offset + 15] ?? '',
      nguoiPheDuyet: r[offset + 16] ?? '',
      tenNguon: r[offset + 17] ?? '',
      maNguon: r[offset + 18] ?? '',
      lo: r[offset + 19] ?? '',
      hangMuc: r[offset + 20] ?? '',
      soHopDong: r[offset + 21] ?? '',
      thuKho: r[offset + 22] ?? '',
      bienSoXe: r[offset + 23] ?? '',
      phanKhu: r[offset + 24] ?? '',
      duAn: r[offset + 25] ?? '',
      tinhTrang: r[offset + 26] ?? '',
      nguoiNhan: r[offset + 27] ?? '',
      maDonLienQuan: r[offset + 28] ?? '',
      nhaCungCap: r[offset + 29] ?? '',
      maDonChuyenTiepLC: r[offset + 30] ?? '',
      maDonChuyenTiepNB: r[offset + 31] ?? '',
      ghiChu: r[offset + 32] ?? '',
      ghiChuVatTu: r[offset + 33] ?? '',
      trangThai: r[offset + 34] ?? '',
      nhanHieu: r[offset + 35] ?? '',
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
