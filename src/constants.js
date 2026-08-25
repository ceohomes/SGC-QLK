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

// Helper to normalize BCH name according to alias rules (old -> canonical standard)
export function normalizeBchName(name, bchAliasMap = {}) {
  if (name === null || name === undefined) return ''
  const clean = String(name).trim().normalize('NFC').replace(/\s+/g, ' ')
  if (!clean) return ''
  const lookupKey = clean.toLowerCase()
  if (bchAliasMap && bchAliasMap[lookupKey]) {
    return bchAliasMap[lookupKey]
  }
  return clean
}

export function parseXlsxToRows(data, bchAliasMap = {}) {
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

    let khoiLuongNhap = r[offset + 9] ?? ''
    let khoiLuongXuat = r[offset + 13] ?? ''

    // Helper to check if a quantity cell is empty/blank
    const isBlankQty = (v) => {
      if (v === null || v === undefined) return true
      const s = String(v).trim()
      return s === '' || s === '-' || s === '—' || s === '–'
    }

    // Logic khi up đơn: Trường hợp khối lượng xuất mà trống thì auto lấy bằng khối lượng nhập, và ngược lại
    if (isBlankQty(khoiLuongXuat) && !isBlankQty(khoiLuongNhap)) {
      khoiLuongXuat = khoiLuongNhap
    } else if (isBlankQty(khoiLuongNhap) && !isBlankQty(khoiLuongXuat)) {
      khoiLuongNhap = khoiLuongXuat
    }

    // Chuẩn hóa tên đơn vị giao & nhận theo bảng quy tắc tên cũ -> tên chuẩn
    const rawDonViGiao = r[offset + 11] ?? ''
    const rawDonViNhan = r[offset + 15] ?? ''
    const donViGiao = normalizeBchName(rawDonViGiao, bchAliasMap)
    const donViNhan = normalizeBchName(rawDonViNhan, bchAliasMap)

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
      khoiLuongNhap,
      maDonViGiao: r[offset + 10] ?? '',
      donViGiao,
      nguoiGiao: r[offset + 12] ?? '',
      khoiLuongXuat,
      maDonViNhan: r[offset + 14] ?? '',
      donViNhan,
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
      if (val === 0) return '-'
      // Format số với dấu phân cách hàng nghìn
      if (Number.isInteger(val)) return val.toLocaleString('vi-VN')
      return val.toLocaleString('vi-VN', { maximumFractionDigits: 2 })
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
      if (num === 0) return '-'
      if (Number.isInteger(num)) return num.toLocaleString('vi-VN')
      return num.toLocaleString('vi-VN', { maximumFractionDigits: 2 })
    }
  }
  return String(val)
}

export function isApprovedStatus(statusStr) {
  if (!statusStr) return true
  const s = String(statusStr).toLowerCase().trim()
  if (!s) return true
  // Exclude explicit pending or negative status
  if (s.includes('chờ') || s.includes('chưa') || s.includes('pending')) return false
  if (s.includes('từ chối') || s.includes('hủy') || s.includes('reject') || s.includes('không')) return false
  // Consider approved/completed by default
  return true
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

export function getUnitCategory(name) {
  if (!name) return 'chuaphanbo'
  const upper = String(name).toUpperCase().trim()
  
  // 1. Nhà cung cấp
  const isNcc = 
    upper.includes('CÔNG TY') || 
    upper.includes('CONG TY') || 
    upper.includes('CTY') || 
    upper.includes('C.TY') ||
    upper.includes('DNTN') ||
    upper.includes('TNHH') ||
    upper.includes('CỔ PHẦN') ||
    upper.includes('CO PHAN') ||
    upper.includes('TẬP ĐOÀN') ||
    upper.includes('TAP DOAN') ||
    upper.includes('COOP') ||
    upper.includes('CO.OP') ||
    upper.includes('HTX') ||
    upper.includes('HỢP TÁC XÃ') ||
    upper.includes('XÍ NGHIỆP') ||
    upper.includes('XI NGHIEP') ||
    upper.includes('NHÀ MÁY') ||
    upper.includes('NHA MAY') ||
    upper.includes('TMDV') ||
    upper.includes('XNK') ||
    upper.startsWith('CT ') ||
    upper.includes(' CT ') ||
    upper.startsWith('CP ') ||
    upper.includes(' CP ') ||
    upper.includes('TRẠM BÊ TÔNG') ||
    upper.includes('TRAM BE TONG') ||
    upper.includes('CỬA HÀNG') ||
    upper.includes('CUA HANG') ||
    upper.includes('ĐẠI LÝ') ||
    upper.includes('DAI LY')

  if (isNcc) {
    return 'ncc'
  }
  
  // 2. Tổ đội
  const hasToDoi = 
    upper.includes('TỔ ĐỘI') || 
    upper.includes('TO DOI') || 
    upper.startsWith('ĐỘI ') || 
    upper.startsWith('DOI ') ||
    upper.includes(' ĐỘI ') || 
    upper.includes(' DOI ') ||
    upper.startsWith('TỔ ') ||
    upper.includes(' TỔ ')
    
  if (hasToDoi) {
    return 'todoi'
  }
  
  // 3. Kho BCH
  const tokens = upper.split(/[^A-ZÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸYĐ]/)
  const hasKho = tokens.includes('KHO')
  const hasSgc = tokens.includes('SGC')
  const hasBch = tokens.includes('BCH')
  
  if (hasKho || hasSgc || hasBch) {
    return 'kho'
  }
  
  // 4. Check if it's a personal full name (which falls into "Tổ đội")
  const trimmed = String(name).trim()
  if (/[-/_[\]()0-9]/.test(trimmed)) {
    return 'chuaphanbo'
  }
  
  const words = trimmed.split(/\s+/)
  if (words.length >= 2 && words.length <= 5) {
    const isCapitalized = words.every(w => {
      if (!w) return true
      const firstChar = w[0]
      return firstChar === firstChar.toUpperCase()
    })
    
    const excludeKeywords = [
      'SGC', 'BCH', 'VFVA', 'HLX', 'TPC', 'TCE', 'HP', 'VP', 'KT', 'GT', 'QD', 'CP', 'TNHH', 'MTV', 
      'TRẠM', 'TRAM', 'BÊ TÔNG', 'BE TONG', 'NHÀ MÁY', 'NHA MAY', 'XÍ NGHIỆP', 'XI NGHIEP', 'DỰ ÁN', 'DU AN',
      'BAN CHỈ HUY', 'BAN CHI HUY', 'HẠ TẦNG', 'HA TANG', 'ĐƯỜNG', 'DUONG', 'CẦU', 'CAU', 'SÂN', 'SAN'
    ]
    const hasExclude = words.some(w => excludeKeywords.includes(w.toUpperCase()))
    
    if (isCapitalized && !hasExclude) {
      return 'todoi'
    }
  }
  
  return 'chuaphanbo'
}

/**
 * Trích xuất danh sách tất cả các Ban Chỉ Huy (Kho BCH) đã chuẩn hóa
 * Đồng bộ chính xác 100% với logic hiển thị của Tab 'Chuẩn hóa tên BCH'
 */
export function getStandardizedBchList({
  sourceRows = [],
  bchAliasRules = [],
  bchAliasMap = {},
  customCategoryMap = {},
  dbCategoryMap = {}
}) {
  const norm = (s) => (s || '').trim().normalize('NFC').replace(/\s+/g, ' ')
  const normKey = (s) => norm(s).toLowerCase()

  const normCustom = {}
  Object.entries(customCategoryMap || {}).forEach(([k, v]) => {
    if (k) normCustom[normKey(k)] = v
  })

  const normDb = {}
  Object.entries(dbCategoryMap || {}).forEach(([k, v]) => {
    if (k) normDb[normKey(k)] = v
  })

  const getUnitCat = (name) => {
    if (!name) return 'chuaphanbo'
    const k = normKey(name)
    if (normCustom[k]) return normCustom[k]
    if (normDb[k]) return normDb[k]

    const upper = name.toUpperCase().trim()
    const isNcc = 
      upper.includes('CÔNG TY') || upper.includes('CONG TY') || upper.includes('CTY') ||
      upper.includes('C.TY') || upper.includes('DNTN') || upper.includes('TNHH') ||
      upper.includes('CỔ PHẦN') || upper.includes('CO PHAN') || upper.includes('TẬP ĐOÀN') ||
      upper.includes('TAP DOAN') || upper.includes('COOP') || upper.includes('CO.OP') ||
      upper.includes('HTX') || upper.includes('HỢP TÁC XÃ') || upper.includes('XÍ NGHIỆP') ||
      upper.includes('XI NGHIEP') || upper.includes('NHÀ MÁY') || upper.includes('NHA MAY') ||
      upper.includes('TMDV') || upper.includes('XNK') || upper.startsWith('CT ') ||
      upper.includes(' CT ') || upper.startsWith('CP ') || upper.includes(' CP ') ||
      upper.includes('TRẠM BÊ TÔNG') || upper.includes('TRAM BE TONG') ||
      upper.includes('CỬA HÀNG') || upper.includes('CUA HANG') || upper.includes('ĐẠI LÝ') || upper.includes('DAI LY')
    if (isNcc) return 'ncc'

    const hasToDoi = 
      upper.includes('TỔ ĐỘI') || upper.includes('TO DOI') || upper.startsWith('ĐỘI ') ||
      upper.startsWith('DOI ') || upper.includes(' ĐỘI ') || upper.includes(' DOI ') ||
      upper.startsWith('TỔ ') || upper.includes(' TỔ ')
    if (hasToDoi) return 'todoi'

    const s = name.toLowerCase().trim()
    if (
      s.startsWith('bch') ||
      s.includes('ban chỉ huy') ||
      s.includes('kho bch') ||
      s.includes('dự án') ||
      s.includes('bch ckn') ||
      s.includes('bch thi công') ||
      s.includes('cọc khoan nhồi') ||
      s.includes('đường sắt') ||
      s.includes('cao tốc') ||
      s.includes('hạ long') ||
      s.includes('thanh hóa')
    ) {
      return 'kho'
    }

    const tokens = upper.split(/[^A-ZÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸYĐ]/)
    if (tokens.includes('BCH') || tokens.includes('SGC')) return 'kho'

    return getUnitCategory(name)
  }

  // 1. Quét toàn bộ đơn vị trong sourceRows
  const unitCounts = new Map()
  sourceRows.forEach(r => {
    const g = norm(r.donViGiao)
    const n = norm(r.donViNhan)
    if (g && !unitCounts.has(normKey(g))) unitCounts.set(normKey(g), g)
    if (n && !unitCounts.has(normKey(n))) unitCounts.set(normKey(n), n)
  })

  // 2. Gom tất cả tên chuẩn từ bchAliasRules
  const stdGroups = new Map() // stdKey -> stdClean
  const mappedOldKeys = new Set()

  ;(bchAliasRules || []).forEach(r => {
    if (r.ten_chuan) {
      const stdClean = norm(r.ten_chuan)
      const stdKey = normKey(stdClean)
      if (!stdGroups.has(stdKey)) {
        stdGroups.set(stdKey, stdClean)
      }
    }
    if (r.ten_cu) {
      mappedOldKeys.add(normKey(r.ten_cu))
    }
  })

  // 3. Quét các đơn vị thực tế thuộc kho (và chưa bị gộp thành tên cũ của rule)
  unitCounts.forEach((unitName, kKey) => {
    if (mappedOldKeys.has(kKey)) return // đã được map vào rule
    const cat = getUnitCat(unitName)
    if (cat === 'kho' || (bchAliasRules || []).some(r => normKey(r.ten_chuan) === kKey)) {
      const stdName = normalizeBchName(unitName, bchAliasMap) || unitName
      const stdClean = norm(stdName)
      const stdKey = normKey(stdClean)
      if (!stdGroups.has(stdKey)) {
        stdGroups.set(stdKey, stdClean)
      }
    }
  })

  return Array.from(stdGroups.values()).sort((a, b) => a.localeCompare(b, 'vi', { sensitivity: 'base' }))
}

// Columns for Tình trạng đơn vật tư
export const COLS_TINH_TRANG_DON = [
  { key: 'stt', label: 'STT', width: 50 },
  { key: 'soDon', label: 'Số đơn', width: 140 },
  { key: 'ngayXuat', label: 'Ngày xuất', width: 100 },
  { key: 'donViGiao', label: 'Đơn vị giao', width: 180 },
  { key: 'nguoiGiao', label: 'Người giao', width: 130 },
  { key: 'donViNhan', label: 'Đơn vị nhận', width: 180 },
  { key: 'nguoiNhan', label: 'Người nhận', width: 130 },
  { key: 'cbPheDuyet1', label: 'Cán bộ phê duyệt 1', width: 140 },
  { key: 'cbPheDuyet2', label: 'Cán bộ phê duyệt 2', width: 140 },
  { key: 'cbPheDuyet3', label: 'Cán bộ phê duyệt 3', width: 140 },
  { key: 'cbPheDuyet4', label: 'Cán bộ phê duyệt 4', width: 140 },
  { key: 'trangThai', label: 'Trạng thái', width: 130 },
  { key: 'soDonSAP', label: 'Số đơn/phiếu SAP', width: 140 },
  { key: 'dongBoSAP', label: 'Đồng bộ SAP', width: 120 },
]

export function parseTinhTrangDonRows(data, bchAliasMap = {}) {
  if (!data || !Array.isArray(data) || data.length === 0) return []

  const formatCellDate = (val) => {
    if (val === null || val === undefined || val === '') return ''
    if (typeof val === 'number' && val > 20000 && val < 60000) {
      const d = new Date(Math.round((val - 25569) * 86400 * 1000))
      const dd = String(d.getUTCDate()).padStart(2, '0')
      const mm = String(d.getUTCMonth() + 1).padStart(2, '0')
      const yyyy = d.getUTCFullYear()
      return `${dd}/${mm}/${yyyy}`
    }
    if (val instanceof Date) {
      const dd = String(val.getDate()).padStart(2, '0')
      const mm = String(val.getMonth() + 1).padStart(2, '0')
      const yyyy = val.getFullYear()
      return `${dd}/${mm}/${yyyy}`
    }
    const str = String(val).trim()
    const isoMatch = str.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/)
    if (isoMatch) {
      return `${isoMatch[3].padStart(2, '0')}/${isoMatch[2].padStart(2, '0')}/${isoMatch[1]}`
    }
    return str
  }

  const cleanStr = (v) => (v === null || v === undefined ? '' : String(v).trim())

  let headerRowIdx = -1
  let colMap = {
    stt: -1,
    soDon: -1,
    ngayXuat: -1,
    donViGiao: -1,
    nguoiGiao: -1,
    donViNhan: -1,
    nguoiNhan: -1,
    cbPheDuyet1: -1,
    cbPheDuyet2: -1,
    cbPheDuyet3: -1,
    cbPheDuyet4: -1,
    trangThai: -1,
    soDonSAP: -1,
    dongBoSAP: -1,
  }

  for (let r = 0; r < Math.min(data.length, 15); r++) {
    const row = data[r]
    if (!Array.isArray(row)) continue

    const matched = {}
    row.forEach((cell, cIdx) => {
      if (!cell) return
      const s = String(cell).toLowerCase().trim().normalize('NFC')
      if (s === 'stt' || s.startsWith('stt')) matched.stt = cIdx
      if (s.includes('số đơn') || s.includes('so don') || s.includes('mã đơn') || s.includes('ma don') || s.includes('số phiếu') || s.includes('so phieu') || s === 'mã đơn hàng' || s === 'order no') {
        if (!s.includes('sap')) matched.soDon = cIdx
      }
      if (s.includes('ngày xuất') || s.includes('ngay xuat') || s.includes('ngày tạo') || (s.includes('ngày') && !s.includes('sinh') && !s.includes('nhập'))) matched.ngayXuat = cIdx
      if (s.includes('đơn vị giao') || s.includes('don vi giao') || s.includes('kho giao') || s.includes('nơi giao')) matched.donViGiao = cIdx
      if (s.includes('người giao') || s.includes('nguoi giao') || s.includes('thủ kho giao')) matched.nguoiGiao = cIdx
      if (s.includes('đơn vị nhận') || s.includes('don vi nhan') || s.includes('kho nhận') || s.includes('nơi nhận')) matched.donViNhan = cIdx
      if (s.includes('người nhận') || s.includes('nguoi nhan') || s.includes('thủ kho nhận')) matched.nguoiNhan = cIdx
      if (s.includes('1') && (s.includes('phê duyệt') || s.includes('cán bộ') || s.includes('cb') || s.includes('duyệt'))) matched.cbPheDuyet1 = cIdx
      if (s.includes('2') && (s.includes('phê duyệt') || s.includes('cán bộ') || s.includes('cb') || s.includes('duyệt'))) matched.cbPheDuyet2 = cIdx
      if (s.includes('3') && (s.includes('phê duyệt') || s.includes('cán bộ') || s.includes('cb') || s.includes('duyệt'))) matched.cbPheDuyet3 = cIdx
      if (s.includes('4') && (s.includes('phê duyệt') || s.includes('cán bộ') || s.includes('cb') || s.includes('duyệt'))) matched.cbPheDuyet4 = cIdx
      if (s.includes('trạng thái') || s.includes('trang thai') || s.includes('tình trạng')) {
        if (!s.includes('sap')) matched.trangThai = cIdx
      }
      if (s.includes('sap') && (s.includes('số') || s.includes('phiếu') || s.includes('đơn') || s.includes('mã'))) matched.soDonSAP = cIdx
      if (s.includes('đồng bộ') || s.includes('dong bo') || (s.includes('sap') && (s.includes('đồng bộ') || s.includes('trạng thái') || s.includes('sync')))) matched.dongBoSAP = cIdx
    })

    if (Object.keys(matched).length >= 3) {
      headerRowIdx = r
      colMap = { ...colMap, ...matched }
      break
    }
  }

  const startRow = headerRowIdx >= 0 ? headerRowIdx + 1 : 0
  const usePositional = headerRowIdx === -1 || colMap.soDon === -1

  const rows = []
  for (let i = startRow; i < data.length; i++) {
    const r = data[i]
    if (!r || !Array.isArray(r)) continue

    let stt = usePositional ? cleanStr(r[0]) : (colMap.stt >= 0 ? cleanStr(r[colMap.stt]) : '')
    let soDon = usePositional ? cleanStr(r[1]) : (colMap.soDon >= 0 ? cleanStr(r[colMap.soDon]) : '')
    let rawNgayXuat = usePositional ? r[2] : (colMap.ngayXuat >= 0 ? r[colMap.ngayXuat] : '')
    let donViGiao = usePositional ? cleanStr(r[3]) : (colMap.donViGiao >= 0 ? cleanStr(r[colMap.donViGiao]) : '')
    let nguoiGiao = usePositional ? cleanStr(r[4]) : (colMap.nguoiGiao >= 0 ? cleanStr(r[colMap.nguoiGiao]) : '')
    let donViNhan = usePositional ? cleanStr(r[5]) : (colMap.donViNhan >= 0 ? cleanStr(r[colMap.donViNhan]) : '')
    let nguoiNhan = usePositional ? cleanStr(r[6]) : (colMap.nguoiNhan >= 0 ? cleanStr(r[colMap.nguoiNhan]) : '')
    let cbPheDuyet1 = usePositional ? cleanStr(r[7]) : (colMap.cbPheDuyet1 >= 0 ? cleanStr(r[colMap.cbPheDuyet1]) : '')
    let cbPheDuyet2 = usePositional ? cleanStr(r[8]) : (colMap.cbPheDuyet2 >= 0 ? cleanStr(r[colMap.cbPheDuyet2]) : '')
    let cbPheDuyet3 = usePositional ? cleanStr(r[9]) : (colMap.cbPheDuyet3 >= 0 ? cleanStr(r[colMap.cbPheDuyet3]) : '')
    let cbPheDuyet4 = usePositional ? cleanStr(r[10]) : (colMap.cbPheDuyet4 >= 0 ? cleanStr(r[colMap.cbPheDuyet4]) : '')
    let trangThai = usePositional ? cleanStr(r[11]) : (colMap.trangThai >= 0 ? cleanStr(r[colMap.trangThai]) : '')
    let soDonSAP = usePositional ? cleanStr(r[12]) : (colMap.soDonSAP >= 0 ? cleanStr(r[colMap.soDonSAP]) : '')
    let dongBoSAP = usePositional ? cleanStr(r[13]) : (colMap.dongBoSAP >= 0 ? cleanStr(r[colMap.dongBoSAP]) : '')

    if (!soDon && !rawNgayXuat && !donViGiao && !donViNhan && !trangThai && !soDonSAP) {
      continue
    }

    const sSoDonLower = String(soDon).toLowerCase().trim()
    const sSttLower = String(stt).toLowerCase().trim()
    if (sSoDonLower.includes('tổng cộng') || sSoDonLower.includes('tong cong') || sSttLower.includes('tổng cộng') || sSoDonLower === 'số đơn' || sSoDonLower === 'mã đơn') {
      continue
    }

    if (!soDon && !donViGiao && !donViNhan && !nguoiNhan && !cbPheDuyet1 && !cbPheDuyet2) {
      continue
    }

    const ngayXuat = formatCellDate(rawNgayXuat)
    // Giữ nguyên 100% dữ liệu gốc của Đơn vị giao và Đơn vị nhận theo file tải lên
    donViGiao = cleanStr(donViGiao)
    donViNhan = cleanStr(donViNhan)

    rows.push({
      id: rows.length + 1,
      stt: stt || String(rows.length + 1),
      soDon,
      ngayXuat,
      donViGiao,
      nguoiGiao,
      donViNhan,
      nguoiNhan,
      cbPheDuyet1,
      cbPheDuyet2,
      cbPheDuyet3,
      cbPheDuyet4,
      trangThai: trangThai || '—',
      soDonSAP: soDonSAP || '—',
      dongBoSAP: dongBoSAP || '—'
    })
  }

  return rows
}


