import React, { useState, useMemo, useEffect } from 'react'
import * as XLSXStyleRaw from 'xlsx-js-style'
const XLSXStyle = XLSXStyleRaw.default || XLSXStyleRaw

import { 
  Package, Search, Filter, Download, ArrowUpDown, Layers, 
  RefreshCw, X, Eye, FileText, CheckCircle2, ChevronLeft, 
  ChevronRight, BarChart3, Database, Building2, Warehouse,
  SlidersHorizontal, CheckSquare, Square
} from 'lucide-react'

import { isApprovedStatus, normalizeBchName, getStandardizedBchList } from '../constants.js'

// Helper to parse double values safely
const parseVal = (val) => {
  if (val === null || val === undefined) return 0
  if (typeof val === 'number') return val
  const cleaned = String(val).replace(/[^\d.-]/g, '').replace(',', '.')
  const num = parseFloat(cleaned)
  return isNaN(num) ? 0 : num
}

// Parse row date safely
const parseRowDate = (dateVal) => {
  if (!dateVal) return null
  if (dateVal instanceof Date) return dateVal
  const s = String(dateVal).trim()
  if (!s) return null
  
  if (/^\d+(\.\d+)?$/.test(s)) {
    const num = parseFloat(s)
    return new Date(Math.round((num - 25569) * 86400 * 1000))
  }
  
  const parts = s.split('/')
  if (parts.length === 3) {
    const d = parseInt(parts[0], 10)
    const m = parseInt(parts[1], 10) - 1
    const y = parseInt(parts[2], 10)
    if (!isNaN(d) && !isNaN(m) && !isNaN(y)) {
      return new Date(y, m, d)
    }
  }
  
  const d = new Date(s)
  return isNaN(d.getTime()) ? null : d
}

// Searchable Select Component
function SearchableSelect({ value, onChange, options, placeholder = 'Tất cả', searchPlaceholder = 'Tìm kiếm...' }) {
  const [isOpen, setIsOpen] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const containerRef = React.useRef(null)

  useEffect(() => {
    const handleClickOutside = (e) => {
      if (containerRef.current && !containerRef.current.contains(e.target)) {
        setIsOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => document.removeEventListener('mousedown', handleClickOutside)
  }, [])

  const filteredOptions = useMemo(() => {
    if (!searchQuery) return options
    const q = searchQuery.toLowerCase().trim()
    return options.filter(opt => {
      const label = typeof opt === 'object' ? (opt.label || opt.name || opt.value) : String(opt)
      return label.toLowerCase().includes(q)
    })
  }, [options, searchQuery])

  const selectedLabel = useMemo(() => {
    if (!value) return placeholder
    const found = options.find(opt => (typeof opt === 'object' ? opt.value : opt) === value)
    if (found) return typeof found === 'object' ? found.label : found
    return value
  }, [value, options, placeholder])

  return (
    <div ref={containerRef} style={{ position: 'relative', width: '100%' }}>
      <div
        onClick={() => setIsOpen(!isOpen)}
        style={{
          padding: '0 5px 0 12px',
          border: '1.5px solid #D5DCEB',
          borderRadius: '12px',
          fontSize: '15px',
          fontWeight: 600,
          background: '#ffffff',
          color: value ? '#16171C' : '#4A4C55',
          cursor: 'pointer',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          gap: 6,
          height: '36px',
          boxSizing: 'border-box',
          userSelect: 'none'
        }}
      >
        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', minWidth: 0 }}>
          {selectedLabel}
        </span>
        <span style={{ flex: 'none', width: 24, height: 24, borderRadius: 7, background: '#E4EEFA', color: '#0F58A7', display: 'grid', placeItems: 'center' }}>
          <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.6" strokeLinecap="round" strokeLinejoin="round" style={{ transform: isOpen ? 'rotate(180deg)' : 'none', transition: 'transform .15s' }}><path d="M6 9l6 6 6-6" /></svg>
        </span>
      </div>

      {isOpen && (
        <div style={{
          position: 'absolute',
          top: '104%',
          left: 0,
          right: 0,
          background: '#ffffff',
          border: '1px solid #D5DCEB',
          borderRadius: '14px',
          minWidth: '260px',
          boxShadow: '0 16px 32px -12px rgba(10, 36, 80, 0.3)',
          zIndex: 9999,
          maxHeight: '280px',
          display: 'flex',
          flexDirection: 'column',
          overflow: 'hidden'
        }}>
          <div style={{ padding: '8px', borderBottom: '1px solid #f1f5f9', background: '#f8fafc' }}>
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder={searchPlaceholder}
              autoFocus
              style={{
                width: '100%',
                height: '34px',
                padding: '0 10px',
                fontSize: '15px',
                border: '1.5px solid #D5DCEB',
                borderRadius: '10px',
                outline: 'none',
                boxSizing: 'border-box'
              }}
              onClick={(e) => e.stopPropagation()}
            />
          </div>
          <div style={{ overflowY: 'auto', flex: 1 }}>
            <div
              onClick={() => {
                onChange('')
                setIsOpen(false)
                setSearchQuery('')
              }}
              style={{
                padding: '8px 12px',
                fontSize: '15px',
                color: '#0F58A7',
                fontWeight: 600,
                cursor: 'pointer',
                borderBottom: '1px solid #f1f5f9',
                background: !value ? '#E4EEFA' : 'transparent'
              }}
            >
              Tất cả (bỏ lọc)
            </div>
            {filteredOptions.length === 0 ? (
              <div style={{ padding: '12px', fontSize: '15px', color: '#8C8E97', textAlign: 'center' }}>
                Không tìm thấy kết quả
              </div>
            ) : (
              filteredOptions.map((opt, idx) => {
                const optVal = typeof opt === 'object' ? opt.value : opt
                const optLabel = typeof opt === 'object' ? opt.label : opt
                const isSelected = value === optVal

                return (
                  <div
                    key={idx}
                    onClick={() => {
                      onChange(optVal)
                      setIsOpen(false)
                      setSearchQuery('')
                    }}
                    style={{
                      padding: '8px 12px',
                      fontSize: '15px',
                      color: isSelected ? '#0A3D73' : '#16171C',
                      fontWeight: isSelected ? 700 : 400,
                      cursor: 'pointer',
                      background: isSelected ? '#E4EEFA' : 'transparent',
                      borderBottom: '1px solid #f8fafc',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between'
                    }}
                    onMouseOver={(e) => {
                      if (!isSelected) e.currentTarget.style.background = '#f8fafc'
                    }}
                    onMouseOut={(e) => {
                      if (!isSelected) e.currentTarget.style.background = 'transparent'
                    }}
                  >
                    <span>{optLabel}</span>
                    {isSelected && <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="#0F58A7" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round"><path d="M5 12.5l4.5 4.5L19 7.5" /></svg>}
                  </div>
                )
              })
            )}
          </div>
        </div>
      )}
    </div>
  )
}

export default function ThongKeTheoVatTuTab({
  chungRows = [],
  giaoRows = [],
  nhanRows = [],
  allProjects = [],
  customCategoryMap = {},
  dbCategoryMap = {},
  materialPriceRows: propMaterialPriceRows = [],
  materialPrices: propMaterialPrices = {},
  materialClassifications: propMaterialClassifications = {},
  depreciationOptions = [],
  bchAliasMap = {},
  bchAliasRules = []
}) {
  const normChungRows = useMemo(() => {
    if (!bchAliasMap || Object.keys(bchAliasMap).length === 0) return chungRows
    return chungRows.map(r => ({
      ...r,
      donViGiao: normalizeBchName(r.donViGiao, bchAliasMap),
      donViNhan: normalizeBchName(r.donViNhan, bchAliasMap)
    }))
  }, [chungRows, bchAliasMap])

  const normGiaoRows = useMemo(() => {
    if (!bchAliasMap || Object.keys(bchAliasMap).length === 0) return giaoRows
    return giaoRows.map(r => ({
      ...r,
      donViGiao: normalizeBchName(r.donViGiao, bchAliasMap),
      donViNhan: normalizeBchName(r.donViNhan, bchAliasMap)
    }))
  }, [giaoRows, bchAliasMap])

  const normNhanRows = useMemo(() => {
    if (!bchAliasMap || Object.keys(bchAliasMap).length === 0) return nhanRows
    return nhanRows.map(r => ({
      ...r,
      donViGiao: normalizeBchName(r.donViGiao, bchAliasMap),
      donViNhan: normalizeBchName(r.donViNhan, bchAliasMap)
    }))
  }, [nhanRows, bchAliasMap])
  // Filter states
  const [searchTerm, setSearchTerm] = useState('')
  const [selectedYear, setSelectedYear] = useState('ALL') // 'ALL' or specific number like 2026
  const [selectedCategory, setSelectedCategory] = useState('ALL') // 'ALL', 'khauhao', 'ccdc', 'tieuhai', or specific
  const [selectedMaterialSap, setSelectedMaterialSap] = useState('') // Specific material selector
  const [stockFilter, setStockFilter] = useState('active') // 'all', 'active' (> 0), 'has_activity', 'negative' (< 0), 'zero' (= 0)
  const [selectedWarehouseFilter, setSelectedWarehouseFilter] = useState('') // Filter by specific warehouse
  const [hideEmptyWarehouses, setHideEmptyWarehouses] = useState(true)

  // Sort states
  const [sortField, setSortField] = useState('totalStock') // 'maSAP', 'tenVatTu', 'dvt', 'classification', 'totalStock', or warehouse name
  const [sortDirection, setSortDirection] = useState('desc') // 'asc' | 'desc'

  // Pagination states
  const [currentPage, setCurrentPage] = useState(1)
  const [pageSize, setPageSize] = useState(50)

  // Material Transaction Detail Modal
  const [detailModalMaterial, setDetailModalMaterial] = useState(null)

  // Phím Esc đóng hộp thoại lịch sử xuất nhập
  useEffect(() => {
    if (!detailModalMaterial) return
    const onKey = (e) => { if (e.key === 'Escape') { e.preventDefault(); setDetailModalMaterial(null) } }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [detailModalMaterial])

  // Unit categorization helper
  const getCategory = (name) => {
    if (!name) return 'chuaphanbo'
    const norm = String(name).trim().replace(/\s+/g, ' ')
    if (customCategoryMap && customCategoryMap[norm]) {
      return customCategoryMap[norm]
    }
    if (dbCategoryMap && dbCategoryMap[norm]) {
      return dbCategoryMap[norm]
    }
    const upper = norm.toUpperCase()
    if (upper.includes('CÔNG TY') || upper.includes('CONG TY') || upper.includes('CTY') || upper.includes('DNTN')) {
      return 'ncc'
    }
    if (upper.includes('TỔ ĐỘI') || upper.includes('TO DOI')) {
      return 'todoi'
    }
    const tokens = upper.split(/[^A-ZÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸYĐ]/)
    if (tokens.includes('KHO') || tokens.includes('SGC') || tokens.includes('BCH')) {
      return 'kho'
    }
    if (/[-/_[\]()0-9]/.test(norm)) {
      return 'chuaphanbo'
    }
    const words = norm.split(/\s+/)
    if (words.length >= 2 && words.length <= 5) {
      const isCapitalized = words.every(w => w && w[0] === w[0].toUpperCase())
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

  // Load prices and classifications
  const materialPriceRows = useMemo(() => {
    if (propMaterialPriceRows && propMaterialPriceRows.length > 0) return propMaterialPriceRows
    try {
      const saved = localStorage.getItem('sgc_report_material_price_rows')
      return saved ? JSON.parse(saved) : []
    } catch (e) {
      return []
    }
  }, [propMaterialPriceRows])

  const materialPrices = useMemo(() => {
    if (propMaterialPrices && Object.keys(propMaterialPrices).length > 0) return propMaterialPrices
    try {
      const saved = localStorage.getItem('sgc_report_material_prices')
      return saved ? JSON.parse(saved) : {}
    } catch (e) {
      return {}
    }
  }, [propMaterialPrices])

  const materialClassifications = useMemo(() => {
    if (propMaterialClassifications && Object.keys(propMaterialClassifications).length > 0) return propMaterialClassifications
    try {
      const saved = localStorage.getItem('sgc_report_material_classifications')
      return saved ? JSON.parse(saved) : {}
    } catch (e) {
      return {}
    }
  }, [propMaterialClassifications])

  // Extract all unique warehouses (Kho BCH / Dự án chuẩn hóa)
  const uniquePhysicalWarehouses = useMemo(() => {
    const sourceRows = (normChungRows && normChungRows.length > 0) ? normChungRows : [...normGiaoRows, ...normNhanRows]
    return getStandardizedBchList({
      sourceRows,
      bchAliasRules,
      bchAliasMap,
      customCategoryMap,
      dbCategoryMap
    })
  }, [normChungRows, normGiaoRows, normNhanRows, bchAliasRules, bchAliasMap, customCategoryMap, dbCategoryMap])

  // Extract all unique years from transactions
  const uniqueYears = useMemo(() => {
    const years = new Set()
    const sourceRows = (normChungRows && normChungRows.length > 0) ? normChungRows : [...normGiaoRows, ...normNhanRows]
    sourceRows.forEach(r => {
      const d = parseRowDate(r.ngayXuatNhap)
      if (d) years.add(d.getFullYear())
    })
    const list = Array.from(years).sort((a, b) => b - a)
    if (list.length === 0) list.push(new Date().getFullYear())
    return list
  }, [normChungRows, normGiaoRows, normNhanRows])

  // List of distinct material classification options
  const categoryOptions = useMemo(() => {
    const set = new Set()
    set.add('Tài sản khấu hao')
    set.add('Công cụ dụng cụ')
    set.add('Vật tư tiêu hao')

    if (Array.isArray(depreciationOptions)) {
      depreciationOptions.forEach(opt => {
        if (opt) set.add(String(opt).trim())
      })
    }

    Object.values(customCategoryMap || {}).forEach(val => {
      if (val && typeof val === 'string' && val.length > 1 && !['kho', 'ncc', 'todoi', 'chuaphanbo'].includes(val)) {
        set.add(val.trim())
      }
    })
    Object.values(materialClassifications || {}).forEach(val => {
      if (val && typeof val === 'string' && val.trim()) {
        set.add(val.trim())
      }
    })

    return Array.from(set).sort((a, b) => a.localeCompare(b, 'vi'))
  }, [depreciationOptions, customCategoryMap, materialClassifications])

  // Compute stock per material per warehouse
  const { projectStocksByMaterial, materialCatalogMap } = useMemo(() => {
    const map = {} // { [maSAP]: { [projName]: stock } }
    const catalog = {} // { [maSAP]: { maSAP, maVatTu, tenVatTu, dvt, thongSoKyThuat, classification, rawTransactions } }

    const cutoffDate = selectedYear !== 'ALL' 
      ? new Date(Number(selectedYear), 12, 0, 23, 59, 59, 999)
      : null

    const sourceRows = (normChungRows && normChungRows.length > 0) ? normChungRows : [...normGiaoRows, ...normNhanRows]

    sourceRows.forEach(r => {
      if (!isApprovedStatus(r.trangThai)) return

      const sap = String(r.maSAP || r.maVatTu || r.tenVatTu || '').trim()
      if (!sap) return

      const classification = String(customCategoryMap[sap] || materialClassifications[sap] || '').trim()

      if (!catalog[sap]) {
        catalog[sap] = {
          maSAP: String(r.maSAP || sap).trim(),
          maVatTu: String(r.maVatTu || '').trim(),
          tenVatTu: String(r.tenVatTu || '').trim(),
          dvt: String(r.dvt || '').trim(),
          thongSoKyThuat: String(r.thongSoKyThuat || '').trim(),
          classification: classification,
          rawTransactions: []
        }
      } else {
        if (!catalog[sap].maVatTu && r.maVatTu) catalog[sap].maVatTu = String(r.maVatTu).trim()
        if (!catalog[sap].tenVatTu && r.tenVatTu) catalog[sap].tenVatTu = String(r.tenVatTu).trim()
        if (!catalog[sap].dvt && r.dvt) catalog[sap].dvt = String(r.dvt).trim()
        if (!catalog[sap].thongSoKyThuat && r.thongSoKyThuat) catalog[sap].thongSoKyThuat = String(r.thongSoKyThuat).trim()
        if (!catalog[sap].classification && classification) catalog[sap].classification = classification
      }

      const rowDate = parseRowDate(r.ngayXuatNhap)
      const nhanUnit = String(r.donViNhan || '').trim()
      const giaoUnit = String(r.donViGiao || '').trim()

      const catNhan = getCategory(nhanUnit)
      const catGiao = getCategory(giaoUnit)

      // Rule: "Đơn liên quan đến tổ đội để giá trị = 0"
      if (catNhan === 'todoi' || catGiao === 'todoi') return

      const importVal = parseVal(r.khoiLuongNhap)
      const exportVal = parseVal(r.khoiLuongXuat)
      const val = importVal || exportVal || 0

      if (val <= 0) return

      // Record transaction for drilldown
      catalog[sap].rawTransactions.push({
        date: rowDate,
        donViGiao: giaoUnit,
        donViNhan: nhanUnit,
        catGiao,
        catNhan,
        khoiLuongNhap: importVal,
        khoiLuongXuat: exportVal,
        val,
        soPhieu: r.soPhieu || r.maPhieu || '',
        nguoiGiao: r.nguoiGiao || '',
        nguoiNhan: r.nguoiNhan || '',
        ghiChu: r.ghiChu || ''
      })

      // Skip date if beyond cutoff
      if (cutoffDate && rowDate && rowDate > cutoffDate) return

      // Adjust stock of nhanUnit and giaoUnit
      uniquePhysicalWarehouses.forEach(pName => {
        const pLower = pName.toLowerCase()
        const isNhan = nhanUnit.toLowerCase() === pLower
        const isGiao = giaoUnit.toLowerCase() === pLower

        if (isNhan) {
          if (!map[sap]) map[sap] = {}
          if (!map[sap][pName]) map[sap][pName] = 0
          map[sap][pName] += val
        }
        if (isGiao) {
          if (!map[sap]) map[sap] = {}
          if (!map[sap][pName]) map[sap][pName] = 0
          map[sap][pName] -= val
        }
      })
    })

    // Clean up small floating point differences
    Object.keys(map).forEach(sap => {
      Object.keys(map[sap]).forEach(pName => {
        map[sap][pName] = Math.round(map[sap][pName] * 1000) / 1000
        if (Math.abs(map[sap][pName]) < 0.0001) {
          map[sap][pName] = 0
        }
      })
    })

    return { projectStocksByMaterial: map, materialCatalogMap: catalog }
  }, [uniquePhysicalWarehouses, selectedYear, normChungRows, normGiaoRows, normNhanRows, customCategoryMap, materialClassifications])

  // Material Stats Array
  const rawMaterialStatsData = useMemo(() => {
    return Object.values(materialCatalogMap).map(g => {
      const sap = g.maSAP
      let totalStock = 0
      uniquePhysicalWarehouses.forEach(pName => {
        const pStock = (projectStocksByMaterial[sap] && projectStocksByMaterial[sap][pName]) || 0
        totalStock += pStock
      })
      totalStock = Math.round(totalStock * 1000) / 1000

      const priceRow = materialPriceRows.find(r => String(r.maSAP || '').trim().toLowerCase() === sap.toLowerCase())
      const averageUnitPrice = priceRow ? (priceRow.donGiaTrungBinh || 0) : (materialPrices[sap] || 0)
      const closingValue = Math.round(totalStock * averageUnitPrice)

      return {
        ...g,
        totalStock,
        averageUnitPrice,
        closingValue
      }
    })
  }, [materialCatalogMap, uniquePhysicalWarehouses, projectStocksByMaterial, materialPriceRows, materialPrices])

  // Fast options list for specific material selector dropdown
  const materialDropdownOptions = useMemo(() => {
    return rawMaterialStatsData.map(item => ({
      value: item.maSAP,
      label: `${item.maSAP} - ${item.tenVatTu || 'Vật tư'}${item.dvt ? ` (${item.dvt})` : ''}`
    })).sort((a, b) => a.label.localeCompare(b.label, 'vi'))
  }, [rawMaterialStatsData])

  // Filter Data
  const filteredMaterialStatsData = useMemo(() => {
    const text = searchTerm.toLowerCase().trim()

    return rawMaterialStatsData.filter(item => {
      // 1. Filter by specific material SAP if chosen
      if (selectedMaterialSap && item.maSAP.toLowerCase() !== selectedMaterialSap.toLowerCase()) {
        return false
      }

      // 2. Filter by Category
      if (selectedCategory !== 'ALL') {
        const c = (item.classification || '').toLowerCase()
        if (selectedCategory === 'khauhao') {
          const isKhauHao = c.includes('khấu hao') || c.includes('tài sản') || c.includes('tskh')
          if (!isKhauHao) return false
        } else if (selectedCategory === 'ccdc') {
          const isCcdc = c.includes('công cụ') || c.includes('ccdc') || c.includes('dụng cụ')
          if (!isCcdc) return false
        } else if (selectedCategory === 'tieuhai') {
          const isTieuHao = c.includes('tiêu hao') || c.includes('vật tư tiêu hao')
          if (!isTieuHao) return false
        } else if (selectedCategory === 'chua_phan_nhom') {
          if (item.classification && item.classification.trim()) return false
        } else {
          if (item.classification.toLowerCase() !== selectedCategory.toLowerCase()) return false
        }
      }

      // 3. Filter by Stock condition
      if (stockFilter === 'active' && item.totalStock <= 0) {
        return false
      } else if (stockFilter === 'negative' && item.totalStock >= 0) {
        return false
      } else if (stockFilter === 'zero' && item.totalStock !== 0) {
        return false
      } else if (stockFilter === 'has_activity' && item.totalStock === 0 && item.rawTransactions.length === 0) {
        return false
      }

      // 4. Filter by specific warehouse
      if (selectedWarehouseFilter) {
        const wStock = (projectStocksByMaterial[item.maSAP] && projectStocksByMaterial[item.maSAP][selectedWarehouseFilter]) || 0
        if (Math.abs(wStock) < 0.0001) return false
      }

      // 5. Filter by Text Search
      if (text) {
        const matchSap = (item.maSAP || '').toLowerCase().includes(text)
        const matchMaVt = (item.maVatTu || '').toLowerCase().includes(text)
        const matchTen = (item.tenVatTu || '').toLowerCase().includes(text)
        const matchSpec = (item.thongSoKyThuat || '').toLowerCase().includes(text)
        const matchDvt = (item.dvt || '').toLowerCase().includes(text)
        const matchCat = (item.classification || '').toLowerCase().includes(text)
        if (!matchSap && !matchMaVt && !matchTen && !matchSpec && !matchDvt && !matchCat) {
          return false
        }
      }

      return true
    })
  }, [rawMaterialStatsData, selectedMaterialSap, selectedCategory, stockFilter, selectedWarehouseFilter, searchTerm, projectStocksByMaterial])

  // Determine which warehouse columns to display (active warehouses having values or all)
  const displayPhysicalWarehouses = useMemo(() => {
    if (selectedWarehouseFilter) {
      return [selectedWarehouseFilter]
    }
    if (!hideEmptyWarehouses) {
      return uniquePhysicalWarehouses
    }
    return uniquePhysicalWarehouses.filter(pName => {
      return filteredMaterialStatsData.some(item => {
        const val = (projectStocksByMaterial[item.maSAP] && projectStocksByMaterial[item.maSAP][pName]) || 0
        return Math.abs(val) > 0.0001
      })
    })
  }, [uniquePhysicalWarehouses, hideEmptyWarehouses, selectedWarehouseFilter, filteredMaterialStatsData, projectStocksByMaterial])

  // Sorting
  const sortedMaterialStatsData = useMemo(() => {
    const sorted = [...filteredMaterialStatsData]
    sorted.sort((a, b) => {
      let valA
      let valB

      if (sortField === 'totalStock' || sortField === 'closingValue' || sortField === 'averageUnitPrice') {
        valA = a[sortField] || 0
        valB = b[sortField] || 0
      } else if (uniquePhysicalWarehouses.includes(sortField)) {
        valA = (projectStocksByMaterial[a.maSAP] && projectStocksByMaterial[a.maSAP][sortField]) || 0
        valB = (projectStocksByMaterial[b.maSAP] && projectStocksByMaterial[b.maSAP][sortField]) || 0
      } else {
        valA = String(a[sortField] || '').toLowerCase()
        valB = String(b[sortField] || '').toLowerCase()
        return sortDirection === 'asc' ? valA.localeCompare(valB, 'vi') : valB.localeCompare(valA, 'vi')
      }

      return sortDirection === 'asc' ? valA - valB : valB - valA
    })
    return sorted
  }, [filteredMaterialStatsData, sortField, sortDirection, uniquePhysicalWarehouses, projectStocksByMaterial])

  // Summary Metrics
  const summaryMetrics = useMemo(() => {
    let totalItems = sortedMaterialStatsData.length
    let grandStock = 0
    let grandValue = 0
    const warehouseTotals = {}

    displayPhysicalWarehouses.forEach(pName => {
      warehouseTotals[pName] = 0
    })

    sortedMaterialStatsData.forEach(item => {
      grandStock += item.totalStock || 0
      grandValue += item.closingValue || 0

      displayPhysicalWarehouses.forEach(pName => {
        const val = (projectStocksByMaterial[item.maSAP] && projectStocksByMaterial[item.maSAP][pName]) || 0
        warehouseTotals[pName] += val
      })
    })

    grandStock = Math.round(grandStock * 1000) / 1000

    return {
      totalItems,
      grandStock,
      grandValue,
      warehouseTotals
    }
  }, [sortedMaterialStatsData, displayPhysicalWarehouses, projectStocksByMaterial])

  // Pagination calculation
  const totalPages = Math.ceil(sortedMaterialStatsData.length / pageSize) || 1
  const paginatedData = useMemo(() => {
    if (pageSize === -1) return sortedMaterialStatsData
    const startIndex = (currentPage - 1) * pageSize
    return sortedMaterialStatsData.slice(startIndex, startIndex + pageSize)
  }, [sortedMaterialStatsData, currentPage, pageSize])

  // Handle Sort Click
  const handleSort = (field) => {
    if (sortField === field) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc')
    } else {
      setSortField(field)
      setSortDirection('desc')
    }
  }

  // Export to Excel
  const handleExportExcel = () => {
    if (sortedMaterialStatsData.length === 0) {
      alert('Không có dữ liệu để xuất.')
      return
    }

    const wb = XLSXStyle.utils.book_new()
    const ws = {}

    const columns = [
      { key: 'STT', label: 'STT', width: 60 },
      { key: 'maSAP', label: 'Mã SAP', width: 120 },
      { key: 'tenVatTu', label: 'Tên vật tư', width: 320 },
      { key: 'dvt', label: 'ĐVT', width: 80 },
      { key: 'classification', label: 'Nhóm vật tư', width: 160 },
      { key: 'totalStock', label: 'Tổng cộng', width: 120 },
      ...displayPhysicalWarehouses.map(pName => ({
        key: pName,
        label: pName,
        width: 110
      }))
    ]

    const headerStyle = {
      font: { name: 'Arial', bold: true, color: { rgb: 'FFFFFF' }, sz: 10 },
      fill: { fgColor: { rgb: '1E3A8A' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: '000000' } },
        bottom: { style: 'medium', color: { rgb: '000000' } },
        left: { style: 'thin', color: { rgb: 'CBD5E1' } },
        right: { style: 'thin', color: { rgb: 'CBD5E1' } }
      }
    }

    const totalHeaderStyle = {
      ...headerStyle,
      fill: { fgColor: { rgb: '0F766E' } }
    }

    // Title Row
    ws['A1'] = {
      v: `BÁO CÁO THỐNG KÊ TỒN KHO THEO VẬT TƯ ${selectedYear !== 'ALL' ? `(NĂM ${selectedYear})` : '(TẤT CẢ CÁC NĂM)'}`,
      t: 's',
      s: {
        font: { name: 'Arial', bold: true, sz: 14, color: { rgb: '1E3A8A' } },
        alignment: { horizontal: 'center', vertical: 'center' }
      }
    }

    // Subtitle Row
    ws['A2'] = {
      v: `Thời gian xuất báo cáo: ${new Date().toLocaleString('vi-VN')} | Tổng số loại vật tư: ${sortedMaterialStatsData.length.toLocaleString('vi-VN')}`,
      t: 's',
      s: {
        font: { name: 'Arial', italic: true, sz: 9.5, color: { rgb: '64748B' } },
        alignment: { horizontal: 'center', vertical: 'center' }
      }
    }

    // Header Row (Row 4)
    columns.forEach((col, cIdx) => {
      const cellRef = XLSXStyle.utils.encode_cell({ r: 3, c: cIdx })
      ws[cellRef] = {
        v: col.label,
        t: 's',
        s: col.key === 'totalStock' ? totalHeaderStyle : headerStyle
      }
    })

    // Data Rows
    let curR = 4
    sortedMaterialStatsData.forEach((item, idx) => {
      const isEven = idx % 2 === 0
      const rowBg = isEven ? 'FFFFFF' : 'F8FAFC'

      columns.forEach((col, cIdx) => {
        const cellRef = XLSXStyle.utils.encode_cell({ r: curR, c: cIdx })
        let v = ''
        let t = 's'
        let align = 'left'
        let isNumber = false

        if (col.key === 'STT') {
          v = idx + 1
          t = 'n'
          align = 'center'
        } else if (col.key === 'maSAP') {
          v = item.maSAP || ''
          align = 'center'
        } else if (col.key === 'tenVatTu') {
          v = item.tenVatTu || ''
        } else if (col.key === 'dvt') {
          v = item.dvt || ''
          align = 'center'
        } else if (col.key === 'classification') {
          v = item.classification || 'Chưa phân nhóm'
        } else if (col.key === 'totalStock') {
          v = item.totalStock || 0
          t = 'n'
          align = 'right'
          isNumber = true
        } else {
          // Warehouse columns
          const val = (projectStocksByMaterial[item.maSAP] && projectStocksByMaterial[item.maSAP][col.key]) || 0
          v = val
          t = 'n'
          align = 'right'
          isNumber = true
        }

        const cellStyle = {
          font: { name: 'Arial', sz: 9.5, color: { rgb: isNumber && v < 0 ? 'DC2626' : '1E293B' }, bold: col.key === 'totalStock' },
          fill: { fgColor: { rgb: col.key === 'totalStock' ? 'EFF6FF' : rowBg } },
          alignment: { horizontal: align, vertical: 'center' },
          border: {
            top: { style: 'thin', color: { rgb: 'E2E8F0' } },
            bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          }
        }

        if (isNumber) {
          cellStyle.numFmt = '#,##0.00'
        }

        ws[cellRef] = { v, t, s: cellStyle }
      })
      curR++
    })

    // Total Row
    const totalRowStyle = {
      font: { name: 'Arial', bold: true, sz: 10, color: { rgb: '1E3A8A' } },
      fill: { fgColor: { rgb: 'E2E8F0' } },
      alignment: { vertical: 'center' },
      border: {
        top: { style: 'medium', color: { rgb: '1E3A8A' } },
        bottom: { style: 'medium', color: { rgb: '1E3A8A' } },
        left: { style: 'thin', color: { rgb: 'CBD5E1' } },
        right: { style: 'thin', color: { rgb: 'CBD5E1' } }
      }
    }

    columns.forEach((col, cIdx) => {
      const cellRef = XLSXStyle.utils.encode_cell({ r: curR, c: cIdx })
      let v = ''
      let t = 's'
      let align = 'right'

      if (cIdx === 0) {
        v = 'TỔNG CỘNG'
        align = 'center'
      } else if (cIdx === 1 || cIdx === 2 || cIdx === 3 || cIdx === 4) {
        v = ''
      } else if (col.key === 'totalStock') {
        v = summaryMetrics.grandStock
        t = 'n'
      } else {
        v = summaryMetrics.warehouseTotals[col.key] || 0
        t = 'n'
      }

      ws[cellRef] = {
        v,
        t,
        s: {
          ...totalRowStyle,
          alignment: { horizontal: align, vertical: 'center' },
          numFmt: t === 'n' ? '#,##0.00' : undefined
        }
      }
    })

    // Merges
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: columns.length - 1 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: columns.length - 1 } },
      { s: { r: curR, c: 0 }, e: { r: curR, c: 4 } }
    ]

    // Column Widths
    ws['!cols'] = columns.map(c => ({ wpx: c.width }))
    ws['!ref'] = XLSXStyle.utils.encode_range({
      s: { r: 0, c: 0 },
      e: { r: curR, c: columns.length - 1 }
    })

    XLSXStyle.utils.book_append_sheet(wb, ws, 'Thống kê theo vật tư')

    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'array' })
    const blob = new Blob([wbout], { type: 'application/octet-stream' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `SGC_ThongKe_TheoVatTu_${selectedYear !== 'ALL' ? `Nam_${selectedYear}` : 'TatCa'}.xlsx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  return (
    <div className="sgk-page">
      <style>{TKVT_CSS}</style>

      {/* Hàng số liệu + nút thao tác */}
      <div className="sgk-bar">
        <div className="sg-kpis" title="Báo cáo tổng hợp số lượng tồn kho chi tiết theo từng vật tư trên tất cả các Ban Chỉ Huy (kho dự án)">
          <div className="sg-kpi sm sg-c-blue" title="Số loại vật tư đang hiển thị theo bộ lọc">
            <span className="sg-kpi-ic"><Layers size={16} /></span>
            <span className="sg-kpi-bd">
              <span className="sg-kpi-lb">Số loại vật tư</span>
              <span className="sg-kpi-v">{summaryMetrics.totalItems.toLocaleString('vi-VN')}</span>
            </span>
          </div>
          <div className="sg-kpi sm sg-c-green" title="Tổng số lượng tồn của toàn công ty theo bộ lọc">
            <span className="sg-kpi-ic"><Package size={16} /></span>
            <span className="sg-kpi-bd">
              <span className="sg-kpi-lb">Tổng tồn toàn công ty</span>
              <span className="sg-kpi-v">{summaryMetrics.grandStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</span>
            </span>
          </div>
          <div className="sg-kpi sm sg-c-orange" title={`Số kho / BCH đang hiển thị trên tổng ${uniquePhysicalWarehouses.length} kho`}>
            <span className="sg-kpi-ic"><Warehouse size={16} /></span>
            <span className="sg-kpi-bd">
              <span className="sg-kpi-lb">Kho / BCH hiển thị</span>
              <span className="sg-kpi-v">{displayPhysicalWarehouses.length}<small>/ {uniquePhysicalWarehouses.length}</small></span>
            </span>
          </div>
          {summaryMetrics.grandValue > 0 && (
            <div className="sg-kpi sm sg-c-violet" title="Giá trị tồn ước tính theo đơn giá vật tư">
              <span className="sg-kpi-ic"><BarChart3 size={16} /></span>
              <span className="sg-kpi-bd">
                <span className="sg-kpi-lb">Giá trị tồn ước tính</span>
                <span className="sg-kpi-v">{Math.round(summaryMetrics.grandValue).toLocaleString('vi-VN')}<small>đ</small></span>
              </span>
            </div>
          )}
        </div>

        <span className="sgk-spacer" />

        <button
          type="button"
          className={'sgk-check' + (hideEmptyWarehouses ? ' on' : '')}
          onClick={() => setHideEmptyWarehouses(!hideEmptyWarehouses)}
          title="Ẩn các cột kho không có số tồn"
        >
          {hideEmptyWarehouses ? <CheckSquare size={18} /> : <Square size={18} />}
          Ẩn kho không có tồn
        </button>

        <button type="button" onClick={handleExportExcel} className="sg-btn green">
          <Download size={16} />
          Xuất Excel
        </button>
      </div>

      {/* Hàng lọc ngay trên bảng */}
      <div className="sgk-bar sgk-filters">
        {/* Tìm nhanh */}
        <div className="sgk-search" title="Tìm nhanh theo mã SAP, tên vật tư, quy cách, ĐVT">
          <span className="sgk-search-ic"><Search size={16} /></span>
          <input
            type="text"
            value={searchTerm}
            onChange={(e) => {
              setSearchTerm(e.target.value)
              setCurrentPage(1)
            }}
            placeholder="Tìm mã SAP, tên vật tư, quy cách, ĐVT..."
          />
          {searchTerm && (
            <button type="button" className="sgk-x" onClick={() => setSearchTerm('')} title="Xoá tìm kiếm" aria-label="Xoá tìm kiếm">
              <IosX />
            </button>
          )}
        </div>

        {/* Chọn đích danh 1 vật tư */}
        <div className="sgk-f sgk-f-mat" title="Chọn đích danh 1 vật tư">
          <SearchableSelect
            value={selectedMaterialSap}
            onChange={(val) => {
              setSelectedMaterialSap(val)
              setCurrentPage(1)
            }}
            options={materialDropdownOptions}
            placeholder="Tất cả các vật tư"
            searchPlaceholder="Tìm kiếm tên hoặc mã SAP..."
          />
        </div>

        {/* Nhóm vật tư */}
        <select
          className="sgk-sel"
          title="Nhóm vật tư"
          value={selectedCategory}
          onChange={(e) => {
            setSelectedCategory(e.target.value)
            setCurrentPage(1)
          }}
        >
          <option value="ALL">Tất cả nhóm vật tư</option>
          <option value="khauhao">Tài sản khấu hao</option>
          <option value="ccdc">Công cụ dụng cụ</option>
          <option value="tieuhai">Vật tư tiêu hao</option>
          <option value="chua_phan_nhom">Chưa phân nhóm</option>
          {categoryOptions.filter(c => !['Tài sản khấu hao', 'Công cụ dụng cụ', 'Vật tư tiêu hao'].includes(c)).map(cat => (
            <option key={cat} value={cat}>{cat}</option>
          ))}
        </select>

        {/* Năm thống kê */}
        <select
          className="sgk-sel"
          title="Năm thống kê"
          value={selectedYear}
          onChange={(e) => {
            setSelectedYear(e.target.value)
            setCurrentPage(1)
          }}
        >
          <option value="ALL">Tất cả năm</option>
          {uniqueYears.map(yr => (
            <option key={yr} value={yr}>Năm {yr}</option>
          ))}
        </select>

        {/* Trạng thái tồn */}
        <select
          className="sgk-sel"
          title="Trạng thái tồn"
          value={stockFilter}
          onChange={(e) => {
            setStockFilter(e.target.value)
            setCurrentPage(1)
          }}
        >
          <option value="all">Tất cả số dư</option>
          <option value="active">Chỉ còn tồn (&gt; 0)</option>
          <option value="has_activity">Có phát sinh</option>
          <option value="negative">Tồn âm (&lt; 0)</option>
          <option value="zero">Hết tồn (= 0)</option>
        </select>

        {/* Lọc theo kho */}
        <div className="sgk-f sgk-f-wh" title="Lọc theo kho">
          <SearchableSelect
            value={selectedWarehouseFilter}
            onChange={(val) => {
              setSelectedWarehouseFilter(val)
              setCurrentPage(1)
            }}
            options={uniquePhysicalWarehouses}
            placeholder="Tất cả các kho"
            searchPlaceholder="Tìm kiếm kho dự án..."
          />
        </div>
      </div>

      {/* Bảng chính */}
      <div className="sgk-tablebox">
        <div className="sgk-scroll">
          <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, textAlign: 'left', fontSize: '12px', minWidth: '1250px' }}>
            <thead style={{ position: 'sticky', top: 0, zIndex: 20 }}>
              <tr style={{ background: '#1e3a8a' }}>
                <th style={{ padding: '12px 8px', fontWeight: 700, color: '#ffffff', width: '55px', minWidth: '55px', textAlign: 'center', borderRight: '1px solid rgba(255,255,255,0.15)', background: '#1e3a8a' }}>
                  STT
                </th>
                <th 
                  onClick={() => handleSort('maSAP')}
                  style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '120px', minWidth: '120px', textAlign: 'center', borderRight: '1px solid rgba(255,255,255,0.15)', background: '#1e3a8a' }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4 }}>
                    <span>Mã SAP</span>
                    <ArrowUpDown size={12} opacity={sortField === 'maSAP' ? 1 : 0.4} />
                  </div>
                </th>
                <th 
                  onClick={() => handleSort('tenVatTu')}
                  style={{ padding: '12px 12px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '280px', minWidth: '280px', borderRight: '1px solid rgba(255,255,255,0.15)', background: '#1e3a8a' }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                    <span>Tên vật tư</span>
                    <ArrowUpDown size={12} opacity={sortField === 'tenVatTu' ? 1 : 0.4} />
                  </div>
                </th>
                <th style={{ padding: '12px 8px', fontWeight: 700, color: '#ffffff', textAlign: 'center', width: '70px', minWidth: '70px', borderRight: '1px solid rgba(255,255,255,0.15)', background: '#1e3a8a' }}>
                  ĐVT
                </th>
                <th 
                  onClick={() => handleSort('classification')}
                  style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '150px', minWidth: '150px', borderRight: '1px solid rgba(255,255,255,0.15)', background: '#1e3a8a' }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                    <span>Nhóm vật tư</span>
                    <ArrowUpDown size={12} opacity={sortField === 'classification' ? 1 : 0.4} />
                  </div>
                </th>
                
                {/* Total Stock Column Header */}
                <th 
                  onClick={() => handleSort('totalStock')}
                  style={{
                    padding: '12px 10px',
                    fontWeight: 700,
                    color: '#facc15',
                    textAlign: 'center',
                    cursor: 'pointer',
                    userSelect: 'none',
                    width: '120px',
                    minWidth: '120px',
                    background: '#172554',
                    borderRight: '2px solid rgba(255,255,255,0.3)',
                    borderLeft: '1px solid rgba(255,255,255,0.15)'
                  }}
                >
                  <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 4 }}>
                    <span>Tổng cộng</span>
                    <ArrowUpDown size={12} opacity={sortField === 'totalStock' ? 1 : 0.4} />
                  </div>
                </th>
                
                {/* Dynamic Warehouse Columns */}
                {displayPhysicalWarehouses.map(pName => (
                  <th 
                    key={pName}
                    onClick={() => handleSort(pName)}
                    style={{ 
                      padding: '12px 8px', 
                      fontWeight: 700, 
                      color: '#ffffff', 
                      textAlign: 'center', 
                      width: '95px',
                      minWidth: '95px', 
                      background: '#1d4ed8', 
                      borderRight: '1px solid rgba(255, 255, 255, 0.2)', 
                      fontSize: '12px',
                      lineHeight: '1.25',
                      cursor: 'pointer',
                      userSelect: 'none'
                    }}
                    title={`Sắp xếp theo số tồn của kho: ${pName}`}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 2 }}>
                      <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', display: '-webkit-box', WebkitLineClamp: 2, WebkitBoxOrient: 'vertical' }}>
                        {pName}
                      </span>
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {sortedMaterialStatsData.length === 0 ? (
                <tr>
                  <td colSpan={6 + displayPhysicalWarehouses.length} style={{ padding: '48px 20px', textAlign: 'center', color: '#94a3b8', fontSize: '15px' }}>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10 }}>
                      <Package size={40} strokeWidth={1.5} color="#cbd5e1" />
                      <div style={{ fontWeight: 600, color: '#64748b' }}>Không tìm thấy vật tư nào khớp với điều kiện lọc hiện tại.</div>
                      <div style={{ fontSize: '12px', color: '#94a3b8' }}>Hãy thử điều chỉnh lại từ khóa tìm kiếm hoặc bỏ bớt các bộ lọc.</div>
                    </div>
                  </td>
                </tr>
              ) : (
                <>
                  {paginatedData.map((item, idx) => {
                    const rowNumber = (currentPage - 1) * (pageSize > 0 ? pageSize : 0) + idx + 1
                    const isEven = idx % 2 === 0

                    return (
                      <tr 
                        key={item.maSAP}
                        title="Nhấp đúp để xem chi tiết lịch sử xuất nhập của vật tư này"
                        onDoubleClick={() => setDetailModalMaterial(item)}
                        style={{ 
                          borderBottom: '1px solid #e2e8f0',
                          background: isEven ? '#ffffff' : '#f8fafc',
                          transition: 'background 0.15s',
                          cursor: 'pointer'
                        }}
                        onMouseOver={(e) => { e.currentTarget.style.background = '#eff6ff' }}
                        onMouseOut={(e) => { e.currentTarget.style.background = isEven ? '#ffffff' : '#f8fafc' }}
                      >
                        <td style={{ padding: '9px 6px', textAlign: 'center', color: '#64748b', fontWeight: 600, borderRight: '1px solid #e2e8f0' }}>
                          {rowNumber}
                        </td>
                        <td style={{ padding: '9px 8px', fontWeight: 700, color: '#1e293b', textAlign: 'center', borderRight: '1px solid #e2e8f0' }}>
                          {item.maSAP}
                        </td>
                        <td style={{ padding: '9px 12px', fontWeight: 600, color: '#1e3a8a', borderRight: '1px solid #e2e8f0' }}>
                          <div>{item.tenVatTu}</div>
                          {item.thongSoKyThuat && (
                            <div style={{ fontSize: '12px', color: '#64748b', fontWeight: 400 }}>{item.thongSoKyThuat}</div>
                          )}
                        </td>
                        <td style={{ padding: '9px 6px', textAlign: 'center', color: '#475569', borderRight: '1px solid #e2e8f0' }}>
                          {item.dvt}
                        </td>
                        <td style={{ padding: '9px 10px', borderRight: '1px solid #e2e8f0' }}>
                          {item.classification ? (
                            <span style={{
                              fontSize: '12px',
                              fontWeight: 600,
                              padding: '2px 8px',
                              borderRadius: '12px',
                              background: item.classification.toLowerCase().includes('khấu hao') ? '#fef3c7' : item.classification.toLowerCase().includes('công cụ') ? '#e0e7ff' : '#f1f5f9',
                              color: item.classification.toLowerCase().includes('khấu hao') ? '#92400e' : item.classification.toLowerCase().includes('công cụ') ? '#3730a3' : '#475569',
                              display: 'inline-block'
                            }}>
                              {item.classification}
                            </span>
                          ) : (
                            <span style={{ fontSize: '12px', color: '#94a3b8', fontStyle: 'italic' }}>Chưa phân nhóm</span>
                          )}
                        </td>
                        
                        {/* Total Stock Cell */}
                        <td style={{
                          padding: '9px 10px',
                          textAlign: 'right',
                          fontWeight: 700,
                          background: 'rgba(219, 234, 254, 0.35)',
                          borderLeft: '1px solid #cbd5e1',
                          borderRight: '2px solid #cbd5e1',
                          color: item.totalStock > 0 ? '#1e3a8a' : (item.totalStock < 0 ? '#ef4444' : '#94a3b8')
                        }}>
                          {item.totalStock !== 0 ? item.totalStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                        </td>
                        
                        {/* Warehouse Stock Cells */}
                        {displayPhysicalWarehouses.map(pName => {
                          const val = (projectStocksByMaterial[item.maSAP] && projectStocksByMaterial[item.maSAP][pName]) || 0
                          const isPos = val > 0
                          const isNeg = val < 0

                          return (
                            <td 
                              key={pName}
                              style={{ 
                                padding: '9px 6px', 
                                textAlign: 'right', 
                                fontWeight: val !== 0 ? 600 : 400, 
                                background: val !== 0 ? 'rgba(219, 234, 254, 0.12)' : 'transparent', 
                                borderRight: '1px solid #e2e8f0', 
                                color: isPos ? '#1e293b' : (isNeg ? '#ef4444' : '#94a3b8'),
                                fontSize: '15px'
                              }}
                            >
                              {val !== 0 ? val.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                            </td>
                          )
                        })}
                      </tr>
                    )
                  })}

                  {/* Grand Total Footer Row */}
                  <tr style={{ background: '#f1f5f9', fontWeight: 700, borderTop: '2px solid #94a3b8', borderBottom: '2px solid #94a3b8', position: 'sticky', bottom: 0, zIndex: 10 }}>
                    <td colSpan="5" style={{ padding: '12px 14px', color: '#1e3a8a', fontSize: '15px', fontWeight: 700, textAlign: 'right', borderRight: '1px solid #cbd5e1', background: '#e2e8f0' }}>
                      Tổng cộng tồn theo kho
                    </td>
                    
                    {/* Total of Totals */}
                    <td style={{
                      padding: '12px 10px',
                      textAlign: 'right',
                      background: '#dbeafe',
                      borderLeft: '1px solid #cbd5e1',
                      borderRight: '2px solid #94a3b8',
                      color: '#1e3a8a',
                      fontSize: '15px',
                      fontWeight: 700
                    }}>
                      {summaryMetrics.grandStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                    </td>
                    
                    {/* Warehouse Column Totals */}
                    {displayPhysicalWarehouses.map(pName => {
                      const totalVal = summaryMetrics.warehouseTotals[pName] || 0
                      return (
                        <td 
                          key={pName}
                          style={{ 
                            padding: '12px 6px', 
                            textAlign: 'right', 
                            background: '#e2e8f0', 
                            borderRight: '1px solid #cbd5e1', 
                            color: totalVal > 0 ? '#1e3a8a' : (totalVal < 0 ? '#ef4444' : '#64748b'),
                            fontWeight: 700,
                            fontSize: '15px'
                          }}
                        >
                          {totalVal !== 0 ? totalVal.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                        </td>
                      )
                    })}
                  </tr>
                </>
              )}
            </tbody>
          </table>
        </div>

        {/* Chân bảng: số dòng + phân trang */}
        <div className="sgk-foot">
          <div style={{ fontSize: '12px', color: '#4A4C55', fontWeight: 600 }}>
            Hiển thị <strong>{paginatedData.length}</strong> / <strong>{sortedMaterialStatsData.length}</strong> vật tư (tổng hệ thống: {rawMaterialStatsData.length})
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, fontSize: '12px', color: '#4A4C55', fontWeight: 600 }}>
              <span>Dòng/trang:</span>
              <select
                className="sgk-sel sm"
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value))
                  setCurrentPage(1)
                }}
              >
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
                <option value={200}>200</option>
                <option value={-1}>Tất cả</option>
              </select>
            </div>

            {pageSize > 0 && totalPages > 1 && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                <button
                  type="button"
                  className="sgk-pg"
                  disabled={currentPage <= 1}
                  onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                  title="Trang trước"
                >
                  <ChevronLeft size={16} />
                </button>

                <span style={{ fontSize: '12px', color: '#16171C', padding: '0 6px', fontWeight: 700 }}>
                  Trang {currentPage} / {totalPages}
                </span>

                <button
                  type="button"
                  className="sgk-pg"
                  disabled={currentPage >= totalPages}
                  onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                  title="Trang sau"
                >
                  <ChevronRight size={16} />
                </button>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Hộp thoại lịch sử xuất nhập của vật tư */}
      {detailModalMaterial && (
        <div className="sgk-ov">
          <div className="sgk-dlg">
            <div className="sgk-dlg-hd">
              <div style={{ minWidth: 0 }}>
                <h3>Lịch sử xuất nhập: {detailModalMaterial.tenVatTu}</h3>
                <div className="sgk-dlg-meta">
                  Mã SAP: <strong>{detailModalMaterial.maSAP}</strong> · ĐVT: <strong>{detailModalMaterial.dvt || '-'}</strong> · Nhóm: <strong>{detailModalMaterial.classification || 'Chưa phân nhóm'}</strong>
                </div>
              </div>
              <button type="button" className="sgk-close" onClick={() => setDetailModalMaterial(null)} title="Đóng (Esc)">
                <IosX />
                <span>Đóng</span>
              </button>
            </div>

            <div style={{ padding: '0 20px 10px', display: 'flex', flexWrap: 'wrap', gap: 8 }}>
              <div className="sg-kpi sm sg-c-blue">
                <span className="sg-kpi-bd" style={{ paddingLeft: 8 }}>
                  <span className="sg-kpi-lb">Giao dịch phát sinh</span>
                  <span className="sg-kpi-v">{detailModalMaterial.rawTransactions.length}</span>
                </span>
              </div>
              <div className="sg-kpi sm sg-c-green">
                <span className="sg-kpi-bd" style={{ paddingLeft: 8 }}>
                  <span className="sg-kpi-lb">Tổng tồn hiện tại</span>
                  <span className="sg-kpi-v">{detailModalMaterial.totalStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}<small>{detailModalMaterial.dvt}</small></span>
                </span>
              </div>
            </div>

            <div style={{ flex: 1, overflowY: 'auto', padding: '0 20px 20px' }}>
              <table style={{ width: '100%', textAlign: 'left', background: '#fff' }}>
                <thead style={{ position: 'sticky', top: 0, zIndex: 2 }}>
                  <tr>
                    <th style={{ width: '56px' }}>STT</th>
                    <th style={{ width: '110px' }}>Ngày</th>
                    <th style={{ width: '120px' }}>Số phiếu</th>
                    <th>Đơn vị giao</th>
                    <th>Đơn vị nhận</th>
                    <th style={{ width: '100px' }}>Nhập</th>
                    <th style={{ width: '100px' }}>Xuất</th>
                    <th style={{ width: '140px' }}>Ghi chú</th>
                  </tr>
                </thead>
                <tbody>
                  {detailModalMaterial.rawTransactions.length === 0 ? (
                    <tr>
                      <td colSpan="8" style={{ padding: '24px', textAlign: 'center', color: '#8C8E97' }}>
                        Chưa có lịch sử giao dịch phát sinh nào được ghi nhận.
                      </td>
                    </tr>
                  ) : (
                    detailModalMaterial.rawTransactions.map((tx, idx) => {
                      const dateStr = tx.date ? tx.date.toLocaleDateString('vi-VN', { day: '2-digit', month: '2-digit', year: 'numeric' }) : '-'
                      return (
                        <tr key={idx}>
                          <td style={{ textAlign: 'center', color: '#4A4C55' }}>{idx + 1}</td>
                          <td className="sg-num" style={{ textAlign: 'center', color: '#16171C' }}>{dateStr}</td>
                          <td style={{ color: '#0F58A7', fontWeight: 600 }}>{tx.soPhieu || '-'}</td>
                          <td style={{ color: '#4A4C55' }}>{tx.donViGiao || '-'}</td>
                          <td style={{ color: '#4A4C55' }}>{tx.donViNhan || '-'}</td>
                          <td style={{ textAlign: 'right', fontWeight: 700, color: tx.khoiLuongNhap > 0 ? '#0D7A3C' : '#8C8E97' }}>
                            {tx.khoiLuongNhap > 0 ? tx.khoiLuongNhap.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                          </td>
                          <td style={{ textAlign: 'right', fontWeight: 700, color: tx.khoiLuongXuat > 0 ? '#D92D3A' : '#8C8E97' }}>
                            {tx.khoiLuongXuat > 0 ? tx.khoiLuongXuat.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                          </td>
                          <td style={{ color: '#4A4C55', fontSize: '12px' }}>{tx.ghiChu || '-'}</td>
                        </tr>
                      )
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

// Dấu X kiểu iPhone (SVG, nét đều đầu tròn)
function IosX() {
  return (
    <svg width="11" height="11" viewBox="0 0 12 12" fill="none" aria-hidden="true">
      <path d="M2.5 2.5l7 7M9.5 2.5l-7 7" stroke="currentColor" strokeWidth="2.3" strokeLinecap="round" vectorEffect="non-scaling-stroke" />
    </svg>
  )
}

const TKVT_CSS = `
.sgk-page { height: 100%; display: flex; flex-direction: column; gap: 10px; padding: 14px 20px 16px; box-sizing: border-box; overflow: hidden; }
.sgk-bar { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; flex: none; }
.sgk-spacer { flex: 1 1 0; }
.sgk-filters { gap: 8px; }
.sgk-search { position: relative; flex: 1 1 260px; min-width: 220px; max-width: 360px; }
.sgk-search-ic { position: absolute; left: 13px; top: 50%; transform: translateY(-50%); color: #6B7AA0; display: flex; pointer-events: none; }
html body .sgk-search input { width: 100%; height: 36px; box-sizing: border-box; padding: 0 38px; font-size: 15px; color: #16171C; background: #fff; border: 1.5px solid #D5DCEB; border-radius: 99px; outline: none; }
.sgk-x { position: absolute; right: 7px; top: 50%; transform: translateY(-50%); width: 24px; height: 24px; padding: 0; border: 0; border-radius: 50%; display: grid; place-items: center; background: #FFE7E5; color: #FF3B30; cursor: pointer; transition: background .15s, color .15s; }
.sgk-x:hover { background: #FF3B30; color: #fff; }
html body .sgk-x:active { transform: translateY(-50%) scale(.94); }
.sgk-f { min-width: 0; }
.sgk-f-mat { flex: 1 1 240px; max-width: 320px; }
.sgk-f-wh { flex: 1 1 200px; max-width: 260px; }
html body .sgk-sel { height: 36px; box-sizing: border-box; border: 1.5px solid #D5DCEB; padding-left: 12px; font-size: 15px; font-weight: 600; color: #16171C; flex: 0 1 auto; max-width: 230px; }
html body .sgk-sel.sm { height: 32px; font-size: 12px; }
.sgk-check { display: inline-flex; align-items: center; gap: 6px; height: 36px; padding: 0 14px 0 10px; border: 0; border-radius: 99px; background: #fff; color: #4A4C55; font-size: 15px; font-weight: 600; box-shadow: inset 0 0 0 1px #D5DCEB; cursor: pointer; white-space: nowrap; }
.sgk-check.on { color: #0F58A7; background: #E4EEFA; box-shadow: inset 0 0 0 1px #B9D0EE; }
.sgk-tablebox { flex: 1 1 auto; min-height: 0; display: flex; flex-direction: column; gap: 8px; }
.sgk-scroll { flex: 1 1 auto; min-height: 0; overflow: auto; position: relative; border-radius: 14px; }
.sgk-scroll > table { background: #fff; }
.sgk-foot { flex: none; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; padding: 6px 14px; border-radius: 99px; background: rgba(255,255,255,.72); }
.sgk-pg { width: 32px; height: 32px; padding: 0; border: 0; border-radius: 50%; display: grid; place-items: center; background: #fff; color: #0F58A7; box-shadow: inset 0 0 0 1px #D5DCEB; cursor: pointer; }
.sgk-pg:hover:not(:disabled) { background: #E4EEFA; }
.sgk-pg:disabled { color: #B5BAC6; cursor: not-allowed; }
.sgk-ov { position: fixed; inset: 0; z-index: 10000; background: rgba(22, 23, 28, .45); backdrop-filter: blur(4px); display: flex; align-items: center; justify-content: center; padding: 20px; }
.sgk-dlg { background: #fff; border-radius: 24px; width: 100%; max-width: 980px; max-height: 90vh; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 30px 60px -20px rgba(10, 36, 80, .45); }
.sgk-dlg-hd { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; padding: 16px 16px 12px 20px; }
.sgk-dlg-hd h3 { margin: 0; font-size: 22px; font-weight: 700; color: #0A3D73; line-height: 1.3; }
.sgk-dlg-meta { margin-top: 4px; font-size: 12px; color: #4A4C55; font-weight: 600; }
.sgk-close { flex: none; display: inline-flex; align-items: center; gap: 6px; height: 32px; padding: 0 12px 0 6px; border: 0; border-radius: 99px; background: #FFE7E5; color: #FF3B30; font-size: 15px; font-weight: 600; cursor: pointer; }
.sgk-close svg { width: 20px; height: 20px; padding: 4.5px; box-sizing: border-box; border-radius: 50%; background: rgba(255,255,255,.7); }
.sgk-close:hover { background: #FF3B30; color: #fff; }
.sgk-close:hover svg { background: rgba(255,255,255,.2); }
`
