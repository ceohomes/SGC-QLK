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
          padding: '8px 12px',
          border: '1px solid #cbd5e1',
          borderRadius: '8px',
          fontSize: '13px',
          fontWeight: 500,
          background: '#ffffff',
          color: value ? '#1e293b' : '#64748b',
          cursor: 'pointer',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          height: '38px',
          boxSizing: 'border-box',
          userSelect: 'none',
          boxShadow: '0 1px 2px rgba(0,0,0,0.05)'
        }}
      >
        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '90%' }}>
          {selectedLabel}
        </span>
        <span style={{ fontSize: '10px', color: '#94a3b8', marginLeft: 4 }}>▼</span>
      </div>

      {isOpen && (
        <div style={{
          position: 'absolute',
          top: '104%',
          left: 0,
          right: 0,
          background: '#ffffff',
          border: '1px solid #cbd5e1',
          borderRadius: '8px',
          boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.15), 0 8px 10px -6px rgba(0, 0, 0, 0.1)',
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
                padding: '6px 10px',
                fontSize: '12.5px',
                border: '1px solid #cbd5e1',
                borderRadius: '6px',
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
                fontSize: '12.5px',
                color: '#ef4444',
                fontWeight: 600,
                cursor: 'pointer',
                borderBottom: '1px solid #f1f5f9',
                background: !value ? '#fef2f2' : 'transparent'
              }}
            >
              -- Tất cả / Mặc định --
            </div>
            {filteredOptions.length === 0 ? (
              <div style={{ padding: '12px', fontSize: '12px', color: '#94a3b8', textAlign: 'center' }}>
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
                      fontSize: '12.5px',
                      color: isSelected ? '#1e40af' : '#334155',
                      fontWeight: isSelected ? 700 : 400,
                      cursor: 'pointer',
                      background: isSelected ? '#eff6ff' : 'transparent',
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
                    {isSelected && <span style={{ color: '#2563eb', fontSize: '12px' }}>✓</span>}
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
    <div style={{ padding: '20px', height: '100%', display: 'flex', flexDirection: 'column', boxSizing: 'border-box', overflow: 'hidden', background: '#f8fafc' }}>
      
      {/* Top Header & Metrics Bar */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', marginBottom: '16px', gap: 16, flexWrap: 'wrap' }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{
              background: 'linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)',
              color: '#ffffff',
              padding: '8px',
              borderRadius: '10px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              boxShadow: '0 4px 6px -1px rgba(37, 99, 235, 0.2)'
            }}>
              <Package size={22} />
            </div>
            <div>
              <h2 style={{ fontSize: '19px', fontWeight: 800, color: '#0f172a', margin: 0, letterSpacing: '-0.02em' }}>
                THỐNG KÊ TỒN KHO THEO VẬT TƯ
              </h2>
              <p style={{ fontSize: '13px', color: '#64748b', margin: '2px 0 0 0' }}>
                Báo cáo tổng hợp số lượng tồn kho chi tiết theo từng vật tư trên tất cả các Ban Chỉ Huy (Kho Dự Án)
              </p>
            </div>
          </div>
        </div>

        {/* Action Buttons */}
        <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
          <button
            onClick={handleExportExcel}
            style={{
              background: 'linear-gradient(135deg, #059669 0%, #10b981 100%)',
              color: '#ffffff',
              border: 'none',
              padding: '9px 18px',
              borderRadius: '8px',
              fontSize: '13.5px',
              fontWeight: 700,
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              boxShadow: '0 4px 10px rgba(16, 185, 129, 0.25)',
              transition: 'all 0.2s'
            }}
            onMouseOver={(e) => e.currentTarget.style.transform = 'translateY(-1px)'}
            onMouseOut={(e) => e.currentTarget.style.transform = 'none'}
          >
            <Download size={16} />
            Xuất file Excel
          </button>
        </div>
      </div>

      {/* KPI Cards Summary Bar */}
      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))',
        gap: '12px',
        marginBottom: '16px'
      }}>
        <div style={{
          background: '#ffffff',
          border: '1px solid #e2e8f0',
          borderRadius: '10px',
          padding: '12px 16px',
          boxShadow: '0 1px 3px rgba(0,0,0,0.04)',
          display: 'flex',
          alignItems: 'center',
          gap: 12
        }}>
          <div style={{ background: '#eff6ff', color: '#2563eb', padding: '10px', borderRadius: '8px' }}>
            <Layers size={20} />
          </div>
          <div>
            <div style={{ fontSize: '11.5px', color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Số loại vật tư</div>
            <div style={{ fontSize: '18px', fontWeight: 800, color: '#1e3a8a' }}>
              {summaryMetrics.totalItems.toLocaleString('vi-VN')}
            </div>
          </div>
        </div>

        <div style={{
          background: '#ffffff',
          border: '1px solid #e2e8f0',
          borderRadius: '10px',
          padding: '12px 16px',
          boxShadow: '0 1px 3px rgba(0,0,0,0.04)',
          display: 'flex',
          alignItems: 'center',
          gap: 12
        }}>
          <div style={{ background: '#ecfdf5', color: '#059669', padding: '10px', borderRadius: '8px' }}>
            <Package size={20} />
          </div>
          <div>
            <div style={{ fontSize: '11.5px', color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Tổng tồn toàn công ty</div>
            <div style={{ fontSize: '18px', fontWeight: 800, color: '#059669' }}>
              {summaryMetrics.grandStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
            </div>
          </div>
        </div>

        <div style={{
          background: '#ffffff',
          border: '1px solid #e2e8f0',
          borderRadius: '10px',
          padding: '12px 16px',
          boxShadow: '0 1px 3px rgba(0,0,0,0.04)',
          display: 'flex',
          alignItems: 'center',
          gap: 12
        }}>
          <div style={{ background: '#fef3c7', color: '#d97706', padding: '10px', borderRadius: '8px' }}>
            <Warehouse size={20} />
          </div>
          <div>
            <div style={{ fontSize: '11.5px', color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Số Kho / BCH hiển thị</div>
            <div style={{ fontSize: '18px', fontWeight: 800, color: '#d97706' }}>
              {displayPhysicalWarehouses.length} <span style={{ fontSize: '12px', fontWeight: 500, color: '#94a3b8' }}>/ {uniquePhysicalWarehouses.length}</span>
            </div>
          </div>
        </div>

        {summaryMetrics.grandValue > 0 && (
          <div style={{
            background: '#ffffff',
            border: '1px solid #e2e8f0',
            borderRadius: '10px',
            padding: '12px 16px',
            boxShadow: '0 1px 3px rgba(0,0,0,0.04)',
            display: 'flex',
            alignItems: 'center',
            gap: 12
          }}>
            <div style={{ background: '#f5f3ff', color: '#7c3aed', padding: '10px', borderRadius: '8px' }}>
              <BarChart3 size={20} />
            </div>
            <div>
              <div style={{ fontSize: '11.5px', color: '#64748b', fontWeight: 600, textTransform: 'uppercase' }}>Giá trị tồn ước tính</div>
              <div style={{ fontSize: '18px', fontWeight: 800, color: '#7c3aed' }}>
                {Math.round(summaryMetrics.grandValue).toLocaleString('vi-VN')} đ
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Modern Filter Toolbar */}
      <div style={{
        background: '#ffffff',
        border: '1px solid #e2e8f0',
        borderRadius: '12px',
        padding: '14px 16px',
        marginBottom: '16px',
        display: 'flex',
        gap: '12px',
        alignItems: 'center',
        flexWrap: 'wrap',
        boxShadow: '0 1px 3px rgba(0,0,0,0.04)'
      }}>
        {/* Quick Search */}
        <div style={{ flex: '2 1 240px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
          <label style={{ fontSize: '11.5px', fontWeight: 700, color: '#475569' }}>Tìm kiếm nhanh vật tư:</label>
          <div style={{ position: 'relative', width: '100%' }}>
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value)
                setCurrentPage(1)
              }}
              placeholder="Nhập mã SAP, tên vật tư, quy cách, ĐVT..."
              style={{
                width: '100%',
                padding: '8px 32px 8px 34px',
                border: '1px solid #cbd5e1',
                borderRadius: '8px',
                fontSize: '13px',
                height: '38px',
                boxSizing: 'border-box',
                outline: 'none',
                color: '#1e293b'
              }}
            />
            <Search size={15} color="#94a3b8" style={{ position: 'absolute', left: '11px', top: '12px' }} />
            {searchTerm && (
              <X
                size={14}
                color="#94a3b8"
                onClick={() => setSearchTerm('')}
                style={{ position: 'absolute', right: '11px', top: '12px', cursor: 'pointer' }}
              />
            )}
          </div>
        </div>

        {/* Select Specific Material Dropdown */}
        <div style={{ flex: '2 1 260px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
          <label style={{ fontSize: '11.5px', fontWeight: 700, color: '#475569' }}>Chọn đích danh 1 vật tư:</label>
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

        {/* Category Classification Filter */}
        <div style={{ flex: '1.2 1 170px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
          <label style={{ fontSize: '11.5px', fontWeight: 700, color: '#475569' }}>Nhóm vật tư:</label>
          <select
            value={selectedCategory}
            onChange={(e) => {
              setSelectedCategory(e.target.value)
              setCurrentPage(1)
            }}
            style={{
              padding: '8px 12px',
              border: '1px solid #cbd5e1',
              borderRadius: '8px',
              fontSize: '13px',
              fontWeight: 500,
              background: '#ffffff',
              color: '#1e293b',
              height: '38px',
              outline: 'none'
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
        </div>

        {/* Year Filter */}
        <div style={{ flex: '0.8 1 120px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
          <label style={{ fontSize: '11.5px', fontWeight: 700, color: '#475569' }}>Năm thống kê:</label>
          <select
            value={selectedYear}
            onChange={(e) => {
              setSelectedYear(e.target.value)
              setCurrentPage(1)
            }}
            style={{
              padding: '8px 10px',
              border: '1px solid #cbd5e1',
              borderRadius: '8px',
              fontSize: '13px',
              fontWeight: 600,
              background: '#ffffff',
              color: '#1e293b',
              height: '38px',
              outline: 'none'
            }}
          >
            <option value="ALL">Tất cả năm</option>
            {uniqueYears.map(yr => (
              <option key={yr} value={yr}>Năm {yr}</option>
            ))}
          </select>
        </div>

        {/* Stock Filter Condition */}
        <div style={{ flex: '1 1 150px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
          <label style={{ fontSize: '11.5px', fontWeight: 700, color: '#475569' }}>Trạng thái tồn:</label>
          <select
            value={stockFilter}
            onChange={(e) => {
              setStockFilter(e.target.value)
              setCurrentPage(1)
            }}
            style={{
              padding: '8px 10px',
              border: '1px solid #cbd5e1',
              borderRadius: '8px',
              fontSize: '13px',
              fontWeight: 500,
              background: '#ffffff',
              color: '#1e293b',
              height: '38px',
              outline: 'none'
            }}
          >
            <option value="all">Tất cả số dư</option>
            <option value="active">Chỉ còn tồn (&gt; 0)</option>
            <option value="has_activity">Có phát sinh</option>
            <option value="negative">Tồn âm (&lt; 0)</option>
            <option value="zero">Hết tồn (= 0)</option>
          </select>
        </div>

        {/* Warehouse Filter */}
        <div style={{ flex: '1.2 1 170px', display: 'flex', flexDirection: 'column', gap: '4px' }}>
          <label style={{ fontSize: '11.5px', fontWeight: 700, color: '#475569' }}>Lọc theo Kho:</label>
          <SearchableSelect
            value={selectedWarehouseFilter}
            onChange={(val) => {
              setSelectedWarehouseFilter(val)
              setCurrentPage(1)
            }}
            options={uniquePhysicalWarehouses}
            placeholder="Tất cả các Kho"
            searchPlaceholder="Tìm kiếm kho dự án..."
          />
        </div>

        {/* Hide Empty Warehouses Checkbox Toggle */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginTop: 18, userSelect: 'none', cursor: 'pointer' }}
          onClick={() => setHideEmptyWarehouses(!hideEmptyWarehouses)}
        >
          {hideEmptyWarehouses ? (
            <CheckSquare size={18} color="#2563eb" />
          ) : (
            <Square size={18} color="#94a3b8" />
          )}
          <span style={{ fontSize: '12.5px', color: '#475569', fontWeight: 500 }}>
            Ẩn các kho không có tồn
          </span>
        </div>
      </div>

      {/* Main Table Container */}
      <div style={{
        flex: 1,
        background: '#ffffff',
        border: '1px solid #cbd5e1',
        borderRadius: '12px',
        overflow: 'hidden',
        display: 'flex',
        flexDirection: 'column',
        boxShadow: '0 4px 6px -1px rgba(0,0,0,0.05)'
      }}>
        <div style={{ flex: 1, overflow: 'auto', position: 'relative' }}>
          <table style={{ width: '100%', borderCollapse: 'separate', borderSpacing: 0, textAlign: 'left', fontSize: '12.5px', minWidth: '1250px' }}>
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
                    fontWeight: 800,
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
                      fontSize: '11px',
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
                  <td colSpan={6 + displayPhysicalWarehouses.length} style={{ padding: '48px 20px', textAlign: 'center', color: '#94a3b8', fontSize: '14px' }}>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10 }}>
                      <Package size={40} strokeWidth={1.5} color="#cbd5e1" />
                      <div style={{ fontWeight: 600, color: '#64748b' }}>Không tìm thấy vật tư nào khớp với điều kiện lọc hiện tại.</div>
                      <div style={{ fontSize: '12.5px', color: '#94a3b8' }}>Hãy thử điều chỉnh lại từ khóa tìm kiếm hoặc bỏ bớt các bộ lọc.</div>
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
                            <div style={{ fontSize: '11px', color: '#64748b', fontWeight: 400 }}>{item.thongSoKyThuat}</div>
                          )}
                        </td>
                        <td style={{ padding: '9px 6px', textAlign: 'center', color: '#475569', borderRight: '1px solid #e2e8f0' }}>
                          {item.dvt}
                        </td>
                        <td style={{ padding: '9px 10px', borderRight: '1px solid #e2e8f0' }}>
                          {item.classification ? (
                            <span style={{
                              fontSize: '11px',
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
                            <span style={{ fontSize: '11px', color: '#94a3b8', fontStyle: 'italic' }}>Chưa phân nhóm</span>
                          )}
                        </td>
                        
                        {/* Total Stock Cell */}
                        <td style={{
                          padding: '9px 10px',
                          textAlign: 'right',
                          fontWeight: 800,
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
                                fontSize: '12px'
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
                  <tr style={{ background: '#f1f5f9', fontWeight: 800, borderTop: '2px solid #94a3b8', borderBottom: '2px solid #94a3b8', position: 'sticky', bottom: 0, zIndex: 10 }}>
                    <td colSpan="5" style={{ padding: '12px 14px', color: '#1e3a8a', fontSize: '13px', fontWeight: 800, textAlign: 'right', borderRight: '1px solid #cbd5e1', background: '#e2e8f0' }}>
                      TỔNG CỘNG TỒN THEO KHO:
                    </td>
                    
                    {/* Total of Totals */}
                    <td style={{
                      padding: '12px 10px',
                      textAlign: 'right',
                      background: '#dbeafe',
                      borderLeft: '1px solid #cbd5e1',
                      borderRight: '2px solid #94a3b8',
                      color: '#1e3a8a',
                      fontSize: '13.5px',
                      fontWeight: 800
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
                            fontSize: '12px'
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

        {/* Table Footer & Pagination Controls */}
        <div style={{
          padding: '10px 16px',
          background: '#ffffff',
          borderTop: '1px solid #e2e8f0',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          flexWrap: 'wrap',
          gap: 10
        }}>
          <div style={{ fontSize: '12.5px', color: '#64748b' }}>
            Hiển thị <strong>{paginatedData.length}</strong> / <strong>{sortedMaterialStatsData.length}</strong> vật tư (Tổng hệ thống: {rawMaterialStatsData.length})
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '12.5px', color: '#475569' }}>
              <span>Dòng/trang:</span>
              <select
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value))
                  setCurrentPage(1)
                }}
                style={{
                  padding: '4px 8px',
                  borderRadius: '6px',
                  border: '1px solid #cbd5e1',
                  fontSize: '12px',
                  outline: 'none',
                  background: '#ffffff'
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
                  disabled={currentPage <= 1}
                  onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
                  style={{
                    padding: '5px 8px',
                    borderRadius: '6px',
                    border: '1px solid #cbd5e1',
                    background: currentPage <= 1 ? '#f1f5f9' : '#ffffff',
                    color: currentPage <= 1 ? '#94a3b8' : '#334155',
                    cursor: currentPage <= 1 ? 'not-allowed' : 'pointer',
                    display: 'flex',
                    alignItems: 'center'
                  }}
                >
                  <ChevronLeft size={14} />
                </button>

                <span style={{ fontSize: '12.5px', color: '#334155', padding: '0 6px', fontWeight: 600 }}>
                  Trang {currentPage} / {totalPages}
                </span>

                <button
                  disabled={currentPage >= totalPages}
                  onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
                  style={{
                    padding: '5px 8px',
                    borderRadius: '6px',
                    border: '1px solid #cbd5e1',
                    background: currentPage >= totalPages ? '#f1f5f9' : '#ffffff',
                    color: currentPage >= totalPages ? '#94a3b8' : '#334155',
                    cursor: currentPage >= totalPages ? 'not-allowed' : 'pointer',
                    display: 'flex',
                    alignItems: 'center'
                  }}
                >
                  <ChevronRight size={14} />
                </button>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Material Transactions Drilldown Modal */}
      {detailModalMaterial && (
        <div style={{
          position: 'fixed',
          inset: 0,
          background: 'rgba(15, 23, 42, 0.65)',
          backdropFilter: 'blur(3px)',
          zIndex: 99999,
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          padding: '20px'
        }}>
          <div style={{
            background: '#ffffff',
            borderRadius: '14px',
            width: '100%',
            maxWidth: '960px',
            maxHeight: '90vh',
            display: 'flex',
            flexDirection: 'column',
            overflow: 'hidden',
            boxShadow: '0 20px 25px -5px rgba(0,0,0,0.2)'
          }}>
            {/* Modal Header */}
            <div style={{
              padding: '16px 20px',
              background: 'linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%)',
              color: '#ffffff',
              display: 'flex',
              justifyContent: 'space-between',
              alignItems: 'center'
            }}>
              <div>
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 800 }}>
                  LỊCH SỬ XUẤT NHẬP: {detailModalMaterial.tenVatTu}
                </h3>
                <div style={{ fontSize: '12.5px', color: '#bfdbfe', marginTop: 2 }}>
                  Mã SAP: <strong>{detailModalMaterial.maSAP}</strong> | ĐVT: <strong>{detailModalMaterial.dvt || '---'}</strong> | Nhóm: <strong>{detailModalMaterial.classification || 'Chưa phân nhóm'}</strong>
                </div>
              </div>
              <button
                onClick={() => setDetailModalMaterial(null)}
                style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: '#ffffff', borderRadius: '6px', width: 30, height: 30, display: 'flex', alignItems: 'center', justifyContent: 'center', cursor: 'pointer' }}
              >
                <X size={16} />
              </button>
            </div>

            {/* Modal Body: Table of Transactions */}
            <div style={{ flex: 1, overflowY: 'auto', padding: '16px 20px' }}>
              <div style={{ marginBottom: 12, display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontSize: '13px', color: '#475569', fontWeight: 600 }}>
                  Tổng cộng: {detailModalMaterial.rawTransactions.length} giao dịch phát sinh
                </span>
                <span style={{ fontSize: '13px', fontWeight: 800, color: '#1e3a8a' }}>
                  Tổng tồn hiện tại: {detailModalMaterial.totalStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })} {detailModalMaterial.dvt}
                </span>
              </div>

              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '12px', textAlign: 'left' }}>
                <thead>
                  <tr style={{ background: '#f1f5f9', borderBottom: '2px solid #cbd5e1' }}>
                    <th style={{ padding: '8px', width: '40px', textAlign: 'center' }}>STT</th>
                    <th style={{ padding: '8px', width: '90px' }}>Ngày</th>
                    <th style={{ padding: '8px', width: '100px' }}>Số phiếu</th>
                    <th style={{ padding: '8px' }}>Đơn vị giao</th>
                    <th style={{ padding: '8px' }}>Đơn vị nhận</th>
                    <th style={{ padding: '8px', textAlign: 'right', width: '90px' }}>Nhập</th>
                    <th style={{ padding: '8px', textAlign: 'right', width: '90px' }}>Xuất</th>
                    <th style={{ padding: '8px', width: '120px' }}>Ghi chú</th>
                  </tr>
                </thead>
                <tbody>
                  {detailModalMaterial.rawTransactions.length === 0 ? (
                    <tr>
                      <td colSpan="8" style={{ padding: '24px', textAlign: 'center', color: '#94a3b8' }}>
                        Chưa có lịch sử giao dịch phát sinh nào được ghi nhận.
                      </td>
                    </tr>
                  ) : (
                    detailModalMaterial.rawTransactions.map((tx, idx) => {
                      const dateStr = tx.date ? tx.date.toLocaleDateString('vi-VN') : '---'
                      return (
                        <tr key={idx} style={{ borderBottom: '1px solid #e2e8f0' }}>
                          <td style={{ padding: '8px', textAlign: 'center', color: '#64748b' }}>{idx + 1}</td>
                          <td style={{ padding: '8px', fontWeight: 600, color: '#334155' }}>{dateStr}</td>
                          <td style={{ padding: '8px', color: '#2563eb', fontWeight: 600 }}>{tx.soPhieu || '-'}</td>
                          <td style={{ padding: '8px', color: '#475569' }}>{tx.donViGiao || '-'}</td>
                          <td style={{ padding: '8px', color: '#475569' }}>{tx.donViNhan || '-'}</td>
                          <td style={{ padding: '8px', textAlign: 'right', fontWeight: 700, color: tx.khoiLuongNhap > 0 ? '#059669' : '#94a3b8' }}>
                            {tx.khoiLuongNhap > 0 ? tx.khoiLuongNhap.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                          </td>
                          <td style={{ padding: '8px', textAlign: 'right', fontWeight: 700, color: tx.khoiLuongXuat > 0 ? '#dc2626' : '#94a3b8' }}>
                            {tx.khoiLuongXuat > 0 ? tx.khoiLuongXuat.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                          </td>
                          <td style={{ padding: '8px', color: '#64748b', fontSize: '11.5px' }}>{tx.ghiChu || '-'}</td>
                        </tr>
                      )
                    })
                  )}
                </tbody>
              </table>
            </div>

            {/* Modal Footer */}
            <div style={{ padding: '12px 20px', background: '#f8fafc', borderTop: '1px solid #e2e8f0', display: 'flex', justifyContent: 'flex-end' }}>
              <button
                onClick={() => setDetailModalMaterial(null)}
                style={{
                  padding: '8px 16px',
                  background: '#e2e8f0',
                  color: '#334155',
                  border: 'none',
                  borderRadius: '6px',
                  fontWeight: 600,
                  fontSize: '13px',
                  cursor: 'pointer'
                }}
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
