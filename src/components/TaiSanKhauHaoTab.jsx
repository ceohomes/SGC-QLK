import React, { useState, useMemo, useEffect } from 'react'
import * as XLSXStyleRaw from 'xlsx-js-style'
const XLSXStyle = XLSXStyleRaw.default || XLSXStyleRaw
const XLSX = XLSXStyle

import { 
  DollarSign, Search, Filter, ArrowUpDown, Download, 
  Layers, Package, AlertTriangle, Calendar, Info, CheckCircle2 
} from 'lucide-react'

import { isApprovedStatus } from '../constants.js'

// Simple styled SearchableSelect for use inside the Tab
function LocalSearchableSelect({ value, onChange, options, placeholder = 'Tất cả', searchPlaceholder = 'Tìm kiếm...' }) {
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
    const q = searchQuery.toLowerCase().trim()
    if (!q) return options
    return options.filter(o => String(o || '').toLowerCase().includes(q))
  }, [options, searchQuery])

  return (
    <div ref={containerRef} style={{ position: 'relative', width: '100%' }}>
      <div 
        onClick={() => setIsOpen(!isOpen)}
        style={{
          padding: '8px 12px',
          background: '#ffffff',
          border: '1px solid #cbd5e1',
          borderRadius: '6px',
          fontSize: '13px',
          fontWeight: 500,
          color: '#1e293b',
          cursor: 'pointer',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          minHeight: '38px',
          boxSizing: 'border-box'
        }}
      >
        <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', maxWidth: '90%' }}>
          {value || placeholder}
        </span>
        <span style={{ fontSize: '10px', color: '#64748b' }}>▼</span>
      </div>

      {isOpen && (
        <div style={{
          position: 'absolute',
          top: '102%',
          left: 0,
          right: 0,
          background: '#ffffff',
          border: '1px solid #cbd5e1',
          borderRadius: '6px',
          boxShadow: '0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -4px rgba(0, 0, 0, 0.1)',
          zIndex: 999,
          maxHeight: '260px',
          display: 'flex',
          flexDirection: 'column',
          overflow: 'hidden'
        }}>
          <div style={{ padding: '6px', borderBottom: '1px solid #f1f5f9', background: '#f8fafc' }}>
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder={searchPlaceholder}
              style={{
                width: '100%',
                padding: '6px 10px',
                fontSize: '12px',
                border: '1px solid #e2e8f0',
                borderRadius: '4px',
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
                background: !value ? '#f1f5f9' : 'transparent'
              }}
            >
              -- Bỏ chọn / Tất cả --
            </div>
            {filteredOptions.length === 0 ? (
              <div style={{ padding: '8px 12px', fontSize: '12px', color: '#94a3b8', textAlign: 'center' }}>
                Không tìm thấy kết quả
              </div>
            ) : (
              filteredOptions.map((opt, idx) => (
                <div
                  key={idx}
                  onClick={() => {
                    onChange(opt)
                    setIsOpen(false)
                    setSearchQuery('')
                  }}
                  style={{
                    padding: '8px 12px',
                    fontSize: '12.5px',
                    color: '#334155',
                    cursor: 'pointer',
                    background: value === opt ? '#e2e8f0' : 'transparent',
                    borderBottom: '1px solid #f8fafc'
                  }}
                  onMouseOver={(e) => e.currentTarget.style.background = '#f1f5f9'}
                  onMouseOut={(e) => e.currentTarget.style.background = value === opt ? '#e2e8f0' : 'transparent'}
                >
                  {opt}
                </div>
              ))
            )}
          </div>
        </div>
      )}
    </div>
  )
}

export default function TaiSanKhauHaoTab({ 
  chungRows = [], 
  giaoRows = [], 
  nhanRows = [], 
  allProjects = [],
  customCategoryMap = {},
  materialPriceRows: propMaterialPriceRows,
  materialPrices: propMaterialPrices,
  materialClassifications: propMaterialClassifications
}) {
  const [localProject, setLocalProject] = useState('')
  const [selectedYear, setSelectedYear] = useState(2026)
  const [statusFilter, setStatusFilter] = useState('all') // 'all' | 'khauhao' | 'active'
  const [searchTerm, setSearchTerm] = useState('')
  const [currentPage, setCurrentPage] = useState(1)
  const [pageSize, setPageSize] = useState(50)
  
  // Sub-tabs Selection
  const [activeSubTab, setActiveSubTab] = useState('detail') // 'detail' | 'summary_all' | 'summary_month' | 'material_stats' | 'reconciliation'
  const [selectedMonth, setSelectedMonth] = useState(new Date().getMonth() + 1)
  const [matSortField, setMatSortField] = useState('totalStock')
  const [matSortDirection, setMatSortDirection] = useState('desc')
  
  // Sorting state for detail tab
  const [sortField, setSortField] = useState('stock') // Default sort by stock
  const [sortDirection, setSortDirection] = useState('desc') // 'asc' | 'desc'

  // Sorting state for summary tab
  const [sumSortField, setSumSortField] = useState('projectName')
  const [sumSortDirection, setSumSortDirection] = useState('asc')

  // Discrepancy report states
  const [showDiscrepancyDetail, setShowDiscrepancyDetail] = useState(false)
  const [discrepancySearch, setDiscrepancySearch] = useState('')
  const [discrepancyPage, setDiscrepancyPage] = useState(1)
  const discrepancyPageSize = 10

  const getCategory = (name) => {
    if (!name) return 'chuaphanbo'
    const norm = String(name).trim().replace(/\s+/g, ' ')
    if (customCategoryMap && customCategoryMap[norm]) {
      return customCategoryMap[norm]
    }
    const upper = norm.toUpperCase();
    if (upper.includes('CÔNG TY') || upper.includes('CONG TY') || upper.includes('CTY') || upper.includes('DNTN')) {
      return 'ncc';
    }
    if (upper.includes('TỔ ĐỘI') || upper.includes('TO DOI')) {
      return 'todoi';
    }
    const tokens = upper.split(/[^A-ZÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸYĐ]/);
    if (tokens.includes('KHO') || tokens.includes('SGC') || tokens.includes('BCH')) {
      return 'kho';
    }
    if (/[-/_[\]()0-9]/.test(norm)) {
      return 'chuaphanbo';
    }
    const words = norm.split(/\s+/);
    if (words.length >= 2 && words.length <= 5) {
      const isCapitalized = words.every(w => w && w[0] === w[0].toUpperCase());
      const excludeKeywords = [
        'SGC', 'BCH', 'VFVA', 'HLX', 'TPC', 'TCE', 'HP', 'VP', 'KT', 'GT', 'QD', 'CP', 'TNHH', 'MTV', 
        'TRẠM', 'TRAM', 'BÊ TÔNG', 'BE TONG', 'NHÀ MÁY', 'NHA MAY', 'XÍ NGHIỆP', 'XI NGHIEP', 'DỰ ÁN', 'DU AN',
        'BAN CHỈ HUY', 'BAN CHI HUY', 'HẠ TẦNG', 'HA TANG', 'ĐƯỜNG', 'DUONG', 'CẦU', 'CAU', 'SÂN', 'SAN'
      ];
      const hasExclude = words.some(w => excludeKeywords.includes(w.toUpperCase()));
      if (isCapitalized && !hasExclude) {
        return 'todoi';
      }
    }
    return 'chuaphanbo';
  }

  // Load prices and classifications from props or localStorage
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

  // Days from target date to today
  const getDaysToToday = (dateVal) => {
    const d = parseRowDate(dateVal)
    if (!d) return 0
    const today = new Date()
    today.setHours(0, 0, 0, 0)
    const target = new Date(d)
    target.setHours(0, 0, 0, 0)
    const diff = today.getTime() - target.getTime()
    return Math.max(0, Math.floor(diff / (1000 * 60 * 60 * 24)))
  }

  // Extract all unique warehouses
  const uniqueWarehouses = useMemo(() => {
    const list = new Set()
    const seen = new Set()
    const addName = (name) => {
      if (!name) return
      const norm = name.trim()
      const lower = norm.toLowerCase()
      if (!seen.has(lower)) {
        seen.add(lower)
        list.add(norm)
      }
    }
    allProjects.forEach(p => { if (p) addName(p) })
    chungRows.forEach(r => {
      addName(r.donViGiao)
      addName(r.donViNhan)
    })
    giaoRows.forEach(r => {
      addName(r.donViGiao)
      addName(r.donViNhan)
    })
    nhanRows.forEach(r => {
      addName(r.donViGiao)
      addName(r.donViNhan)
    })
    return [...list].sort()
  }, [allProjects, chungRows, giaoRows, nhanRows])

  // Extract all physical warehouses (excluding suppliers and crews) for dropdown selection
  const uniquePhysicalWarehouses = useMemo(() => {
    return uniqueWarehouses.filter(name => {
      const cat = getCategory(name)
      return cat !== 'ncc' && cat !== 'todoi'
    })
  }, [uniqueWarehouses, getCategory])

  // Extract all unique years
  const uniqueYears = useMemo(() => {
    const years = new Set()
    const sourceRows = (chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows]
    sourceRows.forEach(r => {
      const d = parseRowDate(r.ngayXuatNhap)
      if (d) {
        years.add(d.getFullYear())
      }
    })
    if (years.size === 0) {
      years.add(new Date().getFullYear())
    }
    return [...years].sort((a, b) => b - a)
  }, [chungRows, giaoRows, nhanRows])

  // Set default selected year
  useEffect(() => {
    if (uniqueYears.length > 0 && !uniqueYears.includes(selectedYear)) {
      setSelectedYear(uniqueYears[0])
    }
  }, [uniqueYears])

  // Process data for Assets with 12-month Stock Breakdown
  const assetReportData = useMemo(() => {
    const groups = {}
    const projLower = localProject ? localProject.trim().toLowerCase() : ''

    const processRow = (r) => {
      const nhanUnit = String(r.donViNhan || '').trim()
      const giaoUnit = String(r.donViGiao || '').trim()

      const catNhan = getCategory(nhanUnit)
      const catGiao = getCategory(giaoUnit)

      // Rule: "Đơn liên quan đến tổ đội để giá trị = 0"
      if (catNhan === 'todoi' || catGiao === 'todoi') {
        return { receivedVal: 0, issuedVal: 0, isRelated: true }
      }

      let receivedVal = 0
      let issuedVal = 0

      if (localProject) {
        const nhanLower = nhanUnit.toLowerCase()
        const giaoLower = giaoUnit.toLowerCase()
        const isNhan = nhanLower === projLower
        const isGiao = giaoLower === projLower

        if (!isNhan && !isGiao) {
          return { receivedVal: 0, issuedVal: 0, isRelated: false } // Not related to localProject
        }

        if (isNhan) {
          const val = parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
          if (catGiao === 'ncc' || catGiao === 'kho') {
            receivedVal = val
          }
        }

        if (isGiao) {
          const val = parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
          if (catNhan === 'ncc' || catNhan === 'kho') {
            issuedVal = val
          }
        }
      } else {
        const importVal = parseVal(r.khoiLuongNhap)
        if (importVal > 0) {
          if (catGiao === 'ncc' || catGiao === 'kho') {
            receivedVal = importVal
          }
        }

        const exportVal = parseVal(r.khoiLuongXuat)
        if (exportVal > 0) {
          if (catNhan === 'ncc' || catNhan === 'kho') {
            issuedVal = exportVal
          }
        }
      }

      return { receivedVal, issuedVal, isRelated: true }
    }

    const sourceRows = (chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows]

    sourceRows.forEach(r => {
      // Filter out non-completed/non-approved orders
      if (!isApprovedStatus(r.trangThai)) return

      // Basic validation: must have SAP code
      const sap = String(r.maSAP || '').trim()
      if (!sap) return

      // Determine if it is an asset (classification has 'khấu hao' or 'tài sản')
      const classification = String(customCategoryMap[sap] || materialClassifications[sap] || '').trim()
      const cLower = classification.toLowerCase()
      const isKhauHao = classification && (cLower.includes('khấu hao') || cLower.includes('tài sản') || cLower.includes('tskh'))

      if (!isKhauHao) return

      const rowRes = processRow(r)
      if (!rowRes.isRelated) return

      if (!groups[sap]) {
        groups[sap] = {
          maSAP: sap,
          maVatTu: String(r.maVatTu || '').trim(),
          tenVatTu: String(r.tenVatTu || '').trim(),
          dvt: String(r.dvt || '').trim(),
          thongSoKyThuat: String(r.thongSoKyThuat || '').trim(),
          received: 0,
          issued: 0,
          latestReceivedDate: null,
          latestIssuedDate: null,
          classification: classification,
          rawTransactions: [] // Store transaction list for monthly calculations
        }
      } else {
        // Fallbacks for empty info
        if (!groups[sap].maVatTu && r.maVatTu) groups[sap].maVatTu = String(r.maVatTu).trim()
        if (!groups[sap].tenVatTu && r.tenVatTu) groups[sap].tenVatTu = String(r.tenVatTu).trim()
        if (!groups[sap].dvt && r.dvt) groups[sap].dvt = String(r.dvt).trim()
        if (!groups[sap].thongSoKyThuat && r.thongSoKyThuat) groups[sap].thongSoKyThuat = String(r.thongSoKyThuat).trim()
      }

      const rowDate = parseRowDate(r.ngayXuatNhap)
      const { receivedVal, issuedVal } = rowRes

      groups[sap].received += receivedVal
      groups[sap].issued += issuedVal
      const qty = receivedVal - issuedVal

      if (receivedVal > 0 && rowDate && (!groups[sap].latestReceivedDate || rowDate > groups[sap].latestReceivedDate)) {
        groups[sap].latestReceivedDate = rowDate
      }
      if (issuedVal > 0 && rowDate && (!groups[sap].latestIssuedDate || rowDate > groups[sap].latestIssuedDate)) {
        groups[sap].latestIssuedDate = rowDate
      }

      // Add to transaction list for monthly breakdown
      groups[sap].rawTransactions.push({
        date: rowDate,
        qty: qty
      })
    })

    // Map and calculate stock, days unused, depreciation pricing
    const list = Object.values(groups).map(g => {
      const stock = g.received - g.issued
      const daysUnused = getDaysToToday(g.latestReceivedDate)
      const unusedStatus = (stock > 0 && daysUnused > 30) ? 'Chưa sử dụng (> 30 ngày)' : 'Đang sử dụng'

      // Pricing details
      const priceRow = materialPriceRows.find(r => String(r.maSAP || '').trim().toLowerCase() === g.maSAP.toLowerCase())
      const estimatedUnitPrice = priceRow ? (priceRow.donGiaTrungBinh1Ngay || 0) : 0
      const averageUnitPrice = priceRow ? (priceRow.donGiaTrungBinh || 0) : (materialPrices[g.maSAP] || 0)

      // Calculate stock at the end of each of the 12 months for the selectedYear
      const monthlyStock = {}
      for (let m = 1; m <= 12; m++) {
        const cutoffDate = new Date(selectedYear, m, 0, 23, 59, 59, 999)
        let monthlyBalance = 0
        g.rawTransactions.forEach(t => {
          if (!t.date || t.date <= cutoffDate) {
            monthlyBalance += t.qty
          }
        })
        monthlyStock[`T${m}`] = Math.round(monthlyBalance * 1000) / 1000
      }

      // Accumulated Depreciation Value for over 30 days unused items
      let valueOver30Days = 0
      if (unusedStatus === 'Chưa sử dụng (> 30 ngày)' && stock > 0) {
        valueOver30Days = Math.round(daysUnused * stock * estimatedUnitPrice)
      }

      const totalAssetValue = Math.round(stock * averageUnitPrice)

      return {
        ...g,
        stock: Math.round(stock * 1000) / 1000,
        daysUnused,
        unusedStatus,
        estimatedUnitPrice,
        averageUnitPrice,
        valueOver30Days,
        totalAssetValue,
        monthlyStock
      }
    })

    return list
  }, [chungRows, giaoRows, nhanRows, localProject, selectedYear, customCategoryMap, materialClassifications, materialPriceRows, materialPrices])

  // Apply client-side filtering on report data in a fast, decoupled useMemo
  const filteredAssetReportData = useMemo(() => {
    const text = searchTerm.toLowerCase().trim()
    return assetReportData.filter(item => {
      // Filter out items that have no monthly data at all (all T1..T12 are 0) unless user is explicitly searching
      const hasMonthlyData = Object.values(item.monthlyStock).some(val => val !== 0)
      if (!hasMonthlyData && !text) return false

      // 1. Text filter
      const matchText = !text || 
        (item.maSAP || '').toLowerCase().includes(text) ||
        (item.maVatTu || '').toLowerCase().includes(text) ||
        (item.tenVatTu || '').toLowerCase().includes(text) ||
        (item.thongSoKyThuat || '').toLowerCase().includes(text)

      // 2. Status filter
      let matchStatus = true
      if (statusFilter === 'khauhao') {
        matchStatus = item.unusedStatus === 'Chưa sử dụng (> 30 ngày)'
      } else if (statusFilter === 'active') {
        matchStatus = item.unusedStatus === 'Đang sử dụng'
      }

      return matchText && matchStatus
    })
  }, [assetReportData, searchTerm, statusFilter])

  // Sort report data
  const sortedData = useMemo(() => {
    const sorted = [...filteredAssetReportData]
    sorted.sort((a, b) => {
      let valA = a[sortField]
      let valB = b[sortField]

      if (sortField.startsWith('T')) {
        valA = a.monthlyStock[sortField] || 0
        valB = b.monthlyStock[sortField] || 0
      }

      if (typeof valA === 'string') {
        return sortDirection === 'asc' 
          ? valA.localeCompare(valB) 
          : valB.localeCompare(valA)
      } else {
        const nA = valA || 0
        const nB = valB || 0
        return sortDirection === 'asc' ? nA - nB : nB - nA
      }
    })
    return sorted
  }, [filteredAssetReportData, sortField, sortDirection])

  // Process and aggregate project-wise data for Company-wide Summary
  const rawCompanyProjectSummaries = useMemo(() => {
    // Filter out todoi (crews) entirely from the company-wide summary list
    const filteredWarehouses = uniqueWarehouses.filter(name => getCategory(name) !== 'todoi')

    const summaries = filteredWarehouses.map(projName => {
      const groups = {}
      const projLower = projName.trim().toLowerCase()
      const projCat = getCategory(projName)
      
      const processRow = (r) => {
        const nhanUnit = String(r.donViNhan || '').trim()
        const giaoUnit = String(r.donViGiao || '').trim()

        const catNhan = getCategory(nhanUnit)
        const catGiao = getCategory(giaoUnit)

        // Rule: "Đơn liên quan đến tổ đội để giá trị = 0"
        if (catNhan === 'todoi' || catGiao === 'todoi') {
          return { receivedVal: 0, issuedVal: 0, isRelated: true }
        }

        let receivedVal = 0
        let issuedVal = 0

        const nhanLower = nhanUnit.toLowerCase()
        const giaoLower = giaoUnit.toLowerCase()

        if (projCat === 'ncc') {
          // For a Supplier:
          // 'received' (supply) is when Supplier issues/supplies to a warehouse/ncc (giaoUnit is Supplier)
          // 'issued' (return) is when Supplier receives/gets returned from a warehouse/ncc (nhanUnit is Supplier)
          const isGiao = giaoLower === projLower
          const isNhan = nhanLower === projLower

          if (!isGiao && !isNhan) {
            return { receivedVal: 0, issuedVal: 0, isRelated: false }
          }

          if (isGiao) {
            receivedVal = parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
          }
          if (isNhan) {
            issuedVal = parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
          }
        } else {
          // For a Warehouse (projCat === 'kho' or any other non-ncc name):
          const isNhan = nhanLower === projLower
          const isGiao = giaoLower === projLower

          if (!isNhan && !isGiao) {
            return { receivedVal: 0, issuedVal: 0, isRelated: false }
          }

          if (isNhan) {
            receivedVal = parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
          }
          if (isGiao) {
            issuedVal = parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
          }
        }

        return { receivedVal, issuedVal, isRelated: true }
      }

      const sourceRows = (chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows]

      sourceRows.forEach(r => {
        // Filter out non-completed/non-approved orders
        if (!isApprovedStatus(r.trangThai)) return

        const sap = String(r.maSAP || '').trim()
        if (!sap) return

        const classification = String(customCategoryMap[sap] || materialClassifications[sap] || '').trim()
        const cLower = classification.toLowerCase()
        const isKhauHao = classification && (cLower.includes('khấu hao') || cLower.includes('tài sản') || cLower.includes('tskh'))
        if (!isKhauHao) return

        const rowRes = processRow(r)
        if (!rowRes.isRelated) return

        if (!groups[sap]) {
          groups[sap] = {
            maSAP: sap,
            received: 0,
            issued: 0,
            latestReceivedDate: null,
            rawTransactions: []
          }
        }

        const rowDate = parseRowDate(r.ngayXuatNhap)
        const { receivedVal, issuedVal } = rowRes

        groups[sap].received += receivedVal
        groups[sap].issued += issuedVal
        const qty = receivedVal - issuedVal

        if (receivedVal > 0 && rowDate && (!groups[sap].latestReceivedDate || rowDate > groups[sap].latestReceivedDate)) {
          groups[sap].latestReceivedDate = rowDate
        }

        groups[sap].rawTransactions.push({
          date: rowDate,
          qty: qty
        })
      })

      let totalAssets = 0
      let totalStock = 0
      let totalValue = 0
      let totalDepreciation = 0

      const monthlyStockSum = {
        T1: 0, T2: 0, T3: 0, T4: 0, T5: 0, T6: 0, T7: 0, T8: 0, T9: 0, T10: 0, T11: 0, T12: 0
      }

      Object.values(groups).forEach(g => {
        const stock = g.received - g.issued
        const priceRow = materialPriceRows.find(pr => String(pr.maSAP || '').trim().toLowerCase() === g.maSAP.toLowerCase())
        const estimatedUnitPrice = priceRow ? (priceRow.donGiaTrungBinh1Ngay || 0) : 0
        const averageUnitPrice = priceRow ? (priceRow.donGiaTrungBinh || 0) : (materialPrices[g.maSAP] || 0)

        const totalAssetValue = Math.round(stock * averageUnitPrice)
        
        const daysUnused = getDaysToToday(g.latestReceivedDate)
        const unusedStatus = (stock > 0 && daysUnused > 30) ? 'Chưa sử dụng (> 30 ngày)' : 'Đang sử dụng'

        let valueOver30Days = 0
        if (unusedStatus === 'Chưa sử dụng (> 30 ngày)' && stock > 0) {
          valueOver30Days = Math.round(daysUnused * stock * estimatedUnitPrice)
        }

        if (stock !== 0 || valueOver30Days > 0) {
          totalAssets++
          totalStock += stock
          totalValue += totalAssetValue
          totalDepreciation += valueOver30Days
        }

        for (let m = 1; m <= 12; m++) {
          const cutoffDate = new Date(selectedYear, m, 0, 23, 59, 59, 999)
          let monthlyBalance = 0
          g.rawTransactions.forEach(t => {
            if (!t.date || t.date <= cutoffDate) {
              monthlyBalance += t.qty
            }
          })
          monthlyStockSum[`T${m}`] += monthlyBalance
        }
      })

      for (let m = 1; m <= 12; m++) {
        monthlyStockSum[`T${m}`] = Math.round(monthlyStockSum[`T${m}`] * 1000) / 1000
      }

      return {
        projectName: projName,
        totalAssets,
        totalStock: Math.round(totalStock * 1000) / 1000,
        totalValue,
        totalDepreciation,
        monthlyStock: monthlyStockSum
      }
    })

    return summaries
  }, [uniqueWarehouses, selectedYear, chungRows, giaoRows, nhanRows, customCategoryMap, materialClassifications, materialPriceRows, materialPrices])

  const companyProjectSummaries = useMemo(() => {
    // Filter projects with actual data
    let filtered = rawCompanyProjectSummaries.filter(s => s.totalAssets > 0 || s.totalStock > 0 || Object.values(s.monthlyStock).some(v => v !== 0))
    
    // Filter by name if search is active and active sub-tab is summary_all
    if (searchTerm.trim() && activeSubTab === 'summary_all') {
      const q = searchTerm.toLowerCase().trim()
      filtered = filtered.filter(s => (s.projectName || '').toLowerCase().includes(q))
    }

    // Sort company-wide summaries
    filtered.sort((a, b) => {
      let valA = a[sumSortField]
      let valB = b[sumSortField]

      if (sumSortField.startsWith('T')) {
        valA = a.monthlyStock[sumSortField] || 0
        valB = b.monthlyStock[sumSortField] || 0
      }

      if (typeof valA === 'string') {
        return sumSortDirection === 'asc' 
          ? valA.localeCompare(valB) 
          : valB.localeCompare(valA)
      } else {
        const nA = valA || 0
        const nB = valB || 0
        return sumSortDirection === 'asc' ? nA - nB : nB - nA
      }
    })

    return filtered
  }, [rawCompanyProjectSummaries, searchTerm, sumSortField, sumSortDirection, activeSubTab])

  // Process and aggregate project-wise data for Monthly Summary (material-wise breakdown)
  const rawCompanyProjectMonthlySummaries = useMemo(() => {
    // Filter out todoi (crews) entirely from the company-wide summary list
    const filteredWarehouses = uniqueWarehouses.filter(name => getCategory(name) !== 'todoi')

    const flatList = []

    filteredWarehouses.forEach(projName => {
      const groups = {}
      const projLower = projName.trim().toLowerCase()
      const projCat = getCategory(projName)
      
      const processRow = (r) => {
        const nhanUnit = String(r.donViNhan || '').trim()
        const giaoUnit = String(r.donViGiao || '').trim()

        const catNhan = getCategory(nhanUnit)
        const catGiao = getCategory(giaoUnit)

        // Rule: "Đơn liên quan đến tổ đội để giá trị = 0"
        if (catNhan === 'todoi' || catGiao === 'todoi') {
          return { receivedVal: 0, issuedVal: 0, isRelated: true }
        }

        let receivedVal = 0
        let issuedVal = 0

        const nhanLower = nhanUnit.toLowerCase()
        const giaoLower = giaoUnit.toLowerCase()

        if (projCat === 'ncc') {
          const isGiao = giaoLower === projLower
          const isNhan = nhanLower === projLower

          if (!isGiao && !isNhan) {
            return { receivedVal: 0, issuedVal: 0, isRelated: false }
          }

          if (isGiao) {
            receivedVal = parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
          }
          if (isNhan) {
            issuedVal = parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
          }
        } else {
          const isNhan = nhanLower === projLower
          const isGiao = giaoLower === projLower

          if (!isNhan && !isGiao) {
            return { receivedVal: 0, issuedVal: 0, isRelated: false }
          }

          if (isNhan) {
            receivedVal = parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
          }
          if (isGiao) {
            issuedVal = parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
          }
        }

        return { receivedVal, issuedVal, isRelated: true }
      }

      const sourceRows = (chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows]

      sourceRows.forEach(r => {
        // Filter out non-completed/non-approved orders
        if (!isApprovedStatus(r.trangThai)) return

        const sap = String(r.maSAP || '').trim()
        if (!sap) return

        const classification = String(customCategoryMap[sap] || materialClassifications[sap] || '').trim()
        const cLower = classification.toLowerCase()
        const isKhauHao = classification && (cLower.includes('khấu hao') || cLower.includes('tài sản') || cLower.includes('tskh'))
        if (!isKhauHao) return

        const rowRes = processRow(r)
        if (!rowRes.isRelated) return

        if (!groups[sap]) {
          groups[sap] = {
            maSAP: sap,
            tenVatTu: String(r.tenVatTu || '').trim(),
            dvt: String(r.dvt || '').trim(),
            received: 0,
            issued: 0,
            rawTransactions: []
          }
        } else {
          if (!groups[sap].tenVatTu && r.tenVatTu) groups[sap].tenVatTu = String(r.tenVatTu).trim()
          if (!groups[sap].dvt && r.dvt) groups[sap].dvt = String(r.dvt).trim()
        }

        const rowDate = parseRowDate(r.ngayXuatNhap)
        const { receivedVal, issuedVal } = rowRes

        groups[sap].received += receivedVal
        groups[sap].issued += issuedVal
        const qty = receivedVal - issuedVal

        groups[sap].rawTransactions.push({
          date: rowDate,
          qty: qty
        })
      })

      Object.values(groups).forEach(g => {
        const stock = g.received - g.issued
        
        const monthlyStockSum = {
          T1: 0, T2: 0, T3: 0, T4: 0, T5: 0, T6: 0, T7: 0, T8: 0, T9: 0, T10: 0, T11: 0, T12: 0
        }

        for (let m = 1; m <= 12; m++) {
          const cutoffDate = new Date(selectedYear, m, 0, 23, 59, 59, 999)
          let monthlyBalance = 0
          g.rawTransactions.forEach(t => {
            if (!t.date || t.date <= cutoffDate) {
              monthlyBalance += t.qty
            }
          })
          monthlyStockSum[`T${m}`] = Math.round(monthlyBalance * 1000) / 1000
        }

        const priceRow = materialPriceRows.find(pr => String(pr.maSAP || '').trim().toLowerCase() === g.maSAP.toLowerCase())
        const averageUnitPrice = priceRow ? (priceRow.donGiaTrungBinh || 0) : (materialPrices[g.maSAP] || 0)
        const totalValue = Math.round(stock * averageUnitPrice)

        const hasActivity = stock !== 0 || Object.values(monthlyStockSum).some(v => v !== 0)
        if (hasActivity) {
          flatList.push({
            projectName: projName,
            maSAP: g.maSAP,
            tenVatTu: g.tenVatTu,
            dvt: g.dvt,
            totalStock: Math.round(stock * 1000) / 1000,
            totalValue,
            unitPrice: averageUnitPrice,
            monthlyStock: monthlyStockSum
          })
        }
      })
    })

    return flatList
  }, [uniqueWarehouses, selectedYear, chungRows, giaoRows, nhanRows, customCategoryMap, materialClassifications, materialPriceRows, materialPrices])

  const companyProjectMonthlySummaries = useMemo(() => {
    let filtered = rawCompanyProjectMonthlySummaries
    
    // Filter by name/SAP/material if search is active and active sub-tab is summary_month
    if (searchTerm.trim() && activeSubTab === 'summary_month') {
      const q = searchTerm.toLowerCase().trim()
      filtered = filtered.filter(s => 
        (s.projectName || '').toLowerCase().includes(q) ||
        (s.maSAP || '').toLowerCase().includes(q) ||
        (s.tenVatTu || '').toLowerCase().includes(q)
      )
    }

    // Sort company-wide summaries
    filtered.sort((a, b) => {
      let valA = a[sumSortField]
      let valB = b[sumSortField]

      if (sumSortField.startsWith('T')) {
        valA = a.monthlyStock[sumSortField] || 0
        valB = b.monthlyStock[sumSortField] || 0
      }

      if (typeof valA === 'string') {
        return sumSortDirection === 'asc' 
          ? valA.localeCompare(valB) 
          : valB.localeCompare(valA)
      } else {
        const nA = valA || 0
        const nB = valB || 0
        return sumSortDirection === 'asc' ? nA - nB : nB - nA
      }
    })

    return filtered
  }, [rawCompanyProjectMonthlySummaries, searchTerm, sumSortField, sumSortDirection, activeSubTab])

  const groupedProjectMonthlySummaries = useMemo(() => {
    const khoBch = []
    const ncc = []

    companyProjectMonthlySummaries.forEach(item => {
      const cat = getCategory(item.projectName)
      if (cat === 'ncc') {
        ncc.push(item)
      } else {
        khoBch.push(item)
      }
    })

    return [
      { id: 'I', title: 'I. KHO DỰ ÁN / KHO BCH (A)', items: khoBch },
      { id: 'II', title: 'II. NHẬP TỪ NHÀ CUNG CẤP (B)', items: ncc }
    ]
  }, [companyProjectMonthlySummaries, getCategory])

  const reconciliationMonthlyDifference = useMemo(() => {
    const groupI = groupedProjectMonthlySummaries.find(g => g.id === 'I')
    const groupII = groupedProjectMonthlySummaries.find(g => g.id === 'II')

    const itemsI = groupI ? groupI.items : []
    const itemsII = groupII ? groupII.items : []

    const stockI = itemsI.reduce((sum, item) => sum + item.totalStock, 0)
    const stockII = itemsII.reduce((sum, item) => sum + item.totalStock, 0)

    const monthlyDiff = {}
    const monthlyAmtDiff = {}
    for (let m = 1; m <= 12; m++) {
      const valI = itemsI.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0), 0)
      const valII = itemsII.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0), 0)
      monthlyDiff[`T${m}`] = Math.round((valI - valII) * 1000) / 1000

      const amtI = itemsI.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0) * (item.unitPrice || 0), 0)
      const amtII = itemsII.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0) * (item.unitPrice || 0), 0)
      monthlyAmtDiff[`T${m}`] = Math.round(amtI - amtII)
    }

    return {
      stockDiff: Math.round((stockI - stockII) * 1000) / 1000,
      monthlyDiff,
      monthlyAmtDiff
    }
  }, [groupedProjectMonthlySummaries])

  const groupedProjectSummaries = useMemo(() => {
    const khoBch = []
    const ncc = []

    companyProjectSummaries.forEach(item => {
      const cat = getCategory(item.projectName)
      if (cat === 'ncc') {
        ncc.push(item)
      } else {
        khoBch.push(item)
      }
    })

    return [
      { id: 'I', title: 'I. KHO DỰ ÁN / KHO BCH (A)', items: khoBch },
      { id: 'II', title: 'II. NHẬP TỪ NHÀ CUNG CẤP (B)', items: ncc }
    ]
  }, [companyProjectSummaries, getCategory])

  const reconciliationDifference = useMemo(() => {
    const groupI = groupedProjectSummaries.find(g => g.id === 'I')
    const groupII = groupedProjectSummaries.find(g => g.id === 'II')

    const itemsI = groupI ? groupI.items : []
    const itemsII = groupII ? groupII.items : []

    const assetsI = itemsI.reduce((sum, item) => sum + item.totalAssets, 0)
    const assetsII = itemsII.reduce((sum, item) => sum + item.totalAssets, 0)

    const stockI = itemsI.reduce((sum, item) => sum + item.totalStock, 0)
    const stockII = itemsII.reduce((sum, item) => sum + item.totalStock, 0)

    const valueI = itemsI.reduce((sum, item) => sum + item.totalValue, 0)
    const valueII = itemsII.reduce((sum, item) => sum + item.totalValue, 0)

    const depreciationI = itemsI.reduce((sum, item) => sum + item.totalDepreciation, 0)
    const depreciationII = itemsII.reduce((sum, item) => sum + item.totalDepreciation, 0)

    const monthlyDiff = {}
    for (let m = 1; m <= 12; m++) {
      const sumI = itemsI.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0), 0)
      const sumII = itemsII.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0), 0)
      monthlyDiff[`T${m}`] = sumI - sumII
    }

    return {
      assetsDiff: assetsI - assetsII,
      stockDiff: stockI - stockII,
      valueDiff: valueI - valueII,
      depreciationDiff: depreciationI - depreciationII,
      monthlyDiff
    }
  }, [groupedProjectSummaries])

  // Calculate material-level stock for Group I (Kho Dự án) and Group II (Nhà cung cấp)
  const materialReconciliationData = useMemo(() => {
    const materialMap = {}

    const sourceRows = (chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows]

    sourceRows.forEach(r => {
      if (!isApprovedStatus(r.trangThai)) return

      const sap = String(r.maSAP || '').trim()
      if (!sap) return

      // Determine if it is an asset (classification has 'khấu hao' or 'tài sản')
      const classification = String(customCategoryMap[sap] || materialClassifications[sap] || '').trim()
      const cLower = classification.toLowerCase()
      const isKhauHao = classification && (cLower.includes('khấu hao') || cLower.includes('tài sản') || cLower.includes('tskh'))
      if (!isKhauHao) return

      const nhanUnit = String(r.donViNhan || '').trim()
      const giaoUnit = String(r.donViGiao || '').trim()

      const catNhan = getCategory(nhanUnit)
      const catGiao = getCategory(giaoUnit)

      // Rule: "Đơn liên quan đến tổ đội để giá trị = 0"
      if (catNhan === 'todoi' || catGiao === 'todoi') {
        return
      }

      if (!materialMap[sap]) {
        materialMap[sap] = {
          maSAP: sap,
          tenVatTu: String(r.tenVatTu || r.maSAP || ''),
          stockA: 0,
          stockB: 0,
          rawTransactions: []
        }
      }

      const valNhan = parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
      const valXuatNhan = parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
      const valGiao = parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
      const valXuatGiao = parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)

      // Group I (Kho Dự án) stock changes
      if (nhanUnit && catNhan !== 'ncc') {
        materialMap[sap].stockA += valNhan
      }
      if (giaoUnit && catGiao !== 'ncc') {
        materialMap[sap].stockA -= valXuatGiao
      }

      // Group II (Nhà cung cấp) stock changes
      if (giaoUnit && catGiao === 'ncc') {
        materialMap[sap].stockB += valGiao
      }
      if (nhanUnit && catNhan === 'ncc') {
        materialMap[sap].stockB -= valXuatNhan
      }

      // Record this transaction for the material's lifecycle
      materialMap[sap].rawTransactions.push({
        maDon: r.maDon || 'N/A',
        date: r.ngayXuatNhap ? String(r.ngayXuatNhap) : 'N/A',
        giaoUnit: giaoUnit || 'N/A',
        nhanUnit: nhanUnit || 'N/A',
        qty: valNhan || valGiao || 0,
        trangThai: r.trangThai || 'N/A',
        catGiao,
        catNhan
      })
    })

    return materialMap
  }, [chungRows, giaoRows, nhanRows, getCategory, customCategoryMap, materialClassifications])

  // Aggregate discrepancies by Material SAP code and sort with negative differences first
  const materialMismatchSummary = useMemo(() => {
    return Object.values(materialReconciliationData).map(item => {
      const priceRow = materialPriceRows?.find(pr => String(pr.maSAP || '').trim().toLowerCase() === item.maSAP.toLowerCase())
      const averageUnitPrice = priceRow ? (priceRow.donGiaTrungBinh || 0) : (materialPrices?.[item.maSAP] || 0)
      
      const diff = item.stockA - item.stockB
      const valueDiff = Math.round(diff * averageUnitPrice)

      return {
        maSAP: item.maSAP,
        tenVatTu: item.tenVatTu,
        effectI: item.stockA,
        effectII: item.stockB,
        diff: Math.round(diff * 1000) / 1000,
        valueDiff
      }
    })
    .filter(item => Math.abs(item.diff) > 0.001) // Show only those with discrepancies
    .sort((a, b) => a.diff - b.diff) // Most negative difference first
  }, [materialReconciliationData, materialPriceRows, materialPrices])

  // Compute precise transaction discrepancies in chronological order (Lifecycle) for discrepant materials
  const mismatchAnalysis = useMemo(() => {
    const discrepantSaps = new Set(materialMismatchSummary.map(m => m.maSAP))
    const list = []

    Object.values(materialReconciliationData).forEach(item => {
      if (!discrepantSaps.has(item.maSAP)) return

      item.rawTransactions.forEach(t => {
        // Calculate the impacts to represent in the detailed spreadsheet/table:
        let effectI = 0
        let effectII = 0

        if (t.nhanUnit && t.catNhan !== 'ncc') {
          effectI += t.qty
        }
        if (t.giaoUnit && t.catGiao !== 'ncc') {
          effectI -= t.qty
        }

        if (t.giaoUnit && t.catGiao === 'ncc') {
          effectII += t.qty
        }
        if (t.nhanUnit && t.catNhan === 'ncc') {
          effectII -= t.qty
        }

        const diff = effectI - effectII

        let reason = ''
        if (!t.nhanUnit || t.nhanUnit === 'N/A') {
          reason = 'Khuyết Đơn vị nhận (vật tư chưa biết nhập về kho nào)'
        } else if (!t.giaoUnit || t.giaoUnit === 'N/A') {
          reason = 'Khuyết Đơn vị giao (chưa rõ nguồn xuất phát)'
        } else if (t.catGiao === 'ncc' && t.catNhan !== 'kho') {
          reason = `Lưu chuyển lệch dòng: Nhà cung cấp cấp phát trực tiếp cho đơn vị phụ (${t.nhanUnit})`
        } else if (t.catGiao === 'kho' && t.catNhan === 'chuaphanbo') {
          reason = `Lưu chuyển lệch dòng: Xuất kho cho đơn vị chưa phân loại (${t.nhanUnit})`
        } else if (Math.abs(diff) > 0.001) {
          reason = 'Chênh lệch do luân chuyển nội bộ lệch dòng'
        } else {
          reason = 'Giao nhận khớp hoàn toàn tại giao dịch này'
        }

        list.push({
          maDon: t.maDon,
          date: t.date,
          maSAP: item.maSAP,
          tenVatTu: item.tenVatTu,
          giaoUnit: t.giaoUnit,
          nhanUnit: t.nhanUnit,
          qty: t.qty,
          effectI,
          effectII,
          diff,
          reason
        })
      })
    })

    return list.sort((a, b) => {
      const dateA = parseRowDate(a.date) || new Date(0)
      const dateB = parseRowDate(b.date) || new Date(0)
      return dateA.getTime() - dateB.getTime()
    })
  }, [materialReconciliationData, materialMismatchSummary])

  const filteredMaterialMismatchSummary = useMemo(() => {
    if (!searchTerm.trim()) return materialMismatchSummary
    const q = searchTerm.toLowerCase().trim()
    return materialMismatchSummary.filter(m => 
      m.maSAP.toLowerCase().includes(q) || 
      m.tenVatTu.toLowerCase().includes(q)
    )
  }, [materialMismatchSummary, searchTerm])

  // Stats calculation
  const stats = useMemo(() => {
    let totalAssets = 0
    let totalStock = 0
    let totalValue = 0
    let totalDepreciation = 0
    let unusedCount = 0

    assetReportData.forEach(item => {
      totalAssets++
      if (item.stock > 0) {
        totalStock += item.stock
        totalValue += item.totalAssetValue || 0
        if (item.unusedStatus === 'Chưa sử dụng (> 30 ngày)') {
          unusedCount++
          totalDepreciation += item.valueOver30Days || 0
        }
      }
    })

    return {
      totalAssets,
      totalStock,
      totalValue,
      totalDepreciation,
      unusedCount
    }
  }, [assetReportData])


  const grandMonthlyTotals = useMemo(() => {
    const totals = { T1: 0, T2: 0, T3: 0, T4: 0, T5: 0, T6: 0, T7: 0, T8: 0, T9: 0, T10: 0, T11: 0, T12: 0 }
    // Only sum physical warehouses (Group I) to avoid double counting
    const physicalProjects = companyProjectSummaries.filter(p => getCategory(p.projectName) !== 'ncc')
    physicalProjects.forEach(p => {
      for (let m = 1; m <= 12; m++) {
        totals[`T${m}`] += p.monthlyStock[`T${m}`] || 0
      }
    })
    for (let m = 1; m <= 12; m++) {
      totals[`T${m}`] = Math.round(totals[`T${m}`] * 1000) / 1000
    }
    return totals
  }, [companyProjectSummaries])

  const projectStocksByMaterial = useMemo(() => {
    const map = {} // { [maSAP]: { [projName]: stock } }
    
    // Dates up to the end of selectedYear
    const cutoffDate = new Date(selectedYear, 12, 0, 23, 59, 59, 999)
    const sourceRows = (chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows]
    
    sourceRows.forEach(r => {
      if (!isApprovedStatus(r.trangThai)) return
      
      const sap = String(r.maSAP || '').trim()
      if (!sap) return
      
      const classification = String(customCategoryMap[sap] || materialClassifications[sap] || '').trim()
      const cLower = classification.toLowerCase()
      const isKhauHao = classification && (cLower.includes('khấu hao') || cLower.includes('tài sản') || cLower.includes('tskh'))
      if (!isKhauHao) return

      const rowDate = parseRowDate(r.ngayXuatNhap)
      if (!rowDate || rowDate > cutoffDate) return

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

    return map
  }, [uniquePhysicalWarehouses, selectedYear, chungRows, giaoRows, nhanRows, customCategoryMap, materialClassifications])

  const materialStatsData = useMemo(() => {
    const groups = {}
    // We ignore localProject because "Thống kê theo vật tư" must always compile across all BCH warehouses
    const targetProject = ''
    const projLower = ''

    const processRow = (r) => {
      const nhanUnit = String(r.donViNhan || '').trim()
      const giaoUnit = String(r.donViGiao || '').trim()

      const catNhan = getCategory(nhanUnit)
      const catGiao = getCategory(giaoUnit)

      // Rule: "Đơn liên quan đến tổ đội để giá trị = 0"
      if (catNhan === 'todoi' || catGiao === 'todoi') {
        return { receivedVal: 0, issuedVal: 0, isRelated: true }
      }

      let receivedVal = 0
      let issuedVal = 0

      if (targetProject) {
        const nhanLower = nhanUnit.toLowerCase()
        const giaoLower = giaoUnit.toLowerCase()
        const isNhan = nhanLower === projLower
        const isGiao = giaoLower === projLower

        if (!isNhan && !isGiao) {
          return { receivedVal: 0, issuedVal: 0, isRelated: false } // Not related to localProject
        }

        if (isNhan) {
          const val = parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
          if (catGiao === 'ncc' || catGiao === 'kho') {
            receivedVal = val
          }
        }

        if (isGiao) {
          const val = parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
          if (catNhan === 'ncc' || catNhan === 'kho') {
            issuedVal = val
          }
        }
      } else {
        const importVal = parseVal(r.khoiLuongNhap)
        if (importVal > 0) {
          if (catGiao === 'ncc' || catGiao === 'kho') {
            receivedVal = importVal
          }
        }

        const exportVal = parseVal(r.khoiLuongXuat)
        if (exportVal > 0) {
          if (catNhan === 'ncc' || catNhan === 'kho') {
            issuedVal = exportVal
          }
        }
      }

      return { receivedVal, issuedVal, isRelated: true }
    }

    const sourceRows = (chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows]

    sourceRows.forEach(r => {
      if (!isApprovedStatus(r.trangThai)) return

      const sap = String(r.maSAP || '').trim()
      if (!sap) return

      const classification = String(customCategoryMap[sap] || materialClassifications[sap] || '').trim()
      const cLower = classification.toLowerCase()
      const isKhauHao = classification && (cLower.includes('khấu hao') || cLower.includes('tài sản') || cLower.includes('tskh'))

      if (!isKhauHao) return

      const rowRes = processRow(r)
      if (!rowRes.isRelated) return

      if (!groups[sap]) {
        groups[sap] = {
          maSAP: sap,
          maVatTu: String(r.maVatTu || '').trim(),
          tenVatTu: String(r.tenVatTu || '').trim(),
          dvt: String(r.dvt || '').trim(),
          thongSoKyThuat: String(r.thongSoKyThuat || '').trim(),
          classification: classification,
          rawTransactions: []
        }
      } else {
        if (!groups[sap].maVatTu && r.maVatTu) groups[sap].maVatTu = String(r.maVatTu).trim()
        if (!groups[sap].tenVatTu && r.tenVatTu) groups[sap].tenVatTu = String(r.tenVatTu).trim()
        if (!groups[sap].dvt && r.dvt) groups[sap].dvt = String(r.dvt).trim()
        if (!groups[sap].thongSoKyThuat && r.thongSoKyThuat) groups[sap].thongSoKyThuat = String(r.thongSoKyThuat).trim()
      }

      const rowDate = parseRowDate(r.ngayXuatNhap)
      const { receivedVal, issuedVal } = rowRes

      groups[sap].rawTransactions.push({
        date: rowDate,
        receivedVal,
        issuedVal
      })
    })

    // Dates for the selectedYear
    const startOfYear = new Date(selectedYear, 0, 1, 0, 0, 0, 0)
    const endOfYear = new Date(selectedYear, 12, 0, 23, 59, 59, 999)

    const list = Object.values(groups).map(g => {
      let openStock = 0
      let inMonthReceived = 0
      let inMonthIssued = 0
      let latestReceivedDate = null

      g.rawTransactions.forEach(t => {
        if (!t.date) return
        if (t.date < startOfYear) {
          openStock += (t.receivedVal - t.issuedVal)
        } else if (t.date <= endOfYear) {
          inMonthReceived += t.receivedVal
          inMonthIssued += t.issuedVal
          if (t.receivedVal > 0 && (!latestReceivedDate || t.date > latestReceivedDate)) {
            latestReceivedDate = t.date
          }
        }
      })

      const closingStock = openStock + inMonthReceived - inMonthIssued
      const priceRow = materialPriceRows.find(r => String(r.maSAP || '').trim().toLowerCase() === g.maSAP.toLowerCase())
      const averageUnitPrice = priceRow ? (priceRow.donGiaTrungBinh || 0) : (materialPrices[g.maSAP] || 0)
      const closingValue = Math.round(closingStock * averageUnitPrice)

      // Unused status
      const daysUnused = latestReceivedDate ? getDaysToToday(latestReceivedDate) : (g.rawTransactions.length > 0 ? getDaysToToday(g.rawTransactions[g.rawTransactions.length - 1].date) : 0)
      const unusedStatus = (closingStock > 0 && daysUnused > 30) ? 'Chưa sử dụng (> 30 ngày)' : 'Đang sử dụng'

      // Calculate total stock across physical warehouses
      let totalStock = 0
      uniquePhysicalWarehouses.forEach(pName => {
        const pStock = (projectStocksByMaterial[g.maSAP] && projectStocksByMaterial[g.maSAP][pName]) || 0
        totalStock += pStock
      })

      return {
        ...g,
        openStock: Math.round(openStock * 1000) / 1000,
        inMonthReceived: Math.round(inMonthReceived * 1000) / 1000,
        inMonthIssued: Math.round(inMonthIssued * 1000) / 1000,
        closingStock: Math.round(closingStock * 1000) / 1000,
        averageUnitPrice,
        closingValue,
        unusedStatus,
        daysUnused,
        totalStock: Math.round(totalStock * 1000) / 1000
      }
    })

    return list
  }, [chungRows, giaoRows, nhanRows, localProject, selectedYear, customCategoryMap, materialClassifications, materialPriceRows, materialPrices, uniquePhysicalWarehouses, projectStocksByMaterial])

  const filteredMaterialStatsData = useMemo(() => {
    const text = searchTerm.toLowerCase().trim()
    return materialStatsData.filter(item => {
      // Show item if it has any active stock/activity or if search is active
      const hasActivity = item.openStock !== 0 || item.inMonthReceived !== 0 || item.inMonthIssued !== 0 || item.closingStock !== 0 || item.totalStock !== 0
      if (!hasActivity && !text) return false

      if (!text) return true

      return (item.maSAP || '').toLowerCase().includes(text) ||
        (item.maVatTu || '').toLowerCase().includes(text) ||
        (item.tenVatTu || '').toLowerCase().includes(text) ||
        (item.thongSoKyThuat || '').toLowerCase().includes(text)
    })
  }, [materialStatsData, searchTerm])

  const activePhysicalWarehouses = useMemo(() => {
    return uniquePhysicalWarehouses.filter(pName => {
      return filteredMaterialStatsData.some(item => {
        const val = (projectStocksByMaterial[item.maSAP] && projectStocksByMaterial[item.maSAP][pName]) || 0
        return Math.abs(val) > 0.0001
      })
    })
  }, [uniquePhysicalWarehouses, filteredMaterialStatsData, projectStocksByMaterial])

  const handleMatSort = (field) => {
    if (matSortField === field) {
      setMatSortDirection(prev => prev === 'asc' ? 'desc' : 'asc')
    } else {
      setMatSortField(field)
      setMatSortDirection('desc')
    }
  }

  const sortedMaterialStatsData = useMemo(() => {
    const sorted = [...filteredMaterialStatsData]
    sorted.sort((a, b) => {
      let valA = a[matSortField]
      let valB = b[matSortField]

      if (typeof valA === 'string') {
        return matSortDirection === 'asc' 
          ? valA.localeCompare(valB) 
          : valB.localeCompare(valA)
      } else {
        const nA = valA || 0
        const nB = valB || 0
        return matSortDirection === 'asc' ? nA - nB : nB - nA
      }
    })
    return sorted
  }, [filteredMaterialStatsData, matSortField, matSortDirection])

  // Pagination calculation
  const totalPages = Math.ceil(sortedData.length / pageSize) || 1
  const paginatedData = useMemo(() => {
    const startIndex = (currentPage - 1) * pageSize
    return sortedData.slice(startIndex, startIndex + pageSize)
  }, [sortedData, currentPage, pageSize])

  // Reset page when filters change
  useEffect(() => {
    setCurrentPage(1)
  }, [localProject, statusFilter, searchTerm, selectedYear])

  const handleSort = (field) => {
    if (sortField === field) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc')
    } else {
      setSortField(field)
      setSortDirection('desc')
    }
  }

  const handleSumSort = (field) => {
    if (sumSortField === field) {
      setSumSortDirection(prev => prev === 'asc' ? 'desc' : 'asc')
    } else {
      setSumSortField(field)
      setSumSortDirection('desc')
    }
  }

  // Format date helper
  const formatDate = (dateObj) => {
    const d = parseRowDate(dateObj)
    if (!d) return '-'
    const day = String(d.getDate()).padStart(2, '0')
    const month = String(d.getMonth() + 1).padStart(2, '0')
    const year = d.getFullYear()
    return `${day}/${month}/${year}`
  }

  // Export to Excel function (maintains beautiful formatting with 12 months)
  const handleExportExcel = () => {
    if (activeSubTab === 'summary_all' || activeSubTab === 'reconciliation') {
      if (companyProjectSummaries.length === 0) {
        alert('Không có dữ liệu để xuất.')
        return
      }
      
      const wb = XLSXStyle.utils.book_new()
      const ws = {}
      
      const columns = [
        { key: 'STT', label: 'STT', width: 50 },
        { key: 'projectName', label: 'Kho / Dự án', width: 280 },
        { key: 'totalAssets', label: 'Số chủng loại', width: 120 },
        { key: 'totalStock', label: 'Tổng tồn kho', width: 130 },
        { key: 'T1', label: 'Tháng 1', width: 85 },
        { key: 'T2', label: 'Tháng 2', width: 85 },
        { key: 'T3', label: 'Tháng 3', width: 85 },
        { key: 'T4', label: 'Tháng 4', width: 85 },
        { key: 'T5', label: 'Tháng 5', width: 85 },
        { key: 'T6', label: 'Tháng 6', width: 85 },
        { key: 'T7', label: 'Tháng 7', width: 85 },
        { key: 'T8', label: 'Tháng 8', width: 85 },
        { key: 'T9', label: 'Tháng 9', width: 85 },
        { key: 'T10', label: 'Tháng 10', width: 85 },
        { key: 'T11', label: 'Tháng 11', width: 85 },
        { key: 'T12', label: 'Tháng 12', width: 85 },
        { key: 'totalValue', label: 'Giá trị tồn tạm tính (đ)', width: 160 }
      ]
      
      ws['!cols'] = columns.map(c => ({ wpx: c.width }))
      
      ws['A1'] = {
        v: `BÁO CÁO TỔNG HỢP TÀI SẢN KHẤU HAO TOÀN CÔNG TY (NĂM ${selectedYear})`,
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '1E3A8A' } },
          alignment: { horizontal: 'left', vertical: 'center' }
        }
      }
      ws['A2'] = {
        v: `Tổng hợp toàn công ty từ tháng 1 đến tháng 12 năm ${selectedYear} được chọn.`,
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, italic: true, color: { rgb: '475569' } },
          alignment: { horizontal: 'left', vertical: 'center' }
        }
      }
      
      let excelRowIdx = 4
      
      function getColLabel(index) {
        let label = ''
        let temp = index
        while (temp >= 0) {
          label = String.fromCharCode((temp % 26) + 65) + label
          temp = Math.floor(temp / 26) - 1
        }
        return label
      }

      // Write headers
      columns.forEach((col, colIdx) => {
        const cellRef = `${getColLabel(colIdx)}${excelRowIdx}`
        const isMonthCol = col.key.startsWith('T') && col.key !== 'totalAssets' && col.key !== 'totalStock' && col.key !== 'totalValue' && col.key !== 'totalDepreciation'
        
        ws[cellRef] = {
          v: col.label,
          t: 's',
          s: {
            fill: { fgColor: { rgb: isMonthCol ? '1D4ED8' : '1E3A8A' } },
            font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: '475569' } },
              bottom: { style: 'medium', color: { rgb: '1E293B' } },
              left: { style: 'thin', color: { rgb: '475569' } },
              right: { style: 'thin', color: { rgb: '475569' } }
            }
          }
        }
      })
      
      // Write data
      const groupedDataForExcel = [
        { id: 'I', title: 'I. KHO DỰ ÁN / KHO BCH (A)', items: companyProjectSummaries.filter(item => getCategory(item.projectName) !== 'ncc') },
        { id: 'II', title: 'II. NHẬP TỪ NHÀ CUNG CẤP (B)', items: companyProjectSummaries.filter(item => getCategory(item.projectName) === 'ncc') }
      ]

      ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 16 } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: 16 } }
      ]

      const groupRowIndices = {}

      groupedDataForExcel.forEach(group => {
        // Write group title row
        excelRowIdx++
        ws[`A${excelRowIdx}`] = {
          v: group.title,
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '1E3A8A' } },
            fill: { fgColor: { rgb: 'E2E8F0' } },
            alignment: { horizontal: 'left', vertical: 'center' }
          }
        }
        ws['!merges'].push({ s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: 16 } })

        // Initialize other cells in merged row to have clean backgrounds
        for (let cIdx = 1; cIdx < 17; cIdx++) {
          ws[`${getColLabel(cIdx)}${excelRowIdx}`] = {
            v: '',
            t: 's',
            s: { fill: { fgColor: { rgb: 'E2E8F0' } } }
          }
        }

        const startItemRow = excelRowIdx + 1

        if (group.items.length === 0) {
          excelRowIdx++
          ws[`A${excelRowIdx}`] = {
            v: 'Không có dữ liệu cho nhóm này.',
            t: 's',
            s: {
              font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '94A3B8' } },
              alignment: { horizontal: 'center', vertical: 'center' }
            }
          }
          ws['!merges'].push({ s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: 16 } })
          for (let cIdx = 1; cIdx < 17; cIdx++) {
            ws[`${getColLabel(cIdx)}${excelRowIdx}`] = { v: '', t: 's' }
          }
        } else {
          group.items.forEach((projSum, idx) => {
            excelRowIdx++
            const cells = [
              idx + 1,
              projSum.projectName,
              projSum.totalAssets,
              projSum.totalStock,
              projSum.monthlyStock.T1,
              projSum.monthlyStock.T2,
              projSum.monthlyStock.T3,
              projSum.monthlyStock.T4,
              projSum.monthlyStock.T5,
              projSum.monthlyStock.T6,
              projSum.monthlyStock.T7,
              projSum.monthlyStock.T8,
              projSum.monthlyStock.T9,
              projSum.monthlyStock.T10,
              projSum.monthlyStock.T11,
              projSum.monthlyStock.T12,
              projSum.totalValue
            ]
            
            cells.forEach((val, colIdx) => {
              const cellRef = `${getColLabel(colIdx)}${excelRowIdx}`
              const isNum = typeof val === 'number'
              const isMonthCol = colIdx >= 4 && colIdx <= 15
              
              const baseStyle = {
                font: { name: 'Segoe UI', sz: 9.5 },
                alignment: {
                  horizontal: colIdx === 1 ? 'left' : (isNum ? 'right' : 'center'),
                  vertical: 'center'
                },
                border: {
                  top: { style: 'thin', color: { rgb: 'E2E8F0' } },
                  bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
                  left: { style: 'thin', color: { rgb: 'E2E8F0' } },
                  right: { style: 'thin', color: { rgb: 'E2E8F0' } }
                }
              }
              
              if (isMonthCol) {
                baseStyle.fill = { fgColor: { rgb: 'F0F9FF' } }
                if (val < 0) {
                  baseStyle.font = { name: 'Segoe UI', sz: 9.5, color: { rgb: 'EF4444' } }
                }
              }
              
              ws[cellRef] = { v: val, t: isNum ? 'n' : 's', s: baseStyle }
              if (isNum) {
                ws[cellRef].z = colIdx >= 16 ? '#,##0' : '#,##0.00'
              }
            })
          })
        }

        const endItemRow = excelRowIdx

        // Write Subtotal Row for this Group
        excelRowIdx++
        groupRowIndices[group.id] = excelRowIdx
        ws[`A${excelRowIdx}`] = {
          v: `CỘNG ${group.id === 'I' ? 'KHO DỰ ÁN (A)' : 'NHẬP TỪ NCC (B)'}`,
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '1E3A8A' } },
            alignment: { horizontal: 'left', vertical: 'center' },
            fill: { fgColor: { rgb: 'F1F5F9' } },
            border: {
              top: { style: 'thin', color: { rgb: 'CBD5E1' } },
              bottom: { style: 'medium', color: { rgb: '1E3A8A' } }
            }
          }
        }
        ws['!merges'].push({ s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: 1 } })

        // Initialize empty cell B for merged range
        ws[`B${excelRowIdx}`] = { v: '', t: 's', s: { fill: { fgColor: { rgb: 'F1F5F9' } } } }

        // Sum column cells for group assets, stock, T1..T12, totalValue, totalDepreciation
        for (let c = 2; c < columns.length; c++) {
          const colLabel = getColLabel(c)
          const cellRef = `${colLabel}${excelRowIdx}`
          
          if (group.items.length === 0) {
            ws[cellRef] = { v: 0, t: 'n', z: c >= 16 ? '#,##0' : '#,##0.00', s: { font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '1E3A8A' } }, alignment: { horizontal: 'right', vertical: 'center' }, fill: { fgColor: { rgb: 'F1F5F9' } } } }
          } else {
            ws[cellRef] = {
              f: `SUM(${colLabel}${startItemRow}:${colLabel}${endItemRow})`,
              t: 'n',
              z: c >= 16 ? '#,##0' : '#,##0.00',
              s: {
                font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '1E3A8A' } },
                alignment: { horizontal: 'right', vertical: 'center' },
                fill: { fgColor: { rgb: 'F1F5F9' } },
                border: {
                  top: { style: 'thin', color: { rgb: 'CBD5E1' } },
                  bottom: { style: 'medium', color: { rgb: '1E3A8A' } }
                }
              }
            }
          }
        }
      })
      
      // Write Reconciled Difference Row
      excelRowIdx++
      ws[`A${excelRowIdx}`] = {
        v: 'CHÊNH LỆCH ĐỐI CHIẾU (A - B)',
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '047857' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          fill: { fgColor: { rgb: 'ECFDF5' } },
          border: {
            top: { style: 'medium', color: { rgb: '10B981' } },
            bottom: { style: 'medium', color: { rgb: '10B981' } }
          }
        }
      }
      
      ws['!merges'].push({ s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: 1 } })
      ws[`B${excelRowIdx}`] = { v: '', t: 's', s: { fill: { fgColor: { rgb: 'ECFDF5' } } } }
      
      // assets col index = 2
      ws[`C${excelRowIdx}`] = {
        v: '-',
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '047857' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          fill: { fgColor: { rgb: 'ECFDF5' } },
          border: {
            top: { style: 'medium', color: { rgb: '10B981' } },
            bottom: { style: 'medium', color: { rgb: '10B981' } }
          }
        }
      }

      // Formula differences: (Row Group I - Row Group II)
      for (let c = 3; c < columns.length; c++) {
        const colLabel = getColLabel(c)
        const cellRef = `${colLabel}${excelRowIdx}`
        const rowI = groupRowIndices['I']
        const rowII = groupRowIndices['II']
        
        ws[cellRef] = {
          f: `${colLabel}${rowI}-${colLabel}${rowII}`,
          t: 'n',
          z: c >= 16 ? '#,##0' : '#,##0.00',
          s: {
            font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '047857' } },
            alignment: { horizontal: 'right', vertical: 'center' },
            fill: { fgColor: { rgb: 'ECFDF5' } },
            border: {
              top: { style: 'medium', color: { rgb: '10B981' } },
              bottom: { style: 'medium', color: { rgb: '10B981' } }
            }
          }
        }
      }
      
      ws['!ref'] = `A1:Q${excelRowIdx}`
      XLSXStyle.utils.book_append_sheet(wb, ws, 'Tổng hợp toàn công ty')

      // Append Detailed Mismatch Worksheet if there are any discrepancies
      if (mismatchAnalysis && mismatchAnalysis.length > 0) {
        const wsMismatch = {}
        const mcols = [
          { key: 'stt', label: 'STT', width: 8 },
          { key: 'maDon', label: 'Mã đơn', width: 15 },
          { key: 'date', label: 'Ngày', width: 14 },
          { key: 'maSAP', label: 'Mã SAP', width: 14 },
          { key: 'tenVatTu', label: 'Tên vật tư (Tài sản)', width: 35 },
          { key: 'giaoUnit', label: 'Đơn vị giao (A)', width: 22 },
          { key: 'nhanUnit', label: 'Đơn vị nhận (B)', width: 22 },
          { key: 'qty', label: 'Số lượng', width: 15 },
          { key: 'effectI', label: 'Tác động Kho Dự án (A)', width: 20 },
          { key: 'effectII', label: 'Tác động NCC (B)', width: 20 },
          { key: 'diff', label: 'Chênh lệch (A - B)', width: 16 },
          { key: 'reason', label: 'Nguyên nhân chênh lệch chi tiết', width: 45 }
        ]

        // Write title of mismatch sheet
        wsMismatch['A1'] = {
          v: `CHI TIẾT CÁC GIAO DỊCH LỆCH DÒNG ĐỐI CHIẾU (${selectedYear})`,
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '1E3A8A' } },
            alignment: { horizontal: 'center', vertical: 'center' }
          }
        }
        wsMismatch['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 11 } }]

        // Write headers
        mcols.forEach((col, cIdx) => {
          const colLabel = getColLabel(cIdx)
          const cellRef = `${colLabel}3`
          wsMismatch[cellRef] = {
            v: col.label,
            t: 's',
            s: {
              font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
              fill: { fgColor: { rgb: '1E3A8A' } },
              alignment: { horizontal: 'center', vertical: 'center' },
              border: {
                top: { style: 'thin', color: { rgb: 'CBD5E1' } },
                bottom: { style: 'medium', color: { rgb: '0F172A' } }
              }
            }
          }
        })

        let mRowIdx = 4
        mismatchAnalysis.forEach((m, mIdx) => {
          wsMismatch[`A${mRowIdx}`] = { v: mIdx + 1, t: 'n', s: { font: { name: 'Segoe UI', sz: 9 }, alignment: { horizontal: 'center' } } }
          wsMismatch[`B${mRowIdx}`] = { v: m.maDon, t: 's', s: { font: { name: 'Segoe UI', sz: 9 }, alignment: { horizontal: 'center' } } }
          wsMismatch[`C${mRowIdx}`] = { v: m.date, t: 's', s: { font: { name: 'Segoe UI', sz: 9 }, alignment: { horizontal: 'center' } } }
          wsMismatch[`D${mRowIdx}`] = { v: m.maSAP, t: 's', s: { font: { name: 'Segoe UI', sz: 9 }, alignment: { horizontal: 'center' } } }
          wsMismatch[`E${mRowIdx}`] = { v: m.tenVatTu, t: 's', s: { font: { name: 'Segoe UI', sz: 9 } } }
          wsMismatch[`F${mRowIdx}`] = { v: m.giaoUnit, t: 's', s: { font: { name: 'Segoe UI', sz: 9 } } }
          wsMismatch[`G${mRowIdx}`] = { v: m.nhanUnit, t: 's', s: { font: { name: 'Segoe UI', sz: 9 } } }
          wsMismatch[`H${mRowIdx}`] = { v: m.qty, t: 'n', z: '#,##0.00', s: { font: { name: 'Segoe UI', sz: 9 }, alignment: { horizontal: 'right' } } }
          wsMismatch[`I${mRowIdx}`] = { v: m.effectI, t: 'n', z: '#,##0.00', s: { font: { name: 'Segoe UI', sz: 9 }, alignment: { horizontal: 'right' } } }
          wsMismatch[`J${mRowIdx}`] = { v: m.effectII, t: 'n', z: '#,##0.00', s: { font: { name: 'Segoe UI', sz: 9 }, alignment: { horizontal: 'right' } } }
          wsMismatch[`K${mRowIdx}`] = { v: m.diff, t: 'n', z: '#,##0.00', s: { font: { name: 'Segoe UI', sz: 9, bold: true, color: { rgb: m.diff < 0 ? 'EF4444' : '047857' } }, alignment: { horizontal: 'right' } } }
          wsMismatch[`L${mRowIdx}`] = { v: m.reason, t: 's', s: { font: { name: 'Segoe UI', sz: 9, color: { rgb: 'B45309' } } } }
          mRowIdx++
        })

        wsMismatch['!ref'] = `A1:L${mRowIdx - 1}`
        wsMismatch['!cols'] = mcols.map(c => ({ wch: c.width }))

        XLSXStyle.utils.book_append_sheet(wb, wsMismatch, 'Chi tiết lệch dòng đối chiếu')
      }

      XLSXStyle.writeFile(wb, `SGC_TongHop_TaiSan_KhauHao_ToanCongTy_${selectedYear}.xlsx`)
      return
    }

    if (activeSubTab === 'summary_month') {
      if (companyProjectMonthlySummaries.length === 0) {
        alert('Không có dữ liệu để xuất.')
        return
      }

      const wb = XLSXStyle.utils.book_new()
      const ws = {}

      const columns = [
        { key: 'STT', width: 50 },
        { key: 'projectName', width: 280 },
        { key: 'maSAP', width: 110 },
        { key: 'tenVatTu', width: 280 },
        { key: 'dvt', width: 80 },
        { key: 'totalStock', width: 130 },
        { key: 'unitPrice', width: 140 },
      ]

      for (let m = 1; m <= 12; m++) {
        columns.push({ key: `T${m}`, width: 65 })
      }

      for (let m = 1; m <= 12; m++) {
        columns.push({ key: `Amt_T${m}`, width: 110 })
      }

      const maxColIdx = columns.length - 1

      ws['!cols'] = columns.map(c => ({ wpx: c.width }))

      ws['A1'] = {
        v: `BÁO CÁO TỔNG HỢP TÀI SẢN KHẤU HAO THEO THÁNG (NĂM ${selectedYear})`,
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '1E3A8A' } },
          alignment: { horizontal: 'left', vertical: 'center' }
        }
      }
      ws['A2'] = {
        v: `Tổng hợp chi tiết số dư tồn kho hàng tháng trong năm ${selectedYear}.`,
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, italic: true, color: { rgb: '475569' } },
          alignment: { horizontal: 'left', vertical: 'center' }
        }
      }
      ws['A3'] = {
        v: `Ngày xuất báo cáo: ${new Date().toLocaleDateString('vi-VN')}`,
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '0F172A' } },
          alignment: { horizontal: 'left', vertical: 'center' }
        }
      }

      ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: maxColIdx } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: maxColIdx } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: maxColIdx } },
        // Header row merges
        { s: { r: 4, c: 0 }, e: { r: 5, c: 0 } }, // STT
        { s: { r: 4, c: 1 }, e: { r: 5, c: 1 } }, // ProjectName
        { s: { r: 4, c: 2 }, e: { r: 5, c: 2 } }, // SAP
        { s: { r: 4, c: 3 }, e: { r: 5, c: 3 } }, // Tên vật tư
        { s: { r: 4, c: 4 }, e: { r: 5, c: 4 } }, // Đvt
        { s: { r: 4, c: 5 }, e: { r: 5, c: 5 } }, // Tổng tồn kho
        { s: { r: 4, c: 6 }, e: { r: 5, c: 6 } }, // Đơn giá tháng
        { s: { r: 4, c: 7 }, e: { r: 4, c: 18 } }, // KHỐI LƯỢNG
        { s: { r: 4, c: 19 }, e: { r: 4, c: 30 } }, // THÀNH TIỀN
      ]

      let excelRowIdx = 5

      function getColLabel(index) {
        let label = ''
        let temp = index
        while (temp >= 0) {
          label = String.fromCharCode((temp % 26) + 65) + label
          temp = Math.floor(temp / 26) - 1
        }
        return label
      }

      // Helper to generate cell with styles
      const createHeaderCell = (val, bg) => ({
        v: val,
        t: 's',
        s: {
          fill: { fgColor: { rgb: bg } },
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: '475569' } },
            bottom: { style: 'thin', color: { rgb: '475569' } },
            left: { style: 'thin', color: { rgb: '475569' } },
            right: { style: 'thin', color: { rgb: '475569' } }
          }
        }
      })

      // Write headers Row 5 (excelRowIdx = 5, which is row index 4 in 0-based index)
      ws['A5'] = createHeaderCell('STT', '1E3A8A')
      ws['B5'] = createHeaderCell('Kho / Dự án', '1E3A8A')
      ws['C5'] = createHeaderCell('Mã SAP', '1E3A8A')
      ws['D5'] = createHeaderCell('Tên vật tư', '1E3A8A')
      ws['E5'] = createHeaderCell('Đvt', '1E3A8A')
      ws['F5'] = createHeaderCell('Tổng tồn kho', '1E3A8A')
      ws['G5'] = createHeaderCell('KHỐI LƯỢNG TỒN KHO THEO THÁNG', '1D4ED8')
      ws['S5'] = createHeaderCell('Đơn giá tháng', '1E3A8A')
      ws['T5'] = createHeaderCell('THÀNH TIỀN TỒN KHO THEO THÁNG', '047857')

      // Fill empty cells for merged ranges in row 5 so borders look nice
      for (let c = 7; c <= 17; c++) ws[`${getColLabel(c)}5`] = createHeaderCell('', '1D4ED8')
      for (let c = 20; c <= 30; c++) ws[`${getColLabel(c)}5`] = createHeaderCell('', '047857')

      // Write headers Row 6 (excelRowIdx = 6, which is row index 5 in 0-based index)
      ws['A6'] = createHeaderCell('', '1E3A8A')
      ws['B6'] = createHeaderCell('', '1E3A8A')
      ws['C6'] = createHeaderCell('', '1E3A8A')
      ws['D6'] = createHeaderCell('', '1E3A8A')
      ws['E6'] = createHeaderCell('', '1E3A8A')
      ws['F6'] = createHeaderCell('', '1E3A8A')
      for (let m = 1; m <= 12; m++) {
        ws[`${getColLabel(5 + m)}6`] = createHeaderCell(`T${m}`, '1D4ED8')
      }
      ws['S6'] = createHeaderCell('', '1E3A8A')
      for (let m = 1; m <= 12; m++) {
        ws[`${getColLabel(18 + m)}6`] = createHeaderCell(`T${m}`, '047857')
      }

      excelRowIdx = 6 // set to 6 since we have written rows 5 and 6

      // Write groups
      groupedProjectMonthlySummaries.forEach((group) => {
        excelRowIdx++
        const groupLabelCell = `A${excelRowIdx}`
        ws[groupLabelCell] = {
          v: group.title,
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 11, bold: true, color: { rgb: '1E3A8A' } },
            fill: { fgColor: { rgb: 'F1F5F9' } },
            alignment: { horizontal: 'left', vertical: 'center' },
            border: {
              top: { style: 'thin', color: { rgb: 'CBD5E1' } },
              bottom: { style: 'thin', color: { rgb: 'CBD5E1' } }
            }
          }
        }
        ws['!merges'].push({ s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: maxColIdx } })

        for (let c = 1; c <= maxColIdx; c++) {
          ws[`${getColLabel(c)}${excelRowIdx}`] = {
            v: '',
            t: 's',
            s: {
              fill: { fgColor: { rgb: 'F1F5F9' } },
              border: {
                top: { style: 'thin', color: { rgb: 'CBD5E1' } },
                bottom: { style: 'thin', color: { rgb: 'CBD5E1' } }
              }
            }
          }
        }

        group.items.forEach((item, itemIdx) => {
          excelRowIdx++
          const rowValues = [
            itemIdx + 1,
            item.projectName,
            item.maSAP,
            item.tenVatTu,
            item.dvt,
            item.totalStock,
            ...[...Array(12)].map((_, i) => item.monthlyStock[`T${i + 1}`] || 0),
            item.unitPrice || 0,
            ...[...Array(12)].map((_, i) => (item.monthlyStock[`T${i + 1}`] || 0) * (item.unitPrice || 0))
          ]

          rowValues.forEach((val, colIdx) => {
            const cellRef = `${getColLabel(colIdx)}${excelRowIdx}`
            const isNum = typeof val === 'number'
            const isQtyMonthCol = colIdx >= 6 && colIdx < 18
            const isAmtMonthCol = colIdx >= 19 && colIdx < 31

            const baseStyle = {
              font: { name: 'Segoe UI', sz: 9.5 },
              alignment: {
                horizontal: (colIdx === 1 || colIdx === 3) ? 'left' : (isNum ? 'right' : 'center'),
                vertical: 'center'
              },
              border: {
                top: { style: 'thin', color: { rgb: 'E2E8F0' } },
                bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
                left: { style: 'thin', color: { rgb: 'E2E8F0' } },
                right: { style: 'thin', color: { rgb: 'E2E8F0' } }
              }
            }

            if (isQtyMonthCol) {
              baseStyle.fill = { fgColor: { rgb: 'F0F9FF' } }
            } else if (isAmtMonthCol) {
              baseStyle.fill = { fgColor: { rgb: 'ECFDF5' } }
            }

            const cellObj = { v: val === 0 && colIdx >= 5 ? '' : val, t: isNum ? 'n' : 's', s: baseStyle }
            if (isNum) {
              const useNoDecimals = colIdx === 18 || isAmtMonthCol
              cellObj.z = useNoDecimals ? '#,##0;(#,##0);"-"' : '#,##0.00;(#,##0.00);"-"'
            }
            ws[cellRef] = cellObj
          })
        })

        // Group total row
        excelRowIdx++
        const groupStock = group.items.reduce((sum, item) => sum + item.totalStock, 0)
        const groupMonthly = {}
        const groupMonthlyAmount = {}
        for (let m = 1; m <= 12; m++) {
          groupMonthly[`T${m}`] = group.items.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0), 0)
          groupMonthlyAmount[`T${m}`] = group.items.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0) * (item.unitPrice || 0), 0)
        }

        const totalRowValues = [
          `CỘNG ${group.id === 'I' ? 'KHO DỰ ÁN (A)' : 'NHẬP TỪ NCC (B)'}`,
          '',
          '',
          '',
          '',
          groupStock,
          ...[...Array(12)].map((_, i) => groupMonthly[`T${i + 1}`] || 0),
          '',
          ...[...Array(12)].map((_, i) => groupMonthlyAmount[`T${i + 1}`] || 0)
        ]

        totalRowValues.forEach((val, colIdx) => {
          const actualColIdx = colIdx
          if (colIdx >= 1 && colIdx <= 4) return // merged

          const cellRef = `${getColLabel(actualColIdx)}${excelRowIdx}`
          const isNum = typeof val === 'number'
          const isQtyMonthCol = actualColIdx >= 6 && actualColIdx < 18
          const isAmtMonthCol = actualColIdx >= 19 && actualColIdx < 31

          const baseStyle = {
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
            alignment: {
              horizontal: actualColIdx === 0 ? 'left' : (isNum ? 'right' : 'center'),
              vertical: 'center'
            },
            fill: { fgColor: { rgb: 'F8FAFC' } },
            border: {
              top: { style: 'thin', color: { rgb: 'CBD5E1' } },
              bottom: { style: 'medium', color: { rgb: '1E293B' } },
              left: { style: 'thin', color: { rgb: 'CBD5E1' } },
              right: { style: 'thin', color: { rgb: 'CBD5E1' } }
            }
          }

          if (isQtyMonthCol) {
            baseStyle.fill = { fgColor: { rgb: 'E0F2FE' } }
          } else if (isAmtMonthCol) {
            baseStyle.fill = { fgColor: { rgb: 'D1FAE5' } }
            baseStyle.font.color = { rgb: '047857' }
          }

          const cellObj = { v: val === 0 && actualColIdx >= 5 ? '' : val, t: isNum ? 'n' : 's', s: baseStyle }
          if (isNum) {
            const useNoDecimals = actualColIdx === 18 || isAmtMonthCol
            cellObj.z = useNoDecimals ? '#,##0;(#,##0);"-"' : '#,##0.00;(#,##0.00);"-"'
          }
          ws[cellRef] = cellObj
        })
        ws['!merges'].push({ s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: 4 } })
      })

      // Grand difference row
      excelRowIdx++
      const diffRowValues = [
        '✓ CHÊNH LỆCH ĐỐI CHIẾU (A - B)',
        '',
        '',
        '',
        '',
        reconciliationMonthlyDifference.stockDiff,
        ...[...Array(12)].map((_, i) => reconciliationMonthlyDifference.monthlyDiff[`T${i + 1}`] || 0),
        '',
        ...[...Array(12)].map((_, i) => reconciliationMonthlyDifference.monthlyAmtDiff[`T${i + 1}`] || 0)
      ]

      diffRowValues.forEach((val, colIdx) => {
        if (colIdx >= 1 && colIdx <= 4) return

        const actualColIdx = colIdx
        const cellRef = `${getColLabel(actualColIdx)}${excelRowIdx}`
        const isNum = typeof val === 'number'
        const isQtyCol = actualColIdx >= 5 && actualColIdx < 18
        const isAmtCol = actualColIdx >= 19 && actualColIdx < 31

        const baseStyle = {
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '047857' } },
          alignment: {
            horizontal: actualColIdx === 0 ? 'center' : 'right',
            vertical: 'center'
          },
          fill: { fgColor: { rgb: 'ECFDF5' } },
          border: {
            top: { style: 'medium', color: { rgb: '10B981' } },
            bottom: { style: 'medium', color: { rgb: '10B981' } },
            left: { style: 'thin', color: { rgb: 'CBD5E1' } },
            right: { style: 'thin', color: { rgb: 'CBD5E1' } }
          }
        }

        const cellObj = { v: val === 0 && actualColIdx >= 5 ? '' : val, t: isNum ? 'n' : 's', s: baseStyle }
        if (isNum) {
          const useNoDecimals = isAmtCol
          cellObj.z = useNoDecimals ? '#,##0;(#,##0);"-"' : '#,##0.00;(#,##0.00);"-"'
        }
        ws[cellRef] = cellObj
      })
      ws['!merges'].push({ s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: 4 } })

      ws['!ref'] = `A1:${getColLabel(maxColIdx)}${excelRowIdx}`
      XLSXStyle.utils.book_append_sheet(wb, ws, 'Tổng hợp tháng')

      XLSXStyle.writeFile(wb, `SGC_TongHop_TaiSan_KhauHao_Nam_${selectedYear}.xlsx`)
      return
    }

    if (activeSubTab === 'material_stats') {
      if (sortedMaterialStatsData.length === 0) {
        alert('Không có dữ liệu để xuất.')
        return
      }

      const wb = XLSXStyle.utils.book_new()
      const ws = {}

      const columns = [
        { key: 'STT', label: 'STT', width: 60 },
        { key: 'maSAP', label: 'Mã SAP', width: 120 },
        { key: 'tenVatTu', label: 'Tên vật tư', width: 300 },
        { key: 'dvt', label: 'ĐVT', width: 80 },
        { key: 'totalStock', label: 'Tổng cộng', width: 110 },
        ...activePhysicalWarehouses.map(pName => ({
          key: pName,
          label: pName,
          width: 85
        }))
      ]

      ws['!cols'] = columns.map(c => ({ wpx: c.width }))

      const maxColIdx = columns.length - 1

      ws['A1'] = {
        v: `BÁO CÁO THỐNG KÊ TÀI SẢN KHẤU HAO THEO VẬT TƯ (NĂM ${selectedYear})`,
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '1E3A8A' } },
          alignment: { horizontal: 'left', vertical: 'center' }
        }
      }
      ws['A2'] = {
        v: `Kho / Dự án: Tất cả Kho / Dự án | Trạng thái: Tất cả tài sản | Năm: ${selectedYear}`,
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, italic: true, color: { rgb: '475569' } },
          alignment: { horizontal: 'left', vertical: 'center' }
        }
      }
      ws['A3'] = {
        v: `Ngày xuất báo cáo: ${new Date().toLocaleDateString('vi-VN')} | Tổng số loại vật tư: ${sortedMaterialStatsData.length}`,
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '0F172A' } },
          alignment: { horizontal: 'left', vertical: 'center' }
        }
      }

      ws['!merges'] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: maxColIdx } },
        { s: { r: 1, c: 0 }, e: { r: 1, c: maxColIdx } },
        { s: { r: 2, c: 0 }, e: { r: 2, c: maxColIdx } }
      ]

      let excelRowIdx = 5

      function getColLabel(index) {
        let label = ''
        let temp = index
        while (temp >= 0) {
          label = String.fromCharCode((temp % 26) + 65) + label
          temp = Math.floor(temp / 26) - 1
        }
        return label
      }

      // Write Headers
      columns.forEach((col, colIdx) => {
        const cellRef = `${getColLabel(colIdx)}${excelRowIdx}`
        const isStockCol = colIdx >= 4
        
        ws[cellRef] = {
          v: col.label,
          t: 's',
          s: {
            fill: { fgColor: { rgb: isStockCol ? '1D4ED8' : '1E3A8A' } },
            font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: '475569' } },
              bottom: { style: 'medium', color: { rgb: '1E293B' } },
              left: { style: 'thin', color: { rgb: '475569' } },
              right: { style: 'thin', color: { rgb: '475569' } }
            }
          }
        }
      })

      // Write Data Rows
      sortedMaterialStatsData.forEach((item, idx) => {
        excelRowIdx++
        const cells = [
          idx + 1,
          item.maSAP,
          item.tenVatTu,
          item.dvt,
          item.totalStock !== 0 ? item.totalStock : '',
          ...activePhysicalWarehouses.map(pName => {
            const val = (projectStocksByMaterial[item.maSAP] && projectStocksByMaterial[item.maSAP][pName]) || 0
            return val !== 0 ? val : ''
          })
        ]

        cells.forEach((val, colIdx) => {
          const cellRef = `${getColLabel(colIdx)}${excelRowIdx}`
          const isNum = typeof val === 'number'
          const isStockCol = colIdx >= 4

          const baseStyle = {
            font: { name: 'Segoe UI', sz: 9.5 },
            alignment: {
              horizontal: colIdx === 2 ? 'left' : (isNum ? 'right' : 'center'),
              vertical: 'center',
              wrapText: true
            },
            border: {
              top: { style: 'thin', color: { rgb: 'E2E8F0' } },
              bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
              left: { style: 'thin', color: { rgb: 'E2E8F0' } },
              right: { style: 'thin', color: { rgb: 'E2E8F0' } }
            }
          }

          if (isStockCol) {
            baseStyle.fill = { fgColor: { rgb: 'F0F9FF' } }
          }

          const cellObj = { v: val ?? '', t: isNum ? 'n' : 's', s: baseStyle }
          if (isNum) {
            cellObj.z = '#,##0.00;(#,##0.00);"-"'
          }
          ws[cellRef] = cellObj
        })
      })

      // Write Summary Row
      excelRowIdx++
      const labelCellRef = `A${excelRowIdx}`
      ws[labelCellRef] = {
        v: 'TỔNG CỘNG',
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '0F172A' } },
          alignment: { horizontal: 'center', vertical: 'center' },
          fill: { fgColor: { rgb: 'E2E8F0' } },
          border: {
            top: { style: 'medium', color: { rgb: '1E293B' } },
            bottom: { style: 'medium', color: { rgb: '1E293B' } },
            left: { style: 'thin', color: { rgb: 'CBD5E1' } },
            right: { style: 'thin', color: { rgb: 'CBD5E1' } }
          }
        }
      }

      ws['!merges'].push({ s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: 3 } })

      // Fill empty cells of merged area in total row
      for (let c = 1; c <= 3; c++) {
        const colLabel = getColLabel(c)
        const cellRef = `${colLabel}${excelRowIdx}`
        ws[cellRef] = {
          v: '',
          t: 's',
          s: {
            fill: { fgColor: { rgb: 'E2E8F0' } },
            border: {
              top: { style: 'medium', color: { rgb: '1E293B' } },
              bottom: { style: 'medium', color: { rgb: '1E293B' } },
              left: { style: 'thin', color: { rgb: 'CBD5E1' } },
              right: { style: 'thin', color: { rgb: 'CBD5E1' } }
            }
          }
        }
      }

      // Sum formulas for stock columns
      for (let c = 4; c < columns.length; c++) {
        const colLabel = getColLabel(c)
        const cellRef = `${colLabel}${excelRowIdx}`
        const baseStyle = {
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '0F172A' } },
          alignment: { horizontal: 'right', vertical: 'center' },
          fill: { fgColor: { rgb: 'E2E8F0' } },
          border: {
            top: { style: 'medium', color: { rgb: '1E293B' } },
            bottom: { style: 'medium', color: { rgb: '1E293B' } },
            left: { style: 'thin', color: { rgb: 'CBD5E1' } },
            right: { style: 'thin', color: { rgb: 'CBD5E1' } }
          }
        }

        ws[cellRef] = {
          f: `SUM(${colLabel}6:${colLabel}${excelRowIdx - 1})`,
          t: 'n',
          z: '#,##0.00;(#,##0.00);"-"',
          s: baseStyle
        }
      }

      ws['!ref'] = `A1:${getColLabel(columns.length - 1)}${excelRowIdx}`
      XLSXStyle.utils.book_append_sheet(wb, ws, 'Thống kê Vật tư')

      const fileProjName = 'Tat_ca'
      XLSXStyle.writeFile(wb, `SGC_ThongKe_VatTu_Nam_${selectedYear}_${fileProjName}.xlsx`)
      return
    }

    if (sortedData.length === 0) {
      alert('Không có dữ liệu để xuất.')
      return
    }

    const toExcelSerial = (dateVal) => {
      const d = (dateVal instanceof Date) ? dateVal : parseRowDate(dateVal)
      if (!d || isNaN(d.getTime())) return null
      return Math.round(d.getTime() / 86400000) + 25569
    }

    const buildDetailSheet = (list, title, qtyLabel) => {
      const dws = {}
      const dcols = [
        { key: 'maDon', label: 'Mã đơn', width: 160 },
        { key: 'ngayXuatNhap', label: 'Ngày xuất nhập', width: 100 },
        { key: 'maSAP', label: 'Mã SAP', width: 100 },
        { key: 'maVatTu', label: 'Mã vật tư', width: 100 },
        { key: 'tenVatTu', label: 'Tên vật tư', width: 220 },
        { key: 'dvt', label: 'ĐVT', width: 70 },
        { key: 'thongSoKyThuat', label: 'Thông số kỹ thuật', width: 160 },
        { key: 'donViGiao', label: 'Đơn vị giao', width: 150 },
        { key: 'donViNhan', label: 'Đơn vị nhận', width: 150 },
        { key: 'khoiLuong', label: qtyLabel, width: 130 },
        { key: 'nguoiGiao', label: 'Người giao', width: 120 },
        { key: 'nguoiNhan', label: 'Người nhận', width: 120 },
        { key: 'duAn', label: 'Dự án', width: 160 },
        { key: 'trangThai', label: 'Trạng thái', width: 120 },
        { key: 'ghiChu', label: 'Ghi chú', width: 200 }
      ]
      dws['!cols'] = dcols.map(c => ({ wpx: c.width }))

      dws['A1'] = {
        v: title,
        t: 's',
        s: { font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '1E3A8A' } }, alignment: { horizontal: 'left', vertical: 'center' } }
      }
      dws['A2'] = {
        v: `Kho / Dự án: ${localProject || 'Tất cả'}  |  Tổng số dòng: ${list.length.toLocaleString('vi-VN')}`,
        t: 's',
        s: { font: { name: 'Segoe UI', sz: 10, italic: true }, alignment: { horizontal: 'left', vertical: 'center' } }
      }

      let dRowIdx = 4
      dcols.forEach((col, colIdx) => {
        const cellRef = `${String.fromCharCode(65 + colIdx)}${dRowIdx}`
        dws[cellRef] = {
          v: col.label,
          t: 's',
          s: {
            fill: { fgColor: { rgb: '1E3A8A' } },
            font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
            alignment: { horizontal: 'center', vertical: 'center' },
            border: {
              top: { style: 'thin', color: { rgb: '1E3A8A' } },
              bottom: { style: 'medium', color: { rgb: '1E3A8A' } },
              left: { style: 'thin', color: { rgb: '1E3A8A' } },
              right: { style: 'thin', color: { rgb: '1E3A8A' } }
            }
          }
        }
      })

      let totalQty = 0
      list.forEach(item => {
        dRowIdx++
        totalQty += item.khoiLuong || 0
        dcols.forEach((col, colIdx) => {
          const val = item[col.key]
          const cellRef = `${String.fromCharCode(65 + colIdx)}${dRowIdx}`

          if (col.key === 'ngayXuatNhap') {
            const serial = toExcelSerial(val)
            dws[cellRef] = {
              v: serial !== null ? serial : (val || ''),
              t: serial !== null ? 'n' : 's',
              s: {
                font: { name: 'Segoe UI', sz: 9.5 },
                alignment: { horizontal: 'center', vertical: 'center' },
                border: {
                  top: { style: 'thin', color: { rgb: 'E2E8F0' } },
                  bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
                  left: { style: 'thin', color: { rgb: 'E2E8F0' } },
                  right: { style: 'thin', color: { rgb: 'E2E8F0' } }
                }
              }
            }
            if (serial !== null) dws[cellRef].z = 'dd/mm/yyyy'
            return
          }

          const isNum = typeof val === 'number'
          dws[cellRef] = {
            v: val ?? '',
            t: isNum ? 'n' : 's',
            s: {
              font: { name: 'Segoe UI', sz: 9.5 },
              alignment: { 
                horizontal: colIdx === 4 || colIdx === 6 ? 'left' : (isNum ? 'right' : 'center'), 
                vertical: 'center' 
              },
              border: {
                top: { style: 'thin', color: { rgb: 'E2E8F0' } },
                bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
                left: { style: 'thin', color: { rgb: 'E2E8F0' } },
                right: { style: 'thin', color: { rgb: 'E2E8F0' } }
              }
            }
          }
          if (isNum) dws[cellRef].z = '#,##0.00'
        })
      })

      // Total row
      dRowIdx++
      dws[`A${dRowIdx}`] = {
        v: 'TỔNG CỘNG',
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, bold: true },
          alignment: { horizontal: 'center', vertical: 'center' },
          fill: { fgColor: { rgb: 'F1F5F9' } },
          border: {
            top: { style: 'medium', color: { rgb: '1E3A8A' } },
            bottom: { style: 'medium', color: { rgb: '1E3A8A' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          }
        }
      }
      dws['!merges'] = [{ s: { r: dRowIdx - 1, c: 0 }, e: { r: dRowIdx - 1, c: 8 } }]
      for (let c = 1; c <= 8; c++) {
        const cellRef = `${String.fromCharCode(65 + c)}${dRowIdx}`
        dws[cellRef] = {
          v: '', t: 's',
          s: {
            fill: { fgColor: { rgb: 'F1F5F9' } },
            border: {
              top: { style: 'medium', color: { rgb: '1E3A8A' } },
              bottom: { style: 'medium', color: { rgb: '1E3A8A' } },
              left: { style: 'thin', color: { rgb: 'E2E8F0' } },
              right: { style: 'thin', color: { rgb: 'E2E8F0' } }
            }
          }
        }
      }
      const totalCellRef = `${String.fromCharCode(65 + 9)}${dRowIdx}`
      dws[totalCellRef] = {
        v: totalQty, t: 'n', z: '#,##0.00',
        s: {
          font: { name: 'Segoe UI', sz: 10, bold: true },
          alignment: { horizontal: 'right', vertical: 'center' },
          fill: { fgColor: { rgb: 'F1F5F9' } },
          border: {
            top: { style: 'medium', color: { rgb: '1E3A8A' } },
            bottom: { style: 'medium', color: { rgb: '1E3A8A' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          }
        }
      }

      dws['!ref'] = `A1:O${dRowIdx}`
      return dws
    }

    const processRow = (r) => {
      const nhanUnit = String(r.donViNhan || '').trim()
      const giaoUnit = String(r.donViGiao || '').trim()

      const catNhan = getCategory(nhanUnit)
      const catGiao = getCategory(giaoUnit)

      // Rule: "Đơn liên quan đến tổ đội để giá trị = 0"
      if (catNhan === 'todoi' || catGiao === 'todoi') {
        return { receivedVal: 0, issuedVal: 0, isRelated: true }
      }

      let receivedVal = 0
      let issuedVal = 0

      if (localProject) {
        const nhanLower = nhanUnit.toLowerCase()
        const giaoLower = giaoUnit.toLowerCase()
        const isNhan = nhanLower === projLower
        const isGiao = giaoLower === projLower

        if (!isNhan && !isGiao) {
          return { receivedVal: 0, issuedVal: 0, isRelated: false } // Not related to localProject
        }

        if (isNhan) {
          const val = parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
          if (catGiao === 'ncc' || catGiao === 'kho') {
            receivedVal = val
          }
        }

        if (isGiao) {
          const val = parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
          if (catNhan === 'ncc' || catNhan === 'kho') {
            issuedVal = val
          }
        }
      } else {
        const importVal = parseVal(r.khoiLuongNhap)
        if (importVal > 0) {
          if (catGiao === 'ncc' || catGiao === 'kho') {
            receivedVal = importVal
          }
        }

        const exportVal = parseVal(r.khoiLuongXuat)
        if (exportVal > 0) {
          if (catNhan === 'ncc' || catNhan === 'kho') {
            issuedVal = exportVal
          }
        }
      }

      return { receivedVal, issuedVal, isRelated: true }
    }

    const nhanList = []
    const xuatList = []
    const projLower = localProject ? localProject.trim().toLowerCase() : ''

    const sourceRows = (chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows]

    sourceRows.forEach(r => {
      const sap = String(r.maSAP || '').trim()
      if (!sap) return

      // Determine if it is an asset (classification has 'khấu hao' or 'tài sản')
      const classification = String(customCategoryMap[sap] || materialClassifications[sap] || '').trim()
      const cLower = classification.toLowerCase()
      const isKhauHao = classification && (cLower.includes('khấu hao') || cLower.includes('tài sản') || cLower.includes('tskh'))
      if (!isKhauHao) return

      const rowRes = processRow(r)
      if (!rowRes.isRelated) return

      const { receivedVal, issuedVal } = rowRes

      const baseInfo = {
        ngayXuatNhap: r.ngayXuatNhap || '',
        maSAP: sap,
        maVatTu: String(r.maVatTu || '').trim(),
        tenVatTu: String(r.tenVatTu || '').trim(),
        dvt: String(r.dvt || '').trim(),
        thongSoKyThuat: String(r.thongSoKyThuat || '').trim(),
        donViGiao: String(r.donViGiao || '').trim(),
        donViNhan: String(r.donViNhan || '').trim(),
        nguoiGiao: String(r.nguoiGiao || '').trim(),
        nguoiNhan: String(r.nguoiNhan || '').trim(),
        trangThai: String(r.trangThai || '').trim(),
        duAn: String(r.duAn || '').trim(),
        ghiChu: String(r.ghiChu || '').trim()
      }

      if (receivedVal > 0) {
        nhanList.push({ 
          ...baseInfo, 
          maDon: String(r.maDonNhapKho || r.maDonXuatKho || '').trim(), 
          khoiLuong: receivedVal 
        })
      }
      if (issuedVal > 0) {
        xuatList.push({ 
          ...baseInfo, 
          maDon: String(r.maDonXuatKho || r.maDonNhapKho || '').trim(), 
          khoiLuong: issuedVal 
        })
      }
    })

    const wb = XLSXStyle.utils.book_new()
    const ws = {}

    const columns = [
      { key: 'STT', label: 'STT', width: 50 },
      { key: 'maSAP', label: 'Mã SAP', width: 100 },
      { key: 'maVatTu', label: 'Mã vật tư', width: 100 },
      { key: 'tenVatTu', label: 'Tên vật tư', width: 250 },
      { key: 'dvt', label: 'ĐVT', width: 60 },
      { key: 'thongSoKyThuat', label: 'Thông số kỹ thuật', width: 150 },
      { key: 'stock', label: 'Tồn hiện tại', width: 80 },
      { key: 'T1', label: 'Tháng 1', width: 60 },
      { key: 'T2', label: 'Tháng 2', width: 60 },
      { key: 'T3', label: 'Tháng 3', width: 60 },
      { key: 'T4', label: 'Tháng 4', width: 60 },
      { key: 'T5', label: 'Tháng 5', width: 60 },
      { key: 'T6', label: 'Tháng 6', width: 60 },
      { key: 'T7', label: 'Tháng 7', width: 60 },
      { key: 'T8', label: 'Tháng 8', width: 60 },
      { key: 'T9', label: 'Tháng 9', width: 60 },
      { key: 'T10', label: 'Tháng 10', width: 60 },
      { key: 'T11', label: 'Tháng 11', width: 60 },
      { key: 'T12', label: 'Tháng 12', width: 60 },
      { key: 'estimatedUnitPrice', label: 'Đơn giá 1 ngày', width: 110 },
      { key: 'unusedStatus', label: 'Trạng thái', width: 120 }
    ]

    function getColLabel(index) {
      let label = ''
      let temp = index
      while (temp >= 0) {
        label = String.fromCharCode((temp % 26) + 65) + label
        temp = Math.floor(temp / 26) - 1
      }
      return label
    }

    ws['!cols'] = columns.map(c => ({ wpx: c.width }))

    let excelRowIdx = 1

    ws['A1'] = {
      v: `BÁO CÁO THỐNG KÊ TÀI SẢN KHẤU HAO THEO CÁC THÁNG (NĂM ${selectedYear})`,
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '1E3A8A' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      }
    }
    ws['A2'] = {
      v: `Kho / Dự án: ${localProject || 'Tất cả'} | Trạng thái: ${statusFilter === 'all' ? 'Tất cả tài sản' : statusFilter === 'khauhao' ? 'Chưa sử dụng (> 30 ngày)' : 'Đang sử dụng'}`,
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 10, italic: true, color: { rgb: '475569' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      }
    }
    ws['A3'] = {
      v: `Ngày xuất báo cáo: ${new Date().toLocaleDateString('vi-VN')} | Tổng số loại: ${stats.totalAssets} | Tổng tồn kho: ${stats.totalStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}`,
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '0F172A' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      }
    }

    excelRowIdx = 5

    // Write Headers
    columns.forEach((col, colIdx) => {
      const cellRef = `${getColLabel(colIdx)}${excelRowIdx}`
      const isMonthCol = col.key.startsWith('T') && col.key !== 'Tên vật tư'
      
      ws[cellRef] = {
        v: col.label,
        t: 's',
        s: {
          fill: { fgColor: { rgb: isMonthCol ? '1D4ED8' : '1E3A8A' } }, // Slightly different shade of blue for months
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: '475569' } },
            bottom: { style: 'medium', color: { rgb: '1E293B' } },
            left: { style: 'thin', color: { rgb: '475569' } },
            right: { style: 'thin', color: { rgb: '475569' } }
          }
        }
      }
    })

    // Write Data Rows
    sortedData.forEach((item, idx) => {
      excelRowIdx++
      const cells = [
        idx + 1,
        item.maSAP,
        item.maVatTu,
        item.tenVatTu,
        item.dvt,
        item.thongSoKyThuat,
        item.stock,
        item.monthlyStock.T1,
        item.monthlyStock.T2,
        item.monthlyStock.T3,
        item.monthlyStock.T4,
        item.monthlyStock.T5,
        item.monthlyStock.T6,
        item.monthlyStock.T7,
        item.monthlyStock.T8,
        item.monthlyStock.T9,
        item.monthlyStock.T10,
        item.monthlyStock.T11,
        item.monthlyStock.T12,
        item.estimatedUnitPrice,
        item.unusedStatus
      ]

      cells.forEach((val, colIdx) => {
        const cellRef = `${getColLabel(colIdx)}${excelRowIdx}`
        const isNum = typeof val === 'number'
        const isMonthCol = colIdx >= 7 && colIdx <= 18

        const baseStyle = {
          font: { name: 'Segoe UI', sz: 9.5 },
          alignment: {
            horizontal: colIdx === 3 || colIdx === 5 ? 'left' : (isNum ? 'right' : 'center'),
            vertical: 'center',
            wrapText: true
          },
          border: {
            top: { style: 'thin', color: { rgb: 'E2E8F0' } },
            bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          }
        }

        // Distinct styling for Monthly stock columns
        if (isMonthCol) {
          baseStyle.fill = { fgColor: { rgb: 'F0F9FF' } } // Very soft blue background for months
          if (val < 0) {
            baseStyle.font = { name: 'Segoe UI', sz: 9.5, color: { rgb: 'EF4444' } } // Red text for negative stock
          }
        }

        if (colIdx === 6) {
          // Tồn hiện tại
          ws[cellRef] = {
            f: `SUMIFS(Don_Nhan!$J:$J, Don_Nhan!$C:$C, $B${excelRowIdx}) - SUMIFS(Don_Giao!$J:$J, Don_Giao!$C:$C, $B${excelRowIdx})`,
            v: val,
            t: 'n',
            z: '#,##0;(#,##0);"-"',
            s: baseStyle
          }
        } else if (isMonthCol) {
          // Tháng 1 -> 12
          const m = colIdx - 6
          ws[cellRef] = {
            f: `SUMIFS(Don_Nhan!$J:$J, Don_Nhan!$C:$C, $B${excelRowIdx}, Don_Nhan!$B:$B, "<="&DATE(${selectedYear}, ${m + 1}, 0)) - SUMIFS(Don_Giao!$J:$J, Don_Giao!$C:$C, $B${excelRowIdx}, Don_Giao!$B:$B, "<="&DATE(${selectedYear}, ${m + 1}, 0))`,
            v: val,
            t: 'n',
            z: '#,##0;(#,##0);"-"',
            s: baseStyle
          }
        } else {
          const cellObj = { v: val, t: isNum ? 'n' : 's', s: baseStyle }
          if (isNum && colIdx === 19) {
            cellObj.z = '#,##0;(#,##0);"-"'
          }
          ws[cellRef] = cellObj
        }
      })
    })

    // Write Summary Row
    excelRowIdx++
    ws[`A${excelRowIdx}`] = {
      v: 'TỔNG CỘNG',
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '0F172A' } },
        alignment: { horizontal: 'center', vertical: 'center' },
        fill: { fgColor: { rgb: 'E2E8F0' } },
        border: {
          top: { style: 'medium', color: { rgb: '1E293B' } },
          bottom: { style: 'medium', color: { rgb: '1E293B' } },
          left: { style: 'thin', color: { rgb: 'CBD5E1' } },
          right: { style: 'thin', color: { rgb: 'CBD5E1' } }
        }
      }
    }

    // Merges for headers
    ws['!merges'] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 20 } }, // Title merge
      { s: { r: 1, c: 0 }, e: { r: 1, c: 20 } }, // Metadata merge
      { s: { r: 2, c: 0 }, e: { r: 2, c: 20 } }, // Summary metadata merge
      { s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: 5 } } // Summary Row label merge
    ]

    // Style the remaining columns of summary row
    for (let c = 6; c < columns.length; c++) {
      const colLabel = getColLabel(c)
      const cellRef = `${colLabel}${excelRowIdx}`
      const baseStyle = {
        font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '0F172A' } },
        alignment: { horizontal: 'right', vertical: 'center' },
        fill: { fgColor: { rgb: 'E2E8F0' } },
        border: {
          top: { style: 'medium', color: { rgb: '1E293B' } },
          bottom: { style: 'medium', color: { rgb: '1E293B' } },
          left: { style: 'thin', color: { rgb: 'CBD5E1' } },
          right: { style: 'thin', color: { rgb: 'CBD5E1' } }
        }
      }

      // Add SUM formulas for Tồn hiện tại (G) and Monthly stocks (H to S)
      if (c >= 6 && c <= 18) {
        ws[cellRef] = {
          f: `SUM(${colLabel}6:${colLabel}${excelRowIdx - 1})`,
          t: 'n',
          z: '#,##0;(#,##0);"-"',
          s: baseStyle
        }
      } else {
        ws[cellRef] = {
          v: '',
          t: 's',
          s: baseStyle
        }
      }
    }

    ws['!ref'] = `A1:U${excelRowIdx}`
    XLSXStyle.utils.book_append_sheet(wb, ws, 'Thống kê Tài sản Khấu hao')

    // Append Don_Nhan and Don_Giao sheets
    const nhanSheet = buildDetailSheet(nhanList, 'CHI TIẾT ĐƠN NHẬN TÀI SẢN KHẤU HAO', 'Khối lượng nhận')
    XLSXStyle.utils.book_append_sheet(wb, nhanSheet, 'Don_Nhan')

    const giaoSheet = buildDetailSheet(xuatList, 'CHI TIẾT ĐƠN GIAO TÀI SẢN KHẤU HAO', 'Khối lượng giao')
    XLSXStyle.utils.book_append_sheet(wb, giaoSheet, 'Don_Giao')

    const fileProjName = localProject ? String(localProject).replace(/[^a-zA-Z0-9_đĐâÂêÊôÔ]/g, '_') : 'Tat_ca'
    XLSXStyle.writeFile(wb, `SGC_ThongKe_TaiSan_KhauHao_TheoThang_${selectedYear}_${fileProjName}.xlsx`)
  }

  return (
    <div style={{ padding: '20px', height: '100%', display: 'flex', flexDirection: 'column', overflowY: 'auto', boxSizing: 'border-box', background: '#f8fafc' }}>
      
      {/* Sub-tabs Selection & Export controls */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '16px', flexWrap: 'wrap', gap: '12px', borderBottom: '2px solid #e2e8f0', paddingBottom: '12px' }}>
        <div style={{ display: 'flex', gap: '8px' }}>
          <button
            onClick={() => {
              setActiveSubTab('detail')
              setSearchTerm('')
            }}
            style={{
              padding: '10px 20px',
              borderRadius: '8px',
              fontSize: '14px',
              fontWeight: 700,
              cursor: 'pointer',
              background: activeSubTab === 'detail' ? '#1e3a8a' : '#ffffff',
              color: activeSubTab === 'detail' ? '#ffffff' : '#64748b',
              border: '1px solid ' + (activeSubTab === 'detail' ? '#1e3a8a' : '#cbd5e1'),
              boxShadow: activeSubTab === 'detail' ? '0 4px 10px rgba(30, 58, 138, 0.15)' : 'none',
              transition: 'all 0.2s',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            <Info size={16} />
            Báo cáo chi tiết
          </button>
          <button
            onClick={() => {
              setActiveSubTab('summary_all')
              setSearchTerm('')
            }}
            style={{
              padding: '10px 20px',
              borderRadius: '8px',
              fontSize: '14px',
              fontWeight: 700,
              cursor: 'pointer',
              background: activeSubTab === 'summary_all' ? '#1e3a8a' : '#ffffff',
              color: activeSubTab === 'summary_all' ? '#ffffff' : '#64748b',
              border: '1px solid ' + (activeSubTab === 'summary_all' ? '#1e3a8a' : '#cbd5e1'),
              boxShadow: activeSubTab === 'summary_all' ? '0 4px 10px rgba(30, 58, 138, 0.15)' : 'none',
              transition: 'all 0.2s',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            <Layers size={16} />
            Tổng hợp toàn công ty
          </button>
          <button
            onClick={() => {
              setActiveSubTab('summary_month')
              setSearchTerm('')
            }}
            style={{
              padding: '10px 20px',
              borderRadius: '8px',
              fontSize: '14px',
              fontWeight: 700,
              cursor: 'pointer',
              background: activeSubTab === 'summary_month' ? '#1e3a8a' : '#ffffff',
              color: activeSubTab === 'summary_month' ? '#ffffff' : '#64748b',
              border: '1px solid ' + (activeSubTab === 'summary_month' ? '#1e3a8a' : '#cbd5e1'),
              boxShadow: activeSubTab === 'summary_month' ? '0 4px 10px rgba(30, 58, 138, 0.15)' : 'none',
              transition: 'all 0.2s',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            <Calendar size={16} />
            Tổng hợp báo cáo tháng
          </button>
          <button
            onClick={() => {
              setActiveSubTab('material_stats')
              setSearchTerm('')
            }}
            style={{
              padding: '10px 20px',
              borderRadius: '8px',
              fontSize: '14px',
              fontWeight: 700,
              cursor: 'pointer',
              background: activeSubTab === 'material_stats' ? '#1e3a8a' : '#ffffff',
              color: activeSubTab === 'material_stats' ? '#ffffff' : '#64748b',
              border: '1px solid ' + (activeSubTab === 'material_stats' ? '#1e3a8a' : '#cbd5e1'),
              boxShadow: activeSubTab === 'material_stats' ? '0 4px 10px rgba(30, 58, 138, 0.15)' : 'none',
              transition: 'all 0.2s',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            <Package size={16} />
            Thống kê theo vật tư
          </button>
          <button
            onClick={() => {
              setActiveSubTab('reconciliation')
              setSearchTerm('')
            }}
            style={{
              padding: '10px 20px',
              borderRadius: '8px',
              fontSize: '14px',
              fontWeight: 700,
              cursor: 'pointer',
              background: activeSubTab === 'reconciliation' ? '#1e3a8a' : '#ffffff',
              color: activeSubTab === 'reconciliation' ? '#ffffff' : '#64748b',
              border: '1px solid ' + (activeSubTab === 'reconciliation' ? '#1e3a8a' : '#cbd5e1'),
              boxShadow: activeSubTab === 'reconciliation' ? '0 4px 10px rgba(30, 58, 138, 0.15)' : 'none',
              transition: 'all 0.2s',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            <CheckCircle2 size={16} />
            Đối chiếu
          </button>
        </div>

        <button 
          onClick={handleExportExcel}
          style={{
            background: '#10b981',
            color: '#ffffff',
            border: 'none',
            padding: '10px 16px',
            borderRadius: '8px',
            fontSize: '13.5px',
            fontWeight: 700,
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            gap: '8px',
            boxShadow: '0 4px 6px -1px rgba(16, 185, 129, 0.2)'
          }}
          onMouseOver={(e) => e.currentTarget.style.background = '#059669'}
          onMouseOut={(e) => e.currentTarget.style.background = '#10b981'}
        >
          <Download size={16} />
          Xuất file Excel
        </button>
      </div>


      {/* Filters Box */}
      <div style={{
        background: '#ffffff',
        border: '1px solid #e2e8f0',
        borderRadius: '12px',
        padding: '16px',
        marginBottom: '16px',
        display: 'flex',
        gap: '12px',
        alignItems: 'center',
        flexWrap: 'wrap',
        boxShadow: '0 1px 3px rgba(0,0,0,0.05)'
      }}>
        {/* Project Selector - Hide on Company-wide Summary and Material Stats */}
        {activeSubTab === 'detail' && (
          <div style={{ flex: '1 1 200px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
            <label style={{ fontSize: '12px', fontWeight: 700, color: '#475569' }}>Chọn Kho / Dự án:</label>
            <LocalSearchableSelect
              value={localProject}
              onChange={(val) => setLocalProject(val)}
              options={uniquePhysicalWarehouses}
              placeholder="Tất cả Kho / Dự án"
              searchPlaceholder="Tìm kiếm dự án..."
            />
          </div>
        )}

        {/* Year Selector */}
        <div style={{ flex: '0 0 120px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
          <label style={{ fontSize: '12px', fontWeight: 700, color: '#475569' }}>Chọn Năm thống kê:</label>
          <select
            value={selectedYear}
            onChange={(e) => setSelectedYear(parseInt(e.target.value, 10))}
            style={{
              padding: '8px 12px',
              border: '1px solid #cbd5e1',
              borderRadius: '6px',
              fontSize: '13px',
              fontWeight: 600,
              background: '#ffffff',
              color: '#1e293b',
              height: '38px',
              outline: 'none'
            }}
          >
            {uniqueYears.map(yr => (
              <option key={yr} value={yr}>Năm {yr}</option>
            ))}
          </select>
        </div>

        {/* Status Selector - Hide on Company-wide Summary */}
        {activeSubTab === 'detail' && (
          <div style={{ flex: '1 1 180px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
            <label style={{ fontSize: '12px', fontWeight: 700, color: '#475569' }}>Bộ lọc trạng thái:</label>
            <select
              value={statusFilter}
              onChange={(e) => setStatusFilter(e.target.value)}
              style={{
                padding: '8px 12px',
                border: '1px solid #cbd5e1',
                borderRadius: '6px',
                fontSize: '13px',
                fontWeight: 500,
                background: '#ffffff',
                color: '#1e293b',
                height: '38px',
                outline: 'none'
              }}
            >
              <option value="all">Tất cả tài sản</option>
              <option value="khauhao">Chưa sử dụng (&gt; 30 ngày)</option>
              <option value="active">Đang sử dụng (&lt;= 30 ngày / Hết tồn)</option>
            </select>
          </div>
        )}

        {/* Text Search */}
        <div style={{ flex: '2 1 240px', display: 'flex', flexDirection: 'column', gap: '6px' }}>
          <label style={{ fontSize: '12px', fontWeight: 700, color: '#475569' }}>
            {activeSubTab === 'summary_all' ? 'Tìm kiếm Kho / Dự án:' : activeSubTab === 'summary_month' ? 'Tìm kiếm Kho / Dự án / Mã SAP / Vật tư:' : activeSubTab === 'reconciliation' ? 'Tìm kiếm nhanh vật tư:' : 'Tìm kiếm nhanh:'}
          </label>
          <div style={{ position: 'relative', width: '100%' }}>
            <input
              type="text"
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              placeholder={activeSubTab === 'summary_all' ? 'Nhập tên kho / dự án cần tìm...' : activeSubTab === 'summary_month' ? 'Tìm theo kho, dự án, mã SAP hoặc tên vật tư...' : activeSubTab === 'reconciliation' ? 'Tìm theo mã SAP, tên vật tư...' : 'Tìm theo mã, tên vật tư...'}
              style={{
                width: '100%',
                padding: '8px 12px 8px 36px',
                border: '1px solid #cbd5e1',
                borderRadius: '6px',
                fontSize: '13px',
                height: '38px',
                boxSizing: 'border-box',
                outline: 'none',
                color: '#1e293b'
              }}
            />
            <Search size={16} color="#94a3b8" style={{ position: 'absolute', left: '12px', top: '11px' }} />
          </div>
        </div>
      </div>

      {/* Alert Banner for Reconciliation */}
      {activeSubTab === 'reconciliation' && (
        <div style={{
          background: '#ecfdf5',
          border: '1px solid #a7f3d0',
          borderRadius: '12px',
          padding: '14px 18px',
          marginBottom: '16px',
          display: 'flex',
          alignItems: 'center',
          gap: '12px',
          color: '#065f46',
          fontSize: '13px',
          lineHeight: '1.5',
          boxShadow: '0 1px 2px rgba(0,0,0,0.02)'
        }}>
          <div style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            background: '#10b981',
            borderRadius: '50%',
            width: '24px',
            height: '24px',
            minWidth: '24px',
            color: '#ffffff',
            fontWeight: 'bold',
            fontSize: '14px'
          }}>
            ✓
          </div>
          <div>
            <strong style={{ fontWeight: 700, color: '#047857', display: 'block', marginBottom: '2px', fontSize: '13.5px' }}>
              Hệ thống đối soát tự động: Phân tích chênh lệch đối chiếu chi tiết giữa Kho Dự án (A) và Nhập từ Nhà cung cấp (B)
            </strong>
            <span>
              Các phát sinh cấp phát/mượn vật tư liên quan đến <strong>Tổ đội (Crews)</strong> được phân loại là lưu chuyển nội bộ tạm thời, 
              không ghi nhận là xuất kho cố định, bảo toàn tính chuẩn xác trong việc đối chiếu khớp số lượng nhập/tồn.
            </span>
          </div>
        </div>
      )}

      {/* Reconciliation Analysis & Discrepancy Cause Report */}
      {activeSubTab === 'reconciliation' && (
        <div style={{ display: 'flex', flexDirection: 'column', gap: '16px', marginBottom: '16px' }}>
          {/* Material-level Discrepancy Summary Table */}
          <div style={{ background: '#ffffff', borderRadius: '12px', border: '1px solid #cbd5e1', padding: '18px', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: '12px', flexWrap: 'wrap', gap: '8px' }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Package size={16} color="#1e3a8a" />
                <span style={{ fontSize: '13.5px', fontWeight: 800, color: '#1e3a8a' }}>
                  📋 DANH SÁCH CÁC LOẠI VẬT TƯ CHÊNH LỆCH / BỊ ÂM TOÀN CÔNG TY
                </span>
              </div>
              <button
                onClick={() => {
                  setShowDiscrepancyDetail(!showDiscrepancyDetail)
                  setDiscrepancyPage(1)
                }}
                style={{
                  padding: '5px 12px',
                  background: showDiscrepancyDetail ? '#475569' : '#1e3a8a',
                  color: '#ffffff',
                  border: 'none',
                  borderRadius: '6px',
                  fontSize: '11px',
                  fontWeight: 600,
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  gap: '4px',
                  transition: 'all 0.15s ease'
                }}
              >
                {showDiscrepancyDetail ? 'Ẩn chi tiết giao dịch lệch' : 'Xem chi tiết giao dịch lệch'}
                <span style={{ fontSize: '9px' }}>{showDiscrepancyDetail ? '▲' : '▼'}</span>
              </button>
            </div>
            <p style={{ fontSize: '12px', color: '#64748b', margin: '0 0 12px 0', lineHeight: '1.4' }}>
              Dưới đây là danh sách tổng hợp các mã vật tư (tài sản khấu hao) bị âm hoặc chênh lệch giữa <strong>Kho Dự án (A)</strong> và <strong>Nhà cung cấp (B)</strong>, được tổng hợp từ các đơn giao nhận chi tiết. Các dòng có chênh lệch âm biểu thị lượng tồn thực tế ở Kho Dự án đang ít hơn lượng Nhà cung cấp bàn giao.
            </p>
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '12px' }}>
                <thead>
                  <tr style={{ background: '#f1f5f9', borderBottom: '2px solid #cbd5e1' }}>
                    <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155', width: '50px', textAlign: 'center' }}>STT</th>
                    <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155', width: '100px' }}>Mã SAP</th>
                    <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155' }}>Tên vật tư (Tài sản)</th>
                    <th style={{ padding: '8px 10px', fontWeight: 700, color: '#0f766e', width: '140px', textAlign: 'right' }}>Kho Dự án (A)</th>
                    <th style={{ padding: '8px 10px', fontWeight: 700, color: '#b45309', width: '140px', textAlign: 'right' }}>Nhà cung cấp (B)</th>
                    <th style={{ padding: '8px 10px', fontWeight: 700, color: '#ef4444', width: '130px', textAlign: 'right' }}>Chênh lệch (A - B)</th>
                    <th style={{ padding: '8px 10px', fontWeight: 700, color: '#1e3a8a', width: '160px', textAlign: 'right' }}>Giá trị lệch tạm tính</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredMaterialMismatchSummary.map((m, idx) => (
                    <tr key={m.maSAP} style={{ borderBottom: '1px solid #e2e8f0', background: idx % 2 === 0 ? '#ffffff' : '#f8fafc' }}>
                      <td style={{ padding: '8px 10px', textAlign: 'center', color: '#64748b' }}>{idx + 1}</td>
                      <td style={{ padding: '8px 10px', fontFamily: 'monospace', fontSize: '11px', color: '#475569', fontWeight: 600 }}>{m.maSAP}</td>
                      <td style={{ padding: '8px 10px', color: '#1e293b', fontWeight: 500 }}>{m.tenVatTu}</td>
                      <td style={{ padding: '8px 10px', textAlign: 'right', color: '#0f766e', fontWeight: 600 }}>{m.effectI.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</td>
                      <td style={{ padding: '8px 10px', textAlign: 'right', color: '#b45309', fontWeight: 600 }}>{m.effectII.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</td>
                      <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 700, color: m.diff < 0 ? '#ef4444' : '#047857' }}>
                        {m.diff > 0 ? `+${m.diff.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}` : m.diff.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                      </td>
                      <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 700, color: m.valueDiff < 0 ? '#ef4444' : '#047857' }}>
                        {m.valueDiff !== 0 ? `${m.valueDiff.toLocaleString('vi-VN')} đ` : '0 đ'}
                      </td>
                    </tr>
                  ))}
                  {filteredMaterialMismatchSummary.length === 0 && (
                    <tr>
                      <td colSpan="7" style={{ padding: '20px', textAlign: 'center', color: '#64748b', fontStyle: 'italic' }}>
                        🎉 Tuyệt vời! Không có chênh lệch tồn kho vật tư nào khớp với từ khóa tìm kiếm hoặc không có chênh lệch nào tồn tại.
                      </td>
                    </tr>
                  )}
                  {materialMismatchSummary.length > 0 && (
                    <tr style={{ background: '#f1f5f9', fontWeight: 800, borderTop: '2px solid #cbd5e1' }}>
                      <td colSpan="3" style={{ padding: '10px 10px', color: '#1e293b', textAlign: 'right' }}>TỔNG CỘNG CHÊNH LỆCH VẬT TƯ:</td>
                      <td style={{ padding: '10px 10px', textAlign: 'right', color: '#0f766e' }}>
                        {materialMismatchSummary.reduce((sum, item) => sum + item.effectI, 0).toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                      </td>
                      <td style={{ padding: '10px 10px', textAlign: 'right', color: '#b45309' }}>
                        {materialMismatchSummary.reduce((sum, item) => sum + item.effectII, 0).toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                      </td>
                      <td style={{ padding: '10px 10px', textAlign: 'right', color: '#ef4444' }}>
                        {reconciliationDifference.stockDiff.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                      </td>
                      <td style={{ padding: '10px 10px', textAlign: 'right', color: reconciliationDifference.valueDiff < 0 ? '#ef4444' : '#047857' }}>
                        {reconciliationDifference.valueDiff !== 0 ? `${reconciliationDifference.valueDiff.toLocaleString('vi-VN')} đ` : '0 đ'}
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>

          {/* Interactive Mismatch Detail Table */}
          {showDiscrepancyDetail && (
            <div style={{ marginTop: '16px', background: '#ffffff', borderRadius: '8px', border: '1px solid #cbd5e1', padding: '14px' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '10px', marginBottom: '12px' }}>
                <span style={{ fontSize: '13px', fontWeight: 700, color: '#334155' }}>
                  Bảng kê chi tiết {mismatchAnalysis.length} giao dịch lệch dòng đối chiếu:
                </span>
                <div style={{ position: 'relative', width: '280px' }}>
                  <input
                    type="text"
                    placeholder="Tìm theo Mã SAP, tên vật tư, đơn vị..."
                    value={discrepancySearch}
                    onChange={(e) => {
                      setDiscrepancySearch(e.target.value)
                      setDiscrepancyPage(1)
                    }}
                    style={{
                      width: '100%',
                      padding: '6px 10px 6px 30px',
                      fontSize: '12px',
                      border: '1px solid #cbd5e1',
                      borderRadius: '6px',
                      boxSizing: 'border-box',
                      outline: 'none'
                    }}
                  />
                  <Search size={14} color="#94a3b8" style={{ position: 'absolute', left: '10px', top: '9px' }} />
                </div>
              </div>

              {/* Mismatch Table */}
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '12px' }}>
                  <thead>
                    <tr style={{ background: '#f1f5f9', borderBottom: '2px solid #cbd5e1' }}>
                      <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155', width: '50px', textAlign: 'center' }}>STT</th>
                      <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155', width: '90px' }}>Mã đơn</th>
                      <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155', width: '90px', textAlign: 'center' }}>Ngày</th>
                      <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155', width: '90px' }}>Mã SAP</th>
                      <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155' }}>Tên vật tư (Tài sản)</th>
                      <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155', width: '130px' }}>Đơn vị giao</th>
                      <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155', width: '130px' }}>Đơn vị nhận</th>
                      <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155', width: '100px', textAlign: 'right' }}>Số lượng</th>
                      <th style={{ padding: '8px 10px', fontWeight: 700, color: '#ef4444', width: '110px', textAlign: 'right' }}>Chênh lệch</th>
                      <th style={{ padding: '8px 10px', fontWeight: 700, color: '#334155', width: '220px' }}>Nguyên nhân thực tế</th>
                    </tr>
                  </thead>
                  <tbody>
                    {mismatchAnalysis
                      .filter(m => {
                        const q = discrepancySearch.toLowerCase().trim()
                        if (!q) return true
                        return m.maSAP.toLowerCase().includes(q) ||
                          m.tenVatTu.toLowerCase().includes(q) ||
                          m.giaoUnit.toLowerCase().includes(q) ||
                          m.nhanUnit.toLowerCase().includes(q) ||
                          m.reason.toLowerCase().includes(q)
                      })
                      .slice((discrepancyPage - 1) * discrepancyPageSize, discrepancyPage * discrepancyPageSize)
                      .map((m, idx) => {
                        const globalIdx = (discrepancyPage - 1) * discrepancyPageSize + idx + 1
                        return (
                          <tr key={m.rowId} style={{ borderBottom: '1px solid #e2e8f0', background: idx % 2 === 0 ? '#ffffff' : '#f8fafc' }}>
                            <td style={{ padding: '8px 10px', textAlign: 'center', color: '#64748b' }}>{globalIdx}</td>
                            <td style={{ padding: '8px 10px', fontWeight: 600, color: '#475569' }}>{m.maDon}</td>
                            <td style={{ padding: '8px 10px', textAlign: 'center', color: '#64748b' }}>{m.date}</td>
                            <td style={{ padding: '8px 10px', fontFamily: 'monospace', fontSize: '11px' }}>{m.maSAP}</td>
                            <td style={{ padding: '8px 10px', fontWeight: 500, color: '#1e293b' }}>{m.tenVatTu}</td>
                            <td style={{ padding: '8px 10px', color: '#475569' }}>{m.giaoUnit || <span style={{ color: '#ef4444', fontStyle: 'italic' }}>Khuyết</span>}</td>
                            <td style={{ padding: '8px 10px', color: '#475569' }}>{m.nhanUnit || <span style={{ color: '#ef4444', fontStyle: 'italic' }}>Khuyết</span>}</td>
                            <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600 }}>{m.qty.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</td>
                            <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 700, color: m.diff < 0 ? '#ef4444' : '#047857' }}>
                              {m.diff > 0 ? `+${m.diff.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}` : m.diff.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                            </td>
                            <td style={{ padding: '8px 10px', color: '#b45309', fontSize: '11.5px', lineHeight: '1.4' }}>{m.reason}</td>
                          </tr>
                        )
                      })}
                    {mismatchAnalysis.filter(m => {
                      const q = discrepancySearch.toLowerCase().trim()
                      if (!q) return true
                      return m.maSAP.toLowerCase().includes(q) ||
                        m.tenVatTu.toLowerCase().includes(q) ||
                        m.giaoUnit.toLowerCase().includes(q) ||
                        m.nhanUnit.toLowerCase().includes(q) ||
                        m.reason.toLowerCase().includes(q)
                    }).length === 0 && (
                      <tr>
                        <td colSpan="10" style={{ padding: '20px', textAlign: 'center', color: '#64748b', fontStyle: 'italic' }}>
                          Không tìm thấy giao dịch chênh lệch nào khớp với từ khóa tìm kiếm.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {/* Discrepancy Pagination */}
              {mismatchAnalysis.filter(m => {
                const q = discrepancySearch.toLowerCase().trim()
                if (!q) return true
                return m.maSAP.toLowerCase().includes(q) ||
                  m.tenVatTu.toLowerCase().includes(q) ||
                  m.giaoUnit.toLowerCase().includes(q) ||
                  m.nhanUnit.toLowerCase().includes(q) ||
                  m.reason.toLowerCase().includes(q)
              }).length > discrepancyPageSize && (
                <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginTop: '12px', fontSize: '12px' }}>
                  <span style={{ color: '#64748b' }}>
                    Hiển thị {Math.min(discrepancyPage * discrepancyPageSize, mismatchAnalysis.filter(m => {
                      const q = discrepancySearch.toLowerCase().trim()
                      if (!q) return true
                      return m.maSAP.toLowerCase().includes(q) ||
                        m.tenVatTu.toLowerCase().includes(q) ||
                        m.giaoUnit.toLowerCase().includes(q) ||
                        m.nhanUnit.toLowerCase().includes(q) ||
                        m.reason.toLowerCase().includes(q)
                    }).length)} / {mismatchAnalysis.filter(m => {
                      const q = discrepancySearch.toLowerCase().trim()
                      if (!q) return true
                      return m.maSAP.toLowerCase().includes(q) ||
                        m.tenVatTu.toLowerCase().includes(q) ||
                        m.giaoUnit.toLowerCase().includes(q) ||
                        m.nhanUnit.toLowerCase().includes(q) ||
                        m.reason.toLowerCase().includes(q)
                    }).length} dòng
                  </span>
                  <div style={{ display: 'flex', gap: '6px' }}>
                    <button
                      onClick={() => setDiscrepancyPage(p => Math.max(1, p - 1))}
                      disabled={discrepancyPage === 1}
                      style={{
                        padding: '4px 10px',
                        background: '#ffffff',
                        border: '1px solid #cbd5e1',
                        borderRadius: '4px',
                        cursor: discrepancyPage === 1 ? 'not-allowed' : 'pointer',
                        opacity: discrepancyPage === 1 ? 0.5 : 1
                      }}
                    >
                      Trước
                    </button>
                    <button
                      onClick={() => setDiscrepancyPage(p => p + 1)}
                      disabled={discrepancyPage * discrepancyPageSize >= mismatchAnalysis.filter(m => {
                        const q = discrepancySearch.toLowerCase().trim()
                        if (!q) return true
                        return m.maSAP.toLowerCase().includes(q) ||
                          m.tenVatTu.toLowerCase().includes(q) ||
                          m.giaoUnit.toLowerCase().includes(q) ||
                          m.nhanUnit.toLowerCase().includes(q) ||
                          m.reason.toLowerCase().includes(q)
                      }).length}
                      style={{
                        padding: '4px 10px',
                        background: '#ffffff',
                        border: '1px solid #cbd5e1',
                        borderRadius: '4px',
                        cursor: discrepancyPage * discrepancyPageSize >= mismatchAnalysis.filter(m => {
                          const q = discrepancySearch.toLowerCase().trim()
                          if (!q) return true
                          return m.maSAP.toLowerCase().includes(q) ||
                            m.tenVatTu.toLowerCase().includes(q) ||
                            m.giaoUnit.toLowerCase().includes(q) ||
                            m.nhanUnit.toLowerCase().includes(q) ||
                            m.reason.toLowerCase().includes(q)
                        }).length ? 'not-allowed' : 'pointer',
                        opacity: discrepancyPage * discrepancyPageSize >= mismatchAnalysis.filter(m => {
                          const q = discrepancySearch.toLowerCase().trim()
                          if (!q) return true
                          return m.maSAP.toLowerCase().includes(q) ||
                            m.tenVatTu.toLowerCase().includes(q) ||
                            m.giaoUnit.toLowerCase().includes(q) ||
                            m.nhanUnit.toLowerCase().includes(q) ||
                            m.reason.toLowerCase().includes(q)
                        }).length ? 0.5 : 1
                      }}
                    >
                      Sau
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {/* Data Table */}
      <div style={{ background: '#ffffff', border: '1px solid #e2e8f0', borderRadius: '12px', overflow: 'hidden', boxShadow: '0 1px 3px rgba(0,0,0,0.05)', display: 'flex', flexDirection: 'column', flex: 1 }}>
        <div style={{ overflowX: 'auto', flex: 1 }}>
          {activeSubTab === 'summary_all' ? (
            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '12.5px', minWidth: '1150px' }}>
              <thead>
                <tr style={{ background: '#1e3a8a', borderBottom: '2px solid #0f172a' }}>
                  <th style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', width: '60px', textAlign: 'center' }}>STT</th>
                  <th onClick={() => handleSumSort('projectName')} style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '250px' }}>
                    Kho / Dự án {sumSortField === 'projectName' && (sumSortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                  <th onClick={() => handleSumSort('totalAssets')} style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '110px', textAlign: 'center' }}>
                    Số chủng loại {sumSortField === 'totalAssets' && (sumSortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                  <th onClick={() => handleSumSort('totalStock')} style={{ padding: '12px 10px', fontWeight: 700, color: '#facc15', cursor: 'pointer', userSelect: 'none', width: '110px', textAlign: 'right' }}>
                    Tổng tồn kho {sumSortField === 'totalStock' && (sumSortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                  
                  {/* 12 Months columns */}
                  {[...Array(12)].map((_, i) => {
                    const m = i + 1
                    const key = `T${m}`
                    return (
                      <th 
                        key={key}
                        onClick={() => handleSumSort(key)}
                        style={{ 
                          padding: '12px 6px', 
                          fontWeight: 700, 
                          color: '#ffffff', 
                          textAlign: 'center', 
                          cursor: 'pointer', 
                          userSelect: 'none',
                          width: '75px',
                          background: '#1d4ed8', // Distinct background color for T1-T12 headers
                          borderLeft: '1px solid rgba(255, 255, 255, 0.25)', // Vertical borders for header
                          borderRight: '1px solid rgba(255, 255, 255, 0.25)'
                        }}
                      >
                        T{m} {sumSortField === key && (sumSortDirection === 'asc' ? '▲' : '▼')}
                      </th>
                    )
                  })}

                  <th onClick={() => handleSumSort('totalValue')} style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '150px', textAlign: 'right' }}>
                    Giá trị tồn tạm tính {sumSortField === 'totalValue' && (sumSortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                </tr>
              </thead>
              <tbody>
                {companyProjectSummaries.length === 0 ? (
                  <tr>
                    <td colSpan="17" style={{ padding: '32px', textTransform: 'none', color: '#94a3b8', textAlign: 'center', fontSize: '14px', fontWeight: 500 }}>
                      Không có dữ liệu tổng hợp toàn công ty.
                    </td>
                  </tr>
                ) : (
                  <>
                    {groupedProjectSummaries.map((group) => {
                      // Calculate group subtotals
                      const groupAssets = group.items.reduce((sum, item) => sum + item.totalAssets, 0)
                      const groupStock = group.items.reduce((sum, item) => sum + item.totalStock, 0)
                      const groupValue = group.items.reduce((sum, item) => sum + item.totalValue, 0)
                      const groupDepreciation = group.items.reduce((sum, item) => sum + item.totalDepreciation, 0)
                      const groupMonthly = {}
                      for (let m = 1; m <= 12; m++) {
                        groupMonthly[`T${m}`] = group.items.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0), 0)
                      }

                      return (
                        <React.Fragment key={group.id}>
                          {/* Group Title Row */}
                          <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1' }}>
                            <td colSpan="17" style={{ padding: '10px 14px', fontSize: '13px', fontWeight: 800, color: '#1e3a8a', background: '#f1f5f9', letterSpacing: '0.5px' }}>
                              {group.title}
                            </td>
                          </tr>

                          {group.items.length === 0 ? (
                            <tr>
                              <td colSpan="17" style={{ padding: '12px', color: '#94a3b8', textAlign: 'center', fontStyle: 'italic' }}>
                                Không có dữ liệu cho nhóm này.
                              </td>
                            </tr>
                          ) : (
                            group.items.map((item, idx) => {
                              return (
                                <tr 
                                  key={item.projectName} 
                                  title="Nhấp đúp để xem chi tiết kho / dự án này"
                                  style={{ 
                                    borderBottom: '1px solid #e2e8f0',
                                    background: '#ffffff',
                                    transition: 'background 0.2s',
                                    cursor: 'pointer'
                                  }}
                                  onDoubleClick={() => {
                                    setLocalProject(item.projectName)
                                    setActiveSubTab('detail')
                                    setSearchTerm('')
                                  }}
                                  onMouseOver={(e) => { e.currentTarget.style.background = '#f8fafc' }}
                                  onMouseOut={(e) => { e.currentTarget.style.background = '#ffffff' }}
                                >
                                  <td style={{ padding: '10px', textAlign: 'center', color: '#64748b', fontWeight: 600 }}>{idx + 1}</td>
                                  <td style={{ padding: '10px', fontWeight: 700, color: '#1e3a8a' }}>{item.projectName}</td>
                                  <td style={{ padding: '10px', textAlign: 'center', fontWeight: 600, color: '#475569' }}>{item.totalAssets}</td>
                                  <td style={{ padding: '10px', textAlign: 'right', fontWeight: 700, color: '#10b981' }}>{item.totalStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</td>
                                  
                                  {/* 12 Months dynamic balances */}
                                  {[...Array(12)].map((_, i) => {
                                    const m = i + 1
                                    const val = item.monthlyStock[`T${m}`]
                                    return (
                                      <td 
                                        key={m}
                                        style={{ 
                                          padding: '10px 6px', 
                                          textAlign: 'right', 
                                          fontWeight: val !== 0 ? '600' : '400',
                                          color: val > 0 ? '#0f172a' : (val < 0 ? '#ef4444' : '#94a3b8'),
                                          background: 'rgba(219, 234, 254, 0.25)', // light blue tint
                                          borderLeft: '1px solid #cbd5e1', // Vertical borders between monthly columns
                                          borderRight: '1px solid #cbd5e1'
                                        }}
                                      >
                                        {val !== 0 ? val.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                                      </td>
                                    )
                                  })}

                                  <td style={{ padding: '10px', textAlign: 'right', fontWeight: 700, color: '#475569' }}>
                                    {item.totalValue > 0 ? `${item.totalValue.toLocaleString('vi-VN')} đ` : '-'}
                                  </td>
                                </tr>
                              )
                            })
                          )}

                          {/* Subtotal Row for this Group */}
                          <tr style={{ background: '#f8fafc', fontWeight: 700, borderBottom: '2px solid #cbd5e1', borderTop: '1px solid #cbd5e1' }}>
                            <td colSpan="2" style={{ padding: '10px 14px', textTransform: 'none', color: '#1e3a8a', fontSize: '12.5px', fontWeight: 800 }}>
                              CỘNG {group.id === 'I' ? 'KHO DỰ ÁN (A)' : 'NHẬP TỪ NCC (B)'}
                            </td>
                            <td style={{ padding: '10px', textAlign: 'center', color: '#475569' }}>{groupAssets}</td>
                            <td style={{ padding: '10px', textAlign: 'right', color: '#10b981' }}>{groupStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</td>
                            
                            {/* Monthly subtotals */}
                            {[...Array(12)].map((_, i) => {
                              const m = i + 1
                              const val = groupMonthly[`T${m}`]
                              return (
                                <td 
                                  key={m}
                                  style={{ 
                                    padding: '10px 6px', 
                                    textAlign: 'right', 
                                    color: val > 0 ? '#0f172a' : (val < 0 ? '#ef4444' : '#64748b'),
                                    background: 'rgba(219, 234, 254, 0.4)',
                                    borderLeft: '1px solid #cbd5e1',
                                    borderRight: '1px solid #cbd5e1'
                                  }}
                                >
                                  {val !== 0 ? val.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                                </td>
                              )
                            })}
                            
                            <td style={{ padding: '10px', textAlign: 'right', color: '#475569' }}>
                              {groupValue > 0 ? `${groupValue.toLocaleString('vi-VN')} đ` : '-'}
                            </td>
                          </tr>
                        </React.Fragment>
                      )
                    })}
                    
                    {/* Reconciled Difference Row */}
                    <tr style={{ background: '#ecfdf5', fontWeight: 800, borderTop: '2px solid #10b981', borderBottom: '2px solid #10b981' }}>
                      <td colSpan="2" style={{ padding: '12px 10px', textAlign: 'center', color: '#047857' }}>
                        ✓ CHÊNH LỆCH ĐỐI CHIẾU (A - B)
                      </td>
                      <td style={{ padding: '12px 10px', textAlign: 'center', color: '#047857' }}>-</td>
                      <td style={{ padding: '12px 10px', textAlign: 'right', color: '#047857' }}>
                        {reconciliationDifference.stockDiff.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                      </td>
                      
                      {/* 12 Months difference calculated dynamically */}
                      {[...Array(12)].map((_, i) => {
                        const m = i + 1
                        const val = reconciliationDifference.monthlyDiff[`T${m}`]
                        return (
                          <td 
                            key={m}
                            style={{ 
                              padding: '12px 6px', 
                              textAlign: 'right', 
                              color: val !== 0 ? (val < 0 ? '#ef4444' : '#047857') : '#047857',
                              background: '#ecfdf5',
                              borderLeft: '1px solid #cbd5e1',
                              borderRight: '1px solid #cbd5e1'
                            }}
                          >
                            {val !== 0 ? val.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                          </td>
                        )
                      })}
                      
                      <td style={{ padding: '12px 10px', textAlign: 'right', color: '#047857' }}>
                        {reconciliationDifference.valueDiff !== 0 ? `${reconciliationDifference.valueDiff.toLocaleString('vi-VN')} đ` : '0 đ'}
                      </td>
                    </tr>
                  </>
                )}
              </tbody>
            </table>
          ) : activeSubTab === 'summary_month' ? (
            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '11.5px', minWidth: '2600px', border: '1px solid #cbd5e1' }}>
              <thead>
                <tr style={{ background: '#1e3a8a', borderBottom: '1px solid #cbd5e1' }}>
                  <th rowSpan={2} style={{ padding: '10px 8px', fontWeight: 700, color: '#ffffff', width: '50px', textAlign: 'center', verticalAlign: 'middle', border: '1px solid #cbd5e1' }}>STT</th>
                  <th rowSpan={2} onClick={() => handleSumSort('projectName')} style={{ padding: '10px 8px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '220px', textAlign: 'center', verticalAlign: 'middle', border: '1px solid #cbd5e1' }}>
                    Kho / Dự án {sumSortField === 'projectName' && (sumSortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                  <th rowSpan={2} onClick={() => handleSumSort('maSAP')} style={{ padding: '10px 8px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '100px', textAlign: 'center', verticalAlign: 'middle', border: '1px solid #cbd5e1' }}>
                    Mã SAP {sumSortField === 'maSAP' && (sumSortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                  <th rowSpan={2} onClick={() => handleSumSort('tenVatTu')} style={{ padding: '10px 8px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '240px', textAlign: 'center', verticalAlign: 'middle', border: '1px solid #cbd5e1' }}>
                    Tên vật tư {sumSortField === 'tenVatTu' && (sumSortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                  <th rowSpan={2} onClick={() => handleSumSort('dvt')} style={{ padding: '10px 8px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '70px', textAlign: 'center', verticalAlign: 'middle', border: '1px solid #cbd5e1' }}>
                    Đvt {sumSortField === 'dvt' && (sumSortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                  <th rowSpan={2} onClick={() => handleSumSort('totalStock')} style={{ padding: '10px 8px', fontWeight: 700, color: '#facc15', cursor: 'pointer', userSelect: 'none', width: '100px', textAlign: 'center', verticalAlign: 'middle', border: '1px solid #cbd5e1' }}>
                    Tổng tồn kho {sumSortField === 'totalStock' && (sumSortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                  
                  <th rowSpan={2} onClick={() => handleSumSort('unitPrice')} style={{ padding: '10px 8px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '130px', textAlign: 'center', verticalAlign: 'middle', border: '1px solid #cbd5e1' }}>
                    Đơn giá tháng {sumSortField === 'unitPrice' && (sumSortDirection === 'asc' ? '▲' : '▼')}
                  </th>

                  <th colSpan={12} style={{ padding: '6px 8px', fontWeight: 700, color: '#ffffff', textAlign: 'center', background: '#1d4ed8', border: '1px solid #cbd5e1', fontSize: '12px' }}>
                    KHỐI LƯỢNG TỒN KHO THEO THÁNG
                  </th>
                  
                  <th colSpan={12} style={{ padding: '6px 8px', fontWeight: 700, color: '#ffffff', textAlign: 'center', background: '#047857', border: '1px solid #cbd5e1', fontSize: '12px' }}>
                    THÀNH TIỀN TÀI SẢN KHẤU HAO
                  </th>
                </tr>
                <tr style={{ background: '#1e3a8a', borderBottom: '1px solid #cbd5e1' }}>
                  {/* Khối lượng subheaders */}
                  {[...Array(12)].map((_, i) => {
                    const m = i + 1
                    const key = `T${m}`
                    return (
                      <th 
                        key={`qty_T${m}`}
                        onClick={() => handleSumSort(key)}
                        style={{ 
                          padding: '6px 4px', 
                          fontWeight: 700, 
                          color: '#ffffff', 
                          textAlign: 'center', 
                          cursor: 'pointer', 
                          userSelect: 'none',
                          width: '60px',
                          background: '#2563eb',
                          border: '1px solid #cbd5e1',
                          fontSize: '10.5px'
                        }}
                      >
                        T{m} {sumSortField === key && (sumSortDirection === 'asc' ? '▲' : '▼')}
                      </th>
                    )
                  })}
                  {/* Thành tiền subheaders */}
                  {[...Array(12)].map((_, i) => {
                    const m = i + 1
                    return (
                      <th 
                        key={`amt_T${m}`}
                        style={{ 
                          padding: '6px 4px', 
                          fontWeight: 700, 
                          color: '#ffffff', 
                          textAlign: 'center', 
                          userSelect: 'none',
                          width: '85px',
                          background: '#059669',
                          border: '1px solid #cbd5e1',
                          fontSize: '10.5px'
                        }}
                      >
                        T{m}
                      </th>
                    )
                  })}
                </tr>
              </thead>
              <tbody>
                {companyProjectMonthlySummaries.length === 0 ? (
                  <tr>
                    <td colSpan={31} style={{ padding: '32px', textTransform: 'none', color: '#94a3b8', textAlign: 'center', fontSize: '14px', fontWeight: 500 }}>
                      Không có dữ liệu tổng hợp tháng.
                    </td>
                  </tr>
                ) : (
                  <>
                    {groupedProjectMonthlySummaries.map((group) => {
                      const groupStock = group.items.reduce((sum, item) => sum + item.totalStock, 0)
                      const groupMonthly = {}
                      const groupMonthlyAmount = {}
                      for (let m = 1; m <= 12; m++) {
                        groupMonthly[`T${m}`] = group.items.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0), 0)
                        groupMonthlyAmount[`T${m}`] = group.items.reduce((sum, item) => sum + (item.monthlyStock[`T${m}`] || 0) * (item.unitPrice || 0), 0)
                      }
                      const groupValue = group.items.reduce((sum, item) => sum + item.totalValue, 0)

                      return (
                        <React.Fragment key={group.id}>
                          <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #cbd5e1' }}>
                            <td colSpan={31} style={{ padding: '10px 14px', fontSize: '13px', fontWeight: 800, color: '#1e3a8a', background: '#f1f5f9', letterSpacing: '0.5px' }}>
                              {group.title}
                            </td>
                          </tr>

                          {group.items.length === 0 ? (
                            <tr>
                              <td colSpan={31} style={{ padding: '12px', color: '#94a3b8', textAlign: 'center', fontStyle: 'italic' }}>
                                Không có dữ liệu cho nhóm này.
                              </td>
                            </tr>
                          ) : (
                            group.items.map((item, idx) => {
                              return (
                                <tr 
                                  key={`${item.projectName}_${item.maSAP}`} 
                                  title="Nhấp đúp để xem chi tiết kho / dự án này"
                                  style={{ 
                                    borderBottom: '1px solid #e2e8f0',
                                    background: '#ffffff',
                                    transition: 'background 0.2s',
                                    cursor: 'pointer'
                                  }}
                                  onDoubleClick={() => {
                                    setLocalProject(item.projectName)
                                    setActiveSubTab('detail')
                                    setSearchTerm('')
                                  }}
                                  onMouseOver={(e) => { e.currentTarget.style.background = '#f8fafc' }}
                                  onMouseOut={(e) => { e.currentTarget.style.background = '#ffffff' }}
                                >
                                  <td style={{ padding: '8px 10px', textAlign: 'center', color: '#64748b', fontWeight: 600, border: '1px solid #cbd5e1' }}>{idx + 1}</td>
                                  <td style={{ padding: '8px 10px', fontWeight: 700, color: '#1e3a8a', border: '1px solid #cbd5e1' }}>{item.projectName}</td>
                                  <td style={{ padding: '8px 10px', textAlign: 'center', fontWeight: 600, color: '#475569', border: '1px solid #cbd5e1' }}>{item.maSAP}</td>
                                  <td style={{ padding: '8px 10px', color: '#1e293b', fontWeight: 500, border: '1px solid #cbd5e1' }}>{item.tenVatTu}</td>
                                  <td style={{ padding: '8px 10px', textAlign: 'center', color: '#64748b', border: '1px solid #cbd5e1' }}>{item.dvt}</td>
                                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 700, color: '#10b981', border: '1px solid #cbd5e1' }}>{item.totalStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</td>
                                  
                                  {/* Unit price */}
                                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 700, color: '#475569', border: '1px solid #cbd5e1' }}>
                                    {item.unitPrice > 0 ? `${item.unitPrice.toLocaleString('vi-VN')} đ` : '-'}
                                  </td>

                                  {/* Quantity months */}
                                  {[...Array(12)].map((_, i) => {
                                    const m = i + 1
                                    const val = item.monthlyStock[`T${m}`] || 0
                                    return (
                                      <td 
                                        key={`qty_${m}`}
                                        style={{ 
                                          padding: '8px 4px', 
                                          textAlign: 'right', 
                                          fontWeight: val !== 0 ? '600' : '400',
                                          color: val > 0 ? '#0f172a' : (val < 0 ? '#ef4444' : '#94a3b8'),
                                          background: 'rgba(219, 234, 254, 0.15)',
                                          border: '1px solid #cbd5e1',
                                          fontSize: '11px'
                                        }}
                                      >
                                        {val !== 0 ? val.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                                      </td>
                                    )
                                  })}

                                  {/* Amount months */}
                                  {[...Array(12)].map((_, i) => {
                                    const m = i + 1
                                    const val = item.monthlyStock[`T${m}`] || 0
                                    const amt = val * (item.unitPrice || 0)
                                    return (
                                      <td 
                                        key={`amt_${m}`}
                                        style={{ 
                                          padding: '8px 4px', 
                                          textAlign: 'right', 
                                          fontWeight: amt !== 0 ? '600' : '400',
                                          color: amt > 0 ? '#047857' : (amt < 0 ? '#ef4444' : '#94a3b8'),
                                          background: 'rgba(209, 250, 229, 0.15)',
                                          border: '1px solid #cbd5e1',
                                          fontSize: '11px'
                                        }}
                                      >
                                        {amt !== 0 ? `${Math.round(amt).toLocaleString('vi-VN')} đ` : '-'}
                                      </td>
                                    )
                                  })}
                                </tr>
                              )
                            })
                          )}

                          <tr style={{ background: '#f8fafc', fontWeight: 700, borderBottom: '2px solid #cbd5e1', borderTop: '1px solid #cbd5e1' }}>
                            <td colSpan="5" style={{ padding: '8px 14px', textTransform: 'none', color: '#1e3a8a', fontSize: '12px', fontWeight: 800, border: '1px solid #cbd5e1' }}>
                              CỘNG {group.id === 'I' ? 'KHO DỰ ÁN (A)' : 'NHẬP TỪ NCC (B)'}
                            </td>
                            <td style={{ padding: '8px 10px', textAlign: 'right', color: '#10b981', border: '1px solid #cbd5e1' }}>{groupStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</td>
                            
                            <td style={{ padding: '8px 10px', textAlign: 'right', color: '#475569', border: '1px solid #cbd5e1' }}>
                              -
                            </td>

                            {[...Array(12)].map((_, i) => {
                              const m = i + 1
                              const val = groupMonthly[`T${m}`] || 0
                              return (
                                <td 
                                  key={`group_qty_${m}`}
                                  style={{ 
                                    padding: '8px 4px', 
                                    textAlign: 'right', 
                                    color: val > 0 ? '#0f172a' : (val < 0 ? '#ef4444' : '#64748b'),
                                    background: 'rgba(219, 234, 254, 0.3)',
                                    border: '1px solid #cbd5e1',
                                    fontSize: '11px'
                                  }}
                                >
                                  {val !== 0 ? val.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                                </td>
                              )
                            })}

                            {[...Array(12)].map((_, i) => {
                              const m = i + 1
                              const val = groupMonthlyAmount[`T${m}`] || 0
                              return (
                                <td 
                                  key={`group_amt_${m}`}
                                  style={{ 
                                    padding: '8px 4px', 
                                    textAlign: 'right', 
                                    color: val > 0 ? '#047857' : (val < 0 ? '#ef4444' : '#64748b'),
                                    background: 'rgba(209, 250, 229, 0.3)',
                                    border: '1px solid #cbd5e1',
                                    fontSize: '11px'
                                  }}
                                >
                                  {val !== 0 ? `${Math.round(val).toLocaleString('vi-VN')} đ` : '-'}
                                </td>
                              )
                            })}
                          </tr>
                        </React.Fragment>
                      )
                    })}
                    
                    <tr style={{ background: '#ecfdf5', fontWeight: 800, borderTop: '2px solid #10b981', borderBottom: '2px solid #10b981' }}>
                      <td colSpan="5" style={{ padding: '10px 10px', textAlign: 'center', color: '#047857', border: '1px solid #cbd5e1' }}>
                        ✓ CHÊNH LỆCH ĐỐI CHIẾU (A - B)
                      </td>
                      <td style={{ padding: '10px 10px', textAlign: 'right', color: '#047857', border: '1px solid #cbd5e1' }}>
                        {reconciliationMonthlyDifference.stockDiff.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                      </td>
                      
                      <td style={{ padding: '10px 10px', textAlign: 'right', color: '#047857', border: '1px solid #cbd5e1' }}>
                        -
                      </td>

                      {[...Array(12)].map((_, i) => {
                        const m = i + 1
                        const val = reconciliationMonthlyDifference.monthlyDiff[`T${m}`] || 0
                        return (
                          <td 
                            key={`diff_qty_${m}`}
                            style={{ 
                              padding: '10px 4px', 
                              textAlign: 'right', 
                              color: val !== 0 ? (val < 0 ? '#ef4444' : '#047857') : '#047857',
                              background: '#ecfdf5',
                              border: '1px solid #cbd5e1',
                              fontSize: '11px'
                            }}
                          >
                            {val !== 0 ? val.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                          </td>
                        )
                      })}
                      
                      {[...Array(12)].map((_, i) => {
                        const m = i + 1
                        const val = reconciliationMonthlyDifference.monthlyAmtDiff[`T${m}`] || 0
                        return (
                          <td 
                            key={`diff_amt_${m}`}
                            style={{ 
                              padding: '10px 4px', 
                              textAlign: 'right', 
                              color: val !== 0 ? (val < 0 ? '#ef4444' : '#047857') : '#047857',
                              background: '#ecfdf5',
                              border: '1px solid #cbd5e1',
                              fontSize: '11px'
                            }}
                          >
                            {val !== 0 ? `${Math.round(val).toLocaleString('vi-VN')} đ` : '-'}
                          </td>
                        )
                      })}
                    </tr>
                  </>
                )}
              </tbody>
            </table>
          ) : activeSubTab === 'material_stats' ? (
            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '12.5px', minWidth: '1350px' }}>
              <thead>
                <tr style={{ background: '#1e3a8a', borderBottom: '2px solid #0f172a' }}>
                  <th style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', width: '75px', minWidth: '75px', textAlign: 'center' }}>STT</th>
                  <th onClick={() => handleMatSort('maSAP')} style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '120px', minWidth: '120px', textAlign: 'center' }}>
                    Mã SAP
                  </th>
                  <th onClick={() => handleMatSort('tenVatTu')} style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '320px', minWidth: '320px' }}>
                    Tên vật tư
                  </th>
                  <th style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', textAlign: 'center', width: '75px', minWidth: '75px' }}>ĐVT</th>
                  
                  <th onClick={() => handleMatSort('totalStock')} style={{ padding: '12px 10px', fontWeight: 700, color: '#facc15', textAlign: 'center', cursor: 'pointer', userSelect: 'none', width: '110px', minWidth: '110px', background: '#1e3a8a', borderLeft: '1px solid rgba(255, 255, 255, 0.25)', borderRight: '1px solid rgba(255, 255, 255, 0.25)' }}>
                    Tổng cộng
                  </th>
                  
                  {/* Dynamic Warehouse Columns with Reduced Width */}
                  {activePhysicalWarehouses.map(pName => (
                    <th 
                      key={pName} 
                      style={{ 
                        padding: '12px 8px', 
                        fontWeight: 700, 
                        color: '#ffffff', 
                        textAlign: 'center', 
                        width: '85px',
                        minWidth: '85px', 
                        background: '#1d4ed8', 
                        borderLeft: '1px solid rgba(255, 255, 255, 0.25)', 
                        borderRight: '1px solid rgba(255, 255, 255, 0.25)',
                        fontSize: '11px',
                        lineHeight: '1.2'
                      }}
                    >
                      {pName}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {sortedMaterialStatsData.length === 0 ? (
                  <tr>
                    <td colSpan={5 + activePhysicalWarehouses.length} style={{ padding: '32px', textTransform: 'none', color: '#94a3b8', textAlign: 'center', fontSize: '14px', fontWeight: 500 }}>
                      Không tìm thấy thống kê vật tư nào khớp với từ khóa tìm kiếm.
                    </td>
                  </tr>
                ) : (
                  <>
                    {sortedMaterialStatsData.map((item, idx) => {
                      return (
                        <tr 
                          key={item.maSAP} 
                          title="Nhấp đúp để lọc chi tiết vật tư này"
                          style={{ 
                            borderBottom: '1px solid #e2e8f0',
                            background: '#ffffff',
                            transition: 'background 0.2s',
                            cursor: 'pointer'
                          }}
                          onDoubleClick={() => {
                            setSearchTerm(item.maSAP)
                            setActiveSubTab('detail')
                          }}
                          onMouseOver={(e) => { e.currentTarget.style.background = '#f8fafc' }}
                          onMouseOut={(e) => { e.currentTarget.style.background = '#ffffff' }}
                        >
                          <td style={{ padding: '10px', textAlign: 'center', color: '#64748b', fontWeight: 600 }}>{idx + 1}</td>
                          <td style={{ padding: '10px', fontWeight: 700, color: '#1e293b', textAlign: 'center' }}>{item.maSAP}</td>
                          <td style={{ padding: '10px', fontWeight: 600, color: '#1e3a8a' }}>{item.tenVatTu}</td>
                          <td style={{ padding: '10px', textAlign: 'center', color: '#475569' }}>{item.dvt}</td>
                          
                          <td style={{ padding: '10px', textAlign: 'right', fontWeight: 700, background: 'rgba(219, 234, 254, 0.25)', borderLeft: '1px solid #cbd5e1', borderRight: '1px solid #cbd5e1', color: item.totalStock > 0 ? '#1e3a8a' : (item.totalStock < 0 ? '#ef4444' : '#64748b') }}>
                            {item.totalStock !== 0 ? item.totalStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                          </td>
                          
                          {/* Warehouse values */}
                          {activePhysicalWarehouses.map(pName => {
                            const val = (projectStocksByMaterial[item.maSAP] && projectStocksByMaterial[item.maSAP][pName]) || 0
                            return (
                              <td 
                                key={pName}
                                style={{ 
                                  padding: '10px 8px', 
                                  textAlign: 'right', 
                                  fontWeight: 600, 
                                  background: 'rgba(219, 234, 254, 0.15)', 
                                  borderLeft: '1px solid #cbd5e1', 
                                  borderRight: '1px solid #cbd5e1', 
                                  color: val > 0 ? '#1e3a8a' : (val < 0 ? '#ef4444' : '#94a3b8') 
                                }}
                              >
                                {val !== 0 ? val.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                              </td>
                            )
                          })}
                        </tr>
                      )
                    })}
                    
                    {/* Grand Total Row */}
                    <tr style={{ background: '#f8fafc', fontWeight: 700, borderTop: '2px solid #cbd5e1', borderBottom: '2px solid #cbd5e1' }}>
                      <td colSpan="4" style={{ padding: '12px 14px', textTransform: 'none', color: '#1e3a8a', fontSize: '12.5px', fontWeight: 800, textAlign: 'right' }}>
                        TỔNG CỘNG THỐNG KÊ VẬT TƯ:
                      </td>
                      
                      <td style={{ padding: '10px', textAlign: 'right', background: 'rgba(219, 234, 254, 0.4)', borderLeft: '1px solid #cbd5e1', borderRight: '1px solid #cbd5e1', color: '#1e3a8a' }}>
                        {sortedMaterialStatsData.reduce((sum, item) => sum + item.totalStock, 0).toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                      </td>
                      
                      {activePhysicalWarehouses.map(pName => {
                        const totalWarehouseStock = sortedMaterialStatsData.reduce((sum, item) => {
                          const val = (projectStocksByMaterial[item.maSAP] && projectStocksByMaterial[item.maSAP][pName]) || 0
                          return sum + val
                        }, 0)
                        return (
                          <td 
                            key={pName}
                            style={{ 
                              padding: '10px 8px', 
                              textAlign: 'right', 
                              background: 'rgba(219, 234, 254, 0.3)', 
                              borderLeft: '1px solid #cbd5e1', 
                              borderRight: '1px solid #cbd5e1', 
                              color: '#1e3a8a' 
                            }}
                          >
                            {totalWarehouseStock !== 0 ? totalWarehouseStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                          </td>
                        )
                      })}
                    </tr>
                  </>
                )}
              </tbody>
            </table>
          ) : (
            <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '12.5px', minWidth: '1300px' }}>
              <thead>
                <tr style={{ background: '#1e3a8a', borderBottom: '2px solid #0f172a' }}>
                  <th style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', width: '60px', textAlign: 'center' }}>STT</th>
                  <th onClick={() => handleSort('maSAP')} style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', cursor: 'pointer', userSelect: 'none', width: '90px', textAlign: 'center' }}>
                    Mã SAP {sortField === 'maSAP' && (sortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                  <th style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', width: '90px', textAlign: 'center' }}>Mã vật tư</th>
                  <th style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', width: '180px', textAlign: 'center' }}>Tên vật tư</th>
                  <th style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', textAlign: 'center', width: '55px' }}>ĐVT</th>
                  <th style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', width: '120px', textAlign: 'center' }}>Thông số kỹ thuật</th>
                  
                  {/* Current Stock */}
                  <th onClick={() => handleSort('stock')} style={{ padding: '12px 10px', fontWeight: 700, color: '#facc15', textAlign: 'center', cursor: 'pointer', userSelect: 'none', width: '95px' }}>
                    Tồn hiện tại {sortField === 'stock' && (sortDirection === 'asc' ? '▲' : '▼')}
                  </th>

                  {/* 12 Months columns */}
                  {[...Array(12)].map((_, i) => {
                    const m = i + 1
                    const key = `T${m}`
                    return (
                      <th 
                        key={key}
                        onClick={() => handleSort(key)}
                        style={{ 
                          padding: '12px 6px', 
                          fontWeight: 700, 
                          color: '#ffffff', // High-contrast text
                          textAlign: 'center', 
                          cursor: 'pointer', 
                          userSelect: 'none',
                          width: '75px',
                          background: '#1d4ed8', // Distinct background color for T1-T12 headers
                          borderLeft: '1px solid rgba(255, 255, 255, 0.25)', // Vertical borders for headers
                          borderRight: '1px solid rgba(255, 255, 255, 0.25)'
                        }}
                      >
                        T{m} {sortField === key && (sortDirection === 'asc' ? '▲' : '▼')}
                      </th>
                    )
                  })}

                  <th onClick={() => handleSort('estimatedUnitPrice')} style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', textAlign: 'center', cursor: 'pointer', userSelect: 'none', width: '100px' }}>
                    Đơn giá 1 ngày {sortField === 'estimatedUnitPrice' && (sortDirection === 'asc' ? '▲' : '▼')}
                  </th>
                  <th style={{ padding: '12px 10px', fontWeight: 700, color: '#ffffff', textAlign: 'center', width: '110px' }}>Trạng thái sử dụng</th>
                </tr>
              </thead>
              <tbody>
                {paginatedData.length === 0 ? (
                  <tr>
                    <td colSpan="21" style={{ padding: '32px', textTransform: 'none', color: '#94a3b8', textAlign: 'center', fontSize: '14px', fontWeight: 500 }}>
                      Không tìm thấy tài sản khấu hao nào khớp với điều kiện lọc.
                    </td>
                  </tr>
                ) : (
                  paginatedData.map((item, idx) => {
                    const globalIdx = (currentPage - 1) * pageSize + idx + 1
                    const isOver30Days = item.unusedStatus === 'Chưa sử dụng (> 30 ngày)'
                    
                    return (
                      <tr 
                        key={item.maSAP} 
                        style={{ 
                          borderBottom: '1px solid #e2e8f0',
                          background: isOver30Days ? '#fffbeb' : '#ffffff',
                          transition: 'background 0.2s'
                        }}
                        onMouseOver={(e) => { e.currentTarget.style.background = isOver30Days ? '#fef3c7' : '#f8fafc' }}
                        onMouseOut={(e) => { e.currentTarget.style.background = isOver30Days ? '#fffbeb' : '#ffffff' }}
                      >
                        <td style={{ padding: '10px', textAlign: 'center', color: '#64748b', fontWeight: 600 }}>{globalIdx}</td>
                        <td style={{ padding: '10px', fontWeight: 700, color: '#1e293b' }}>{item.maSAP}</td>
                        <td style={{ padding: '10px', color: '#475569', fontFamily: 'monospace' }}>{item.maVatTu || '-'}</td>
                        <td style={{ padding: '10px', fontWeight: 600, color: '#1e3a8a' }}>{item.tenVatTu}</td>
                        <td style={{ padding: '10px', textAlign: 'center', color: '#475569' }}>{item.dvt}</td>
                        <td style={{ padding: '10px', color: '#64748b', fontSize: '12px' }}>{item.thongSoKyThuat || '-'}</td>
                        
                        {/* Current Stock */}
                        <td style={{ padding: '10px', textAlign: 'right', fontWeight: 700, color: item.stock > 0 ? '#10b981' : (item.stock < 0 ? '#ef4444' : '#64748b') }}>
                          {item.stock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                        </td>

                        {/* 12 Months dynamic balances */}
                        {[...Array(12)].map((_, i) => {
                          const m = i + 1
                          const val = item.monthlyStock[`T${m}`]
                          return (
                            <td 
                              key={m}
                              style={{ 
                                padding: '10px 6px', 
                                textAlign: 'right', 
                                fontWeight: val !== 0 ? '600' : '400',
                                color: val > 0 ? '#0f172a' : (val < 0 ? '#ef4444' : '#94a3b8'),
                                background: 'rgba(219, 234, 254, 0.25)', // light blue tint for monthly cols
                                borderLeft: '1px solid #cbd5e1', // Vertical borders between columns
                                borderRight: '1px solid #cbd5e1'
                              }}
                            >
                              {val !== 0 ? val.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '-'}
                            </td>
                          )
                        })}

                        <td style={{ padding: '10px', textAlign: 'right', color: '#475569', fontFamily: 'monospace' }}>
                          {item.estimatedUnitPrice > 0 ? `${item.estimatedUnitPrice.toLocaleString('vi-VN')} đ` : '-'}
                        </td>
                        
                        <td style={{ padding: '10px', textAlign: 'center' }}>
                          {isOver30Days ? (
                            <span style={{ background: '#fef3c7', color: '#92400e', border: '1px solid #fde68a', padding: '2px 8px', borderRadius: '12px', fontSize: '10.5px', fontWeight: 700 }}>
                              Chưa dùng &gt; 30 ngày
                            </span>
                          ) : (
                            <span style={{ background: '#dcfce7', color: '#15803d', border: '1px solid #bbf7d0', padding: '2px 8px', borderRadius: '12px', fontSize: '10.5px', fontWeight: 700 }}>
                              Đang sử dụng
                            </span>
                          )}
                        </td>
                      </tr>
                    )
                  })
                )}
              </tbody>
            </table>
          )}
        </div>

        {/* Pagination & Display size controls - Only show on detail tab */}
        {activeSubTab === 'detail' && sortedData.length > 0 && (
          <div style={{
            padding: '12px 16px',
            borderTop: '1px solid #e2e8f0',
            background: '#f8fafc',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flexWrap: 'wrap',
            gap: '12px'
          }}>
            <div style={{ fontSize: '13px', color: '#64748b' }}>
              Hiển thị từ <strong>{Math.min(sortedData.length, (currentPage - 1) * pageSize + 1)}</strong> đến{' '}
              <strong>{Math.min(sortedData.length, currentPage * pageSize)}</strong> trong tổng số{' '}
              <strong>{sortedData.length}</strong> dòng kết quả.
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '16px' }}>
              {/* Page size selector */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <span style={{ fontSize: '13px', color: '#64748b' }}>Số dòng/trang:</span>
                <select
                  value={pageSize}
                  onChange={(e) => {
                    setPageSize(parseInt(e.target.value, 10))
                    setCurrentPage(1)
                  }}
                  style={{
                    padding: '4px 8px',
                    border: '1px solid #cbd5e1',
                    borderRadius: '4px',
                    fontSize: '12.5px',
                    fontWeight: 500,
                    outline: 'none',
                    background: '#ffffff'
                  }}
                >
                  <option value={20}>20</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                  <option value={200}>200</option>
                </select>
              </div>

              {/* Navigation buttons */}
              <div style={{ display: 'flex', gap: '4px' }}>
                <button
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage(1)}
                  style={{
                    padding: '6px 10px',
                    border: '1px solid #cbd5e1',
                    borderRadius: '6px',
                    background: '#ffffff',
                    fontSize: '12px',
                    fontWeight: 600,
                    cursor: currentPage === 1 ? 'not-allowed' : 'pointer',
                    opacity: currentPage === 1 ? 0.5 : 1
                  }}
                >
                  « Đầu
                </button>
                <button
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage(prev => prev - 1)}
                  style={{
                    padding: '6px 10px',
                    border: '1px solid #cbd5e1',
                    borderRadius: '6px',
                    background: '#ffffff',
                    fontSize: '12px',
                    fontWeight: 600,
                    cursor: currentPage === 1 ? 'not-allowed' : 'pointer',
                    opacity: currentPage === 1 ? 0.5 : 1
                  }}
                >
                  ‹ Trước
                </button>
                <span style={{
                  padding: '6px 12px',
                  fontSize: '13px',
                  fontWeight: 700,
                  color: '#1e3a8a',
                  background: '#eff6ff',
                  borderRadius: '6px',
                  border: '1px solid #bfdbfe'
                }}>
                  {currentPage} / {totalPages}
                </span>
                <button
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage(prev => prev + 1)}
                  style={{
                    padding: '6px 10px',
                    border: '1px solid #cbd5e1',
                    borderRadius: '6px',
                    background: '#ffffff',
                    fontSize: '12px',
                    fontWeight: 600,
                    cursor: currentPage === totalPages ? 'not-allowed' : 'pointer',
                    opacity: currentPage === totalPages ? 0.5 : 1
                  }}
                >
                  Sau ›
                </button>
                <button
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage(totalPages)}
                  style={{
                    padding: '6px 10px',
                    border: '1px solid #cbd5e1',
                    borderRadius: '6px',
                    background: '#ffffff',
                    fontSize: '12px',
                    fontWeight: 600,
                    cursor: currentPage === totalPages ? 'not-allowed' : 'pointer',
                    opacity: currentPage === totalPages ? 0.5 : 1
                  }}
                >
                  Cuối »
                </button>
              </div>
            </div>
          </div>
        )}
      </div>

    </div>
  )
}
