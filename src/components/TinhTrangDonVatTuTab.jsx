import React, { useState, useMemo, useCallback, useRef } from 'react'
import * as XLSXStyleRaw from 'xlsx-js-style'
import {
  FileSpreadsheet,
  Upload,
  Search,
  Filter,
  Download,
  Trash2,
  RefreshCw,
  Plus,
  CheckCircle2,
  Clock,
  AlertCircle,
  ClipboardList,
  Check,
  ChevronLeft,
  ChevronRight,
  ChevronsLeft,
  ChevronsRight,
  Database,
  ArrowUpDown,
  FileCheck,
  Building2,
  User,
  Calendar,
  Layers
} from 'lucide-react'
import {
  COLS_TINH_TRANG_DON,
  parseTinhTrangDonRows,
  isApprovedStatus,
  isPendingStatus,
  isRejectedStatus
} from '../constants.js'

const XLSXStyle = XLSXStyleRaw.default || XLSXStyleRaw

// Helper to parse date string or excel serial to Date object
function parseDateForSort(dateVal) {
  if (!dateVal) return null
  const str = String(dateVal).trim()
  if (!str) return null

  // DD/MM/YYYY or DD-MM-YYYY
  const dmyMatch = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})/)
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10)
    const month = parseInt(dmyMatch[2], 10) - 1
    const year = parseInt(dmyMatch[3], 10)
    return new Date(year, month, day)
  }

  // YYYY/MM/DD or YYYY-MM-DD
  const ymdMatch = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})/)
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10)
    const month = parseInt(ymdMatch[2], 10) - 1
    const day = parseInt(ymdMatch[3], 10)
    return new Date(year, month, day)
  }

  const parsed = new Date(str)
  if (!isNaN(parsed.getTime())) return parsed
  return null
}

export default function TinhTrangDonVatTuTab({
  rows = [],
  setRows,
  fileName = '',
  setFileName,
  selectedProject = '',
  bchAliasMap = {}
}) {
  const [loading, setLoading] = useState(false)
  const [search, setSearch] = useState('')
  const [trangThaiFilter, setTrangThaiFilter] = useState('')
  const [donViGiaoFilter, setDonViGiaoFilter] = useState('')
  const [donViNhanFilter, setDonViNhanFilter] = useState('')
  const [dongBoSAPFilter, setDongBoSAPFilter] = useState('')
  const [startDate, setStartDate] = useState('')
  const [endDate, setEndDate] = useState('')
  const [currentPage, setCurrentPage] = useState(1)
  const [pageSize, setPageSize] = useState(50)
  const [sortField, setSortField] = useState('ngayXuat')
  const [sortAsc, setSortAsc] = useState(false)
  const [isDragOver, setIsDragOver] = useState(false)

  const fileInputRef = useRef(null)
  const appendFileInputRef = useRef(null)

  // Parse Excel file handler
  const processFiles = useCallback((fileList, isAppend = false) => {
    if (!fileList || fileList.length === 0) return
    setLoading(true)

    const files = Array.from(fileList)
    let processedFiles = 0
    let allNewRows = []
    let fileNames = []

    files.forEach(file => {
      const reader = new FileReader()
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result)
          const workbook = XLSXStyle.read(data, { type: 'array' })
          const firstSheetName = workbook.SheetNames[0]
          const worksheet = workbook.Sheets[firstSheetName]
          const sheetJson = XLSXStyle.utils.sheet_to_json(worksheet, { header: 1 })
          
          const parsed = parseTinhTrangDonRows(sheetJson, bchAliasMap)
          allNewRows = allNewRows.concat(parsed)
          fileNames.push(file.name)
        } catch (err) {
          console.error(`Lỗi đọc file ${file.name}:`, err)
        } finally {
          processedFiles++
          if (processedFiles === files.length) {
            const combinedName = fileNames.join(' + ')
            if (isAppend) {
              setRows(prev => {
                const startIdx = prev.length
                const indexedNewRows = allNewRows.map((r, i) => ({
                  ...r,
                  id: startIdx + i + 1,
                  stt: r.stt || String(startIdx + i + 1)
                }))
                return [...prev, ...indexedNewRows]
              })
              setFileName(prev => prev ? `${prev} + ${combinedName}` : combinedName)
            } else {
              const indexedRows = allNewRows.map((r, i) => ({
                ...r,
                id: i + 1,
                stt: r.stt || String(i + 1)
              }))
              setRows(indexedRows)
              setFileName(combinedName)
            }
            setLoading(false)
            setCurrentPage(1)
          }
        }
      }
      reader.readAsArrayBuffer(file)
    })
  }, [bchAliasMap, setRows, setFileName])

  // DropZone handlers
  const handleDrop = (e) => {
    e.preventDefault()
    setIsDragOver(false)
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      processFiles(e.dataTransfer.files, false)
    }
  }

  const handleClearData = () => {
    if (window.confirm('Bạn có chắc chắn muốn xóa toàn bộ dữ liệu sheet Tình trạng đơn vật tư không?')) {
      setRows([])
      setFileName('')
      setCurrentPage(1)
    }
  }

  // Filter options
  const trangThaiOptions = useMemo(() => {
    return [...new Set(rows.map(r => r.trangThai).filter(Boolean))].sort()
  }, [rows])

  const donViGiaoOptions = useMemo(() => {
    return [...new Set(rows.map(r => r.donViGiao).filter(Boolean))].sort()
  }, [rows])

  const donViNhanOptions = useMemo(() => {
    return [...new Set(rows.map(r => r.donViNhan).filter(Boolean))].sort()
  }, [rows])

  const dongBoSAPOptions = useMemo(() => {
    return [...new Set(rows.map(r => r.dongBoSAP).filter(Boolean))].sort()
  }, [rows])

  // Filtered & Sorted Rows
  const filteredRows = useMemo(() => {
    let result = rows

    // Project filter (from header)
    if (selectedProject) {
      const proj = selectedProject.trim().toLowerCase()
      result = result.filter(r => {
        const g = String(r.donViGiao || '').trim().toLowerCase()
        const n = String(r.donViNhan || '').trim().toLowerCase()
        return g === proj || n === proj
      })
    }

    // Search query
    if (search) {
      const q = search.toLowerCase().trim()
      result = result.filter(r => {
        return (
          String(r.soDon || '').toLowerCase().includes(q) ||
          String(r.donViGiao || '').toLowerCase().includes(q) ||
          String(r.nguoiGiao || '').toLowerCase().includes(q) ||
          String(r.donViNhan || '').toLowerCase().includes(q) ||
          String(r.nguoiNhan || '').toLowerCase().includes(q) ||
          String(r.cbPheDuyet1 || '').toLowerCase().includes(q) ||
          String(r.cbPheDuyet2 || '').toLowerCase().includes(q) ||
          String(r.cbPheDuyet3 || '').toLowerCase().includes(q) ||
          String(r.cbPheDuyet4 || '').toLowerCase().includes(q) ||
          String(r.trangThai || '').toLowerCase().includes(q) ||
          String(r.soDonSAP || '').toLowerCase().includes(q) ||
          String(r.dongBoSAP || '').toLowerCase().includes(q)
        )
      })
    }

    // Dropdown filters
    if (trangThaiFilter) {
      result = result.filter(r => r.trangThai === trangThaiFilter)
    }
    if (donViGiaoFilter) {
      result = result.filter(r => r.donViGiao === donViGiaoFilter)
    }
    if (donViNhanFilter) {
      result = result.filter(r => r.donViNhan === donViNhanFilter)
    }
    if (dongBoSAPFilter) {
      result = result.filter(r => r.dongBoSAP === dongBoSAPFilter)
    }

    // Date range filter
    if (startDate || endDate) {
      const start = startDate ? new Date(startDate) : null
      if (start) start.setHours(0, 0, 0, 0)
      const end = endDate ? new Date(endDate) : null
      if (end) end.setHours(23, 59, 59, 999)

      result = result.filter(r => {
        const rowDate = parseDateForSort(r.ngayXuat)
        if (!rowDate) return false
        if (start && rowDate < start) return false
        if (end && rowDate > end) return false
        return true
      })
    }

    // Sorting
    return [...result].sort((a, b) => {
      let valA = a[sortField] || ''
      let valB = b[sortField] || ''

      if (sortField === 'ngayXuat') {
        const dateA = parseDateForSort(valA)
        const dateB = parseDateForSort(valB)
        if (!dateA && !dateB) return 0
        if (!dateA) return 1
        if (!dateB) return -1
        return sortAsc ? dateA.getTime() - dateB.getTime() : dateB.getTime() - dateA.getTime()
      }

      if (sortField === 'stt' || sortField === 'id') {
        const numA = Number(valA) || 0
        const numB = Number(valB) || 0
        return sortAsc ? numA - numB : numB - numA
      }

      const strComp = String(valA).localeCompare(String(valB), 'vi', { sensitivity: 'base' })
      return sortAsc ? strComp : -strComp
    })
  }, [
    rows,
    selectedProject,
    search,
    trangThaiFilter,
    donViGiaoFilter,
    donViNhanFilter,
    dongBoSAPFilter,
    startDate,
    endDate,
    sortField,
    sortAsc
  ])

  // Dynamic breakdown of exact statuses in dataset
  const dynamicStatusList = useMemo(() => {
    const counts = {}
    rows.forEach(r => {
      const st = (r.trangThai || '').trim()
      if (st && st !== '—' && st !== '-') {
        counts[st] = (counts[st] || 0) + 1
      }
    })

    const statusOrder = [
      'chưa xác nhận',
      'chưa phê duyệt 1',
      'chưa phê duyệt 2',
      'chưa phê duyệt 3',
      'chưa phê duyệt 4',
      'chờ phê duyệt',
      'đang xử lý',
      'đã phê duyệt',
      'hoàn thành',
      'từ chối',
      'đã hủy'
    ]

    const keys = Object.keys(counts).sort((a, b) => {
      const aLower = a.toLowerCase()
      const bLower = b.toLowerCase()
      const idxA = statusOrder.findIndex(item => item === aLower)
      const idxB = statusOrder.findIndex(item => item === bLower)
      if (idxA !== -1 && idxB !== -1) return idxA - idxB
      if (idxA !== -1) return -1
      if (idxB !== -1) return 1
      return counts[b] - counts[a]
    })

    return keys.map(statusName => {
      const count = counts[statusName]
      const s = statusName.toLowerCase().trim()
      let bg = '#ffffff'
      let color = '#334155'
      let iconBg = '#f1f5f9'
      let iconColor = '#64748b'
      let border = '#e2e8f0'
      let activeBorder = '#94a3b8'
      let IconComponent = Clock

      if (s.includes('chưa xác nhận') || s.includes('chờ xác nhận')) {
        color = '#c2410c'
        iconBg = '#fff7ed'
        iconColor = '#ea580c'
        border = '#ffedd5'
        activeBorder = '#ea580c'
        IconComponent = Clock
      } else if (s.includes('chưa') || s.includes('chờ') || s.includes('pending')) {
        color = '#b45309'
        iconBg = '#fffbeb'
        iconColor = '#d97706'
        border = '#fef3c7'
        activeBorder = '#d97706'
        IconComponent = Clock
      } else if (s.includes('từ chối') || s.includes('hủy') || s.includes('reject') || s.includes('không')) {
        color = '#b91c1c'
        iconBg = '#fef2f2'
        iconColor = '#dc2626'
        border = '#fee2e2'
        activeBorder = '#dc2626'
        IconComponent = AlertCircle
      } else if (s.includes('đã') || s.includes('phê duyệt') || s.includes('hoàn thành') || s.includes('duyệt') || s.includes('approved')) {
        color = '#047857'
        iconBg = '#ecfdf5'
        iconColor = '#059669'
        border = '#d1fae5'
        activeBorder = '#059669'
        IconComponent = CheckCircle2
      }

      return {
        name: statusName,
        count,
        bg,
        color,
        iconBg,
        iconColor,
        border,
        activeBorder,
        IconComponent
      }
    })
  }, [rows])

  // KPIs
  const stats = useMemo(() => {
    const total = rows.length
    let daPheDuyet = 0
    let choPheDuyet = 0
    let tuChoi = 0
    let dongBoThanhCong = 0
    let chuaDongBo = 0

    rows.forEach(r => {
      const tt = (r.trangThai || '').toLowerCase().trim()
      if (!tt || tt === '—' || tt === '-') {
        // empty status
        return
      }

      if (tt.includes('chờ') || tt.includes('chưa') || tt.includes('pending')) {
        choPheDuyet++
      } else if (tt.includes('từ chối') || tt.includes('hủy') || tt.includes('reject') || tt.includes('không')) {
        tuChoi++
      } else if (tt.includes('đã') || tt.includes('phê duyệt') || tt.includes('hoàn thành') || tt.includes('duyệt') || tt.includes('approved')) {
        daPheDuyet++
      } else {
        choPheDuyet++
      }

      const db = (r.dongBoSAP || '').toLowerCase().trim()
      if (db.includes('đã') || db.includes('thành công') || db.includes('success') || db.includes('ok') || db.includes('đồng bộ')) {
        dongBoThanhCong++
      } else if (db && db !== '—' && db !== '-') {
        chuaDongBo++
      }
    })

    return {
      total,
      daPheDuyet,
      choPheDuyet,
      tuChoi,
      dongBoThanhCong,
      chuaDongBo
    }
  }, [rows])

  // Pagination
  const totalPages = pageSize === 0 ? 1 : Math.ceil(filteredRows.length / pageSize)
  const paginatedRows = useMemo(() => {
    if (pageSize === 0) return filteredRows
    const start = (currentPage - 1) * pageSize
    return filteredRows.slice(start, start + pageSize)
  }, [filteredRows, currentPage, pageSize])

  const handleSort = (field) => {
    if (sortField === field) {
      setSortAsc(prev => !prev)
    } else {
      setSortField(field)
      setSortAsc(true)
    }
  }

  // Export Excel with multi-sheet comprehensive report
  const handleExportExcel = () => {
    if (filteredRows.length === 0) {
      alert('Không có dữ liệu để xuất Excel.')
      return
    }

    const now = new Date()
    const wb = XLSXStyle.utils.book_new()

    // Helper for Excel column letter
    const getExcelColName = (colIdx) => {
      let temp = ''
      let letter = ''
      let c = colIdx + 1
      while (c > 0) {
        temp = (c - 1) % 26
        letter = String.fromCharCode(65 + temp) + letter
        c = Math.floor((c - temp - 1) / 26)
      }
      return letter
    }

    // Helper to sanitize & normalize person names for exact 1-to-1 match
    const cleanPersonName = (name, fallback = 'Chưa phân công người duyệt cụ thể') => {
      if (!name) return fallback
      let s = String(name).normalize('NFC').replace(/\s+/g, ' ').trim()
      if (!s || s === '—' || s === '-' || s === 'null' || s === 'undefined' || s === 'Chưa rõ người nhận' || s === 'Chưa phân công CB1' || s === 'Chưa phân công CB2' || s === 'Chưa phân công CB3' || s === 'Chưa phân công CB4') {
        return fallback
      }
      // Remove trailing slashes, colons, spaces or separators (e.g. "Trần Vĩnh Hùng /" -> "Trần Vĩnh Hùng")
      s = s.replace(/[\s\/:|-]+$/, '').trim()
      // Remove Excel formula wildcard triggers
      s = s.replace(/[*?~]/g, '').trim()
      return s || fallback
    }

    // Helper to analyze timeline & delay for each row
    const analyzeRowTimeline = (row) => {
      const rowDate = parseDateForSort(row.ngayXuat)
      let daysElapsed = 0
      if (rowDate) {
        const rTime = new Date(rowDate)
        rTime.setHours(0, 0, 0, 0)
        const nTime = new Date(now)
        nTime.setHours(0, 0, 0, 0)
        const diff = Math.floor((nTime.getTime() - rTime.getTime()) / (1000 * 60 * 60 * 24))
        daysElapsed = diff > 0 ? diff : 0
      }

      const tt = (row.trangThai || '').toLowerCase().trim()
      let statusGroup = 'other'
      let pendingRole = '—'
      let pendingPerson = '—'
      let isCompleted = false
      let isRejected = false
      let isPending = false
      let delayDays = 0
      let timelineStatus = ''
      let warningLevel = 'Đã hoàn tất'

      if (tt.includes('đã') || tt.includes('hoàn thành') || (tt.includes('phê duyệt') && !tt.includes('chưa') && !tt.includes('chờ'))) {
        statusGroup = 'completed'
        isCompleted = true
        timelineStatus = 'Đã hoàn thành phê duyệt'
        pendingRole = 'Đã hoàn tất'
        pendingPerson = '—'
        warningLevel = 'Đã hoàn tất'
      } else if (tt.includes('từ chối') || tt.includes('hủy') || tt.includes('reject') || tt.includes('không')) {
        statusGroup = 'rejected'
        isRejected = true
        timelineStatus = 'Đã từ chối / Hủy'
        pendingRole = 'Từ chối / Hủy'
        pendingPerson = '—'
        warningLevel = 'Đã hủy / Từ chối'
      } else if (tt.includes('chưa xác nhận') || tt.includes('chờ xác nhận')) {
        statusGroup = 'unconfirmed'
        isPending = true
        delayDays = daysElapsed
        pendingRole = 'Người nhận'
        pendingPerson = cleanPersonName(row.nguoiNhan, 'Chưa rõ người nhận')
        timelineStatus = daysElapsed > 0 
          ? `Chờ Người nhận (${pendingPerson}) xác nhận [Chậm ${daysElapsed} ngày]` 
          : `Chờ Người nhận (${pendingPerson}) xác nhận [Trong ngày]`
      } else if (tt.includes('chưa phê duyệt 1') || tt.includes('chờ duyệt 1')) {
        statusGroup = 'pending_cb1'
        isPending = true
        delayDays = daysElapsed
        pendingRole = 'CB phê duyệt 1'
        pendingPerson = cleanPersonName(row.cbPheDuyet1, 'Chưa phân công CB1')
        timelineStatus = daysElapsed > 0 
          ? `Chờ CB duyệt 1 (${pendingPerson}) phê duyệt [Chậm ${daysElapsed} ngày]` 
          : `Chờ CB duyệt 1 (${pendingPerson}) phê duyệt [Trong ngày]`
      } else if (tt.includes('chưa phê duyệt 2') || tt.includes('chờ duyệt 2')) {
        statusGroup = 'pending_cb2'
        isPending = true
        delayDays = daysElapsed
        pendingRole = 'CB phê duyệt 2'
        pendingPerson = cleanPersonName(row.cbPheDuyet2, 'Chưa phân công CB2')
        timelineStatus = daysElapsed > 0 
          ? `Chờ CB duyệt 2 (${pendingPerson}) phê duyệt [Chậm ${daysElapsed} ngày]` 
          : `Chờ CB duyệt 2 (${pendingPerson}) phê duyệt [Trong ngày]`
      } else if (tt.includes('chưa phê duyệt 3') || tt.includes('chờ duyệt 3')) {
        statusGroup = 'pending_cb3'
        isPending = true
        delayDays = daysElapsed
        pendingRole = 'CB phê duyệt 3'
        pendingPerson = cleanPersonName(row.cbPheDuyet3, 'Chưa phân công CB3')
        timelineStatus = daysElapsed > 0 
          ? `Chờ CB duyệt 3 (${pendingPerson}) phê duyệt [Chậm ${daysElapsed} ngày]` 
          : `Chờ CB duyệt 3 (${pendingPerson}) phê duyệt [Trong ngày]`
      } else if (tt.includes('chưa phê duyệt 4') || tt.includes('chờ duyệt 4')) {
        statusGroup = 'pending_cb4'
        isPending = true
        delayDays = daysElapsed
        pendingRole = 'CB phê duyệt 4'
        pendingPerson = cleanPersonName(row.cbPheDuyet4, 'Chưa phân công CB4')
        timelineStatus = daysElapsed > 0 
          ? `Chờ CB duyệt 4 (${pendingPerson}) phê duyệt [Chậm ${daysElapsed} ngày]` 
          : `Chờ CB duyệt 4 (${pendingPerson}) phê duyệt [Trong ngày]`
      } else {
        statusGroup = 'pending_other'
        isPending = true
        delayDays = daysElapsed
        pendingRole = 'Đang xử lý'
        pendingPerson = cleanPersonName(row.cbPheDuyet1 || row.nguoiNhan, 'Chưa phân công người duyệt cụ thể')
        timelineStatus = daysElapsed > 0 
          ? `Đang chờ xử lý [Chậm ${daysElapsed} ngày]` 
          : `Đang chờ xử lý [Trong ngày]`
      }

      if (isPending) {
        if (delayDays > 2) {
          warningLevel = `Quá hạn (> 2 ngày) - Trễ ${delayDays} ngày`
        } else if (delayDays >= 1) {
          warningLevel = `Chậm ${delayDays} ngày`
        } else {
          warningLevel = 'Đúng hạn (Trong ngày)'
        }
      }

      return {
        daysElapsed,
        statusGroup,
        pendingRole,
        pendingPerson,
        isCompleted,
        isRejected,
        isPending,
        delayDays,
        timelineStatus,
        warningLevel
      }
    }

    // Common styling presets
    const borderAll = {
      top: { style: 'thin', color: { rgb: 'CBD5E1' } },
      bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
      left: { style: 'thin', color: { rgb: 'CBD5E1' } },
      right: { style: 'thin', color: { rgb: 'CBD5E1' } }
    }

    const headerStyle = {
      fill: { patternType: 'solid', fgColor: { rgb: '0A3D73' } },
      font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: '08325E' } },
        bottom: { style: 'medium', color: { rgb: '08325E' } },
        left: { style: 'thin', color: { rgb: '08325E' } },
        right: { style: 'thin', color: { rgb: '08325E' } }
      }
    }

    /* =========================================================================
       SHEET 1: DỮ LIỆU CHI TIẾT CỦA BẢNG TÌNH TRẠNG ĐƠN VẬT TƯ
       ========================================================================= */
    const wsDetail = {}
    const detailCols = [
      { key: 'stt', label: 'STT', width: 50, align: 'center' },
      { key: 'soDon', label: 'Số đơn', width: 230, align: 'left' },
      { key: 'ngayXuat', label: 'Ngày xuất', width: 105, align: 'center' },
      { key: 'donViGiao', label: 'Đơn vị giao', width: 250, align: 'left' },
      { key: 'nguoiGiao', label: 'Người giao', width: 150, align: 'left' },
      { key: 'donViNhan', label: 'Đơn vị nhận', width: 250, align: 'left' },
      { key: 'nguoiNhan', label: 'Người nhận', width: 150, align: 'left' },
      { key: 'cbPheDuyet1', label: 'Cán bộ phê duyệt 1', width: 170, align: 'left' },
      { key: 'cbPheDuyet2', label: 'Cán bộ phê duyệt 2', width: 170, align: 'left' },
      { key: 'cbPheDuyet3', label: 'Cán bộ phê duyệt 3', width: 150, align: 'left' },
      { key: 'cbPheDuyet4', label: 'Cán bộ phê duyệt 4', width: 150, align: 'left' },
      { key: 'trangThai', label: 'Trạng thái', width: 140, align: 'center' },
      { key: 'soDonSAP', label: 'Số đơn/phiếu SAP', width: 140, align: 'center' },
      { key: 'dongBoSAP', label: 'Đồng bộ SAP', width: 120, align: 'center' },
      { key: 'soNgayCham', label: 'Số ngày chậm', width: 110, align: 'right', isNumber: true },
      { key: 'nguoiChoXuLy', label: 'Người đang chờ xử lý', width: 200, align: 'left' },
      { key: 'capDuyetCho', label: 'Cấp duyệt đang chờ', width: 140, align: 'center' },
      { key: 'tienDoXuLy', label: 'Tiến độ xử lý', width: 140, align: 'center' }
    ]

    wsDetail['!cols'] = detailCols.map(c => ({ wpx: c.width }))

    // Sheet 1 Title & Info
    wsDetail['A1'] = {
      v: 'BÁO CÁO CHI TIẾT TÌNH TRẠNG ĐƠN VẬT TƯ',
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '0A3D73' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      }
    }
    wsDetail['A2'] = {
      v: `Tổng cộng: ${filteredRows.length.toLocaleString('vi-VN')} đơn vật tư | Ngày xuất báo cáo: ${now.toLocaleDateString('vi-VN')}`,
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '64748B' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      }
    }

    let detailRowIdx = 4
    detailCols.forEach((col, cIdx) => {
      const cellRef = `${getExcelColName(cIdx)}${detailRowIdx}`
      wsDetail[cellRef] = {
        v: col.label,
        t: 's',
        s: headerStyle
      }
    })

    filteredRows.forEach((row, rIdx) => {
      detailRowIdx++
      const isEven = rIdx % 2 === 1
      const defaultBg = isEven ? 'F8FAFC' : 'FFFFFF'
      const rowAnalysis = analyzeRowTimeline(row)

      const finalPendingPerson = rowAnalysis.isPending ? rowAnalysis.pendingPerson : '—'
      const finalPendingRole = rowAnalysis.isPending ? rowAnalysis.pendingRole : '—'

      const rowValues = {
        stt: rIdx + 1,
        soDon: row.soDon || '—',
        ngayXuat: row.ngayXuat || '—',
        donViGiao: row.donViGiao || '—',
        nguoiGiao: row.nguoiGiao || '—',
        donViNhan: row.donViNhan || '—',
        nguoiNhan: row.nguoiNhan || '—',
        cbPheDuyet1: row.cbPheDuyet1 || '—',
        cbPheDuyet2: row.cbPheDuyet2 || '—',
        cbPheDuyet3: row.cbPheDuyet3 || '—',
        cbPheDuyet4: row.cbPheDuyet4 || '—',
        trangThai: row.trangThai || '—',
        soDonSAP: row.soDonSAP || '—',
        dongBoSAP: row.dongBoSAP || '—',
        soNgayCham: rowAnalysis.isPending ? rowAnalysis.delayDays : 0,
        nguoiChoXuLy: finalPendingPerson,
        capDuyetCho: finalPendingRole,
        tienDoXuLy: rowAnalysis.isCompleted ? 'Đã hoàn thành' : (rowAnalysis.isRejected ? 'Từ chối / Hủy' : 'Đang chờ duyệt')
      }

      detailCols.forEach((col, cIdx) => {
        const cellRef = `${getExcelColName(cIdx)}${detailRowIdx}`
        const val = rowValues[col.key] ?? ''

        let cellBg = defaultBg
        let fontColor = '1E293B'
        let fontBold = false

        if (col.key === 'soDon') {
          fontBold = true
          fontColor = '0A3D73'
        } else if (col.key === 'trangThai') {
          const s = String(val).toLowerCase().trim()
          if (s.includes('chưa xác nhận') || s.includes('chờ xác nhận')) {
            cellBg = 'FFF7ED'
            fontColor = 'C2410C'
            fontBold = true
          } else if (s.includes('chờ') || s.includes('chưa')) {
            cellBg = 'FFFBEB'
            fontColor = 'B45309'
            fontBold = true
          } else if (s.includes('từ chối') || s.includes('hủy') || s.includes('reject')) {
            cellBg = 'FEF2F2'
            fontColor = 'B91C1C'
            fontBold = true
          } else if (s.includes('đã') || s.includes('phê duyệt') || s.includes('hoàn thành')) {
            cellBg = 'ECFDF5'
            fontColor = '047857'
            fontBold = true
          }
        } else if (col.key === 'dongBoSAP') {
          const s = String(val).toLowerCase().trim()
          if (s.includes('đã') || s.includes('thành công') || s.includes('success') || s.includes('ok') || s.includes('đồng bộ')) {
            cellBg = 'EFF6FF'
            fontColor = '1E40AF'
          } else if (s.includes('chưa') || s.includes('lỗi') || s.includes('fail')) {
            cellBg = 'FFF1F2'
            fontColor = '9F1239'
          }
        } else if (col.key === 'soNgayCham' && val > 0) {
          cellBg = val > 2 ? 'FEF2F2' : 'FFFBEB'
          fontColor = val > 2 ? 'DC2626' : 'B45309'
          fontBold = true
        } else if (col.key === 'tienDoXuLy') {
          if (val === 'Đã hoàn thành') {
            cellBg = 'ECFDF5'
            fontColor = '047857'
            fontBold = true
          } else if (val === 'Đang chờ duyệt') {
            cellBg = 'FFF7ED'
            fontColor = 'C2410C'
            fontBold = true
          }
        }

        wsDetail[cellRef] = {
          v: col.isNumber && typeof val === 'number' ? val : (col.key === 'stt' ? val : String(val)),
          t: (col.isNumber || col.key === 'stt') && typeof val === 'number' ? 'n' : 's',
          s: {
            fill: { patternType: 'solid', fgColor: { rgb: cellBg } },
            font: { name: 'Segoe UI', sz: 9.5, bold: fontBold, color: { rgb: fontColor } },
            alignment: {
              horizontal: col.align || 'left',
              vertical: 'center'
            },
            border: borderAll
          }
        }
      })
    })

    wsDetail['!ref'] = `A1:${getExcelColName(detailCols.length - 1)}${detailRowIdx}`
    // Bổ sung công cụ lọc AutoFilter cho bảng dữ liệu chi tiết từ dòng tiêu đề (Dòng 4)
    wsDetail['!autofilter'] = { ref: `A4:${getExcelColName(detailCols.length - 1)}${detailRowIdx}` }

    /* =========================================================================
       SHEET: BẢNG TỔNG HỢP THEO TÊN NGƯỜI PHÊ DUYỆT KÈM SỐ NGÀY DUYỆT CHẬM
       ========================================================================= */
    const wsApprover = {}
    const approverMap = {}

    // Helper: Trích xuất chính xác tên Ban Chỉ Huy từ Đơn vị giao hoặc Đơn vị nhận
    const getBchListFromRow = (row) => {
      const isBch = (str) => {
        if (!str) return false
        const s = String(str).trim()
        if (!s || s === '—' || s === '-') return false
        const lower = s.toLowerCase()
        return (
          lower.startsWith('bch') ||
          lower.includes('bch ') ||
          lower.includes('bch-') ||
          lower.includes('bch_') ||
          lower.includes('-bch') ||
          lower.includes('_bch') ||
          lower.includes('ban chỉ huy') ||
          lower.includes('ban chi huy') ||
          lower.includes('ban điều hành') ||
          lower.includes('ban dieu hanh') ||
          lower.includes('ban đh') ||
          lower.includes('ban dh') ||
          lower.includes('ban qlda') ||
          lower.includes('ban ql') ||
          lower.includes('ban quản lý') ||
          lower.includes('ban quan ly') ||
          lower.includes('sgc-bch') ||
          lower.includes('sgc_bch') ||
          lower.includes('công trường') ||
          lower.includes('cong truong')
        )
      }

      const results = []
      if (isBch(row.donViGiao)) results.push(String(row.donViGiao).trim())
      if (isBch(row.donViNhan)) {
        const nhan = String(row.donViNhan).trim()
        if (!results.includes(nhan)) results.push(nhan)
      }

      // Nếu cả 2 đều không chứa chữ BCH rõ ràng, kiểm tra đơn vị doanh nghiệp/dự án thay vì tên cá nhân
      if (results.length === 0) {
        const isOrg = (str) => {
          if (!str) return false
          const s = String(str).trim()
          const lower = s.toLowerCase()
          return (
            lower.startsWith('công ty') ||
            lower.startsWith('cong ty') ||
            lower.startsWith('cty') ||
            lower.startsWith('g ty') ||
            lower.startsWith('nhà thầu') ||
            lower.startsWith('nha thau') ||
            lower.startsWith('xí nghiệp') ||
            lower.startsWith('xi nghiep') ||
            lower.startsWith('chi nhánh') ||
            lower.startsWith('chi nhanh') ||
            lower.startsWith('phân khu') ||
            lower.startsWith('dự án') ||
            lower.startsWith('da ')
          )
        }
        if (isOrg(row.donViGiao)) results.push(String(row.donViGiao).trim())
        if (isOrg(row.donViNhan)) {
          const nhan = String(row.donViNhan).trim()
          if (!results.includes(nhan)) results.push(nhan)
        }
      }

      // Nếu cả 2 đều không chứa chữ BCH/Org rõ ràng, thử trích xuất từ Mã/Số đơn (ví dụ BCH-VHDP-..., SGC-VHDP_...)
      if (results.length === 0 && row.soDon) {
        const strSoDon = String(row.soDon).trim()
        const bchCodeMatch = strSoDon.match(/(BCH[-_][A-Za-z0-9_&+-]+)/i) || strSoDon.match(/^(SGC[-_][A-Za-z0-9]+)/i)
        if (bchCodeMatch && bchCodeMatch[1]) {
          results.push(bchCodeMatch[1])
        }
      }

      return results
    }

    filteredRows.forEach(row => {
      const analysis = analyzeRowTimeline(row)
      const bchList = getBchListFromRow(row)

      // Identify people involved in this row
      const rolesInvolved = [
        { role: 'Người nhận (Xác nhận)', person: row.nguoiNhan, isCurrentPending: analysis.statusGroup === 'unconfirmed' },
        { role: 'CB phê duyệt 1', person: row.cbPheDuyet1, isCurrentPending: analysis.statusGroup === 'pending_cb1' },
        { role: 'CB phê duyệt 2', person: row.cbPheDuyet2, isCurrentPending: analysis.statusGroup === 'pending_cb2' },
        { role: 'CB phê duyệt 3', person: row.cbPheDuyet3, isCurrentPending: analysis.statusGroup === 'pending_cb3' },
        { role: 'CB phê duyệt 4', person: row.cbPheDuyet4, isCurrentPending: analysis.statusGroup === 'pending_cb4' }
      ]

      // 1. If currently pending approval/confirmation, record into current pending holder
      if (analysis.isPending) {
        let currentPersonName = analysis.pendingPerson
        let currentRole = analysis.pendingRole || 'Người phê duyệt'

        // Normalize map key (lowercase + trimmed) to avoid case-sensitivity split in JS while Excel COUNTIFS is case-insensitive
        const key = `${currentPersonName.toLowerCase().trim()}___${currentRole.toLowerCase().trim()}`
        if (!approverMap[key]) {
          approverMap[key] = {
            name: currentPersonName,
            role: currentRole,
            units: new Set(),
            pendingCount: 0,
            delay0Days: 0,
            delay1to2Days: 0,
            delayedOver2Days: 0,
            maxDelayDays: 0,
            totalDelaySum: 0,
            pendingOrders: [],
            completedCount: 0,
            totalInvolved: 0
          }
        }

        const item = approverMap[key]
        item.pendingCount++
        item.totalInvolved++
        bchList.forEach(bch => item.units.add(bch))

        const d = analysis.delayDays
        if (d > 2) {
          item.delayedOver2Days++
        } else if (d >= 1) {
          item.delay1to2Days++
        } else {
          item.delay0Days++
        }

        if (d > item.maxDelayDays) {
          item.maxDelayDays = d
        }
        item.totalDelaySum += d

        item.pendingOrders.push({
          soDon: row.soDon,
          delayDays: d,
          status: row.trangThai
        })
      }
    })

    const approverList = Object.values(approverMap)
      .filter(item => item.pendingCount > 0)
      .sort((a, b) => {
        // Prioritize who has overdue orders > 2 days
        if (b.delayedOver2Days !== a.delayedOver2Days) return b.delayedOver2Days - a.delayedOver2Days
        // Then who has most pending orders
        if (b.pendingCount !== a.pendingCount) return b.pendingCount - a.pendingCount
        // Then highest delay days
        if (b.maxDelayDays !== a.maxDelayDays) return b.maxDelayDays - a.maxDelayDays
        // Then by alphabetical name
        return a.name.localeCompare(b.name, 'vi')
      })

    const approverCols = [
      { key: 'stt', label: 'STT', width: 45, align: 'center' },
      { key: 'name', label: 'Tên người phê duyệt / xác nhận', width: 230, align: 'left' },
      { key: 'role', label: 'Vị trí / Cấp duyệt', width: 160, align: 'center' },
      { key: 'unitsStr', label: 'Ban chỉ huy liên quan', width: 360, align: 'left' },
      { key: 'pendingCount', label: 'Tổng đơn chờ duyệt', width: 150, align: 'right', isNumber: true },
      { key: 'maxDelayDays', label: 'Chậm nhất (ngày)', width: 140, align: 'right', isNumber: true },
      { key: 'avgDelayDays', label: 'Trung bình chậm', width: 140, align: 'right' }
    ]

    wsApprover['!cols'] = approverCols.map(c => ({ wpx: c.width }))

    // Sheet Title & Subtitle
    wsApprover['A1'] = {
      v: 'BẢNG TỔNG HỢP THEO TÊN NGƯỜI PHÊ DUYỆT & TIẾN ĐỘ DUYỆT CHẬM',
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '0A3D73' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      }
    }
    wsApprover['A2'] = {
      v: `Thống kê số lượng đơn đang chờ duyệt, phân loại số ngày duyệt chậm theo từng người phê duyệt | Ngày xuất: ${now.toLocaleDateString('vi-VN')}`,
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '64748B' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      }
    }

    let appRowIdx = 4
    approverCols.forEach((col, cIdx) => {
      const cellRef = `${getExcelColName(cIdx)}${appRowIdx}`
      wsApprover[cellRef] = {
        v: col.label,
        t: 's',
        s: headerStyle
      }
    })

    // Totals for Approver Sheet
    let sumPending = 0
    let sumMaxDelay = 0

    approverList.forEach((item, rIdx) => {
      appRowIdx++

      sumPending += item.pendingCount
      if (item.maxDelayDays > sumMaxDelay) sumMaxDelay = item.maxDelayDays

      const unitsArr = Array.from(item.units)
      const unitsStr = unitsArr.length > 0 ? unitsArr.join(', ') : '—'

      const avgDelayNum = item.pendingCount > 0 
        ? Number((item.totalDelaySum / item.pendingCount).toFixed(1))
        : 0

      const rowValues = {
        stt: rIdx + 1,
        name: item.name,
        role: item.role,
        unitsStr: unitsStr,
        pendingCount: item.pendingCount,
        maxDelayDays: item.pendingCount > 0 ? item.maxDelayDays : 0,
        avgDelayDays: avgDelayNum
      }

      approverCols.forEach((col, cIdx) => {
        const cellRef = `${getExcelColName(cIdx)}${appRowIdx}`
        const val = rowValues[col.key] ?? ''

        let fontColor = '1E293B'
        let fontBold = false
        let cellFill = 'FFFFFF'

        if (col.key === 'name') {
          fontBold = true
          fontColor = '0A3D73'
        } else if (col.key === 'pendingCount' && item.pendingCount > 0) {
          fontBold = true
          fontColor = item.delayedOver2Days > 0 ? 'DC2626' : 'C2410C'
        } else if (col.key === 'maxDelayDays' && item.maxDelayDays > 0) {
          fontBold = true
          fontColor = item.maxDelayDays > 2 ? 'DC2626' : 'D97706'
        }

        const cellObj = {
          v: col.isNumber && typeof val === 'number' ? val : (col.key === 'avgDelayDays' ? val : String(val)),
          t: col.isNumber || col.key === 'avgDelayDays' ? 'n' : 's',
          s: {
            fill: { patternType: 'solid', fgColor: { rgb: cellFill } },
            font: { name: 'Segoe UI', sz: 9.5, bold: fontBold, color: { rgb: fontColor } },
            alignment: {
              horizontal: col.align || 'left',
              vertical: 'center',
              wrapText: col.key === 'unitsStr'
            },
            border: borderAll
          }
        }

        // Bổ sung công thức Excel tính từ Sheet Dữ liệu chi tiết
        if (col.key === 'pendingCount') {
          cellObj.f = `COUNTIFS('Dữ liệu chi tiết'!$P$5:$P$${detailRowIdx}, B${appRowIdx}, 'Dữ liệu chi tiết'!$Q$5:$Q$${detailRowIdx}, C${appRowIdx})`
        } else if (col.key === 'maxDelayDays') {
          cellObj.f = `_xlfn.MAXIFS('Dữ liệu chi tiết'!$O$5:$O$${detailRowIdx}, 'Dữ liệu chi tiết'!$P$5:$P$${detailRowIdx}, B${appRowIdx}, 'Dữ liệu chi tiết'!$Q$5:$Q$${detailRowIdx}, C${appRowIdx})`
        } else if (col.key === 'avgDelayDays') {
          cellObj.f = `IFERROR(ROUND(AVERAGEIFS('Dữ liệu chi tiết'!$O$5:$O$${detailRowIdx}, 'Dữ liệu chi tiết'!$P$5:$P$${detailRowIdx}, B${appRowIdx}, 'Dữ liệu chi tiết'!$Q$5:$Q$${detailRowIdx}, C${appRowIdx}), 1), 0)`
          cellObj.z = '0.0" ngày"'
        }

        wsApprover[cellRef] = cellObj
      })
    })

    // Total Row for Summary Sheet
    appRowIdx++
    const totalAvgDelay = sumPending > 0 ? Number((approverList.reduce((acc, curr) => acc + curr.totalDelaySum, 0) / sumPending).toFixed(1)) : 0
    const totalApproverValues = {
      stt: '',
      name: 'TỔNG CỘNG',
      role: '',
      unitsStr: '',
      pendingCount: sumPending,
      maxDelayDays: sumMaxDelay,
      avgDelayDays: totalAvgDelay
    }

    const startDataRow = 5
    const endDataRow = appRowIdx - 1

    approverCols.forEach((col, cIdx) => {
      const cellRef = `${getExcelColName(cIdx)}${appRowIdx}`
      const val = totalApproverValues[col.key] ?? ''
      const cellObj = {
        v: col.isNumber || col.key === 'avgDelayDays' ? (typeof val === 'number' ? val : 0) : String(val),
        t: col.isNumber || col.key === 'avgDelayDays' ? 'n' : 's',
        s: {
          fill: { patternType: 'solid', fgColor: { rgb: 'E2E8F0' } },
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: '0A3D73' } },
          alignment: { horizontal: col.align || 'left', vertical: 'center' },
          border: {
            top: { style: 'medium', color: { rgb: '0A3D73' } },
            bottom: { style: 'medium', color: { rgb: '0A3D73' } },
            left: { style: 'thin', color: { rgb: 'CBD5E1' } },
            right: { style: 'thin', color: { rgb: 'CBD5E1' } }
          }
        }
      }

      // Bổ sung công thức Excel cho dòng TỔNG CỘNG
      if (endDataRow >= startDataRow) {
        if (col.key === 'pendingCount') {
          cellObj.f = `SUM(E${startDataRow}:E${endDataRow})`
        } else if (col.key === 'maxDelayDays') {
          cellObj.f = `MAX(F${startDataRow}:F${endDataRow})`
        } else if (col.key === 'avgDelayDays') {
          cellObj.f = `IFERROR(ROUND(AVERAGEIF(G${startDataRow}:G${endDataRow}, ">0"), 1), 0)`
          cellObj.z = '0.0" ngày"'
        }
      }

      wsApprover[cellRef] = cellObj
    })

    wsApprover['!ref'] = `A1:${getExcelColName(approverCols.length - 1)}${appRowIdx}`
    // Bổ sung công cụ lọc AutoFilter cho bảng tổng hợp từ dòng tiêu đề (Dòng 4)
    wsApprover['!autofilter'] = { ref: `A4:${getExcelColName(approverCols.length - 1)}${endDataRow}` }

    // Đổi tên thành "Tổng hợp" và xếp lên đầu tiên bên trái sheet "Dữ liệu chi tiết"
    XLSXStyle.utils.book_append_sheet(wb, wsApprover, 'Tổng hợp')
    XLSXStyle.utils.book_append_sheet(wb, wsDetail, 'Dữ liệu chi tiết')

    // Export Workbook
    const dateStr = now.toISOString().slice(0, 10).replace(/-/g, '')
    XLSXStyle.writeFile(wb, `SGC_Bao_Cao_Tinh_Trang_Don_Vat_Tu_${dateStr}.xlsx`)
  }

  // Status Badge Component
  const renderStatusBadge = (val) => {
    if (!val || val === '—' || val === '-') return <span style={{ color: '#94a3b8' }}>—</span>
    const s = String(val).toLowerCase().trim()

    let bg = '#f1f5f9'
    let color = '#475569'
    let border = '#cbd5e1'

    if (s.includes('chờ') || s.includes('chưa') || s.includes('pending')) {
      bg = '#fffbeb'
      color = '#b45309'
      border = '#fde68a'
    } else if (s.includes('từ chối') || s.includes('hủy') || s.includes('reject') || s.includes('không')) {
      bg = '#fef2f2'
      color = '#b91c1c'
      border = '#fecaca'
    } else if (s.includes('đã') || s.includes('phê duyệt') || s.includes('hoàn thành') || s.includes('duyệt') || s.includes('approved')) {
      bg = '#ecfdf5'
      color = '#047857'
      border = '#a7f3d0'
    }

    return (
      <span style={{
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '3px 10px',
        borderRadius: '6px',
        fontSize: '12px',
        fontWeight: 600,
        backgroundColor: bg,
        color: color,
        border: `1px solid ${border}`,
        whiteSpace: 'nowrap',
        lineHeight: 1.3
      }}>
        {val}
      </span>
    )
  }

  // SAP Sync Badge Component
  const renderSapBadge = (val) => {
    if (!val || val === '—' || val === '-') return <span style={{ color: '#94a3b8' }}>—</span>
    const s = String(val).toLowerCase().trim()

    let bg = '#f1f5f9'
    let color = '#475569'
    let border = '#cbd5e1'

    if (s.includes('đã') || s.includes('thành công') || s.includes('success') || s.includes('ok') || s.includes('đồng bộ')) {
      bg = '#eff6ff'
      color = '#1e40af'
      border = '#bfdbfe'
    } else if (s.includes('chưa') || s.includes('lỗi') || s.includes('fail')) {
      bg = '#fff1f2'
      color = '#9f1239'
      border = '#fecdd3'
    }

    return (
      <span style={{
        display: 'inline-flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '3px 10px',
        borderRadius: '6px',
        fontSize: '12px',
        fontWeight: 600,
        backgroundColor: bg,
        color: color,
        border: `1px solid ${border}`,
        whiteSpace: 'nowrap',
        lineHeight: 1.3
      }}>
        {val}
      </span>
    )
  }

  // Màu viên trạng thái: mỗi viên một màu theo ý nghĩa, không trùng nhau trong dãy
  const statusTones = (() => {
    const used = new Set(['blue', 'indigo'])
    const pick = (prefs) => {
      const all = [...prefs, 'orange', 'rose', 'magenta', 'violet', 'cyan', 'teal', 'green', 'red']
      const c = all.find(x => !used.has(x)) || prefs[0]
      used.add(c)
      return c
    }
    const out = {}
    dynamicStatusList.forEach(item => {
      const s = item.name.toLowerCase().trim()
      if (s.includes('chưa xác nhận') || s.includes('chờ xác nhận')) out[item.name] = pick(['orange'])
      else if (s.includes('chưa') || s.includes('chờ') || s.includes('pending')) out[item.name] = pick(['orange', 'rose', 'magenta', 'violet', 'cyan'])
      else if (s.includes('từ chối') || s.includes('hủy') || s.includes('reject') || s.includes('không')) out[item.name] = pick(['red', 'rose'])
      else if (s.includes('đã') || s.includes('phê duyệt') || s.includes('hoàn thành') || s.includes('duyệt') || s.includes('approved')) out[item.name] = pick(['green', 'teal'])
      else out[item.name] = pick(['cyan', 'teal', 'violet'])
    })
    return out
  })()

  const hasAnyFilter = !!(search || trangThaiFilter || donViGiaoFilter || donViNhanFilter || dongBoSAPFilter || startDate || endDate)

  return (
    <div className="sgt-page">
      <style>{TTDV_CSS}</style>

      {/* Hidden file inputs */}
      <input
        ref={fileInputRef}
        type="file"
        accept=".xlsx, .xls"
        multiple
        style={{ display: 'none' }}
        onChange={(e) => processFiles(e.target.files, false)}
      />
      <input
        ref={appendFileInputRef}
        type="file"
        accept=".xlsx, .xls"
        multiple
        style={{ display: 'none' }}
        onChange={(e) => processFiles(e.target.files, true)}
      />

      {rows.length > 0 && (
        <>
          {/* Hàng viên số liệu (bấm để lọc) + nút thao tác */}
          <div className="sgt-bar">
            <div className="sg-kpis">
              <button
                type="button"
                className={'sg-kpi sm sg-c-blue' + (!trangThaiFilter ? ' on' : ' dim')}
                onClick={() => {
                  setTrangThaiFilter('')
                  setCurrentPage(1)
                }}
                title="Bấm để xem tất cả đơn"
              >
                <span className="sg-kpi-ic"><FileSpreadsheet size={15} /></span>
                <span className="sg-kpi-bd"><span className="sg-kpi-lb">Tổng số đơn</span><span className="sg-kpi-v">{stats.total.toLocaleString('vi-VN')}</span></span>
              </button>

              {dynamicStatusList.map(item => {
                const isSelected = trangThaiFilter === item.name
                const Icon = item.IconComponent
                return (
                  <button
                    type="button"
                    key={item.name}
                    className={`sg-kpi sm sg-c-${statusTones[item.name]}` + (isSelected ? ' on' : (trangThaiFilter ? ' dim' : ''))}
                    onClick={() => {
                      if (isSelected) {
                        setTrangThaiFilter('')
                      } else {
                        setTrangThaiFilter(item.name)
                      }
                      setCurrentPage(1)
                    }}
                    title={`Bấm để lọc theo trạng thái: ${item.name}`}
                  >
                    <span className="sg-kpi-ic"><Icon size={15} /></span>
                    <span className="sg-kpi-bd"><span className="sg-kpi-lb">{item.name}</span><span className="sg-kpi-v">{item.count.toLocaleString('vi-VN')}</span></span>
                  </button>
                )
              })}

              {stats.dongBoThanhCong > 0 && (
                <button
                  type="button"
                  className={'sg-kpi sm sg-c-indigo' + (dongBoSAPFilter ? ' on' : '')}
                  onClick={() => {
                    if (dongBoSAPFilter) {
                      setDongBoSAPFilter('')
                    } else {
                      const firstOk = dongBoSAPOptions.find(o => o.toLowerCase().includes('đã') || o.toLowerCase().includes('đồng bộ'))
                      setDongBoSAPFilter(firstOk || '')
                    }
                    setCurrentPage(1)
                  }}
                  title="Bấm để lọc đơn đã đồng bộ SAP"
                >
                  <span className="sg-kpi-ic"><Database size={15} /></span>
                  <span className="sg-kpi-bd"><span className="sg-kpi-lb">Đã đồng bộ SAP</span><span className="sg-kpi-v">{stats.dongBoThanhCong.toLocaleString('vi-VN')}</span></span>
                </button>
              )}
            </div>

            <span className="sgt-spacer" />

            {fileName && (
              <span className="sgt-file" title={`Tệp đang xem: ${fileName} · ${rows.length.toLocaleString('vi-VN')} đơn`}>
                <FileCheck size={15} />
                <span>{fileName}</span>
              </span>
            )}

            <button
              type="button"
              className="sg-btn ghost"
              onClick={() => appendFileInputRef.current && appendFileInputRef.current.click()}
              title="Tải thêm file dữ liệu nối tiếp"
            >
              <Plus size={15} />
              <span>Tải thêm file</span>
            </button>

            <button
              type="button"
              className="sg-btn"
              onClick={() => fileInputRef.current && fileInputRef.current.click()}
              title="Tải lại từ file Excel mới (thay dữ liệu đang xem)"
            >
              <Upload size={15} />
              <span>Tải lại file mới</span>
            </button>
          </div>

          {/* Hàng lọc / tìm ngay trên bảng */}
          <div className="sgt-bar sgt-filters">
            <div className="sgt-search" title="Tìm theo số đơn, đơn vị, người duyệt, số SAP">
              <span className="sgt-search-ic"><Search size={16} /></span>
              <input
                type="text"
                placeholder="Tìm số đơn, đơn vị, người duyệt, SAP..."
                value={search}
                onChange={(e) => { setSearch(e.target.value); setCurrentPage(1) }}
              />
              {search && (
                <button type="button" className="sgt-x" onClick={() => setSearch('')} title="Xoá tìm kiếm" aria-label="Xoá tìm kiếm">
                  <IosX />
                </button>
              )}
            </div>

            <select
              className="sgt-sel"
              title="Trạng thái"
              value={trangThaiFilter}
              onChange={(e) => { setTrangThaiFilter(e.target.value); setCurrentPage(1) }}
            >
              <option value="">Tất cả trạng thái</option>
              {trangThaiOptions.map(t => (
                <option key={t} value={t}>{t}</option>
              ))}
            </select>

            <select
              className="sgt-sel"
              title="Đơn vị giao"
              value={donViGiaoFilter}
              onChange={(e) => { setDonViGiaoFilter(e.target.value); setCurrentPage(1) }}
            >
              <option value="">Tất cả đơn vị giao</option>
              {donViGiaoOptions.map(g => (
                <option key={g} value={g}>{g}</option>
              ))}
            </select>

            <select
              className="sgt-sel"
              title="Đơn vị nhận"
              value={donViNhanFilter}
              onChange={(e) => { setDonViNhanFilter(e.target.value); setCurrentPage(1) }}
            >
              <option value="">Tất cả đơn vị nhận</option>
              {donViNhanOptions.map(n => (
                <option key={n} value={n}>{n}</option>
              ))}
            </select>

            <select
              className="sgt-sel"
              title="Đồng bộ SAP"
              value={dongBoSAPFilter}
              onChange={(e) => { setDongBoSAPFilter(e.target.value); setCurrentPage(1) }}
            >
              <option value="">Đồng bộ SAP: tất cả</option>
              {dongBoSAPOptions.map(s => (
                <option key={s} value={s}>{s}</option>
              ))}
            </select>

            <div className="sgt-dates" title="Khoảng ngày xuất">
              <Calendar size={15} color="#6B7AA0" />
              <input
                type="date"
                value={startDate}
                onChange={(e) => { setStartDate(e.target.value); setCurrentPage(1) }}
                title="Từ ngày xuất"
              />
              <span>đến</span>
              <input
                type="date"
                value={endDate}
                onChange={(e) => { setEndDate(e.target.value); setCurrentPage(1) }}
                title="Đến ngày xuất"
              />
            </div>

            {hasAnyFilter && (
              <button
                type="button"
                className="sg-btn danger"
                onClick={() => {
                  setSearch('')
                  setTrangThaiFilter('')
                  setDonViGiaoFilter('')
                  setDonViNhanFilter('')
                  setDongBoSAPFilter('')
                  setStartDate('')
                  setEndDate('')
                  setCurrentPage(1)
                }}
                title="Bỏ toàn bộ bộ lọc"
              >
                <IosX />
                Bỏ lọc
              </button>
            )}

            <span className="sgt-spacer" />

            <button
              type="button"
              className="sg-btn green"
              onClick={handleExportExcel}
              title="Xuất bảng dữ liệu đang lọc ra file Excel"
            >
              <Download size={15} />
              <span>Xuất Excel ({filteredRows.length.toLocaleString('vi-VN')})</span>
            </button>

            <button
              type="button"
              className="sg-btn danger"
              onClick={handleClearData}
              title="Xoá toàn bộ dữ liệu đang xem"
            >
              <Trash2 size={15} />
              <span>Xoá dữ liệu</span>
            </button>
          </div>
        </>
      )}

      {/* Vùng nội dung chính */}
      {rows.length === 0 ? (
        /* Khối thả file tải lên */
        <div className="sgt-dropwrap">
          <div
            className={'sgt-drop' + (isDragOver ? ' over' : '')}
            onDragOver={(e) => { e.preventDefault(); setIsDragOver(true) }}
            onDragLeave={() => setIsDragOver(false)}
            onDrop={handleDrop}
            onClick={() => fileInputRef.current && fileInputRef.current.click()}
            title="Tải lên file Excel theo dõi tình trạng đơn vật tư và đồng bộ SAP"
          >
            <div className="sgt-drop-ic"><Upload size={28} /></div>
            <h2>Tải lên file Excel tình trạng đơn vật tư</h2>
            <p>Kéo thả file Excel vào đây hoặc bấm để chọn từ máy</p>
            <span className="sg-btn" style={{ pointerEvents: 'none' }}>
              <Upload size={15} />
              Chọn file Excel
            </span>

            <div className="sgt-drop-note">
              <div className="sgt-drop-note-hd">
                <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9.5" /><path d="M12 11v6" /><path d="M12 7.5h.01" /></svg>
                Cấu trúc file Excel hỗ trợ gồm các cột:
              </div>
              <div>
                <strong>STT</strong> · <strong>Số đơn</strong> · <strong>Ngày xuất</strong> · <strong>Đơn vị giao</strong> · <strong>Người giao</strong> · <strong>Đơn vị nhận</strong> · <strong>Người nhận</strong> · <strong>Cán bộ phê duyệt 1 → 4</strong> · <strong>Trạng thái</strong> · <strong>Số đơn/phiếu SAP</strong> · <strong>Đồng bộ SAP</strong>
              </div>
            </div>
          </div>
        </div>
      ) : (
        /* Bảng dữ liệu */
        <div className="sgt-tablebox">
          <div className="sgt-scroll">
            <table style={{ width: '100%', minWidth: '1750px', borderCollapse: 'separate', borderSpacing: 0, fontSize: '15px', textAlign: 'left' }}>
              <thead style={{ position: 'sticky', top: 0, zIndex: 20 }}>
                <tr style={{ background: '#0a3d73', color: '#ffffff' }}>
                  <th style={{
                    padding: '8px 8px',
                    width: '55px',
                    minWidth: '55px',
                    textAlign: 'center',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    STT
                  </th>
                  <th
                    onClick={() => handleSort('soDon')}
                    style={{
                      padding: '8px 12px',
                      width: '210px',
                      minWidth: '190px',
                      maxWidth: '240px',
                      textAlign: 'left',
                      fontWeight: 700,
                      fontSize: '15px',
                      background: '#0a3d73',
                      color: '#ffffff',
                      cursor: 'pointer',
                      userSelect: 'none',
                      whiteSpace: 'nowrap',
                      wordBreak: 'normal',
                      borderRight: '1px solid rgba(255,255,255,0.3)',
                      borderBottom: '2px solid #062648'
                    }}
                    title="Nhấp để sắp xếp theo Số đơn"
                  >
                    <div style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'flex-start', gap: 5 }}>
                      <span>Số đơn</span>
                      <ArrowUpDown size={12} opacity={sortField === 'soDon' ? 1 : 0.6} />
                    </div>
                  </th>
                  <th
                    onClick={() => handleSort('ngayXuat')}
                    style={{
                      padding: '8px 10px',
                      width: '115px',
                      minWidth: '115px',
                      textAlign: 'center',
                      fontWeight: 700,
                      fontSize: '15px',
                      background: '#0a3d73',
                      color: '#ffffff',
                      cursor: 'pointer',
                      userSelect: 'none',
                      whiteSpace: 'nowrap',
                      wordBreak: 'normal',
                      borderRight: '1px solid rgba(255,255,255,0.3)',
                      borderBottom: '2px solid #062648'
                    }}
                    title="Nhấp để sắp xếp theo Ngày xuất"
                  >
                    <div style={{ display: 'inline-flex', alignItems: 'center', justifyContent: 'center', gap: 5 }}>
                      <span>Ngày xuất</span>
                      <ArrowUpDown size={12} opacity={sortField === 'ngayXuat' ? 1 : 0.6} />
                    </div>
                  </th>
                  <th style={{
                    padding: '8px 12px',
                    width: '230px',
                    minWidth: '230px',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    Đơn vị giao
                  </th>
                  <th style={{
                    padding: '8px 12px',
                    width: '150px',
                    minWidth: '150px',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    Người giao
                  </th>
                  <th style={{
                    padding: '8px 12px',
                    width: '230px',
                    minWidth: '230px',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    Đơn vị nhận
                  </th>
                  <th style={{
                    padding: '8px 12px',
                    width: '150px',
                    minWidth: '150px',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    Người nhận
                  </th>
                  <th style={{
                    padding: '8px 10px',
                    width: '170px',
                    minWidth: '170px',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    CB phê duyệt 1
                  </th>
                  <th style={{
                    padding: '8px 10px',
                    width: '170px',
                    minWidth: '170px',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    CB phê duyệt 2
                  </th>
                  <th style={{
                    padding: '8px 10px',
                    width: '160px',
                    minWidth: '160px',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    CB phê duyệt 3
                  </th>
                  <th style={{
                    padding: '8px 10px',
                    width: '160px',
                    minWidth: '160px',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    CB phê duyệt 4
                  </th>
                  <th style={{
                    padding: '8px 10px',
                    width: '140px',
                    minWidth: '140px',
                    textAlign: 'center',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    Trạng thái
                  </th>
                  <th style={{
                    padding: '8px 10px',
                    width: '150px',
                    minWidth: '150px',
                    textAlign: 'center',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderRight: '1px solid rgba(255,255,255,0.3)',
                    borderBottom: '2px solid #062648'
                  }}>
                    Số đơn/phiếu SAP
                  </th>
                  <th style={{
                    padding: '8px 10px',
                    width: '140px',
                    minWidth: '140px',
                    textAlign: 'center',
                    fontWeight: 700,
                    fontSize: '15px',
                    background: '#0a3d73',
                    color: '#ffffff',
                    whiteSpace: 'nowrap',
                    wordBreak: 'normal',
                    borderBottom: '2px solid #062648'
                  }}>
                    Đồng bộ SAP
                  </th>
                </tr>
              </thead>
              <tbody>
                {paginatedRows.length === 0 ? (
                  <tr>
                    <td colSpan={14} style={{ padding: '36px', textAlign: 'center', color: '#94a3b8' }}>
                      Không tìm thấy bản ghi nào phù hợp với bộ lọc hiện tại.
                    </td>
                  </tr>
                ) : (
                  paginatedRows.map((row, idx) => {
                    const rowNumber = pageSize === 0 ? idx + 1 : (currentPage - 1) * pageSize + idx + 1
                    const isEven = idx % 2 === 1
                    return (
                      <tr
                        key={row.id || idx}
                        style={{
                          backgroundColor: isEven ? '#f8fafc' : '#ffffff',
                          borderBottom: '1px solid #cbd5e1',
                          transition: 'background-color 0.1s ease'
                        }}
                        onMouseEnter={(e) => { e.currentTarget.style.backgroundColor = '#f1f5f9' }}
                        onMouseLeave={(e) => { e.currentTarget.style.backgroundColor = isEven ? '#f8fafc' : '#ffffff' }}
                      >
                        <td style={{ padding: '7px 8px', textAlign: 'center', color: '#64748b', fontWeight: 500, borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {rowNumber}
                        </td>
                        <td style={{
                          padding: '7px 12px',
                          textAlign: 'left',
                          fontWeight: 700,
                          color: '#0a3d73',
                          borderRight: '1px solid #cbd5e1',
                          borderBottom: '1px solid #cbd5e1',
                          fontFamily: "'Roboto', sans-serif",
                          whiteSpace: 'normal',
                          wordBreak: 'break-word',
                          overflowWrap: 'break-word',
                          width: '210px',
                          minWidth: '190px',
                          maxWidth: '240px',
                          lineHeight: 1.35
                        }}>
                          {row.soDon || '—'}
                        </td>
                        <td style={{ padding: '7px 10px', textAlign: 'center', color: '#334155', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {row.ngayXuat || '—'}
                        </td>
                        <td style={{ padding: '7px 12px', color: '#1e293b', fontWeight: 500, borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {row.donViGiao || '—'}
                        </td>
                        <td style={{ padding: '7px 12px', color: '#475569', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {row.nguoiGiao || '—'}
                        </td>
                        <td style={{ padding: '7px 12px', color: '#1e293b', fontWeight: 500, borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {row.donViNhan || '—'}
                        </td>
                        <td style={{ padding: '7px 12px', color: '#475569', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {row.nguoiNhan || '—'}
                        </td>
                        <td style={{ padding: '7px 10px', color: '#475569', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {row.cbPheDuyet1 || '—'}
                        </td>
                        <td style={{ padding: '7px 10px', color: '#475569', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {row.cbPheDuyet2 || '—'}
                        </td>
                        <td style={{ padding: '7px 10px', color: '#475569', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {row.cbPheDuyet3 || '—'}
                        </td>
                        <td style={{ padding: '7px 10px', color: '#475569', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {row.cbPheDuyet4 || '—'}
                        </td>
                        <td style={{ padding: '7px 10px', textAlign: 'center', borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {renderStatusBadge(row.trangThai)}
                        </td>
                        <td style={{ padding: '7px 10px', textAlign: 'center', color: '#0f172a', fontWeight: 600, borderRight: '1px solid #cbd5e1', borderBottom: '1px solid #cbd5e1', fontFamily: "'Roboto', sans-serif", whiteSpace: 'nowrap' }}>
                          {row.soDonSAP || '—'}
                        </td>
                        <td style={{ padding: '7px 10px', textAlign: 'center', borderBottom: '1px solid #cbd5e1', whiteSpace: 'nowrap' }}>
                          {renderSapBadge(row.dongBoSAP)}
                        </td>
                      </tr>
                    )
                  })
                )}
              </tbody>
            </table>
          </div>

          {/* Chân bảng: số dòng + phân trang */}
          <div className="sgt-foot">
            <div style={{ display: 'flex', alignItems: 'center', gap: 12, flexWrap: 'wrap' }}>
              <span>
                Hiển thị <strong>{filteredRows.length === 0 ? 0 : (currentPage - 1) * pageSize + 1}</strong> - <strong>{Math.min(currentPage * pageSize, filteredRows.length)}</strong> trên tổng <strong>{filteredRows.length.toLocaleString('vi-VN')}</strong> đơn
              </span>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <span>Số dòng/trang:</span>
                <select
                  className="sgt-sel sm"
                  value={pageSize}
                  onChange={(e) => { setPageSize(Number(e.target.value)); setCurrentPage(1) }}
                >
                  <option value={20}>20</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                  <option value={200}>200</option>
                  <option value={0}>Tất cả</option>
                </select>
              </div>
            </div>

            {pageSize > 0 && totalPages > 1 && (
              <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                <button type="button" className="sgt-pg" onClick={() => setCurrentPage(1)} disabled={currentPage === 1} title="Trang đầu">
                  <ChevronsLeft size={16} />
                </button>
                <button type="button" className="sgt-pg" onClick={() => setCurrentPage(prev => Math.max(prev - 1, 1))} disabled={currentPage === 1} title="Trang trước">
                  <ChevronLeft size={16} />
                </button>
                <span style={{ padding: '0 8px', fontWeight: 700, color: '#16171C' }}>
                  {currentPage} / {totalPages}
                </span>
                <button type="button" className="sgt-pg" onClick={() => setCurrentPage(prev => Math.min(prev + 1, totalPages))} disabled={currentPage === totalPages} title="Trang sau">
                  <ChevronRight size={16} />
                </button>
                <button type="button" className="sgt-pg" onClick={() => setCurrentPage(totalPages)} disabled={currentPage === totalPages} title="Trang cuối">
                  <ChevronsRight size={16} />
                </button>
              </div>
            )}
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

const TTDV_CSS = `
.sgt-page { display: flex; flex-direction: column; gap: 10px; height: 100%; overflow: hidden; padding: 14px 20px 16px; box-sizing: border-box; }
.sgt-bar { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; flex: none; }
.sgt-spacer { flex: 1 1 0; }
.sgt-file { display: inline-flex; align-items: center; gap: 6px; height: 30px; max-width: 280px; padding: 0 12px; border-radius: 99px; background: #fff; color: #0A3D73; font-size: 12px; font-weight: 700; box-shadow: inset 0 0 0 1px #D5DCEB; }
.sgt-file > span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.sgt-search { position: relative; flex: 1 1 260px; min-width: 220px; max-width: 360px; }
.sgt-search-ic { position: absolute; left: 13px; top: 50%; transform: translateY(-50%); color: #6B7AA0; display: flex; pointer-events: none; }
html body .sgt-search input { width: 100%; height: 36px; box-sizing: border-box; padding: 0 38px; font-size: 15px; color: #16171C; background: #fff; border: 1.5px solid #D5DCEB; border-radius: 99px; outline: none; }
.sgt-x { position: absolute; right: 7px; top: 50%; transform: translateY(-50%); width: 24px; height: 24px; padding: 0; border: 0; border-radius: 50%; display: grid; place-items: center; background: #FFE7E5; color: #FF3B30; cursor: pointer; transition: background .15s, color .15s; }
.sgt-x:hover { background: #FF3B30; color: #fff; }
html body .sgt-x:active { transform: translateY(-50%) scale(.94); }
html body .sgt-sel { height: 36px; box-sizing: border-box; border: 1.5px solid #D5DCEB; padding-left: 12px; font-size: 15px; font-weight: 600; color: #16171C; max-width: 210px; flex: 0 1 auto; }
html body .sgt-sel.sm { height: 32px; font-size: 12px; }
.sgt-dates { display: inline-flex; align-items: center; gap: 6px; height: 36px; padding: 0 6px 0 12px; border-radius: 12px; background: #fff; box-shadow: inset 0 0 0 1.5px #D5DCEB; font-size: 12px; font-weight: 600; color: #4A4C55; }
html body .sgt-dates input { height: 28px; border: 0 !important; box-shadow: none !important; background: transparent; font-size: 15px; color: #16171C; padding: 0 4px; outline: none; }
.sgt-tablebox { flex: 1 1 auto; min-height: 0; display: flex; flex-direction: column; gap: 8px; }
.sgt-scroll { flex: 1 1 auto; min-height: 0; overflow: auto; position: relative; border-radius: 14px; }
.sgt-scroll > table { background: #fff; }
.sgt-foot { flex: none; display: flex; align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 10px; padding: 6px 14px; border-radius: 99px; background: rgba(255,255,255,.72); font-size: 12px; font-weight: 600; color: #4A4C55; }
.sgt-pg { width: 32px; height: 32px; padding: 0; border: 0; border-radius: 50%; display: grid; place-items: center; background: #fff; color: #0F58A7; box-shadow: inset 0 0 0 1px #D5DCEB; cursor: pointer; }
.sgt-pg:hover:not(:disabled) { background: #E4EEFA; }
.sgt-pg:disabled { color: #B5BAC6; cursor: not-allowed; }
.sgt-dropwrap { flex: 1; display: flex; align-items: center; justify-content: center; min-height: 360px; }
.sgt-drop { width: 100%; max-width: 640px; box-sizing: border-box; padding: 32px 28px 26px; border-radius: 24px; border: 2px dashed #C9D3E6; background: rgba(255,255,255,.72); backdrop-filter: blur(10px); display: flex; flex-direction: column; align-items: center; gap: 10px; text-align: center; cursor: pointer; transition: border-color .2s, background .2s, box-shadow .2s; box-shadow: 0 20px 40px -28px rgba(10, 61, 115, .35); }
.sgt-drop:hover, .sgt-drop.over { border-color: #0F58A7; background: rgba(228, 238, 250, .85); }
.sgt-drop-ic { width: 60px; height: 60px; border-radius: 18px; display: grid; place-items: center; color: #fff; background: linear-gradient(135deg, #0A3D73, #0F58A7 55%, #2F86E0); box-shadow: 0 12px 24px -12px rgba(15, 88, 167, .9); margin-bottom: 4px; }
.sgt-drop h2 { margin: 0; font-size: 22px; font-weight: 700; color: #16171C; }
.sgt-drop p { margin: 0 0 4px; font-size: 15px; color: #4A4C55; }
.sgt-drop-note { margin-top: 10px; width: 100%; box-sizing: border-box; padding: 12px 16px; border-radius: 16px; background: rgba(245, 246, 248, .9); text-align: left; font-size: 12px; color: #4A4C55; line-height: 1.6; }
.sgt-drop-note-hd { display: flex; align-items: center; gap: 6px; font-weight: 700; color: #0A3D73; margin-bottom: 4px; }
`
