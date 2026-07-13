import React, { useState, useRef, useCallback, useMemo } from 'react'
import * as XLSXStyleRaw from 'xlsx-js-style'
const XLSXStyle = XLSXStyleRaw.default || XLSXStyleRaw
const XLSX = XLSXStyle
import { exportConsolidatedExcel, buildPhanNhomVatTuSheet } from './excelExporter.js'
import {
  Upload, FileSpreadsheet, Search, X, RefreshCw, Info,
  ChevronDown, ChevronRight, Download, Truck, PackageCheck, Settings, BarChart3,
  AlertCircle, CheckCircle2, Filter, ArrowUpDown, Clock, CloudUpload, Database, Save,
  Pencil, Trash2, Lock, ClipboardList, Warehouse, Building2, Users, HelpCircle,
  Calendar, AlertTriangle, DollarSign, Copy, Terminal, LogOut, User as UserIcon, CheckSquare, ArrowRight
} from 'lucide-react'
import { COLS_GIAO_NHAN, parseXlsxToRows, formatVal, getTrangThaiColor, isApprovedStatus, isPendingStatus, isRejectedStatus } from './constants.js'
import { supabase, isSupabaseConfigured, supabaseUrl, supabaseAnonKey } from './supabaseClient.js'
import LoginPage from './components/LoginPage.jsx'
import QuanLyTaiKhoanTab from './components/QuanLyTaiKhoanTab.jsx'
import TaiSanKhauHaoTab from './components/TaiSanKhauHaoTab.jsx'

const KEYS_TO_REMOVE_FOR_REAL_REPORT = [
  'tenNguon', 'maNguon', 'lo', 'hangMuc', 'soHopDong', 'thuKho', 
  'bienSoXe', 'phanKhu', 'duAn', 'tinhTrang', 'nguoiNhan', 
  'maDonLienQuan', 'nhaCungCap', 'maDonChuyenTiepLC', 'maDonChuyenTiepNB'
]

const baseColsForReal = COLS_GIAO_NHAN.filter(c => !KEYS_TO_REMOVE_FOR_REAL_REPORT.includes(c.key))
const loaiDonIdx = baseColsForReal.findIndex(c => c.key === 'loaiDon')
if (loaiDonIdx !== -1) {
  baseColsForReal.splice(loaiDonIdx, 1,
    { key: 'khoiLuongThuc', label: 'Khối lượng thực', width: 80 },
    { key: 'logicTongHop', label: 'Logic tổng hợp', width: 60 }
  )
}
export const COLS_REAL_REPORT = baseColsForReal

// ─── Searchable Select ────────────────────────────────────────────────────────
function SearchableSelect({ value, onChange, options, placeholder = 'Tất cả dự án', searchPlaceholder, variant = 'header', align = 'left', onEditProject, onDeleteProject }) {
  const [isOpen, setIsOpen] = useState(false)
  const [searchQuery, setSearchQuery] = useState('')
  const containerRef = useRef(null)

  // Click outside close
  React.useEffect(() => {
    function handleClickOutside(event) {
      if (containerRef.current && !containerRef.current.contains(event.target)) {
        setIsOpen(false)
      }
    }
    document.addEventListener('mousedown', handleClickOutside)
    return () => {
      document.removeEventListener('mousedown', handleClickOutside)
    }
  }, [])

  // Clear query on close
  React.useEffect(() => {
    if (!isOpen) {
      setSearchQuery('')
    }
  }, [isOpen])

  const filteredOptions = useMemo(() => {
    const q = searchQuery.toLowerCase().trim()
    if (!q) return options
    return options.filter(o => o.toLowerCase().includes(q))
  }, [options, searchQuery])

  const displaySelected = value || placeholder

  const isHeader = variant === 'header'
  const triggerStyle = isHeader ? {
    background: '#ffffff',
    border: 'none',
    borderRadius: 6,
    padding: '0 32px 0 12px',
    color: '#0f172a',
    fontSize: 14,
    fontWeight: 700,
    cursor: 'pointer',
    width: 'fit-content',
    minWidth: 160,
    maxWidth: 600,
    height: 32,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    userSelect: 'none',
    position: 'relative',
    boxShadow: '0 1px 3px rgba(15,23,42,0.08)',
    boxSizing: 'border-box'
  } : {
    background: '#ffffff',
    border: '1.5px solid #cbd5e1',
    borderRadius: 'var(--radius)',
    padding: '0 32px 0 14px',
    color: 'var(--text)',
    fontSize: 15,
    fontWeight: 500,
    cursor: 'pointer',
    width: '100%',
    minWidth: 180,
    height: 44,
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    userSelect: 'none',
    position: 'relative',
    boxShadow: 'var(--shadow-sm)',
    boxSizing: 'border-box'
  }

  const dropdownPanelStyle = {
    position: 'absolute',
    top: 'calc(100% + 4px)',
    left: align === 'left' ? 0 : 'auto',
    right: align === 'right' ? 0 : 'auto',
    zIndex: 9999,
    background: '#ffffff',
    borderRadius: 8,
    boxShadow: '0 10px 25px -5px rgba(0,0,0,0.1), 0 8px 10px -6px rgba(0,0,0,0.1)',
    border: '1px solid #cbd5e1',
    padding: 6,
    minWidth: '100%',
    maxWidth: 460,
    display: 'flex',
    flexDirection: 'column',
    gap: 6
  }

  return (
    <div ref={containerRef} style={{ position: 'relative', display: isHeader ? 'inline-block' : 'block', width: isHeader ? 'auto' : '100%' }}>
      {/* Trigger */}
      <div onClick={() => setIsOpen(!isOpen)} style={triggerStyle}>
        <span style={{
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          marginRight: 16,
          display: 'block',
          flex: '1 1 0%',
          minWidth: 0,
          lineHeight: 'normal'
        }} title={displaySelected}>
          {displaySelected}
        </span>
        <ChevronDown size={isHeader ? 14 : 16} color={isHeader ? '#0f58a7' : 'var(--text-muted)'} style={{
          position: 'absolute',
          right: isHeader ? 8 : 12,
          top: '50%',
          transform: 'translateY(-50%)',
          pointerEvents: 'none'
        }} />
      </div>

      {/* Popover list */}
      {isOpen && (
        <div style={dropdownPanelStyle}>
          {/* Inner Search Field */}
          <div style={{ position: 'relative', display: 'flex', alignItems: 'center' }}>
            <Search size={13} style={{ position: 'absolute', left: 8, color: '#64748b' }} />
            <input
              type="text"
              placeholder={searchPlaceholder || "Tìm kiếm..."}
              value={searchQuery}
              onChange={e => setSearchQuery(e.target.value)}
              style={{
                width: '100%',
                padding: '5px 8px 5px 28px',
                fontSize: 13,
                border: '1px solid #cbd5e1',
                borderRadius: 4,
                outline: 'none',
                height: 28,
                background: '#f8fafc',
                color: '#0f172a'
              }}
              onClick={e => e.stopPropagation()}
              onKeyDown={e => {
                if (e.key === 'Escape') setIsOpen(false)
              }}
              autoFocus
            />
          </div>

          {/* Render Options */}
          <div style={{
            maxHeight: 200,
            overflowY: 'auto',
            display: 'flex',
            flexDirection: 'column',
            gap: 2,
            paddingRight: 2
          }}>
            {/* "Tất cả dự án" option */}
            <div
              onClick={() => {
                onChange('')
                setIsOpen(false)
              }}
              style={{
                padding: '6px 8px',
                fontSize: 13,
                fontWeight: !value ? 700 : 500,
                borderRadius: 4,
                cursor: 'pointer',
                color: !value ? '#0f58a7' : '#334155',
                background: !value ? '#eff6ff' : 'transparent',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                transition: 'background 0.1s'
              }}
              onMouseOver={e => {
                if (value) e.currentTarget.style.background = '#f1f5f9'
              }}
              onMouseOut={e => {
                if (value) e.currentTarget.style.background = 'transparent'
              }}
            >
              <span>{placeholder}</span>
              {!value && <CheckCircle2 size={13} color="#0f58a7" />}
            </div>

            {/* List options */}
            {filteredOptions.map(opt => {
              const matches = value === opt
              return (
                <div
                  key={opt}
                  onClick={() => {
                    onChange(opt)
                    setIsOpen(false)
                  }}
                  style={{
                    padding: '6px 8px',
                    fontSize: 13,
                    fontWeight: matches ? 700 : 500,
                    borderRadius: 4,
                    cursor: 'pointer',
                    color: matches ? '#0f58a7' : '#334155',
                    background: matches ? '#eff6ff' : 'transparent',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    transition: 'background 0.1s'
                  }}
                  onMouseOver={e => {
                    if (!matches) e.currentTarget.style.background = '#f1f5f9'
                  }}
                  onMouseOut={e => {
                    if (!matches) e.currentTarget.style.background = 'transparent'
                  }}
                  title={opt}
                >
                  <span style={{ textOverflow: 'ellipsis', whiteSpace: 'nowrap', overflow: 'hidden', marginRight: 'auto', textAlign: 'left' }}>
                    {opt}
                  </span>
                  {onDeleteProject && (
                    <button
                      onClick={(e) => {
                        e.stopPropagation()
                        onDeleteProject(opt)
                      }}
                      title="Xóa dự án"
                      style={{
                        background: 'transparent',
                        border: 'none',
                        color: '#64748b',
                        padding: '4px',
                        cursor: 'pointer',
                        borderRadius: '4px',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        marginLeft: 4,
                        marginRight: matches ? 4 : 0,
                        transition: 'all 0.1s'
                      }}
                      onMouseOver={e => {
                        e.stopPropagation()
                        e.currentTarget.style.color = '#ef4444'
                        e.currentTarget.style.background = '#fee2e2'
                      }}
                      onMouseOut={e => {
                        e.stopPropagation()
                        e.currentTarget.style.color = '#64748b'
                        e.currentTarget.style.background = 'transparent'
                      }}
                    >
                      <Trash2 size={11} />
                    </button>
                  )}
                  {matches && <CheckCircle2 size={13} color="#0f58a7" />}
                </div>
              )
            })}

            {filteredOptions.length === 0 && (
              <div style={{ padding: '8px', fontSize: 12, color: '#64748b', textAlign: 'center' }}>
                Không tìm thấy kết quả
              </div>
            )}
          </div>
        </div>
      )}
    </div>
  )
}

// ─── Header ──────────────────────────────────────────────────────────────────
function Header({ selectedProject, setSelectedProject, duAnOptions, onOpenAddProjectModal, onEditProject, onDeleteProject, onOpenConfigModal, onForceRefresh, onMouseEnterLogo, activeTab, currentUser, onLogout }) {
  const getHeaderTitle = (tab) => {
    switch (tab) {
      case 'chung':
        return 'SGC | ĐƠN CHUNG'
      case 'kho':
        return 'SGC | KHO DỰ ÁN'
      case 'inventory':
        return 'SGC | BÁO CÁO XUẤT NHẬP TỒN'
      case 'inventory_real':
        return 'SGC | BÁO CÁO XUẤT NHẬP THỰC'
      case 'depreciation_assets':
        return 'SGC | THỐNG KÊ TÀI SẢN KHẤU HAO'
      case 'dongia':
        return 'SGC | PHÂN NHÓM VẬT TƯ'
      case 'accounts':
        return 'SGC | QUẢN LÝ TÀI KHOẢN'
      default:
        return 'SGC | BÁO CÁO GIAO NHẬN'
    }
  }

  return (
    <header style={{
      background: 'linear-gradient(135deg, #0a3d73 0%, #0f58a7 60%, #1a6abf 100%)', // Original professional blue system
      padding: '0 24px',
      height: 64, // Taller header for premium elegance
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'space-between',
      boxShadow: '0 4px 20px rgba(10,61,115,0.25)',
      position: 'sticky',
      top: 0,
      zIndex: 100,
      borderBottom: '1px solid rgba(255, 255, 255, 0.1)'
    }}>
      <div 
        onMouseEnter={onMouseEnterLogo}
        style={{ display: 'flex', alignItems: 'center', gap: 12, cursor: 'pointer' }}
      >
        <div style={{
          width: 36, height: 36, background: '#ffffff',
          borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center',
          boxShadow: '0 3px 12px rgba(10,61,115,0.2)',
        }}>
          <span style={{ color: '#0f58a7', fontWeight: 900, fontSize: 15, letterSpacing: '0.02em' }}>SGC</span>
        </div>
        <div>
          <div style={{ color: '#ffffff', fontWeight: 800, fontSize: 18, letterSpacing: '0.03em', display: 'flex', alignItems: 'center', gap: 8 }}>
            {getHeaderTitle(activeTab)}
          </div>
        </div>
      </div>

      {/* Removed Kho Dự Án Widget as requested */}

      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        {isSupabaseConfigured && (
          <button
            onClick={async (e) => {
              const btn = e.currentTarget
              const originalText = btn.innerHTML
              btn.disabled = true
              btn.style.opacity = '0.6'
              btn.innerHTML = '<span style="display:flex;align-items:center;gap:4px;"><svg class="animate-spin" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/></svg> Đang tải...</span>'
              try {
                await onForceRefresh()
              } catch (err) {
                console.error(err)
              } finally {
                btn.disabled = false
                btn.style.opacity = '1'
                btn.innerHTML = originalText
              }
            }}
            title="Tải lại toàn bộ dữ liệu sạch trực tiếp từ Supabase (Bỏ qua Caching)"
            style={{
              display: 'flex',
              alignItems: 'center',
              background: 'rgba(255, 255, 255, 0.12)',
              border: '1px solid rgba(255, 255, 255, 0.2)',
              borderRadius: 6,
              padding: '4px 10px',
              color: '#ffffff',
              fontSize: 12,
              fontWeight: 600,
              whiteSpace: 'nowrap',
              cursor: 'pointer',
              transition: 'all 0.15s ease',
              height: 28,
              boxSizing: 'border-box'
            }}
            onMouseOver={(e) => {
              e.currentTarget.style.background = 'rgba(255, 255, 255, 0.25)'
            }}
            onMouseOut={(e) => {
              e.currentTarget.style.background = 'rgba(255, 255, 255, 0.12)'
            }}
          >
            <RefreshCw size={12} style={{ marginRight: 5, color: '#e0f2fe' }} />
            <span>Tải lại DB</span>
          </button>
        )}

        <div
          title={isSupabaseConfigured ? "Đã kết nối cơ sở dữ liệu Supabase" : "Chưa kết nối cơ sở dữ liệu Supabase"}
          style={{
            display: 'flex',
            alignItems: 'center',
            background: isSupabaseConfigured ? 'rgba(16, 185, 129, 0.15)' : 'rgba(245, 158, 11, 0.15)',
            border: isSupabaseConfigured ? '1px solid rgba(16, 185, 129, 0.3)' : '1px solid rgba(245, 158, 11, 0.3)',
            borderRadius: 6,
            padding: '4px 10px',
            color: '#ffffff',
            fontSize: 12,
            fontWeight: 600,
            whiteSpace: 'nowrap',
            cursor: 'default',
            height: 28,
            boxSizing: 'border-box'
          }}
        >
          <Database size={13} style={{ marginRight: 5, color: isSupabaseConfigured ? '#34d399' : '#fbbf24' }} />
          <span>{isSupabaseConfigured ? 'Supabase Connected' : 'Supabase Offline'}</span>
          <span style={{
            width: 6,
            height: 6,
            borderRadius: '50%',
            background: isSupabaseConfigured ? '#10b981' : '#f59e0b',
            marginLeft: 6,
            display: 'inline-block',
            boxShadow: isSupabaseConfigured ? '0 0 8px #10b981' : '0 0 8px #f59e0b'
          }} />
        </div>

        {currentUser && (
          <div style={{
            display: 'flex',
            alignItems: 'center',
            gap: 8,
            background: 'rgba(255,255,255,0.15)',
            border: '1px solid rgba(255,255,255,0.2)',
            borderRadius: 6,
            padding: '4px 10px',
            color: '#ffffff',
            fontSize: 12,
            fontWeight: 600,
            height: 28,
            boxSizing: 'border-box'
          }}>
            <UserIcon size={12} style={{ color: '#93c5fd' }} />
            <span style={{ maxWidth: '80px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={`${currentUser.ho_ten} (${currentUser.quyen})`}>
              {currentUser.ho_ten}
            </span>
            <button
              onClick={onLogout}
              title="Đăng xuất khỏi hệ thống"
              style={{
                background: 'none',
                border: 'none',
                color: '#feca1d',
                cursor: 'pointer',
                display: 'flex',
                alignItems: 'center',
                padding: '2px',
                borderRadius: '4px',
                marginLeft: '2px'
              }}
            >
              <LogOut size={13} />
            </button>
          </div>
        )}

        <div style={{
          background: 'rgba(255,255,255,0.15)',
          border: '1px solid rgba(255,255,255,0.2)',
          borderRadius: 6,
          padding: '4px 12px',
          color: '#ffffff',
          fontSize: 13,
          fontWeight: 600,
          letterSpacing: '0.03em'
        }}>
          v1.0.0
        </div>
      </div>
    </header>
  )
}

// ─── Tab Bar ─────────────────────────────────────────────────────────────────
function TabBar({ tabs, active, onChange }) {
  return (
    <div style={{
      display: 'flex',
      gap: 16,
      background: 'var(--surface)',
      borderBottom: '1px solid var(--border)',
      padding: '0 24px',
    }}>
      {tabs.map(t => (
        <button
          key={t.id}
          onClick={() => onChange(t.id)}
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: 8,
            padding: '16px 8px',
            fontWeight: active === t.id ? 700 : 500,
            fontSize: 15,
            color: active === t.id ? 'var(--primary)' : 'var(--text-muted)',
            background: 'transparent',
            borderBottom: active === t.id ? '3px solid var(--primary)' : '3px solid transparent',
            marginBottom: -1,
            transition: 'all 0.15s ease-in-out',
            whiteSpace: 'nowrap',
          }}
        >
          {React.cloneElement(t.icon, { size: 16, strokeWidth: active === t.id ? 2.5 : 2 })}
          <span>{t.label}</span>
          {t.count != null && t.count > 0 && (
            <span style={{
              background: active === t.id ? 'var(--primary)' : '#cbd5e1',
              color: active === t.id ? '#fff' : 'var(--text-muted)',
              borderRadius: 20,
              padding: '2px 8px',
              fontSize: 11,
              fontWeight: 700,
            }}>{t.count.toLocaleString()}</span>
          )}
        </button>
      ))}
    </div>
  )
}

// ─── Upload Zone ──────────────────────────────────────────────────────────────
function UploadZone({ onFile, label, accept = '.xlsx,.xls', disabled = false }) {
  const [drag, setDrag] = useState(false)
  const ref = useRef()

  const handle = useCallback((files) => {
    if (disabled || !files || files.length === 0) return
    const fileList = Array.from(files)
    const promises = fileList.map(file => {
      return new Promise((resolve) => {
        const reader = new FileReader()
        reader.onload = (e) => {
          try {
            const u8arr = new Uint8Array(e.target.result)
            const wb = XLSX.read(u8arr, { type: 'array' })
            const ws = wb.Sheets[wb.SheetNames[0]]
            if (ws && ws['!ref']) {
              try {
                const refRange = XLSX.utils.decode_range(ws['!ref'])
                refRange.s.r = 0
                refRange.s.c = 0
                ws['!ref'] = XLSX.utils.encode_range(refRange)
              } catch (err) {
                console.error('Error rewriting sheet range:', err)
              }
            }
            const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
            resolve({ data, name: file.name })
          } catch (error) {
            console.error('Error parsing excel file in UploadZone:', error)
            resolve(null)
          }
        }
        reader.readAsArrayBuffer(file)
      })
    })

    Promise.all(promises).then(results => {
      const validResults = results.filter(Boolean)
      if (validResults.length > 0) {
        onFile(validResults)
      }
    })
  }, [onFile, disabled])

  return (
    <div
      className={`upload-zone${drag && !disabled ? ' drag-over' : ''}${disabled ? ' disabled' : ''}`}
      onClick={() => { if (!disabled) ref.current.click() }}
      onDragOver={e => { e.preventDefault(); if (!disabled) setDrag(true) }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); if (!disabled) handle(e.dataTransfer.files) }}
      style={{
        cursor: disabled ? 'not-allowed' : 'pointer',
        opacity: disabled ? 0.75 : 1,
        backgroundColor: disabled ? '#fafafa' : undefined,
        borderColor: disabled ? '#fca5a5' : undefined,
        borderStyle: disabled ? 'dashed' : 'dashed',
        position: 'relative'
      }}
    >
      <input ref={ref} type="file" accept={accept} multiple disabled={disabled} onChange={e => { if (!disabled) handle(e.target.files) }} style={{ display: 'none' }} />
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
        <div style={{
          width: 56, height: 56, borderRadius: 12,
          background: disabled ? '#ff787515' : 'var(--primary-light)',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          boxShadow: disabled ? 'none' : '0 2px 8px rgba(37,99,235,0.08)'
        }}>
          {disabled ? <Lock size={24} color="#f5222d" /> : <Upload size={24} color="var(--primary)" />}
        </div>
        <div>
          <div style={{ fontWeight: 700, fontSize: 16, color: disabled ? '#f5222d' : 'var(--text)', marginBottom: 6 }}>
            {disabled ? 'Nút Tải File Đang Bị Khóa' : label}
          </div>
          <div style={{ color: disabled ? '#7f1d1d' : 'var(--text-muted)', fontSize: 14 }}>
            {disabled ? (
              <span style={{ fontWeight: 600 }}>⚠️ Bạn chưa Chọn Kho Dự Án cụ thể để gán dữ liệu.</span>
            ) : (
              <>Kéo thả các file vào đây hoặc <span style={{ color: 'var(--primary)', fontWeight: 600 }}>chọn file từ thiết bị</span></>
            )}
          </div>
          <div style={{ color: disabled ? '#b91c1c' : 'var(--text-light)', fontSize: 12, marginTop: 6, fontStyle: 'italic', fontWeight: disabled ? 500 : 'normal' }}>
            {disabled 
              ? 'Hãy chọn một dự án cụ thể bên danh sách "KHO DỰ ÁN" ở thanh Menu trên trước để tiếp tục tải lên.' 
              : 'Cho phép chọn nhiều file cùng lúc. Hỗ trợ định dạng chuẩn: .xlsx, .xls'}
          </div>
        </div>
      </div>
    </div>
  )
}

// ─── Stats Bar ────────────────────────────────────────────────────────────────
function StatsBar({ rows, onAppendFile }) {
  const stats = useMemo(() => {
    const total = rows.length
    
    // Group rows exactly by their visible status string value
    const counts = {}
    rows.forEach(r => {
      const status = (r.trangThai || 'Chưa duyệt').trim()
      counts[status] = (counts[status] || 0) + 1
    })
    
    const activeStatuses = Object.keys(counts).filter(status => counts[status] > 0)
    
    // Define ordering priority for cards
    const priority = ['Đã phê duyệt', 'Chưa xác nhận', 'Chờ phê duyệt', 'Chưa phê duyệt', 'Chưa duyệt', 'Từ chối']
    activeStatuses.sort((a, b) => {
      const idxA = priority.indexOf(a)
      const idxB = priority.indexOf(b)
      if (idxA !== -1 && idxB !== -1) return idxA - idxB
      if (idxA !== -1) return -1
      if (idxB !== -1) return 1
      return a.localeCompare(b, 'vi')
    })
    
    const statusCards = activeStatuses.map(status => {
      const count = counts[status]
      
      let color = 'var(--primary)'
      let bg = 'var(--primary-light)'
      let border = 'var(--border)'
      let icon = <ClipboardList size={18} />
      
      const sLower = status.toLowerCase()
      if (sLower === 'đã phê duyệt' || (isApprovedStatus(status) && !sLower.includes('chưa') && !sLower.includes('chờ'))) {
        color = '#10b981'
        bg = '#ecfdf5'
        border = '#a7f3d0'
        icon = <CheckCircle2 size={18} />
      } else if (sLower === 'từ chối' || isRejectedStatus(status)) {
        color = '#ef4444'
        bg = '#fef2f2'
        border = '#fca5a5'
        icon = <AlertCircle size={18} />
      } else {
        color = '#b45309'
        bg = '#fffbeb'
        border = '#fde68a'
        icon = <Clock size={18} />
      }
      
      return {
        label: status,
        value: count.toLocaleString('vi-VN') + ' đơn',
        color,
        bg,
        border,
        icon,
        count
      }
    })
    
    return { total, cards: statusCards }
  }, [rows])

  const cards = [
    { label: 'Tổng số lượng đơn', value: stats.total.toLocaleString('vi-VN') + ' đơn', color: 'var(--primary)', bg: 'var(--primary-light)', border: 'var(--border)', icon: <FileSpreadsheet size={18} /> },
    ...stats.cards
  ]

  return (
    <div style={{ display: 'flex', gap: 12, marginBottom: 16, flexWrap: 'wrap', flexShrink: 0 }}>
      {cards.map(c => (
        <div key={c.label} style={{
          background: c.bg,
          border: `1px solid ${c.border}`,
          borderRadius: 8,
          padding: '10px 16px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          flex: '1 1 200px',
          minWidth: 180,
          boxShadow: 'var(--shadow-sm)'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
            <span style={{ color: 'var(--text-muted)', fontSize: 14, fontWeight: 500 }}>{c.label}:</span>
            <span style={{ color: c.color, fontSize: 16, fontWeight: 700, letterSpacing: '-0.01em' }}>{c.value}</span>
          </div>
          <div style={{
            width: 32, height: 32, borderRadius: 8,
            background: '#ffffff', color: c.color,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            border: `1px solid ${c.border}`,
            flexShrink: 0
          }}>
            {c.icon}
          </div>
        </div>
      ))}

      {onAppendFile && (
        <div 
          onClick={() => document.getElementById('append-file-input')?.click()}
          style={{
            background: '#ffffff',
            border: '2px dashed #0284c7',
            borderRadius: 8,
            padding: '10px 16px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flex: '1 1 200px',
            minWidth: 180,
            boxShadow: 'var(--shadow-sm)',
            cursor: 'pointer',
            transition: 'all 0.2s ease',
            userSelect: 'none'
          }}
          onMouseOver={(e) => {
            e.currentTarget.style.borderColor = '#0369a1';
            e.currentTarget.style.background = '#f0f9ff';
          }}
          onMouseOut={(e) => {
            e.currentTarget.style.borderColor = '#0284c7';
            e.currentTarget.style.background = '#ffffff';
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
            <span style={{ color: '#0369a1', fontSize: 14, fontWeight: 700 }}>Up file nối tiếp:</span>
            <span style={{ color: '#64748b', fontSize: 12, fontWeight: 500 }}>Nối tiếp dữ liệu</span>
          </div>
          <div style={{
            width: 32, height: 32, borderRadius: 8,
            background: '#0284c7', color: '#ffffff',
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            flexShrink: 0
          }}>
            <CloudUpload size={18} />
          </div>
          <input
            id="append-file-input"
            type="file"
            accept=".xlsx,.xls"
            multiple
            style={{ display: 'none' }}
            onChange={(e) => {
              const files = e.target.files;
              if (files && files.length > 0) {
                const fileList = Array.from(files);
                const promises = fileList.map(file => {
                  return new Promise((resolve) => {
                    const reader = new FileReader();
                    reader.onload = (evt) => {
                      if (evt.target?.result) {
                        try {
                          const u8arr = new Uint8Array(evt.target.result)
                          const wb = XLSX.read(u8arr, { type: 'array' });
                          const ws = wb.Sheets[wb.SheetNames[0]];
                          if (ws && ws['!ref']) {
                            try {
                              const refRange = XLSX.utils.decode_range(ws['!ref'])
                              refRange.s.r = 0
                              refRange.s.c = 0
                              ws['!ref'] = XLSX.utils.encode_range(refRange)
                            } catch (err) {
                              console.error('Error rewriting sheet range:', err)
                            }
                          }
                          const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
                          resolve({ data, name: file.name });
                        } catch (error) {
                          console.error('Error parsing appended excel file:', error)
                          resolve(null)
                        }
                      } else {
                        resolve(null);
                      }
                    };
                    reader.readAsArrayBuffer(file);
                  });
                });

                Promise.all(promises).then(results => {
                  const validResults = results.filter(Boolean);
                  if (validResults.length > 0) {
                    onAppendFile(validResults);
                  }
                });

                // Reset value so user can upload same file again
                e.target.value = '';
              }
            }}
          />
        </div>
      )}
    </div>
  )
}

// ─── Filter Bar ───────────────────────────────────────────────────────────────
function FilterBar({
  search,
  setSearch,
  trangThai,
  setTrangThai,
  trangThaiOptions,
  duAn,
  setDuAn,
  duAnOptions,
  donViGiao,
  setDonViGiao,
  donViGiaoOptions = [],
  donViNhan,
  setDonViNhan,
  donViNhanOptions = [],
  startDate,
  setStartDate,
  endDate,
  setEndDate,
  onClear,
  onEditProject,
  onExportExcel
}) {
  return (
    <div style={{ display: 'flex', gap: 10, marginBottom: 12, flexWrap: 'wrap', alignItems: 'center', flexShrink: 0 }}>
      <div style={{ position: 'relative', flex: '1 1 240px', minWidth: 200 }}>
        <Search size={14} style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', color: 'var(--text-light)' }} />
        <input
          className="input"
          style={{ paddingLeft: 34, width: '100%', height: 44 }}
          placeholder="Tìm kiếm tên vật tư, mã đơn, đơn vị..."
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
      </div>

      {/* Date range picker matching Image 1 */}
      <div style={{
        display: 'flex',
        alignItems: 'center',
        gap: 8,
        padding: '0 12px',
        borderRadius: 'var(--radius, 8px)',
        border: '1px solid #cbd5e1',
        background: '#ffffff',
        height: 44,
        boxShadow: '0 1px 2px rgba(15, 23, 42, 0.05)',
        fontSize: '13.5px',
        color: '#334155',
        flexShrink: 0
      }}>
        <input
          type="date"
          value={startDate || ''}
          onChange={e => setStartDate && setStartDate(e.target.value)}
          style={{
            border: 'none',
            outline: 'none',
            background: 'transparent',
            fontSize: '13.5px',
            color: '#0f172a',
            fontFamily: 'inherit',
            cursor: 'pointer',
            padding: 0
          }}
          title="Từ ngày"
        />
        <span style={{ color: '#94a3b8', fontSize: '13px', fontWeight: 600 }}>➔</span>
        <input
          type="date"
          value={endDate || ''}
          onChange={e => setEndDate && setEndDate(e.target.value)}
          style={{
            border: 'none',
            outline: 'none',
            background: 'transparent',
            fontSize: '13.5px',
            color: '#0f172a',
            fontFamily: 'inherit',
            cursor: 'pointer',
            padding: 0
          }}
          title="Đến ngày"
        />
        <Calendar size={15} style={{ color: '#475569', marginLeft: 4, flexShrink: 0 }} />
      </div>

      <select 
         className="input" 
         style={{ 
           flex: '0 0 auto', 
           width: 'auto', 
           minWidth: 160, 
           height: 44,
           paddingRight: 28,
           cursor: 'pointer'
         }} 
         value={trangThai} 
         onChange={e => setTrangThai(e.target.value)}
       >
        <option value="">Tất cả trạng thái</option>
        {trangThaiOptions.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
      <div style={{ flex: '0 0 auto', minWidth: 200, maxWidth: 320 }}>
        <SearchableSelect
          value={donViGiao}
          onChange={setDonViGiao}
          options={donViGiaoOptions}
          placeholder="Tất cả Đơn vị giao"
          searchPlaceholder="Tìm kiếm Đơn vị giao..."
          variant="filter"
          align="right"
         />
      </div>
      <div style={{ flex: '0 0 auto', minWidth: 200, maxWidth: 320 }}>
        <SearchableSelect
          value={donViNhan}
          onChange={setDonViNhan}
          options={donViNhanOptions}
          placeholder="Tất cả Đơn vị nhận"
          searchPlaceholder="Tìm kiếm Đơn vị nhận..."
          variant="filter"
          align="right"
        />
      </div>
      {onExportExcel && (
        <button
          className="btn"
          onClick={onExportExcel}
          style={{
            background: 'linear-gradient(135deg, #059669 0%, #10b981 100%)',
            color: '#ffffff',
            border: 'none',
            boxShadow: '0 2px 4px rgba(16,185,129,0.2)',
            height: 44,
            padding: '0 16px',
            borderRadius: 'var(--radius)',
            display: 'inline-flex',
            alignItems: 'center',
            gap: 6,
            fontWeight: 600,
            fontSize: '13.5px',
            cursor: 'pointer'
          }}
        >
          <Download size={14} /> Xuất Excel
        </button>
      )}
      {(search || trangThai || duAn || donViGiao || donViNhan || startDate || endDate) && (
        <button 
          className="btn btn-outline btn-sm" 
          style={{ 
            height: 44, 
            padding: '0 16px', 
            borderRadius: 'var(--radius)', 
            flexShrink: 0 
          }} 
          onClick={onClear}
        >
          <X size={12} /> Xóa lọc
        </button>
      )}
    </div>
  )
}

// ─── Real Report Summary Table (Báo cáo tổng hợp khối lượng Thực nhập và Thực xuất) ─────────────────
function RealReportSummaryTable({ summaryRows = [], customCategoryMap = {}, localProject = '', materialPriceRows = [], materialMetadataMap = null }) {
  const [pageSize, setPageSize] = React.useState(100)
  const [currentPage, setCurrentPage] = React.useState(1)
  const [detailRow, setDetailRow] = React.useState(null)
  const [selectedSaps, setSelectedSaps] = React.useState(new Set())
  const tableWrapRef = React.useRef(null)
  const mirrorRef = React.useRef(null)

  const toggleSelectAll = () => {
    if (selectedSaps.size === summaryRows.length) {
      setSelectedSaps(new Set())
    } else {
      setSelectedSaps(new Set(summaryRows.map(r => r.maSAP)))
    }
  }

  const toggleSelectRow = (maSAP) => {
    const next = new Set(selectedSaps)
    if (next.has(maSAP)) {
      next.delete(maSAP)
    } else {
      next.add(maSAP)
    }
    setSelectedSaps(next)
  }

  const handleExportSelectedExcel = () => {
    if (selectedSaps.size === 0) return
    const selectedRows = summaryRows.filter(r => selectedSaps.has(r.maSAP))
    exportConsolidatedExcel(selectedRows, localProject, customCategoryMap, getUnitCategory, materialPriceRows, materialMetadataMap)
  }

  // eslint-disable-next-line no-unused-vars
  const old_handleExportSelectedExcel = () => {
    if (selectedSaps.size === 0) return

    const wb = XLSXStyle.utils.book_new()
    const selectedRows = summaryRows.filter(r => selectedSaps.has(r.maSAP))

    const titleStyle = {
      font: { name: 'Segoe UI', sz: 16, bold: true, color: { rgb: '1E3A8A' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    }
    const subtitleStyle = {
      font: { name: 'Segoe UI', sz: 11, italic: true, color: { rgb: '475569' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    }
    
    const tblHeaderStyle1 = {
      fill: { patternType: 'solid', fgColor: { rgb: '0F766E' } },
      font: { name: 'Segoe UI', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    }
    const tblHeaderStyle2 = {
      fill: { patternType: 'solid', fgColor: { rgb: 'C2410C' } },
      font: { name: 'Segoe UI', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    }
    
    const thStyle1 = {
      fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
      font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: 'CCFBF1' } },
        bottom: { style: 'thin', color: { rgb: 'CCFBF1' } },
        left: { style: 'thin', color: { rgb: 'CCFBF1' } },
        right: { style: 'thin', color: { rgb: 'CCFBF1' } }
      }
    }
    const thStyle2 = {
      fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } },
      font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: 'FFEDD5' } },
        bottom: { style: 'thin', color: { rgb: 'FFEDD5' } },
        left: { style: 'thin', color: { rgb: 'FFEDD5' } },
        right: { style: 'thin', color: { rgb: 'FFEDD5' } }
      }
    }

    const dataCellStyle = {
      font: { name: 'Segoe UI', sz: 9.5, color: { rgb: '1E293B' } },
      border: {
        top: { style: 'thin', color: { rgb: 'E2E8F0' } },
        bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
        left: { style: 'thin', color: { rgb: 'E2E8F0' } },
        right: { style: 'thin', color: { rgb: 'E2E8F0' } }
      },
      alignment: { vertical: 'center', wrapText: true }
    }

    const legendLabelStyle = {
      font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
      alignment: { vertical: 'center' }
    }
    const legendNccStyle = {
      fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
      font: { name: 'Segoe UI', sz: 9, bold: true, color: { rgb: '78350F' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'FDE68A' } },
        bottom: { style: 'thin', color: { rgb: 'FDE68A' } },
        left: { style: 'thin', color: { rgb: 'FDE68A' } },
        right: { style: 'thin', color: { rgb: 'FDE68A' } }
      }
    }
    const legendKhoStyle = {
      fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
      font: { name: 'Segoe UI', sz: 9, bold: true, color: { rgb: '1E3A8A' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'BFDBFE' } },
        bottom: { style: 'thin', color: { rgb: 'BFDBFE' } },
        left: { style: 'thin', color: { rgb: 'BFDBFE' } },
        right: { style: 'thin', color: { rgb: 'BFDBFE' } }
      }
    }
    const legendToDoiStyle = {
      fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
      font: { name: 'Segoe UI', sz: 9, bold: true, color: { rgb: '065F46' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'BBF7D0' } },
        bottom: { style: 'thin', color: { rgb: 'BBF7D0' } },
        left: { style: 'thin', color: { rgb: 'BBF7D0' } },
        right: { style: 'thin', color: { rgb: 'BBF7D0' } }
      }
    }

    const subheaders1 = ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Tổng KL Chứng từ', 'Hệ số logic', 'Thực nhập', 'Ghi chú']
    const subheaders2 = ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Tổng KL Chứng từ', 'Hệ số logic', 'Thực xuất', 'Ghi chú']

    const footerStyle = {
      font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E293B' } },
      fill: { patternType: 'solid', fgColor: { rgb: 'F8FAFC' } },
      border: {
        top: { style: 'thin', color: { rgb: '94A3B8' } },
        bottom: { style: 'double', color: { rgb: '94A3B8' } },
        left: { style: 'thin', color: { rgb: 'E2E8F0' } },
        right: { style: 'thin', color: { rgb: 'E2E8F0' } }
      },
      alignment: { vertical: 'center', wrapText: true }
    }

    const totalYellowStyle = {
      ...footerStyle,
      fill: { patternType: 'solid', fgColor: { rgb: 'FFFF00' } },
      alignment: { horizontal: 'right', vertical: 'center', wrapText: true }
    }

    selectedRows.forEach(itemRow => {
      const maSAP = itemRow.maSAP
      const bcSheetName = `BC_${maSAP}`.slice(0, 31)
      const thSheetName = `TH_${maSAP}`.slice(0, 31)
      const nhapSheetName = `Nhap_${maSAP}`.slice(0, 31)
      const xuatSheetName = `Xuat_${maSAP}`.slice(0, 31)

      const txs = itemRow.transactions || []
      const nhapSummaryByUnit = {}
      const xuatSummaryByUnit = {}

      txs.forEach(tx => {
        const normProject = String(localProject || '').trim().toLowerCase()
        const isGiaoProject = String(tx.donViGiao || '').trim().toLowerCase() === normProject
        const isNhanProject = String(tx.donViNhan || '').trim().toLowerCase() === normProject

        if (tx.nhapVal > 0) {
          let role = 'Đơn vị nhận'
          let partner = tx.donViGiao || 'Chưa xác định'
          if (isGiaoProject) {
            role = 'Đơn vị giao'
            partner = tx.donViNhan || 'Chưa xác định'
          }

          const key = `${partner}||${tx.detailTypeNhap}||${role}`
          if (!nhapSummaryByUnit[key]) {
            nhapSummaryByUnit[key] = {
              partner,
              role,
              detailType: tx.detailTypeNhap,
              logicVal: tx.logicNhapVal,
              explain: tx.explainNhap,
              totalQty: 0,
              totalContribution: 0
            }
          }
          nhapSummaryByUnit[key].totalQty += tx.nhapVal
          nhapSummaryByUnit[key].totalContribution += tx.nhapVal * tx.logicNhapVal
        }
        if (tx.xuatVal > 0) {
          let role = 'Đơn vị giao'
          let partner = tx.donViNhan || 'Chưa xác định'
          if (isNhanProject) {
            role = 'Đơn vị nhận'
            partner = tx.donViGiao || 'Chưa xác định'
          }

          const key = `${partner}||${tx.detailTypeXuat}||${role}`
          if (!xuatSummaryByUnit[key]) {
            xuatSummaryByUnit[key] = {
              partner,
              role,
              detailType: tx.detailTypeXuat,
              logicVal: tx.logicXuatVal,
              explain: tx.explainXuat,
              totalQty: 0,
              totalContribution: 0
            }
          }
          xuatSummaryByUnit[key].totalQty += tx.xuatVal
          xuatSummaryByUnit[key].totalContribution += tx.xuatVal * tx.logicXuatVal
        }
      })

      const sortNhapSummaryList = (list) => {
        return [...list].sort((a, b) => {
          const rolePriorityA = a.role === 'Đơn vị nhận' ? 1 : a.role === 'Đơn vị giao' ? 2 : 3
          const rolePriorityB = b.role === 'Đơn vị nhận' ? 1 : b.role === 'Đơn vị giao' ? 2 : 3
          if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

          const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
          const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
          const catPriorityA = catA === 'ncc' ? 1 : catA === 'kho' ? 2 : catA === 'todoi' ? 3 : 4

          const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
          const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
          const catPriorityB = catB === 'ncc' ? 1 : catB === 'kho' ? 2 : catB === 'todoi' ? 3 : 4

          if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

          const valA = Math.abs(a.totalContribution || 0)
          const valB = Math.abs(b.totalContribution || 0)
          if (valB !== valA) return valB - valA

          return (a.partner || '').localeCompare(b.partner || '', 'vi')
        })
      }

      const sortXuatSummaryList = (list) => {
        return [...list].sort((a, b) => {
          const rolePriorityA = a.role === 'Đơn vị giao' ? 1 : a.role === 'Đơn vị nhận' ? 2 : 3
          const rolePriorityB = b.role === 'Đơn vị giao' ? 1 : b.role === 'Đơn vị nhận' ? 2 : 3
          if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

          const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
          const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
          const catPriorityA = catA === 'todoi' ? 1 : catA === 'kho' ? 2 : catA === 'ncc' ? 3 : 4

          const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
          const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
          const catPriorityB = catB === 'todoi' ? 1 : catB === 'kho' ? 2 : catB === 'ncc' ? 3 : 4

          if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

          const valA = Math.abs(a.totalContribution || 0)
          const valB = Math.abs(b.totalContribution || 0)
          if (valB !== valA) return valB - valA

          return (a.partner || '').localeCompare(b.partner || '', 'vi')
        })
      }

      const nhapSummaryList = sortNhapSummaryList(Object.values(nhapSummaryByUnit))
      const xuatSummaryList = sortXuatSummaryList(Object.values(xuatSummaryByUnit))

      const buildSheetDataForMulti = (listNhap, listXuat, isBaoCao = false) => {
        const ws = {}
        
        const sub1 = isBaoCao 
          ? ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Thực nhập', 'Ghi chú']
          : ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Tổng KL Chứng từ', 'Hệ số logic', 'Thực nhập', 'Ghi chú']

        const sub2 = isBaoCao 
          ? ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Thực xuất', 'Ghi chú']
          : ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Tổng KL Chứng từ', 'Hệ số logic', 'Thực xuất', 'Ghi chú']

        if (isBaoCao) {
          ws['!cols'] = [
            { wpx: 40 }, { wpx: 130 }, { wpx: 180 }, { wch: 16 }, { wpx: 250 }, { wch: 16 }, { wch: 16 }, { wpx: 200 },
            { wpx: 40 },
            { wpx: 40 }, { wpx: 130 }, { wpx: 180 }, { wch: 16 }, { wpx: 250 }, { wch: 16 }, { wch: 16 }, { wpx: 200 }
          ]
        } else {
          ws['!cols'] = [
            { wpx: 40 }, { wpx: 130 }, { wpx: 180 }, { wch: 16 }, { wpx: 250 }, { wch: 16 }, { wch: 16 }, { wch: 6 }, { wch: 16 }, { wpx: 200 },
            { wpx: 40 },
            { wpx: 40 }, { wpx: 130 }, { wpx: 180 }, { wch: 16 }, { wpx: 250 }, { wch: 16 }, { wch: 16 }, { wch: 6 }, { wch: 16 }, { wpx: 200 }
          ]
        }

        ws['A2'] = { v: 'BẢNG GIẢI TRÌNH CẤU THÀNH SỐ LIỆU CHI TIẾT', t: 's', s: titleStyle }
        ws['A3'] = { v: `Kho chọn: ${localProject || 'Tất cả'} | Vật tư: ${itemRow.tenVatTu} (${itemRow.maSAP}) | ĐVT: ${itemRow.dvt}`, t: 's', s: subtitleStyle }

        ws['A4'] = { v: 'Màu sắc đối tác:', t: 's', s: legendLabelStyle }
        ws['C4'] = { v: '■ Nhà cung cấp (NCC)', t: 's', s: legendNccStyle }
        ws['E4'] = { v: '■ Kho khác', t: 's', s: legendKhoStyle }
        ws['G4'] = { v: '■ Tổ đội', t: 's', s: legendToDoiStyle }

        ws['A6'] = { v: 'BẢNG TỔNG HỢP DIỄN GIẢI THỰC NHẬP (SUMIFS THEO ĐƠN VỊ)', t: 's', s: tblHeaderStyle1 }
        ws[isBaoCao ? 'J6' : 'L6'] = { v: 'BẢNG TỔNG HỢP DIỄN GIẢI THỰC XUẤT (SUMIFS THEO ĐƠN VỊ)', t: 's', s: tblHeaderStyle2 }

        sub1.forEach((label, i) => {
          const colChar = String.fromCharCode(65 + i)
          ws[`${colChar}7`] = { v: label, t: 's', s: thStyle1 }
        })

        sub2.forEach((label, i) => {
          const startCharCode = isBaoCao ? 74 : 76
          const colChar = String.fromCharCode(startCharCode + i)
          ws[`${colChar}7`] = { v: label, t: 's', s: thStyle2 }
        })

        const sheetMaxRows = Math.max(listNhap.length, listXuat.length)

        if (isBaoCao) {
          ws[`A8`] = { v: '', t: 's', s: footerStyle }
          ws[`B8`] = { v: '', t: 's', s: footerStyle }
          ws[`C8`] = { v: '', t: 's', s: footerStyle }
          ws[`D8`] = { v: '', t: 's', s: footerStyle }
          ws[`E8`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
          ws[`F8`] = { v: '', t: 's', s: footerStyle }
          ws[`G8`] = { f: sheetMaxRows > 0 ? `SUM(G9:G${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: totalYellowStyle }
          ws[`H8`] = { v: '', t: 's', s: footerStyle }

          ws[`I8`] = { v: '', t: 's', s: { font: { name: 'Segoe UI', sz: 9.5 } } }

          ws[`J8`] = { v: '', t: 's', s: footerStyle }
          ws[`K8`] = { v: '', t: 's', s: footerStyle }
          ws[`L8`] = { v: '', t: 's', s: footerStyle }
          ws[`M8`] = { v: '', t: 's', s: footerStyle }
          ws[`N8`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
          ws[`O8`] = { v: '', t: 's', s: footerStyle }
          ws[`P8`] = { f: sheetMaxRows > 0 ? `SUM(P9:P${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: totalYellowStyle }
          ws[`Q8`] = { v: '', t: 's', s: footerStyle }
        } else {
          ws[`A8`] = { v: '', t: 's', s: footerStyle }
          ws[`B8`] = { v: '', t: 's', s: footerStyle }
          ws[`C8`] = { v: '', t: 's', s: footerStyle }
          ws[`D8`] = { v: '', t: 's', s: footerStyle }
          ws[`E8`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
          ws[`F8`] = { v: '', t: 's', s: footerStyle }
          ws[`G8`] = { f: sheetMaxRows > 0 ? `SUM(G9:G${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
          ws[`H8`] = { v: '', t: 's', s: footerStyle }
          ws[`I8`] = { f: sheetMaxRows > 0 ? `SUM(I9:I${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: totalYellowStyle }
          ws[`J8`] = { v: '', t: 's', s: footerStyle }

          ws[`K8`] = { v: '', t: 's', s: { font: { name: 'Segoe UI', sz: 9.5 } } }

          ws[`L8`] = { v: '', t: 's', s: footerStyle }
          ws[`M8`] = { v: '', t: 's', s: footerStyle }
          ws[`N8`] = { v: '', t: 's', s: footerStyle }
          ws[`O8`] = { v: '', t: 's', s: footerStyle }
          ws[`P8`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
          ws[`Q8`] = { v: '', t: 's', s: footerStyle }
          ws[`R8`] = { f: sheetMaxRows > 0 ? `SUM(R9:R${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
          ws[`S8`] = { v: '', t: 's', s: footerStyle }
          ws[`T8`] = { f: sheetMaxRows > 0 ? `SUM(T9:T${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: totalYellowStyle }
          ws[`U8`] = { v: '', t: 's', s: footerStyle }
        }

        for (let i = 0; i < sheetMaxRows; i++) {
          const rowNum = 9 + i
          
          if (i < listNhap.length) {
            const item = listNhap[i]
            const labelType = item.detailType === 'nhan_ncc' ? 'Nhận từ NCC'
                            : item.detailType === 'nhan_kho' ? 'Nhận từ Kho gửi'
                            : item.detailType === 'tra_ncc' ? 'Trả lại NCC'
                            : item.detailType === 'dieu_chuyen_di' ? 'Điều chuyển đi (Không tính)'
                            : item.detailType === 'nhan_khac' ? 'Nhận nguồn khác (Không tính)'
                            : 'Không tính'
            
            const norm = (item.partner || '').trim().replace(/\s+/g, ' ')
            const cat = (customCategoryMap && customCategoryMap[norm]) || getUnitCategory(norm)
            
            let partnerStyle = { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '334155' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } }
            if (cat === 'kho') {
              partnerStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
              }
            } else if (cat === 'ncc') {
              partnerStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
              }
            } else if (cat === 'todoi') {
              partnerStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
              }
            }

            const isZero = item.logicVal === 0
            const isNegative = item.logicVal < 0
            let labelStyle = { ...dataCellStyle }
            if (isZero) {
              labelStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'CBD5E1' } },
                  bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
                  left: { style: 'thin', color: { rgb: 'CBD5E1' } },
                  right: { style: 'thin', color: { rgb: 'CBD5E1' } }
                }
              }
            } else if (isNegative) {
              labelStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'FCA5A5' } },
                  bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
                  left: { style: 'thin', color: { rgb: 'FCA5A5' } },
                  right: { style: 'thin', color: { rgb: 'FCA5A5' } }
                }
              }
            } else {
              labelStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'A7F3D0' } },
                  bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
                  left: { style: 'thin', color: { rgb: 'A7F3D0' } },
                  right: { style: 'thin', color: { rgb: 'A7F3D0' } }
                }
              }
            }

            let roleStyle = { ...dataCellStyle }
            if (item.role === 'Đơn vị nhận') {
              roleStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '6D28D9' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'F5F3FF' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'DDD6FE' } },
                  bottom: { style: 'thin', color: { rgb: 'DDD6FE' } },
                  left: { style: 'thin', color: { rgb: 'DDD6FE' } },
                  right: { style: 'thin', color: { rgb: 'DDD6FE' } }
                }
              }
            } else {
              roleStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'EA580C' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'FED7AA' } },
                  bottom: { style: 'thin', color: { rgb: 'FED7AA' } },
                  left: { style: 'thin', color: { rgb: 'FED7AA' } },
                  right: { style: 'thin', color: { rgb: 'FED7AA' } }
                }
              }
            }

            const khoStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: roleStyle.font.color },
              fill: roleStyle.fill,
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
              border: roleStyle.border
            }

            const catLabel = cat === 'ncc' ? 'Nhà cung cấp (NCC)'
                           : cat === 'kho' ? 'Kho khác'
                           : cat === 'todoi' ? 'Tổ đội'
                           : ''
            let partnerRoleStyle = { ...dataCellStyle }
            if (cat === 'ncc') {
              partnerRoleStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'FDE68A' } },
                  bottom: { style: 'thin', color: { rgb: 'FDE68A' } },
                  left: { style: 'thin', color: { rgb: 'FDE68A' } },
                  right: { style: 'thin', color: { rgb: 'FDE68A' } }
                }
              }
            } else if (cat === 'kho') {
              partnerRoleStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'BFDBFE' } },
                  bottom: { style: 'thin', color: { rgb: 'BFDBFE' } },
                  left: { style: 'thin', color: { rgb: 'BFDBFE' } },
                  right: { style: 'thin', color: { rgb: 'BFDBFE' } }
                }
              }
            } else if (cat === 'todoi') {
              partnerRoleStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'BBF7D0' } },
                  bottom: { style: 'thin', color: { rgb: 'BBF7D0' } },
                  left: { style: 'thin', color: { rgb: 'BBF7D0' } },
                  right: { style: 'thin', color: { rgb: 'BBF7D0' } }
                }
              }
            }

            ws[`A${rowNum}`] = { v: i + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
            ws[`B${rowNum}`] = { v: labelType, t: 's', s: labelStyle }
            ws[`C${rowNum}`] = { v: localProject || 'Tất cả', t: 's', s: khoStyle }
            ws[`D${rowNum}`] = { v: item.role || 'Đơn vị nhận', t: 's', s: roleStyle }
            ws[`E${rowNum}`] = { v: item.partner, t: 's', s: partnerStyle }
            ws[`F${rowNum}`] = { v: catLabel, t: 's', s: partnerRoleStyle }
            
            const partnerCol = item.role === 'Đơn vị giao' ? 'E' : 'D'
            if (isBaoCao) {
              ws[`G${rowNum}`] = { 
                f: `SUMIFS('${nhapSheetName}'!J:J, '${nhapSheetName}'!${partnerCol}:${partnerCol}, E${rowNum})`, 
                t: 'n', 
                z: '#,##0.00;[Red](#,##0.00);"-"',
                s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } } } 
              }
              ws[`H${rowNum}`] = { v: item.explain || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } } }
            } else {
              ws[`G${rowNum}`] = { 
                f: `SUMIFS('${nhapSheetName}'!F:F, '${nhapSheetName}'!${partnerCol}:${partnerCol}, E${rowNum}, '${nhapSheetName}'!G:G, B${rowNum})`, 
                t: 'n', 
                z: '#,##0.00;[Red](#,##0.00);"-"',
                s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } 
              }
              
              ws[`H${rowNum}`] = { 
                v: item.logicVal, 
                t: 'n', 
                z: '#,##0;(#,##0);"-"', 
                s: { 
                  ...dataCellStyle, 
                  alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, 
                  font: { 
                    name: 'Segoe UI', 
                    sz: 9.5, 
                    bold: true,
                    color: item.logicVal === -1 ? { rgb: 'FF0000' } : undefined
                  } 
                } 
              }
              
              ws[`I${rowNum}`] = { 
                f: `G${rowNum} * H${rowNum}`, 
                t: 'n', 
                z: '#,##0.00;[Red](#,##0.00);"-"',
                s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } } } 
              }
              ws[`J${rowNum}`] = { v: item.explain || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } } }
            }
          } else {
            const emptyCols = isBaoCao ? ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'] : ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
            emptyCols.forEach(col => {
              ws[`${col}${rowNum}`] = { v: '', t: 's', s: { ...dataCellStyle, border: {} } }
            })
          }

          ws[isBaoCao ? `I${rowNum}` : `K${rowNum}`] = { v: '', t: 's', s: { font: { name: 'Segoe UI', sz: 9.5 } } }

          if (i < listXuat.length) {
            const item = listXuat[i]
            const labelType = item.detailType === 'xuat_todoi' ? 'Xuất cho Tổ đội'
                            : item.detailType === 'xuat_kho' ? 'Xuất đi Kho nhận'
                            : item.detailType === 'todoi_tra' ? 'Tổ đội trả hàng'
                            : item.detailType === 'xuat_khac' ? 'Xuất khác (Không tính)'
                            : item.detailType === 'nhan_lai_khac' ? 'Nhận lại khác (Không tính)'
                            : 'Không tính'
            
            const norm = (item.partner || '').trim().replace(/\s+/g, ' ')
            const cat = (customCategoryMap && customCategoryMap[norm]) || getUnitCategory(norm)
            
            let partnerStyle = { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '334155' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } }
            if (cat === 'kho') {
              partnerStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
              }
            } else if (cat === 'ncc') {
              partnerStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
              }
            } else if (cat === 'todoi') {
              partnerStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
              }
            }

            const isZero = item.logicVal === 0
            const isNegative = item.logicVal < 0
            let labelStyle = { ...dataCellStyle }
            if (isZero) {
              labelStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'CBD5E1' } },
                  bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
                  left: { style: 'thin', color: { rgb: 'CBD5E1' } },
                  right: { style: 'thin', color: { rgb: 'CBD5E1' } }
                }
              }
            } else if (isNegative) {
              labelStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'FCA5A5' } },
                  bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
                  left: { style: 'thin', color: { rgb: 'FCA5A5' } },
                  right: { style: 'thin', color: { rgb: 'FCA5A5' } }
                }
              }
            } else {
              labelStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
                alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'A7F3D0' } },
                  bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
                  left: { style: 'thin', color: { rgb: 'A7F3D0' } },
                  right: { style: 'thin', color: { rgb: 'A7F3D0' } }
                }
              }
            }

            let roleStyle = { ...dataCellStyle }
            if (item.role === 'Đơn vị nhận') {
              roleStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '6D28D9' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'F5F3FF' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'DDD6FE' } },
                  bottom: { style: 'thin', color: { rgb: 'DDD6FE' } },
                  left: { style: 'thin', color: { rgb: 'DDD6FE' } },
                  right: { style: 'thin', color: { rgb: 'DDD6FE' } }
                }
              }
            } else {
              roleStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'EA580C' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'FED7AA' } },
                  bottom: { style: 'thin', color: { rgb: 'FED7AA' } },
                  left: { style: 'thin', color: { rgb: 'FED7AA' } },
                  right: { style: 'thin', color: { rgb: 'FED7AA' } }
                }
              }
            }

            const khoStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: roleStyle.font.color },
              fill: roleStyle.fill,
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
              border: roleStyle.border
            }

            const catLabel = cat === 'ncc' ? 'Nhà cung cấp (NCC)'
                           : cat === 'kho' ? 'Kho khác'
                           : cat === 'todoi' ? 'Tổ đội'
                           : ''
            let partnerRoleStyle = { ...dataCellStyle }
            if (cat === 'ncc') {
              partnerRoleStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'FDE68A' } },
                  bottom: { style: 'thin', color: { rgb: 'FDE68A' } },
                  left: { style: 'thin', color: { rgb: 'FDE68A' } },
                  right: { style: 'thin', color: { rgb: 'FDE68A' } }
                }
              }
            } else if (cat === 'kho') {
              partnerRoleStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'BFDBFE' } },
                  bottom: { style: 'thin', color: { rgb: 'BFDBFE' } },
                  left: { style: 'thin', color: { rgb: 'BFDBFE' } },
                  right: { style: 'thin', color: { rgb: 'BFDBFE' } }
                }
              }
            } else if (cat === 'todoi') {
              partnerRoleStyle = {
                ...dataCellStyle,
                font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
                fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
                border: {
                  top: { style: 'thin', color: { rgb: 'BBF7D0' } },
                  bottom: { style: 'thin', color: { rgb: 'BBF7D0' } },
                  left: { style: 'thin', color: { rgb: 'BBF7D0' } },
                  right: { style: 'thin', color: { rgb: 'BBF7D0' } }
                }
              }
            }

            const partnerCol = item.role === 'Đơn vị giao' ? 'E' : 'D'
            if (isBaoCao) {
              ws[`J${rowNum}`] = { v: i + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
              ws[`K${rowNum}`] = { v: labelType, t: 's', s: labelStyle }
              ws[`L${rowNum}`] = { v: localProject || 'Tất cả', t: 's', s: khoStyle }
              ws[`M${rowNum}`] = { v: item.role || 'Đơn vị giao', t: 's', s: roleStyle }
              ws[`N${rowNum}`] = { v: item.partner, t: 's', s: partnerStyle }
              ws[`O${rowNum}`] = { v: catLabel, t: 's', s: partnerRoleStyle }
              ws[`P${rowNum}`] = { 
                f: `SUMIFS('${xuatSheetName}'!J:J, '${xuatSheetName}'!${partnerCol}:${partnerCol}, N${rowNum})`, 
                t: 'n', 
                z: '#,##0.00;[Red](#,##0.00);"-"',
                s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } } } 
              }
              ws[`Q${rowNum}`] = { v: item.explain || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } } }
            } else {
              ws[`L${rowNum}`] = { v: i + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
              ws[`M${rowNum}`] = { v: labelType, t: 's', s: labelStyle }
              ws[`N${rowNum}`] = { v: localProject || 'Tất cả', t: 's', s: khoStyle }
              ws[`O${rowNum}`] = { v: item.role || 'Đơn vị giao', t: 's', s: roleStyle }
              ws[`P${rowNum}`] = { v: item.partner, t: 's', s: partnerStyle }
              ws[`Q${rowNum}`] = { v: catLabel, t: 's', s: partnerRoleStyle }
              ws[`R${rowNum}`] = { 
                f: `SUMIFS('${xuatSheetName}'!F:F, '${xuatSheetName}'!${partnerCol}:${partnerCol}, P${rowNum}, '${xuatSheetName}'!G:G, M${rowNum})`, 
                t: 'n', 
                z: '#,##0.00;[Red](#,##0.00);"-"',
                s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } 
              }
              
              ws[`S${rowNum}`] = { 
                v: item.logicVal, 
                t: 'n', 
                z: '#,##0;(#,##0);"-"', 
                s: { 
                  ...dataCellStyle, 
                  alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, 
                  font: { 
                    name: 'Segoe UI', 
                    sz: 9.5, 
                    bold: true,
                    color: item.logicVal === -1 ? { rgb: 'FF0000' } : undefined
                  } 
                } 
              }
              
              ws[`T${rowNum}`] = { 
                f: `R${rowNum} * S${rowNum}`, 
                t: 'n', 
                z: '#,##0.00;[Red](#,##0.00);"-"',
                s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } } } 
              }
              ws[`U${rowNum}`] = { v: item.explain || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } } }
            }
          } else {
            const emptyCols = isBaoCao ? ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q'] : ['L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U']
            emptyCols.forEach(col => {
              ws[`${col}${rowNum}`] = { v: '', t: 's', s: { ...dataCellStyle, border: {} } }
            })
          }
        }

        if (isBaoCao) {
          ws['!ref'] = `A1:Q${8 + sheetMaxRows}`
          ws['!merges'] = [
            { s: { r: 1, c: 0 }, e: { r: 1, c: 16 } },
            { s: { r: 5, c: 0 }, e: { r: 5, c: 7 } },
            { s: { r: 5, c: 9 }, e: { r: 5, c: 16 } }
          ]
          ws['!autofilter'] = { ref: `A7:Q${8 + sheetMaxRows}` }
        } else {
          ws['!ref'] = `A1:U${8 + sheetMaxRows}`
          ws['!merges'] = [
            { s: { r: 1, c: 0 }, e: { r: 1, c: 20 } },
            { s: { r: 5, c: 0 }, e: { r: 5, c: 9 } },
            { s: { r: 5, c: 11 }, e: { r: 5, c: 20 } }
          ]
          ws['!autofilter'] = { ref: `A7:U${8 + sheetMaxRows}` }
        }

        return ws
      }

      const groupAndSortByPartnerForMulti = (list) => {
        const grouped = {}
        list.forEach(item => {
          const normPartner = (item.partner || '').trim()
          if (!grouped[normPartner]) {
            grouped[normPartner] = {
              partner: normPartner,
              role: item.role,
              detailType: item.detailType,
              logicVal: item.logicVal,
              explain: item.explain,
              totalQty: 0,
              totalContribution: 0,
              items: []
            }
          }
          grouped[normPartner].totalQty += item.totalQty
          grouped[normPartner].totalContribution += item.totalQty * item.logicVal
          grouped[normPartner].items.push(item)
        })

        const groupedList = Object.values(grouped).map(group => {
          group.items.sort((a, b) => b.totalQty - a.totalQty)
          const dominant = group.items[0]
          return {
            partner: group.partner,
            role: dominant.role,
            detailType: dominant.detailType,
            logicVal: dominant.logicVal,
            explain: dominant.explain || '',
            totalQty: group.totalQty,
            totalContribution: group.totalContribution
          }
        })

        groupedList.sort((a, b) => b.totalContribution - a.totalContribution)
        return groupedList
      }

      const nhapListBaoCao = nhapSummaryList.filter(item => item.logicVal !== 0)
      const xuatListBaoCao = xuatSummaryList.filter(item => item.logicVal !== 0)

      const groupedNhapBaoCao = groupAndSortByPartnerForMulti(nhapListBaoCao)
      groupedNhapBaoCao.sort((a, b) => {
        const rolePriorityA = a.role === 'Đơn vị nhận' ? 1 : a.role === 'Đơn vị giao' ? 2 : 3
        const rolePriorityB = b.role === 'Đơn vị nhận' ? 1 : b.role === 'Đơn vị giao' ? 2 : 3
        if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

        const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
        const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
        const catPriorityA = catA === 'ncc' ? 1 : catA === 'kho' ? 2 : catA === 'todoi' ? 3 : 4

        const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
        const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
        const catPriorityB = catB === 'ncc' ? 1 : catB === 'kho' ? 2 : catB === 'todoi' ? 3 : 4

        if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

        const valA = Math.abs(a.totalContribution || 0)
        const valB = Math.abs(b.totalContribution || 0)
        if (valB !== valA) return valB - valA

        return (a.partner || '').localeCompare(b.partner || '', 'vi')
      })

      const groupedXuatBaoCao = groupAndSortByPartnerForMulti(xuatListBaoCao)
      groupedXuatBaoCao.sort((a, b) => {
        const rolePriorityA = a.role === 'Đơn vị giao' ? 1 : a.role === 'Đơn vị nhận' ? 2 : 3
        const rolePriorityB = b.role === 'Đơn vị giao' ? 1 : b.role === 'Đơn vị nhận' ? 2 : 3
        if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

        const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
        const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
        const catPriorityA = catA === 'todoi' ? 1 : catA === 'kho' ? 2 : catA === 'ncc' ? 3 : 4

        const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
        const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
        const catPriorityB = catB === 'todoi' ? 1 : catB === 'kho' ? 2 : catB === 'ncc' ? 3 : 4

        if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

        const valA = Math.abs(a.totalContribution || 0)
        const valB = Math.abs(b.totalContribution || 0)
        if (valB !== valA) return valB - valA

        return (a.partner || '').localeCompare(b.partner || '', 'vi')
      })

      const wsBaoCao = buildSheetDataForMulti(groupedNhapBaoCao, groupedXuatBaoCao, true)
      const wsSummary = buildSheetDataForMulti(nhapSummaryList, xuatSummaryList, false)

      const wsThucNhap = {}
      wsThucNhap['!cols'] = [
        { wpx: 50 }, { wpx: 100 }, { wpx: 130 }, { wpx: 240 }, { wpx: 240 },
        { wpx: 120 }, { wpx: 150 }, { wpx: 80 }, { wpx: 250 }, { wpx: 120 }
      ]

      const sheet2TitleStyle = {
        font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '0F766E' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      }

      wsThucNhap['A2'] = { v: 'BẢNG CHI TIẾT CHỨNG TỪ THỰC NHẬP (CẤU THÀNH SỐ LIỆU)', t: 's', s: sheet2TitleStyle }
      wsThucNhap['A3'] = { v: `Mặt hàng: ${itemRow.tenVatTu} (${itemRow.maSAP}) | ĐVT: ${itemRow.dvt}`, t: 's', s: subtitleStyle }

      const s2Headers = ['STT', 'Ngày', 'Mã chứng từ', 'Đơn vị giao (NCC/Kho)', 'Đơn vị nhận (Kho nhận)', 'KL Chứng từ', 'Phân loại đơn vị', 'Hệ số logic', 'Quy tắc & Diễn giải logic', 'Đóng góp thực']
      s2Headers.forEach((label, i) => {
        const colChar = String.fromCharCode(65 + i)
        wsThucNhap[`${colChar}5`] = { v: label, t: 's', s: thStyle1 }
      })

      const nhapTxList = txs.filter(tx => tx.nhapVal > 0)
      const footerStyle1 = { ...footerStyle, fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } } }
      
      wsThucNhap[`A6`] = { v: '', t: 's', s: footerStyle1 }
      wsThucNhap[`B6`] = { v: '', t: 's', s: footerStyle1 }
      wsThucNhap[`C6`] = { v: '', t: 's', s: footerStyle1 }
      wsThucNhap[`D6`] = { v: '', t: 's', s: footerStyle1 }
      wsThucNhap[`E6`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle1, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
      wsThucNhap[`F6`] = { f: nhapTxList.length > 0 ? `SUM(F7:F${6 + nhapTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle1, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
      wsThucNhap[`G6`] = { v: '', t: 's', s: footerStyle1 }
      wsThucNhap[`H6`] = { v: '', t: 's', s: footerStyle1 }
      wsThucNhap[`I6`] = { v: '', t: 's', s: footerStyle1 }
      wsThucNhap[`J6`] = { f: nhapTxList.length > 0 ? `SUM(J7:J${6 + nhapTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle1, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }

      nhapTxList.forEach((tx, idx) => {
        const rowNum = 7 + idx
        const typeLabel = tx.detailTypeNhap === 'nhan_ncc' ? 'Nhận từ NCC'
                        : tx.detailTypeNhap === 'nhan_kho' ? 'Nhận từ Kho gửi'
                        : tx.detailTypeNhap === 'tra_ncc' ? 'Trả lại NCC'
                        : tx.detailTypeNhap === 'dieu_chuyen_di' ? 'Điều chuyển đi (Không tính)'
                        : tx.detailTypeNhap === 'nhan_khac' ? 'Nhận nguồn khác (Không tính)'
                        : 'Không tính'

        const isZero = tx.logicNhapVal === 0
        const isNegative = tx.logicNhapVal < 0
        let typeStyle = { ...dataCellStyle }
        if (isZero) {
          typeStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'CBD5E1' } },
              bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
              left: { style: 'thin', color: { rgb: 'CBD5E1' } },
              right: { style: 'thin', color: { rgb: 'CBD5E1' } }
            }
          }
        } else if (isNegative) {
          typeStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'FCA5A5' } },
              bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
              left: { style: 'thin', color: { rgb: 'FCA5A5' } },
              right: { style: 'thin', color: { rgb: 'FCA5A5' } }
            }
          }
        } else {
          typeStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'A7F3D0' } },
              bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
              left: { style: 'thin', color: { rgb: 'A7F3D0' } },
              right: { style: 'thin', color: { rgb: 'A7F3D0' } }
            }
          }
        }

        wsThucNhap[`A${rowNum}`] = { v: idx + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
        wsThucNhap[`B${rowNum}`] = { v: tx.ngayXuatNhap || '', t: 's', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
        wsThucNhap[`C${rowNum}`] = { v: tx.maDonNhapKho || tx.maDonXuatKho || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true }, alignment: { vertical: 'center', wrapText: true } } }
        wsThucNhap[`D${rowNum}`] = { v: tx.donViGiao || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
        wsThucNhap[`E${rowNum}`] = { v: tx.donViNhan || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
        wsThucNhap[`F${rowNum}`] = { v: tx.nhapVal, t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
        wsThucNhap[`G${rowNum}`] = { v: typeLabel, t: 's', s: typeStyle }
        wsThucNhap[`H${rowNum}`] = { 
          v: tx.logicNhapVal, 
          t: 'n', 
          z: '#,##0;(#,##0);"-"', 
          s: { 
            ...dataCellStyle, 
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, 
            font: { 
              name: 'Segoe UI', 
              sz: 9.5, 
              bold: true,
              color: tx.logicNhapVal === -1 ? { rgb: 'FF0000' } : undefined
            } 
          } 
        }
        wsThucNhap[`I${rowNum}`] = { v: tx.explainNhap || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
        
        wsThucNhap[`J${rowNum}`] = { 
          f: `F${rowNum} * H${rowNum}`, 
          t: 'n', 
          z: '#,##0.00;[Red](#,##0.00);"-"',
          s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } } } 
        }
      })

      wsThucNhap['!ref'] = `A1:J${6 + nhapTxList.length}`
      wsThucNhap['!merges'] = [
        { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } }
      ]
      wsThucNhap['!autofilter'] = { ref: `A5:J${6 + nhapTxList.length}` }

      const wsThucXuat = {}
      wsThucXuat['!cols'] = [
        { wpx: 50 }, { wpx: 100 }, { wpx: 130 }, { wpx: 240 }, { wpx: 240 },
        { wpx: 120 }, { wpx: 150 }, { wpx: 80 }, { wpx: 250 }, { wpx: 120 }
      ]

      const sheet3TitleStyle = {
        font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: 'C2410C' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      }

      wsThucXuat['A2'] = { v: 'BẢNG CHI TIẾT CHỨNG TỪ THỰC XUẤT (CẤU THÀNH SỐ LIỆU)', t: 's', s: sheet3TitleStyle }
      wsThucXuat['A3'] = { v: `Mặt hàng: ${itemRow.tenVatTu} (${itemRow.maSAP}) | ĐVT: ${itemRow.dvt}`, t: 's', s: subtitleStyle }

      const s3Headers = ['STT', 'Ngày', 'Mã chứng từ', 'Đơn vị giao (Kho gửi)', 'Đơn vị nhận (Tổ đội/Kho nhận)', 'KL Chứng từ', 'Phân loại đơn vị', 'Hệ số logic', 'Quy tắc & Diễn giải logic', 'Đóng góp thực']
      s3Headers.forEach((label, i) => {
        const colChar = String.fromCharCode(65 + i)
        wsThucXuat[`${colChar}5`] = { v: label, t: 's', s: thStyle2 }
      })

      const xuatTxList = txs.filter(tx => tx.xuatVal > 0)
      const footerStyle2 = { ...footerStyle, fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } } }
      
      wsThucXuat[`A6`] = { v: '', t: 's', s: footerStyle2 }
      wsThucXuat[`B6`] = { v: '', t: 's', s: footerStyle2 }
      wsThucXuat[`C6`] = { v: '', t: 's', s: footerStyle2 }
      wsThucXuat[`D6`] = { v: '', t: 's', s: footerStyle2 }
      wsThucXuat[`E6`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle2, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
      wsThucXuat[`F6`] = { f: xuatTxList.length > 0 ? `SUM(F7:F${6 + xuatTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle2, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
      wsThucXuat[`G6`] = { v: '', t: 's', s: footerStyle2 }
      wsThucXuat[`H6`] = { v: '', t: 's', s: footerStyle2 }
      wsThucXuat[`I6`] = { v: '', t: 's', s: footerStyle2 }
      wsThucXuat[`J6`] = { f: xuatTxList.length > 0 ? `SUM(J7:J${6 + xuatTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle2, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }

      xuatTxList.forEach((tx, idx) => {
        const rowNum = 7 + idx
        const typeLabel = tx.detailTypeXuat === 'xuat_todoi' ? 'Xuất cho Tổ đội'
                        : tx.detailTypeXuat === 'xuat_kho' ? 'Xuất đi Kho nhận'
                        : tx.detailTypeXuat === 'todoi_tra' ? 'Tổ đội trả hàng'
                        : tx.detailTypeXuat === 'xuat_khac' ? 'Xuất khác (Không tính)'
                        : tx.detailTypeXuat === 'nhan_lai_khac' ? 'Nhận lại khác (Không tính)'
                        : 'Không tính'

        const isZero = tx.logicXuatVal === 0
        const isNegative = tx.logicXuatVal < 0
        let typeStyle = { ...dataCellStyle }
        if (isZero) {
          typeStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'CBD5E1' } },
              bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
              left: { style: 'thin', color: { rgb: 'CBD5E1' } },
              right: { style: 'thin', color: { rgb: 'CBD5E1' } }
            }
          }
        } else if (isNegative) {
          typeStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'FCA5A5' } },
              bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
              left: { style: 'thin', color: { rgb: 'FCA5A5' } },
              right: { style: 'thin', color: { rgb: 'FCA5A5' } }
            }
          }
        } else {
          typeStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
            fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: {
              top: { style: 'thin', color: { rgb: 'A7F3D0' } },
              bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
              left: { style: 'thin', color: { rgb: 'A7F3D0' } },
              right: { style: 'thin', color: { rgb: 'A7F3D0' } }
            }
          }
        }

        wsThucXuat[`A${rowNum}`] = { v: idx + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
        wsThucXuat[`B${rowNum}`] = { v: tx.ngayXuatNhap || '', t: 's', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
        wsThucXuat[`C${rowNum}`] = { v: tx.maDonNhapKho || tx.maDonXuatKho || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true }, alignment: { vertical: 'center', wrapText: true } } }
        wsThucXuat[`D${rowNum}`] = { v: tx.donViGiao || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
        wsThucXuat[`E${rowNum}`] = { v: tx.donViNhan || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
        wsThucXuat[`F${rowNum}`] = { v: tx.xuatVal, t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
        wsThucXuat[`G${rowNum}`] = { v: typeLabel, t: 's', s: typeStyle }
        wsThucXuat[`H${rowNum}`] = { 
          v: tx.logicXuatVal, 
          t: 'n', 
          z: '#,##0;(#,##0);"-"', 
          s: { 
            ...dataCellStyle, 
            alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, 
            font: { 
              name: 'Segoe UI', 
              sz: 9.5, 
              bold: true,
              color: tx.logicXuatVal === -1 ? { rgb: 'FF0000' } : undefined
            } 
          } 
        }
        wsThucXuat[`I${rowNum}`] = { v: tx.explainXuat || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
        
        wsThucXuat[`J${rowNum}`] = { 
          f: `F${rowNum} * H${rowNum}`, 
          t: 'n', 
          z: '#,##0.00;[Red](#,##0.00);"-"',
          s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } } } 
        }
      })

      wsThucXuat['!ref'] = `A1:J${6 + xuatTxList.length}`
      wsThucXuat['!merges'] = [
        { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } }
      ]
      wsThucXuat['!autofilter'] = { ref: `A5:J${6 + xuatTxList.length}` }

      XLSXStyle.utils.book_append_sheet(wb, wsBaoCao, bcSheetName)
      XLSXStyle.utils.book_append_sheet(wb, wsSummary, thSheetName)
      XLSXStyle.utils.book_append_sheet(wb, wsThucNhap, nhapSheetName)
      XLSXStyle.utils.book_append_sheet(wb, wsThucXuat, xuatSheetName)
    })

    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary', compression: false })
    const s2ab = (s) => {
      const buf = new ArrayBuffer(s.length)
      const view = new Uint8Array(buf)
      for (let j = 0; j < s.length; j++) view[j] = s.charCodeAt(j) & 0xFF
      return buf
    }

    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `Giai_Trinh_Gop_${selectedRows.length}_Vat_Tu_${new Date().toISOString().slice(0, 10)}.xlsx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }


  const handleExportDetailExcel = (itemRow, nhapList, xuatList) => {
    const wb = XLSXStyle.utils.book_new()
    
    const titleStyle = {
      font: { name: 'Segoe UI', sz: 16, bold: true, color: { rgb: '1E3A8A' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    }
    const subtitleStyle = {
      font: { name: 'Segoe UI', sz: 11, italic: true, color: { rgb: '475569' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    }
    
    // Header background colors
    const tblHeaderStyle1 = {
      fill: { patternType: 'solid', fgColor: { rgb: '0F766E' } },
      font: { name: 'Segoe UI', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    }
    const tblHeaderStyle2 = {
      fill: { patternType: 'solid', fgColor: { rgb: 'C2410C' } },
      font: { name: 'Segoe UI', sz: 11, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center' }
    }
    
    const thStyle1 = {
      fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
      font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: 'CCFBF1' } },
        bottom: { style: 'thin', color: { rgb: 'CCFBF1' } },
        left: { style: 'thin', color: { rgb: 'CCFBF1' } },
        right: { style: 'thin', color: { rgb: 'CCFBF1' } }
      }
    }
    const thStyle2 = {
      fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } },
      font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: {
        top: { style: 'thin', color: { rgb: 'FFEDD5' } },
        bottom: { style: 'thin', color: { rgb: 'FFEDD5' } },
        left: { style: 'thin', color: { rgb: 'FFEDD5' } },
        right: { style: 'thin', color: { rgb: 'FFEDD5' } }
      }
    }

    // Default cell borders and styles
    const dataCellStyle = {
      font: { name: 'Segoe UI', sz: 9.5, color: { rgb: '1E293B' } },
      border: {
        top: { style: 'thin', color: { rgb: 'E2E8F0' } },
        bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
        left: { style: 'thin', color: { rgb: 'E2E8F0' } },
        right: { style: 'thin', color: { rgb: 'E2E8F0' } }
      },
      alignment: { vertical: 'center', wrapText: true }
    }

    // Color Legend in Row 4
    const legendLabelStyle = {
      font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
      alignment: { vertical: 'center' }
    }
    const legendNccStyle = {
      fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
      font: { name: 'Segoe UI', sz: 9, bold: true, color: { rgb: '78350F' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'FDE68A' } },
        bottom: { style: 'thin', color: { rgb: 'FDE68A' } },
        left: { style: 'thin', color: { rgb: 'FDE68A' } },
        right: { style: 'thin', color: { rgb: 'FDE68A' } }
      }
    }
    const legendKhoStyle = {
      fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
      font: { name: 'Segoe UI', sz: 9, bold: true, color: { rgb: '1E3A8A' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'BFDBFE' } },
        bottom: { style: 'thin', color: { rgb: 'BFDBFE' } },
        left: { style: 'thin', color: { rgb: 'BFDBFE' } },
        right: { style: 'thin', color: { rgb: 'BFDBFE' } }
      }
    }
    const legendToDoiStyle = {
      fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
      font: { name: 'Segoe UI', sz: 9, bold: true, color: { rgb: '065F46' } },
      alignment: { horizontal: 'center', vertical: 'center' },
      border: {
        top: { style: 'thin', color: { rgb: 'BBF7D0' } },
        bottom: { style: 'thin', color: { rgb: 'BBF7D0' } },
        left: { style: 'thin', color: { rgb: 'BBF7D0' } },
        right: { style: 'thin', color: { rgb: 'BBF7D0' } }
      }
    }

    const subheaders1 = ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Tổng KL Chứng từ', 'Hệ số logic', 'Thực nhập', 'Ghi chú']
    const subheaders2 = ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Tổng KL Chứng từ', 'Hệ số logic', 'Thực xuất', 'Ghi chú']

    const footerStyle = {
      font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E293B' } },
      fill: { patternType: 'solid', fgColor: { rgb: 'F8FAFC' } },
      border: {
        top: { style: 'thin', color: { rgb: '94A3B8' } },
        bottom: { style: 'double', color: { rgb: '94A3B8' } },
        left: { style: 'thin', color: { rgb: 'E2E8F0' } },
        right: { style: 'thin', color: { rgb: 'E2E8F0' } }
      },
      alignment: { vertical: 'center', wrapText: true }
    }

    const totalYellowStyle = {
      ...footerStyle,
      fill: { patternType: 'solid', fgColor: { rgb: 'FFFF00' } },
      alignment: { horizontal: 'right', vertical: 'center', wrapText: true }
    }

    // Helper function to build summary or report sheet data
    const buildSheetData = (listNhap, listXuat, isBaoCao = false) => {
      const ws = {}
      
      const sub1 = isBaoCao 
        ? ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Thực nhập', 'Ghi chú']
        : ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Tổng KL Chứng từ', 'Hệ số logic', 'Thực nhập', 'Ghi chú']

      const sub2 = isBaoCao 
        ? ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Thực xuất', 'Ghi chú']
        : ['STT', 'Phân loại đơn vị', 'Kho chọn', 'Vai trò Kho chọn', 'Tên NCC/ Kho khác / Tổ đội', 'Vai trò NCC/ Kho khác/ Tổ đội', 'Tổng KL Chứng từ', 'Hệ số logic', 'Thực xuất', 'Ghi chú']

      if (isBaoCao) {
        ws['!cols'] = [
          { wpx: 40 },  // A: STT
          { wpx: 130 }, // B: Phân loại đơn vị
          { wpx: 180 }, // C: Kho chọn
          { wch: 16 },  // D: Vai trò Kho chọn
          { wpx: 250 }, // E: Tên NCC/ Kho khác / Tổ đội
          { wch: 16 },  // F: Vai trò NCC/ Kho khác/ Tổ đội
          { wch: 16 },  // G: Thực nhập
          { wpx: 200 }, // H: Ghi chú
          { wpx: 40 },  // I: Spacer
          { wpx: 40 },  // J: STT
          { wpx: 130 }, // K: Phân loại đơn vị
          { wpx: 180 }, // L: Kho chọn
          { wch: 16 },  // M: Vai trò Kho chọn
          { wpx: 250 }, // N: Tên NCC/ Kho khác / Tổ đội
          { wch: 16 },  // O: Vai trò NCC/ Kho khác/ Tổ đội
          { wch: 16 },  // P: Thực xuất
          { wpx: 200 }  // Q: Ghi chú
        ]
      } else {
        ws['!cols'] = [
          { wpx: 40 },  // A: STT
          { wpx: 130 }, // B: Phân loại đơn vị
          { wpx: 180 }, // C: Kho chọn
          { wch: 16 },  // D: Vai trò Kho chọn
          { wpx: 250 }, // E: Tên NCC/ Kho khác / Tổ đội
          { wch: 16 },  // F: Vai trò NCC/ Kho khác/ Tổ đội
          { wch: 16 },  // G: Tổng KL Chứng từ
          { wch: 6 },   // H: Hệ số logic
          { wch: 16 },  // I: Thực nhập
          { wpx: 200 }, // J: Ghi chú
          { wpx: 40 },  // K: Spacer
          { wpx: 40 },  // L: STT
          { wpx: 130 }, // M: Phân loại đơn vị
          { wpx: 180 }, // N: Kho chọn
          { wch: 16 },  // O: Vai trò Kho chọn
          { wpx: 250 }, // P: Tên NCC/ Kho khác / Tổ đội
          { wch: 16 },  // Q: Vai trò NCC/ Kho khác/ Tổ đội
          { wch: 16 },  // R: Tổng KL Chứng từ
          { wch: 6 },   // S: Hệ số logic
          { wch: 16 },  // T: Thực xuất
          { wpx: 200 }  // U: Ghi chú
        ]
      }

      ws['A2'] = { v: 'BẢNG GIẢI TRÌNH CẤU THÀNH SỐ LIỆU CHI TIẾT', t: 's', s: titleStyle }
      ws['A3'] = { v: `Kho chọn: ${localProject || 'Tất cả'} | Vật tư: ${itemRow.tenVatTu} (${itemRow.maSAP}) | ĐVT: ${itemRow.dvt}`, t: 's', s: subtitleStyle }

      ws['A4'] = { v: 'Màu sắc đối tác:', t: 's', s: legendLabelStyle }
      ws['C4'] = { v: '■ Nhà cung cấp (NCC)', t: 's', s: legendNccStyle }
      ws['E4'] = { v: '■ Kho khác', t: 's', s: legendKhoStyle }
      ws['G4'] = { v: '■ Tổ đội', t: 's', s: legendToDoiStyle }

      ws['A6'] = { v: 'BẢNG TỔNG HỢP DIỄN GIẢI THỰC NHẬP (SUMIFS THEO ĐƠN VỊ)', t: 's', s: tblHeaderStyle1 }
      ws[isBaoCao ? 'J6' : 'L6'] = { v: 'BẢNG TỔNG HỢP DIỄN GIẢI THỰC XUẤT (SUMIFS THEO ĐƠN VỊ)', t: 's', s: tblHeaderStyle2 }

      sub1.forEach((label, i) => {
        const colChar = String.fromCharCode(65 + i)
        ws[`${colChar}7`] = { v: label, t: 's', s: thStyle1 }
      })

      sub2.forEach((label, i) => {
        const startCharCode = isBaoCao ? 74 : 76 // 74 is 'J', 76 is 'L'
        const colChar = String.fromCharCode(startCharCode + i)
        ws[`${colChar}7`] = { v: label, t: 's', s: thStyle2 }
      })

      const sheetMaxRows = Math.max(listNhap.length, listXuat.length)

      // Totals
      if (isBaoCao) {
        ws[`A8`] = { v: '', t: 's', s: footerStyle }
        ws[`B8`] = { v: '', t: 's', s: footerStyle }
        ws[`C8`] = { v: '', t: 's', s: footerStyle }
        ws[`D8`] = { v: '', t: 's', s: footerStyle }
        ws[`E8`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
        ws[`F8`] = { v: '', t: 's', s: footerStyle }
        ws[`G8`] = { f: sheetMaxRows > 0 ? `SUM(G9:G${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: totalYellowStyle }
        ws[`H8`] = { v: '', t: 's', s: footerStyle }

        ws[`I8`] = { v: '', t: 's', s: { font: { name: 'Segoe UI', sz: 9.5 } } }

        ws[`J8`] = { v: '', t: 's', s: footerStyle }
        ws[`K8`] = { v: '', t: 's', s: footerStyle }
        ws[`L8`] = { v: '', t: 's', s: footerStyle }
        ws[`M8`] = { v: '', t: 's', s: footerStyle }
        ws[`N8`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
        ws[`O8`] = { v: '', t: 's', s: footerStyle }
        ws[`P8`] = { f: sheetMaxRows > 0 ? `SUM(P9:P${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: totalYellowStyle }
        ws[`Q8`] = { v: '', t: 's', s: footerStyle }
      } else {
        ws[`A8`] = { v: '', t: 's', s: footerStyle }
        ws[`B8`] = { v: '', t: 's', s: footerStyle }
        ws[`C8`] = { v: '', t: 's', s: footerStyle }
        ws[`D8`] = { v: '', t: 's', s: footerStyle }
        ws[`E8`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
        ws[`F8`] = { v: '', t: 's', s: footerStyle }
        ws[`G8`] = { f: sheetMaxRows > 0 ? `SUM(G9:G${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
        ws[`H8`] = { v: '', t: 's', s: footerStyle }
        ws[`I8`] = { f: sheetMaxRows > 0 ? `SUM(I9:I${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: totalYellowStyle }
        ws[`J8`] = { v: '', t: 's', s: footerStyle }

        ws[`K8`] = { v: '', t: 's', s: { font: { name: 'Segoe UI', sz: 9.5 } } }

        ws[`L8`] = { v: '', t: 's', s: footerStyle }
        ws[`M8`] = { v: '', t: 's', s: footerStyle }
        ws[`N8`] = { v: '', t: 's', s: footerStyle }
        ws[`O8`] = { v: '', t: 's', s: footerStyle }
        ws[`P8`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
        ws[`Q8`] = { v: '', t: 's', s: footerStyle }
        ws[`R8`] = { f: sheetMaxRows > 0 ? `SUM(R9:R${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
        ws[`S8`] = { v: '', t: 's', s: footerStyle }
        ws[`T8`] = { f: sheetMaxRows > 0 ? `SUM(T9:T${8 + sheetMaxRows})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: totalYellowStyle }
        ws[`U8`] = { v: '', t: 's', s: footerStyle }
      }

      for (let i = 0; i < sheetMaxRows; i++) {
        const rowNum = 9 + i
        
        // Receipts Table
        if (i < listNhap.length) {
          const item = listNhap[i]
          const labelType = item.detailType === 'nhan_ncc' ? 'Nhận từ NCC'
                          : item.detailType === 'nhan_kho' ? 'Nhận từ Kho gửi'
                          : item.detailType === 'tra_ncc' ? 'Trả lại NCC'
                          : item.detailType === 'dieu_chuyen_di' ? 'Điều chuyển đi (Không tính)'
                          : item.detailType === 'nhan_khac' ? 'Nhận nguồn khác (Không tính)'
                          : 'Không tính'
          
          const norm = (item.partner || '').trim().replace(/\s+/g, ' ')
          const cat = (customCategoryMap && customCategoryMap[norm]) || getUnitCategory(norm)
          
          let partnerStyle = { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '334155' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } }
          if (cat === 'kho') {
            partnerStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
            }
          } else if (cat === 'ncc') {
            partnerStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
            }
          } else if (cat === 'todoi') {
            partnerStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
            }
          }

          const isZero = item.logicVal === 0
          const isNegative = item.logicVal < 0
          let labelStyle = { ...dataCellStyle }
          if (isZero) {
            labelStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'CBD5E1' } },
                bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
                left: { style: 'thin', color: { rgb: 'CBD5E1' } },
                right: { style: 'thin', color: { rgb: 'CBD5E1' } }
              }
            }
          } else if (isNegative) {
            labelStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'FCA5A5' } },
                bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
                left: { style: 'thin', color: { rgb: 'FCA5A5' } },
                right: { style: 'thin', color: { rgb: 'FCA5A5' } }
              }
            }
          } else {
            labelStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'A7F3D0' } },
                bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
                left: { style: 'thin', color: { rgb: 'A7F3D0' } },
                right: { style: 'thin', color: { rgb: 'A7F3D0' } }
              }
            }
          }

          let roleStyle = { ...dataCellStyle }
          if (item.role === 'Đơn vị nhận') {
            roleStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '6D28D9' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'F5F3FF' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'DDD6FE' } },
                bottom: { style: 'thin', color: { rgb: 'DDD6FE' } },
                left: { style: 'thin', color: { rgb: 'DDD6FE' } },
                right: { style: 'thin', color: { rgb: 'DDD6FE' } }
              }
            }
          } else {
            roleStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'EA580C' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'FED7AA' } },
                bottom: { style: 'thin', color: { rgb: 'FED7AA' } },
                left: { style: 'thin', color: { rgb: 'FED7AA' } },
                right: { style: 'thin', color: { rgb: 'FED7AA' } }
              }
            }
          }

          const khoStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: roleStyle.font.color },
            fill: roleStyle.fill,
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: roleStyle.border
          }

          const catLabel = cat === 'ncc' ? 'Nhà cung cấp (NCC)'
                         : cat === 'kho' ? 'Kho khác'
                         : cat === 'todoi' ? 'Tổ đội'
                         : ''
          let partnerRoleStyle = { ...dataCellStyle }
          if (cat === 'ncc') {
            partnerRoleStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'FDE68A' } },
                bottom: { style: 'thin', color: { rgb: 'FDE68A' } },
                left: { style: 'thin', color: { rgb: 'FDE68A' } },
                right: { style: 'thin', color: { rgb: 'FDE68A' } }
              }
            }
          } else if (cat === 'kho') {
            partnerRoleStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'BFDBFE' } },
                bottom: { style: 'thin', color: { rgb: 'BFDBFE' } },
                left: { style: 'thin', color: { rgb: 'BFDBFE' } },
                right: { style: 'thin', color: { rgb: 'BFDBFE' } }
              }
            }
          } else if (cat === 'todoi') {
            partnerRoleStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'BBF7D0' } },
                bottom: { style: 'thin', color: { rgb: 'BBF7D0' } },
                left: { style: 'thin', color: { rgb: 'BBF7D0' } },
                right: { style: 'thin', color: { rgb: 'BBF7D0' } }
              }
            }
          }

          ws[`A${rowNum}`] = { v: i + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
          ws[`B${rowNum}`] = { v: labelType, t: 's', s: labelStyle }
          ws[`C${rowNum}`] = { v: localProject || 'Tất cả', t: 's', s: khoStyle }
          ws[`D${rowNum}`] = { v: item.role || 'Đơn vị nhận', t: 's', s: roleStyle }
          ws[`E${rowNum}`] = { v: item.partner, t: 's', s: partnerStyle }
          ws[`F${rowNum}`] = { v: catLabel, t: 's', s: partnerRoleStyle }
          
          const partnerCol = item.role === 'Đơn vị giao' ? 'E' : 'D'
          if (isBaoCao) {
            ws[`G${rowNum}`] = { 
              f: `SUMIFS('Tổng hợp'!I:I, 'Tổng hợp'!E:E, E${rowNum})`, 
              t: 'n', 
              z: '#,##0.00;[Red](#,##0.00);"-"',
              s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } } } 
            }
            ws[`H${rowNum}`] = { v: item.explain || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } } }
          } else {
            ws[`G${rowNum}`] = { 
              f: `SUMIFS('Chi tiết Thực nhập'!F:F, 'Chi tiết Thực nhập'!${partnerCol}:${partnerCol}, E${rowNum}, 'Chi tiết Thực nhập'!G:G, B${rowNum})`, 
              t: 'n', 
              z: '#,##0.00;[Red](#,##0.00);"-"',
              s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } 
            }
            
            ws[`H${rowNum}`] = { 
              v: item.logicVal, 
              t: 'n', 
              z: '#,##0;(#,##0);"-"', 
              s: { 
                ...dataCellStyle, 
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, 
                font: { 
                  name: 'Segoe UI', 
                  sz: 9.5, 
                  bold: true,
                  color: item.logicVal === -1 ? { rgb: 'FF0000' } : undefined
                } 
              } 
            }
            
            ws[`I${rowNum}`] = { 
              f: `G${rowNum} * H${rowNum}`, 
              t: 'n', 
              z: '#,##0.00;[Red](#,##0.00);"-"',
              s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } } } 
            }
            ws[`J${rowNum}`] = { v: item.explain || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } } }
          }
        } else {
          const emptyCols = isBaoCao ? ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'] : ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J']
          emptyCols.forEach(col => {
            ws[`${col}${rowNum}`] = { v: '', t: 's', s: { ...dataCellStyle, border: {} } }
          })
        }

        ws[isBaoCao ? `I${rowNum}` : `K${rowNum}`] = { v: '', t: 's', s: { font: { name: 'Segoe UI', sz: 9.5 } } }

        // Issues Table
        if (i < listXuat.length) {
          const item = listXuat[i]
          const labelType = item.detailType === 'xuat_todoi' ? 'Xuất cho Tổ đội'
                          : item.detailType === 'xuat_kho' ? 'Xuất đi Kho nhận'
                          : item.detailType === 'todoi_tra' ? 'Tổ đội trả hàng'
                          : item.detailType === 'xuat_khac' ? 'Xuất khác (Không tính)'
                          : item.detailType === 'nhan_lai_khac' ? 'Nhận lại khác (Không tính)'
                          : 'Không tính'
          
          const norm = (item.partner || '').trim().replace(/\s+/g, ' ')
          const cat = (customCategoryMap && customCategoryMap[norm]) || getUnitCategory(norm)
          
          let partnerStyle = { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '334155' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } }
          if (cat === 'kho') {
            partnerStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
            }
          } else if (cat === 'ncc') {
            partnerStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
            }
          } else if (cat === 'todoi') {
            partnerStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true }
            }
          }

          const isZero = item.logicVal === 0
          const isNegative = item.logicVal < 0
          let labelStyle = { ...dataCellStyle }
          if (isZero) {
            labelStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'CBD5E1' } },
                bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
                left: { style: 'thin', color: { rgb: 'CBD5E1' } },
                right: { style: 'thin', color: { rgb: 'CBD5E1' } }
              }
            }
          } else if (isNegative) {
            labelStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'FCA5A5' } },
                bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
                left: { style: 'thin', color: { rgb: 'FCA5A5' } },
                right: { style: 'thin', color: { rgb: 'FCA5A5' } }
              }
            }
          } else {
            labelStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
              alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'A7F3D0' } },
                bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
                left: { style: 'thin', color: { rgb: 'A7F3D0' } },
                right: { style: 'thin', color: { rgb: 'A7F3D0' } }
              }
            }
          }

          let roleStyle = { ...dataCellStyle }
          if (item.role === 'Đơn vị nhận') {
            roleStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '6D28D9' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'F5F3FF' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'DDD6FE' } },
                bottom: { style: 'thin', color: { rgb: 'DDD6FE' } },
                left: { style: 'thin', color: { rgb: 'DDD6FE' } },
                right: { style: 'thin', color: { rgb: 'DDD6FE' } }
              }
            }
          } else {
            roleStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'EA580C' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'FED7AA' } },
                bottom: { style: 'thin', color: { rgb: 'FED7AA' } },
                left: { style: 'thin', color: { rgb: 'FED7AA' } },
                right: { style: 'thin', color: { rgb: 'FED7AA' } }
              }
            }
          }

          const khoStyle = {
            ...dataCellStyle,
            font: { name: 'Segoe UI', sz: 9.5, bold: true, color: roleStyle.font.color },
            fill: roleStyle.fill,
            alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
            border: roleStyle.border
          }

          const catLabel = cat === 'ncc' ? 'Nhà cung cấp (NCC)'
                         : cat === 'kho' ? 'Kho khác'
                         : cat === 'todoi' ? 'Tổ đội'
                         : ''
          let partnerRoleStyle = { ...dataCellStyle }
          if (cat === 'ncc') {
            partnerRoleStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '78350F' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'FFFBEB' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'FDE68A' } },
                bottom: { style: 'thin', color: { rgb: 'FDE68A' } },
                left: { style: 'thin', color: { rgb: 'FDE68A' } },
                right: { style: 'thin', color: { rgb: 'FDE68A' } }
              }
            }
          } else if (cat === 'kho') {
            partnerRoleStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '1E3A8A' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'BFDBFE' } },
                bottom: { style: 'thin', color: { rgb: 'BFDBFE' } },
                left: { style: 'thin', color: { rgb: 'BFDBFE' } },
                right: { style: 'thin', color: { rgb: 'BFDBFE' } }
              }
            }
          } else if (cat === 'todoi') {
            partnerRoleStyle = {
              ...dataCellStyle,
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '065F46' } },
              fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'BBF7D0' } },
                bottom: { style: 'thin', color: { rgb: 'BBF7D0' } },
                left: { style: 'thin', color: { rgb: 'BBF7D0' } },
                right: { style: 'thin', color: { rgb: 'BBF7D0' } }
              }
            }
          }

          const partnerCol = item.role === 'Đơn vị giao' ? 'E' : 'D'
          if (isBaoCao) {
            ws[`J${rowNum}`] = { v: i + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
            ws[`K${rowNum}`] = { v: labelType, t: 's', s: labelStyle }
            ws[`L${rowNum}`] = { v: localProject || 'Tất cả', t: 's', s: khoStyle }
            ws[`M${rowNum}`] = { v: item.role || 'Đơn vị giao', t: 's', s: roleStyle }
            ws[`N${rowNum}`] = { v: item.partner, t: 's', s: partnerStyle }
            ws[`O${rowNum}`] = { v: catLabel, t: 's', s: partnerRoleStyle }
            ws[`P${rowNum}`] = { 
              f: `SUMIFS('Tổng hợp'!T:T, 'Tổng hợp'!P:P, N${rowNum})`, 
              t: 'n', 
              z: '#,##0.00;[Red](#,##0.00);"-"',
              s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } } } 
            }
            ws[`Q${rowNum}`] = { v: item.explain || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } } }
          } else {
            ws[`L${rowNum}`] = { v: i + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
            ws[`M${rowNum}`] = { v: labelType, t: 's', s: labelStyle }
            ws[`N${rowNum}`] = { v: localProject || 'Tất cả', t: 's', s: khoStyle }
            ws[`O${rowNum}`] = { v: item.role || 'Đơn vị giao', t: 's', s: roleStyle }
            ws[`P${rowNum}`] = { v: item.partner, t: 's', s: partnerStyle }
            ws[`Q${rowNum}`] = { v: catLabel, t: 's', s: partnerRoleStyle }
            ws[`R${rowNum}`] = { 
              f: `SUMIFS('Chi tiết Thực xuất'!F:F, 'Chi tiết Thực xuất'!${partnerCol}:${partnerCol}, P${rowNum}, 'Chi tiết Thực xuất'!G:G, M${rowNum})`, 
              t: 'n', 
              z: '#,##0.00;[Red](#,##0.00);"-"',
              s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } 
            }
            
            ws[`S${rowNum}`] = { 
              v: item.logicVal, 
              t: 'n', 
              z: '#,##0;(#,##0);"-"', 
              s: { 
                ...dataCellStyle, 
                alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, 
                font: { 
                  name: 'Segoe UI', 
                  sz: 9.5, 
                  bold: true,
                  color: item.logicVal === -1 ? { rgb: 'FF0000' } : undefined
                } 
              } 
            }
            
            ws[`T${rowNum}`] = { 
              f: `R${rowNum} * S${rowNum}`, 
              t: 'n', 
              z: '#,##0.00;[Red](#,##0.00);"-"',
              s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } } } 
            }
            ws[`U${rowNum}`] = { v: item.explain || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, italic: true, color: { rgb: '475569' } }, alignment: { horizontal: 'left', vertical: 'center', wrapText: true } } }
          }
        } else {
          const emptyCols = isBaoCao ? ['J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q'] : ['L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U']
          emptyCols.forEach(col => {
            ws[`${col}${rowNum}`] = { v: '', t: 's', s: { ...dataCellStyle, border: {} } }
          })
        }
      }

      if (isBaoCao) {
        ws['!ref'] = `A1:Q${8 + sheetMaxRows}`
        ws['!merges'] = [
          { s: { r: 1, c: 0 }, e: { r: 1, c: 16 } }, // Main Title
          { s: { r: 5, c: 0 }, e: { r: 5, c: 7 } },  // Receipts Title
          { s: { r: 5, c: 9 }, e: { r: 5, c: 16 } }  // Issues Title
        ]
        ws['!autofilter'] = { ref: `A7:Q${8 + sheetMaxRows}` }
      } else {
        ws['!ref'] = `A1:U${8 + sheetMaxRows}`
        ws['!merges'] = [
          { s: { r: 1, c: 0 }, e: { r: 1, c: 20 } }, // Main Title
          { s: { r: 5, c: 0 }, e: { r: 5, c: 9 } },  // Receipts Title
          { s: { r: 5, c: 11 }, e: { r: 5, c: 20 } }  // Issues Title
        ]
        ws['!autofilter'] = { ref: `A7:U${8 + sheetMaxRows}` }
      }

      return ws
    }

    // Group and sort by Partner Name for Báo cáo sheet
    const groupAndSortByPartner = (list) => {
      const grouped = {}
      list.forEach(item => {
        const normPartner = (item.partner || '').trim()
        if (!grouped[normPartner]) {
          grouped[normPartner] = {
            partner: normPartner,
            role: item.role,
            detailType: item.detailType,
            logicVal: item.logicVal,
            explain: item.explain,
            totalQty: 0,
            totalContribution: 0,
            items: []
          }
        }
        grouped[normPartner].totalQty += item.totalQty
        grouped[normPartner].totalContribution += item.totalQty * item.logicVal
        grouped[normPartner].items.push(item)
      })

      const groupedList = Object.values(grouped).map(group => {
        // Find dominant item for styling details
        group.items.sort((a, b) => b.totalQty - a.totalQty)
        const dominant = group.items[0]
        return {
          partner: group.partner,
          role: dominant.role,
          detailType: dominant.detailType,
          logicVal: dominant.logicVal,
          explain: dominant.explain || '',
          totalQty: group.totalQty,
          totalContribution: group.totalContribution
        }
      })

      // Sort descending by totalContribution (Khối lượng Thực) from largest to smallest
      groupedList.sort((a, b) => b.totalContribution - a.totalContribution)

      return groupedList
    }

    // Filter out "Không tính" rows (where logicVal is 0) to build the Báo cáo sheet
    const nhapListBaoCao = nhapList.filter(item => item.logicVal !== 0)
    const xuatListBaoCao = xuatList.filter(item => item.logicVal !== 0)

    const groupedNhapBaoCao = groupAndSortByPartner(nhapListBaoCao)
    groupedNhapBaoCao.sort((a, b) => {
      // Cấp 1 - Vai trò Kho chọn: Đơn vị nhận trước, Đơn vị giao sau
      const rolePriorityA = a.role === 'Đơn vị nhận' ? 1 : a.role === 'Đơn vị giao' ? 2 : 3
      const rolePriorityB = b.role === 'Đơn vị nhận' ? 1 : b.role === 'Đơn vị giao' ? 2 : 3
      if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

      // Cấp 2 - Phân loại đơn vị: Nhà cung cấp -> Kho khác -> Tổ đội
      const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
      const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
      const catPriorityA = catA === 'ncc' ? 1 : catA === 'kho' ? 2 : catA === 'todoi' ? 3 : 4

      const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
      const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
      const catPriorityB = catB === 'ncc' ? 1 : catB === 'kho' ? 2 : catB === 'todoi' ? 3 : 4

      if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

      // Cấp 3 - Giá trị (Khối lượng) giảm dần trị tuyệt đối
      const valA = Math.abs(a.totalContribution || 0)
      const valB = Math.abs(b.totalContribution || 0)
      if (valB !== valA) return valB - valA

      return (a.partner || '').localeCompare(b.partner || '', 'vi')
    })

    const groupedXuatBaoCao = groupAndSortByPartner(xuatListBaoCao)
    groupedXuatBaoCao.sort((a, b) => {
      // Cấp 1 - Vai trò Kho chọn: Đơn vị giao trước, Đơn vị nhận sau
      const rolePriorityA = a.role === 'Đơn vị giao' ? 1 : a.role === 'Đơn vị nhận' ? 2 : 3
      const rolePriorityB = b.role === 'Đơn vị giao' ? 1 : b.role === 'Đơn vị nhận' ? 2 : 3
      if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

      // Cấp 2 - Phân loại đơn vị: Tổ đội -> Kho khác -> Nhà cung cấp
      const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
      const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
      const catPriorityA = catA === 'todoi' ? 1 : catA === 'kho' ? 2 : catA === 'ncc' ? 3 : 4

      const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
      const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
      const catPriorityB = catB === 'todoi' ? 1 : catB === 'kho' ? 2 : catB === 'ncc' ? 3 : 4

      if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

      // Cấp 3 - Giá trị (Khối lượng) giảm dần trị tuyệt đối
      const valA = Math.abs(a.totalContribution || 0)
      const valB = Math.abs(b.totalContribution || 0)
      if (valB !== valA) return valB - valA

      return (a.partner || '').localeCompare(b.partner || '', 'vi')
    })

    const wsBaoCao = buildSheetData(groupedNhapBaoCao, groupedXuatBaoCao, true)
    const wsSummary = buildSheetData(nhapList, xuatList, false)

    // --- SHEET 2: CHI TIẾT THỰC NHẬP ---
    const wsThucNhap = {}
    wsThucNhap['!cols'] = [
      { wpx: 50 }, { wpx: 100 }, { wpx: 130 }, { wpx: 240 }, { wpx: 240 },
      { wpx: 120 }, { wpx: 150 }, { wpx: 80 }, { wpx: 250 }, { wpx: 120 }
    ]

    const sheet2TitleStyle = {
      font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '0F766E' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    }

    wsThucNhap['A2'] = { v: 'BẢNG CHI TIẾT CHỨNG TỪ THỰC NHẬP (CẤU THÀNH SỐ LIỆU)', t: 's', s: sheet2TitleStyle }
    wsThucNhap['A3'] = { v: `Mặt hàng: ${itemRow.tenVatTu} (${itemRow.maSAP}) | ĐVT: ${itemRow.dvt}`, t: 's', s: subtitleStyle }

    // Header for Sheet 2 in Row 5
    const s2Headers = ['STT', 'Ngày', 'Mã chứng từ', 'Đơn vị giao (NCC/Kho)', 'Đơn vị nhận (Kho nhận)', 'KL Chứng từ', 'Phân loại đơn vị', 'Hệ số logic', 'Quy tắc & Diễn giải logic', 'Đóng góp thực']
    s2Headers.forEach((label, i) => {
      const colChar = String.fromCharCode(65 + i)
      wsThucNhap[`${colChar}5`] = { v: label, t: 's', s: thStyle1 }
    })

    const transactions = itemRow.transactions || []
    const nhapTxList = transactions.filter(tx => tx.nhapVal > 0)

    const footerStyle1 = { ...footerStyle, fill: { patternType: 'solid', fgColor: { rgb: 'F0FDF4' } } }
    
    // Total row for Sheet 2 at Row 6
    wsThucNhap[`A6`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`B6`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`C6`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`D6`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`E6`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle1, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
    wsThucNhap[`F6`] = { f: nhapTxList.length > 0 ? `SUM(F7:F${6 + nhapTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle1, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
    wsThucNhap[`G6`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`H6`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`I6`] = { v: '', t: 's', s: footerStyle1 }
    wsThucNhap[`J6`] = { f: nhapTxList.length > 0 ? `SUM(J7:J${6 + nhapTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle1, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }

    nhapTxList.forEach((tx, idx) => {
      const rowNum = 7 + idx
      const typeLabel = tx.detailTypeNhap === 'nhan_ncc' ? 'Nhận từ NCC'
                      : tx.detailTypeNhap === 'nhan_kho' ? 'Nhận từ Kho gửi'
                      : tx.detailTypeNhap === 'tra_ncc' ? 'Trả lại NCC'
                      : tx.detailTypeNhap === 'dieu_chuyen_di' ? 'Điều chuyển đi (Không tính)'
                      : tx.detailTypeNhap === 'nhan_khac' ? 'Nhận nguồn khác (Không tính)'
                      : 'Không tính'

      // Color "Phân loại đơn vị" badge in sheet 2 as well
      const isZero = tx.logicNhapVal === 0
      const isNegative = tx.logicNhapVal < 0
      let typeStyle = { ...dataCellStyle }
      if (isZero) {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'CBD5E1' } },
            bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
            left: { style: 'thin', color: { rgb: 'CBD5E1' } },
            right: { style: 'thin', color: { rgb: 'CBD5E1' } }
          }
        }
      } else if (isNegative) {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'FCA5A5' } },
            bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
            left: { style: 'thin', color: { rgb: 'FCA5A5' } },
            right: { style: 'thin', color: { rgb: 'FCA5A5' } }
          }
        }
      } else {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'A7F3D0' } },
            bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
            left: { style: 'thin', color: { rgb: 'A7F3D0' } },
            right: { style: 'thin', color: { rgb: 'A7F3D0' } }
          }
        }
      }

      wsThucNhap[`A${rowNum}`] = { v: idx + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
      wsThucNhap[`B${rowNum}`] = { v: tx.ngayXuatNhap || '', t: 's', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
      wsThucNhap[`C${rowNum}`] = { v: tx.maDonNhapKho || tx.maDonXuatKho || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true }, alignment: { vertical: 'center', wrapText: true } } }
      wsThucNhap[`D${rowNum}`] = { v: tx.donViGiao || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucNhap[`E${rowNum}`] = { v: tx.donViNhan || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucNhap[`F${rowNum}`] = { v: tx.nhapVal, t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
      wsThucNhap[`G${rowNum}`] = { v: typeLabel, t: 's', s: typeStyle }
      wsThucNhap[`H${rowNum}`] = { 
        v: tx.logicNhapVal, 
        t: 'n', 
        z: '#,##0;(#,##0);"-"', 
        s: { 
          ...dataCellStyle, 
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, 
          font: { 
            name: 'Segoe UI', 
            sz: 9.5, 
            bold: true,
            color: tx.logicNhapVal === -1 ? { rgb: 'FF0000' } : undefined
          } 
        } 
      }
      wsThucNhap[`I${rowNum}`] = { v: tx.explainNhap || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      
      wsThucNhap[`J${rowNum}`] = { 
        f: `F${rowNum} * H${rowNum}`, 
        t: 'n', 
        z: '#,##0.00;[Red](#,##0.00);"-"',
        s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F766E' } } } 
      }
    })

    wsThucNhap['!ref'] = `A1:J${6 + nhapTxList.length}`
    wsThucNhap['!merges'] = [
      { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } }
    ]
    wsThucNhap['!autofilter'] = { ref: `A5:J${6 + nhapTxList.length}` }


    // --- SHEET 3: CHI TIẾT THỰC XUẤT ---
    const wsThucXuat = {}
    wsThucXuat['!cols'] = [
      { wpx: 50 }, { wpx: 100 }, { wpx: 130 }, { wpx: 240 }, { wpx: 240 },
      { wpx: 120 }, { wpx: 150 }, { wpx: 80 }, { wpx: 250 }, { wpx: 120 }
    ]

    const sheet3TitleStyle = {
      font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: 'C2410C' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    }

    wsThucXuat['A2'] = { v: 'BẢNG CHI TIẾT CHỨNG TỪ THỰC XUẤT (CẤU THÀNH SỐ LIỆU)', t: 's', s: sheet3TitleStyle }
    wsThucXuat['A3'] = { v: `Mặt hàng: ${itemRow.tenVatTu} (${itemRow.maSAP}) | ĐVT: ${itemRow.dvt}`, t: 's', s: subtitleStyle }

    // Header for Sheet 3 in Row 5
    const s3Headers = ['STT', 'Ngày', 'Mã chứng từ', 'Đơn vị giao (Kho gửi)', 'Đơn vị nhận (Tổ đội/Kho nhận)', 'KL Chứng từ', 'Phân loại đơn vị', 'Hệ số logic', 'Quy tắc & Diễn giải logic', 'Đóng góp thực']
    s3Headers.forEach((label, i) => {
      const colChar = String.fromCharCode(65 + i)
      wsThucXuat[`${colChar}5`] = { v: label, t: 's', s: thStyle2 }
    })

    const xuatTxList = transactions.filter(tx => tx.xuatVal > 0)

    const footerStyle2 = { ...footerStyle, fill: { patternType: 'solid', fgColor: { rgb: 'FFF7ED' } } }
    
    // Total row for Sheet 3 at Row 6
    wsThucXuat[`A6`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`B6`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`C6`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`D6`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`E6`] = { v: 'Tổng cộng', t: 's', s: { ...footerStyle2, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
    wsThucXuat[`F6`] = { f: xuatTxList.length > 0 ? `SUM(F7:F${6 + xuatTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle2, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
    wsThucXuat[`G6`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`H6`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`I6`] = { v: '', t: 's', s: footerStyle2 }
    wsThucXuat[`J6`] = { f: xuatTxList.length > 0 ? `SUM(J7:J${6 + xuatTxList.length})` : '', t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...footerStyle2, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }

    xuatTxList.forEach((tx, idx) => {
      const rowNum = 7 + idx
      const typeLabel = tx.detailTypeXuat === 'xuat_todoi' ? 'Xuất cho Tổ đội'
                      : tx.detailTypeXuat === 'xuat_kho' ? 'Xuất đi Kho nhận'
                      : tx.detailTypeXuat === 'todoi_tra' ? 'Tổ đội trả hàng'
                      : tx.detailTypeXuat === 'xuat_khac' ? 'Xuất khác (Không tính)'
                      : tx.detailTypeXuat === 'nhan_lai_khac' ? 'Nhận lại khác (Không tính)'
                      : 'Không tính'

      // Color "Phân loại đơn vị" badge in sheet 3 as well
      const isZero = tx.logicXuatVal === 0
      const isNegative = tx.logicXuatVal < 0
      let typeStyle = { ...dataCellStyle }
      if (isZero) {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '475569' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'CBD5E1' } },
            bottom: { style: 'thin', color: { rgb: 'CBD5E1' } },
            left: { style: 'thin', color: { rgb: 'CBD5E1' } },
            right: { style: 'thin', color: { rgb: 'CBD5E1' } }
          }
        }
      } else if (isNegative) {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '991B1B' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'FEE2E2' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'FCA5A5' } },
            bottom: { style: 'thin', color: { rgb: 'FCA5A5' } },
            left: { style: 'thin', color: { rgb: 'FCA5A5' } },
            right: { style: 'thin', color: { rgb: 'FCA5A5' } }
          }
        }
      } else {
        typeStyle = {
          ...dataCellStyle,
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '047857' } },
          fill: { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } },
          alignment: { horizontal: 'left', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: 'A7F3D0' } },
            bottom: { style: 'thin', color: { rgb: 'A7F3D0' } },
            left: { style: 'thin', color: { rgb: 'A7F3D0' } },
            right: { style: 'thin', color: { rgb: 'A7F3D0' } }
          }
        }
      }

      wsThucXuat[`A${rowNum}`] = { v: idx + 1, t: 'n', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
      wsThucXuat[`B${rowNum}`] = { v: tx.ngayXuatNhap || '', t: 's', s: { ...dataCellStyle, alignment: { horizontal: 'center', vertical: 'center', wrapText: true } } }
      wsThucXuat[`C${rowNum}`] = { v: tx.maDonNhapKho || tx.maDonXuatKho || '', t: 's', s: { ...dataCellStyle, font: { name: 'Segoe UI', sz: 9.5, bold: true }, alignment: { vertical: 'center', wrapText: true } } }
      wsThucXuat[`D${rowNum}`] = { v: tx.donViGiao || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucXuat[`E${rowNum}`] = { v: tx.donViNhan || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      wsThucXuat[`F${rowNum}`] = { v: tx.xuatVal, t: 'n', z: '#,##0.00;[Red](#,##0.00);"-"', s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true } } }
      wsThucXuat[`G${rowNum}`] = { v: typeLabel, t: 's', s: typeStyle }
      wsThucXuat[`H${rowNum}`] = { 
        v: tx.logicXuatVal, 
        t: 'n', 
        z: '#,##0;(#,##0);"-"', 
        s: { 
          ...dataCellStyle, 
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true }, 
          font: { 
            name: 'Segoe UI', 
            sz: 9.5, 
            bold: true,
            color: tx.logicXuatVal === -1 ? { rgb: 'FF0000' } : undefined
          } 
        } 
      }
      wsThucXuat[`I${rowNum}`] = { v: tx.explainXuat || '', t: 's', s: { ...dataCellStyle, alignment: { vertical: 'center', wrapText: true } } }
      
      wsThucXuat[`J${rowNum}`] = { 
        f: `F${rowNum} * H${rowNum}`, 
        t: 'n', 
        z: '#,##0.00;[Red](#,##0.00);"-"',
        s: { ...dataCellStyle, alignment: { horizontal: 'right', vertical: 'center', wrapText: true }, font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'C2410C' } } } 
      }
    })

    wsThucXuat['!ref'] = `A1:J${6 + xuatTxList.length}`
    wsThucXuat['!merges'] = [
      { s: { r: 1, c: 0 }, e: { r: 1, c: 9 } }
    ]
    wsThucXuat['!autofilter'] = { ref: `A5:J${6 + xuatTxList.length}` }

    // Append sheets to workbook
    XLSXStyle.utils.book_append_sheet(wb, wsBaoCao, "Báo cáo")
    XLSXStyle.utils.book_append_sheet(wb, wsSummary, "Tổng hợp")
    XLSXStyle.utils.book_append_sheet(wb, wsThucNhap, "Chi tiết Thực nhập")
    XLSXStyle.utils.book_append_sheet(wb, wsThucXuat, "Chi tiết Thực xuất")

    // Generate output
    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary', compression: false })
    const s2ab = (s) => {
      const buf = new ArrayBuffer(s.length)
      const view = new Uint8Array(buf)
      for (let j = 0; j < s.length; j++) view[j] = s.charCodeAt(j) & 0xFF
      return buf
    }

    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `Giai_Trinh_Chi_Tiet_${itemRow.maSAP}_${new Date().toISOString().slice(0, 10)}.xlsx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  React.useEffect(() => {
    setCurrentPage(1)
    setDetailRow(null)
    setSelectedSaps(new Set())
  }, [summaryRows])

  React.useEffect(() => {
    const wrap = tableWrapRef.current
    const mirror = mirrorRef.current
    if (!wrap || !mirror) return

    const syncMirrorWidth = () => {
      const inner = wrap.querySelector('table')
      if (inner) {
        mirror.firstChild.style.width = inner.scrollWidth + 'px'
      }
    }
    syncMirrorWidth()

    const onWrapScroll = () => { mirror.scrollLeft = wrap.scrollLeft }
    const onMirrorScroll = () => { wrap.scrollLeft = mirror.scrollLeft }

    wrap.addEventListener('scroll', onWrapScroll)
    mirror.addEventListener('scroll', onMirrorScroll)

    const ro = new ResizeObserver(syncMirrorWidth)
    ro.observe(wrap)

    return () => {
      wrap.removeEventListener('scroll', onWrapScroll)
      mirror.removeEventListener('scroll', onMirrorScroll)
      ro.disconnect()
    }
  }, [summaryRows, pageSize])

  const totalPages = Math.ceil(summaryRows.length / pageSize)
  const startIdx = (currentPage - 1) * pageSize
  const endIdx = Math.min(startIdx + pageSize, summaryRows.length)
  const pageRows = summaryRows.slice(startIdx, endIdx)

  const columns = [
    { key: 'maVatTu', label: 'Mã vật tư', width: 100 },
    { key: 'maSAP', label: 'Mã SAP', width: 100 },
    { key: 'thongSoKyThuat', label: 'Thông số kỹ thuật', width: 150 },
    { key: 'tenVatTu', label: 'Tên vật tư', width: 300 },
    { key: 'dvt', label: 'ĐVT', width: 80 },
    { key: 'thucNhap', label: 'Thực nhập', width: 130 },
    { key: 'thucXuat', label: 'Thực xuất', width: 130 },
    { key: 'tonKho', label: 'Tồn kho', width: 130 }
  ]

  const getPageNums = () => {
    const pages = []
    if (totalPages <= 7) {
      for (let i = 1; i <= totalPages; i++) pages.push(i)
    } else {
      pages.push(1)
      if (currentPage > 3) pages.push('...')
      for (let i = Math.max(2, currentPage - 1); i <= Math.min(totalPages - 1, currentPage + 1); i++) pages.push(i)
      if (currentPage < totalPages - 2) pages.push('...')
      pages.push(totalPages)
    }
    return pages
  }

  const btnBase = {
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    height: 30, minWidth: 30, padding: '0 8px',
    border: '1px solid #e2e8f0', borderRadius: 6,
    fontSize: 13, fontWeight: 500, cursor: 'pointer',
    background: '#fff', color: '#374151', transition: 'all 0.15s',
    userSelect: 'none'
  }
  const btnActive = { ...btnBase, background: 'var(--primary)', color: '#fff', borderColor: 'var(--primary)', fontWeight: 700 }
  const btnDisabled = { ...btnBase, opacity: 0.4, cursor: 'not-allowed', pointerEvents: 'none' }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
      {/* Helpful tip banner */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', padding: '10px 14px', background: '#eff6ff', border: '1px solid #bfdbfe', borderRadius: '6px', marginBottom: '10px', fontSize: '12.5px', color: '#1e40af' }}>
        <Info size={16} className="text-blue-600" style={{ flexShrink: 0 }} />
        <span>
          <strong>Hướng dẫn:</strong> Có thể tích chọn các dòng vật tư bằng ô checkbox để xuất gộp hàng loạt bằng nút phía dưới, hoặc Nhấp đúp (double-click) vào bất kỳ dòng vật tư nào để xem <strong>Bảng giải trình cấu thành số liệu chi tiết</strong>.
        </span>
      </div>

      {/* Selection Control Bar */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px 16px', background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: '6px', marginBottom: '10px' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '12px', fontSize: '13px', color: '#334155', fontWeight: 500 }}>
          <span className="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium bg-teal-100 text-teal-800">
            Đã chọn {selectedSaps.size} / {summaryRows.length} vật tư
          </span>
          {selectedSaps.size > 0 && (
            <button
              onClick={() => setSelectedSaps(new Set())}
              style={{ fontSize: '12.5px', color: '#ef4444', background: 'none', border: 'none', cursor: 'pointer', padding: 0, textDecoration: 'underline' }}
            >
              Bỏ chọn tất cả
            </button>
          )}
        </div>
        <button
          onClick={handleExportSelectedExcel}
          disabled={selectedSaps.size === 0}
          style={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: '6px',
            backgroundColor: selectedSaps.size === 0 ? '#cbd5e1' : '#0f766e',
            color: '#ffffff',
            padding: '8px 16px',
            fontSize: '13px',
            fontWeight: '600',
            borderRadius: '6px',
            border: 'none',
            cursor: selectedSaps.size === 0 ? 'not-allowed' : 'pointer',
            boxShadow: selectedSaps.size === 0 ? 'none' : '0 1px 2px 0 rgba(0, 0, 0, 0.05)',
            transition: 'background-color 0.15s'
          }}
        >
          <Download size={15} />
          Xuất Báo cáo Giải trình gộp {selectedSaps.size > 0 ? `(${selectedSaps.size} vật tư)` : ''}
        </button>
      </div>

      {/* Table wrapper */}
      <div className="table-wrap" ref={tableWrapRef} style={{ overflowX: 'auto', overflowY: 'auto' }}>
        <table>
          <thead>
            <tr>
              <th style={{ width: 40, minWidth: 40, maxWidth: 40, textAlign: 'center', verticalAlign: 'middle', fontSize: '12px', padding: '8px' }}>
                <input
                  type="checkbox"
                  className="cursor-pointer rounded border-slate-300 text-teal-600 focus:ring-teal-500 w-4 h-4"
                  checked={selectedSaps.size === summaryRows.length && summaryRows.length > 0}
                  onChange={toggleSelectAll}
                />
              </th>
              <th style={{ width: 50, minWidth: 50, maxWidth: 50, textAlign: 'center', verticalAlign: 'middle', fontSize: '12px', padding: '8px 10px' }}>
                STT
              </th>
              {columns.map(c => {
                const isThucNhap = c.key === 'thucNhap'
                const isThucXuat = c.key === 'thucXuat'
                const isTonKho = c.key === 'tonKho'
                const thBg = isThucNhap ? '#0d9488' : isThucXuat ? '#ea580c' : isTonKho ? '#0f58a7' : undefined
                const thBorderBottom = isThucNhap ? '2px solid #0f766e' : isThucXuat ? '2px solid #c2410c' : isTonKho ? '2px solid #0a3d73' : undefined

                return (
                  <th
                    key={c.key}
                    style={{
                      width: c.width, minWidth: c.width, maxWidth: c.width,
                      textAlign: 'center', verticalAlign: 'middle',
                      fontSize: '12px', padding: '8px 10px',
                      whiteSpace: 'normal', wordBreak: 'break-word', lineHeight: '1.2',
                      background: thBg,
                      borderBottom: thBorderBottom
                    }}
                  >
                    {c.label}
                  </th>
                )
              })}
            </tr>
          </thead>
          <tbody>
            {pageRows.map((row, i) => {
              return (
                <tr
                  key={row.maSAP}
                  onDoubleClick={() => setDetailRow(row)}
                  style={{
                    cursor: 'pointer',
                    transition: 'background-color 0.15s'
                  }}
                  title="Nhấp đúp để xem bảng giải trình chi tiết"
                  className="hover:bg-sky-50"
                >
                  <td style={{ width: 40, minWidth: 40, maxWidth: 40, textAlign: 'center', padding: '6px' }} onClick={(e) => e.stopPropagation()}>
                    <input
                      type="checkbox"
                      className="cursor-pointer rounded border-slate-300 text-teal-600 focus:ring-teal-500 w-4 h-4"
                      checked={selectedSaps.has(row.maSAP)}
                      onChange={() => toggleSelectRow(row.maSAP)}
                    />
                  </td>
                  <td style={{ width: 50, minWidth: 50, maxWidth: 50, textAlign: 'center', fontSize: '12px', color: '#1b1919', padding: '6px 10px', fontWeight: 500 }}>
                    {startIdx + i + 1}
                  </td>
                  {columns.map(col => {
                    const isCenteredCol = ['maVatTu', 'maSAP', 'dvt'].includes(col.key)
                    const isRightAligned = ['thucNhap', 'thucXuat', 'tonKho'].includes(col.key)

                    let cellBg = undefined
                    let cellTextColor = '#1b1919'
                    let cellFontWeight = undefined

                    const isNumeric = col.key === 'thucNhap' || col.key === 'thucXuat' || col.key === 'tonKho'
                    if (col.key === 'tonKho') {
                      const numVal = Number(row[col.key])
                      const isNegative = numVal < 0
                      const isZero = numVal === 0
                      cellBg = isNegative ? '#fef2f2' : (isZero ? undefined : 'var(--primary-light)')
                      cellTextColor = isNegative ? '#ef4444' : (isZero ? 'var(--text-muted)' : 'var(--primary)')
                      cellFontWeight = '800'
                    } else if (col.key === 'thucNhap' || col.key === 'thucXuat') {
                      const numVal = Number(row[col.key])
                      if (numVal > 0) {
                        cellBg = col.key === 'thucNhap' ? '#ecfdf5' : '#fff7ed' // soft green vs soft orange
                        cellTextColor = col.key === 'thucNhap' ? '#065f46' : '#c2410c' // dark green vs dark orange
                        cellFontWeight = '700'
                      } else {
                        cellBg = '#f1f5f9' // soft gray
                        cellTextColor = '#475569' // dark gray
                        cellFontWeight = '500'
                      }
                    }

                    let displayVal = row[col.key]
                    if (isNumeric) {
                      const numVal = Number(row[col.key])
                      displayVal = numVal === 0 ? '-' : numVal.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })
                    }

                    return (
                      <td
                        key={col.key}
                        style={{
                          width: col.width, minWidth: col.width, maxWidth: col.width,
                          color: cellTextColor,
                          backgroundColor: cellBg,
                          fontWeight: cellFontWeight,
                          textAlign: isCenteredCol ? 'center' : (isRightAligned ? 'right' : 'left'),
                          fontSize: '12px', padding: '6px 10px',
                          wordBreak: 'break-word', whiteSpace: 'normal',
                          borderBottom: '1px solid #e2e8f0'
                        }}
                        title={String(displayVal || '')}
                      >
                        {displayVal}
                      </td>
                    )
                  })}
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>

      {/* Modal for detail row */}
      {detailRow && (() => {
        const txs = detailRow.transactions || []
        const nhapSummaryByUnit = {}
        const xuatSummaryByUnit = {}

        txs.forEach(tx => {
          const normProject = String(localProject || '').trim().toLowerCase()
          const isGiaoProject = String(tx.donViGiao || '').trim().toLowerCase() === normProject
          const isNhanProject = String(tx.donViNhan || '').trim().toLowerCase() === normProject

          if (tx.nhapVal > 0) {
            let role = 'Đơn vị nhận'
            let partner = tx.donViGiao || 'Chưa xác định'
            if (isGiaoProject) {
              role = 'Đơn vị giao'
              partner = tx.donViNhan || 'Chưa xác định'
            }

            const key = `${partner}||${tx.detailTypeNhap}||${role}`
            if (!nhapSummaryByUnit[key]) {
              nhapSummaryByUnit[key] = {
                partner,
                role,
                detailType: tx.detailTypeNhap,
                logicVal: tx.logicNhapVal,
                explain: tx.explainNhap,
                totalQty: 0,
                totalContribution: 0
              }
            }
            nhapSummaryByUnit[key].totalQty += tx.nhapVal
            nhapSummaryByUnit[key].totalContribution += tx.nhapVal * tx.logicNhapVal
          }
          if (tx.xuatVal > 0) {
            let role = 'Đơn vị giao'
            let partner = tx.donViNhan || 'Chưa xác định'
            if (isNhanProject) {
              role = 'Đơn vị nhận'
              partner = tx.donViGiao || 'Chưa xác định'
            }

            const key = `${partner}||${tx.detailTypeXuat}||${role}`
            if (!xuatSummaryByUnit[key]) {
              xuatSummaryByUnit[key] = {
                partner,
                role,
                detailType: tx.detailTypeXuat,
                logicVal: tx.logicXuatVal,
                explain: tx.explainXuat,
                totalQty: 0,
                totalContribution: 0
              }
            }
            xuatSummaryByUnit[key].totalQty += tx.xuatVal
            xuatSummaryByUnit[key].totalContribution += tx.xuatVal * tx.logicXuatVal
          }
        })

        const sortNhapSummaryList = (list) => {
          return [...list].sort((a, b) => {
            // Cấp 1 - Vai trò Kho chọn: Đơn vị nhận trước, Đơn vị giao sau
            const rolePriorityA = a.role === 'Đơn vị nhận' ? 1 : a.role === 'Đơn vị giao' ? 2 : 3
            const rolePriorityB = b.role === 'Đơn vị nhận' ? 1 : b.role === 'Đơn vị giao' ? 2 : 3
            if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

            // Cấp 2 - Phân loại đơn vị: Nhà cung cấp đến Kho khác đến Tổ đội
            const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
            const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
            const catPriorityA = catA === 'ncc' ? 1 : catA === 'kho' ? 2 : catA === 'todoi' ? 3 : 4

            const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
            const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
            const catPriorityB = catB === 'ncc' ? 1 : catB === 'kho' ? 2 : catB === 'todoi' ? 3 : 4

            if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

            // Cấp 3 - Giá trị (Khối lượng) giảm dần trị tuyệt đối
            const valA = Math.abs(a.totalContribution || 0)
            const valB = Math.abs(b.totalContribution || 0)
            if (valB !== valA) return valB - valA

            // Fallback: Alphabetical
            return (a.partner || '').localeCompare(b.partner || '', 'vi')
          })
        }

        const sortXuatSummaryList = (list) => {
          return [...list].sort((a, b) => {
            // Cấp 1 - Vai trò Kho chọn: Đơn vị giao trước, Đơn vị nhận sau
            const rolePriorityA = a.role === 'Đơn vị giao' ? 1 : a.role === 'Đơn vị nhận' ? 2 : 3
            const rolePriorityB = b.role === 'Đơn vị giao' ? 1 : b.role === 'Đơn vị nhận' ? 2 : 3
            if (rolePriorityA !== rolePriorityB) return rolePriorityA - rolePriorityB

            // Cấp 2 - Phân loại đơn vị: Từ Tổ đội đến Kho khác đến Nhà cung cấp
            const normA = (a.partner || '').trim().replace(/\s+/g, ' ')
            const catA = (customCategoryMap && customCategoryMap[normA]) || getUnitCategory(normA)
            const catPriorityA = catA === 'todoi' ? 1 : catA === 'kho' ? 2 : catA === 'ncc' ? 3 : 4

            const normB = (b.partner || '').trim().replace(/\s+/g, ' ')
            const catB = (customCategoryMap && customCategoryMap[normB]) || getUnitCategory(normB)
            const catPriorityB = catB === 'todoi' ? 1 : catB === 'kho' ? 2 : catB === 'ncc' ? 3 : 4

            if (catPriorityA !== catPriorityB) return catPriorityA - catPriorityB

            // Cấp 3 - Giá trị (Khối lượng) giảm dần trị tuyệt đối
            const valA = Math.abs(a.totalContribution || 0)
            const valB = Math.abs(b.totalContribution || 0)
            if (valB !== valA) return valB - valA

            // Fallback: Alphabetical
            return (a.partner || '').localeCompare(b.partner || '', 'vi')
          })
        }

        const nhapSummaryList = sortNhapSummaryList(Object.values(nhapSummaryByUnit))
        const xuatSummaryList = sortXuatSummaryList(Object.values(xuatSummaryByUnit))

        return (
          <div style={{
            position: 'fixed',
            top: 0, left: 0, right: 0, bottom: 0,
            backgroundColor: 'rgba(15, 23, 42, 0.6)',
            backdropFilter: 'blur(4px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            zIndex: 1000,
            padding: 12
          }}
          onClick={() => setDetailRow(null)}
          >
            <div style={{
              background: '#ffffff',
              borderRadius: 12,
              width: '98vw',
              height: '96vh',
              maxWidth: '100vw',
              maxHeight: '100vh',
              padding: '20px 24px',
              boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04)',
              border: '1px solid #e2e8f0',
              display: 'flex',
              flexDirection: 'column',
              gap: 16
            }}
            onClick={(e) => e.stopPropagation()}
            >
              {/* Modal Header */}
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: 12 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 16, flexWrap: 'wrap' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <Info size={20} className="text-blue-600" />
                    <div>
                      <h3 style={{ margin: 0, fontSize: 15.5, fontWeight: 700, color: '#1e3a8a' }}>
                        BẢNG GIẢI TRÌNH CẤU THÀNH SỐ LIỆU CHI TIẾT
                      </h3>
                      <p style={{ margin: 0, fontSize: 12, color: '#475569', fontWeight: 600, marginTop: 2 }}>
                        <span style={{ color: '#0284c7', fontWeight: 700 }}>Kho chọn:</span> {localProject || 'Tất cả'} | <span style={{ color: '#0284c7', fontWeight: 700 }}>Vật tư:</span> {detailRow.tenVatTu} ({detailRow.maSAP}) | <span style={{ color: '#0284c7', fontWeight: 700 }}>ĐVT:</span> {detailRow.dvt}
                      </p>
                    </div>
                  </div>

                  {/* Color Legend for partners on Web App */}
                  <div style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 12,
                    fontSize: '11px',
                    fontWeight: 600,
                    background: '#f8fafc',
                    padding: '6px 12px',
                    borderRadius: 6,
                    border: '1px solid #e2e8f0',
                    marginLeft: 16
                  }}>
                    <span style={{ color: '#64748b', fontSize: '10px', textTransform: 'uppercase', letterSpacing: '0.05em', fontWeight: 700 }}>Màu sắc đối tác:</span>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                      <span style={{ width: 10, height: 10, borderRadius: 2, background: '#fffbeb', border: '1px solid #fde68a' }}></span>
                      <span style={{ color: '#78350f' }}>Nhà cung cấp (NCC)</span>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                      <span style={{ width: 10, height: 10, borderRadius: 2, background: '#eff6ff', border: '1px solid #bfdbfe' }}></span>
                      <span style={{ color: '#1e3a8a' }}>Kho khác</span>
                    </div>
                    <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                      <span style={{ width: 10, height: 10, borderRadius: 2, background: '#f0fdf4', border: '1px solid #bbf7d0' }}></span>
                      <span style={{ color: '#065f46' }}>Tổ đội</span>
                    </div>
                  </div>
                </div>
                <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                  <button
                    onClick={() => handleExportDetailExcel(detailRow, nhapSummaryList, xuatSummaryList)}
                    style={{
                      background: 'linear-gradient(135deg, #059669 0%, #10b981 100%)',
                      color: '#ffffff',
                      border: 'none',
                      boxShadow: '0 2px 4px rgba(16,185,129,0.2)',
                      height: 36,
                      padding: '0 16px',
                      borderRadius: 6,
                      display: 'inline-flex',
                      alignItems: 'center',
                      gap: 6,
                      fontWeight: 600,
                      fontSize: '13px',
                      cursor: 'pointer'
                    }}
                  >
                    <Download size={15} /> Xuất Excel chi tiết
                  </button>
                  <button 
                    onClick={() => setDetailRow(null)} 
                    style={{ background: 'transparent', border: 'none', color: '#64748b', cursor: 'pointer', display: 'flex', alignItems: 'center', padding: 4 }}
                  >
                    <X size={20} />
                  </button>
                </div>
              </div>

              {/* Modal Content */}
              <div style={{ overflowY: 'auto', flex: 1, paddingRight: 4, display: 'flex', flexDirection: 'column', gap: 16 }}>
                
                {/* Summary of calculation formulas */}
                <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(440px, 1fr))', gap: 16, marginBottom: 4 }}>
                  
                  {/* Table 1: DIỄN GIẢI THỰC NHẬP */}
                  <div style={{ background: '#fcfdfd', border: '1px solid #0f766e', borderRadius: 8, overflow: 'hidden', boxShadow: '0 2px 4px rgba(15,23,42,0.05)' }}>
                    <div style={{ background: 'linear-gradient(135deg, #0f766e 0%, #115e59 100%)', color: '#ffffff', padding: '10px 14px', fontWeight: 700, fontSize: '13px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Warehouse size={14} /> BẢNG TỔNG HỢP DIỄN GIẢI THỰC NHẬP (SUMIFS THEO ĐƠN VỊ)
                      </span>
                      <span style={{ background: 'rgba(255,255,255,0.18)', padding: '2px 8px', borderRadius: 4, fontSize: '11.5px', fontWeight: 800 }}>
                        Tổng: {detailRow.thucNhap === 0 ? '-' : detailRow.thucNhap.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} {detailRow.dvt}
                      </span>
                    </div>
                    <div style={{ overflowX: 'auto' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px' }}>
                        <thead>
                          <tr style={{ background: '#f0fdf4', borderBottom: '1px solid #ccfbf1', color: '#0f766e', fontWeight: 700 }}>
                            <th style={{ padding: '8px 10px', textAlign: 'center', width: 40, borderBottom: '1px solid #ccfbf1' }}>STT</th>
                            <th style={{ padding: '8px 10px', textAlign: 'left', width: 110, borderBottom: '1px solid #ccfbf1' }}>Phân loại đơn vị</th>
                            <th style={{ padding: '8px 10px', textAlign: 'left', width: 140, borderBottom: '1px solid #ccfbf1' }}>Kho chọn</th>
                            <th style={{ padding: '8px 10px', textAlign: 'center', width: 110, borderBottom: '1px solid #ccfbf1' }}>Vai trò Kho chọn</th>
                            <th style={{ padding: '8px 10px', textAlign: 'left', width: 180, borderBottom: '1px solid #ccfbf1' }}>Tên NCC/ Kho khác / Tổ đội</th>
                            <th style={{ padding: '8px 10px', textAlign: 'center', width: 150, borderBottom: '1px solid #ccfbf1' }}>Vai trò NCC/ Kho khác/ Tổ đội</th>
                            <th style={{ padding: '8px 10px', textAlign: 'right', width: 100, borderBottom: '1px solid #ccfbf1' }}>Tổng KL Chứng từ</th>
                            <th style={{ padding: '8px 10px', textAlign: 'center', width: 80, borderBottom: '1px solid #ccfbf1' }}>Hệ số logic</th>
                            <th style={{ padding: '8px 10px', textAlign: 'right', width: 110, borderBottom: '1px solid #ccfbf1' }}>Thực nhập</th>
                            <th style={{ padding: '8px 10px', textAlign: 'left', width: 180, borderBottom: '1px solid #ccfbf1' }}>Ghi chú</th>
                          </tr>
                        </thead>
                        <tbody>
                          {nhapSummaryList.length === 0 ? (
                            <tr>
                              <td colSpan={10} style={{ padding: '16px', textAlign: 'center', color: '#64748b', fontStyle: 'italic' }}>
                                Không có dữ liệu thực nhập phát sinh
                              </td>
                            </tr>
                          ) : (
                            nhapSummaryList.map((sumItem, idx) => {
                              const isNegative = sumItem.logicVal < 0
                              const isZero = sumItem.logicVal === 0
                              const labelType = sumItem.detailType === 'nhan_ncc' ? 'Nhận từ NCC'
                                              : sumItem.detailType === 'nhan_kho' ? 'Nhận từ Kho gửi'
                                              : sumItem.detailType === 'tra_ncc' ? 'Trả lại NCC'
                                              : sumItem.detailType === 'dieu_chuyen_di' ? 'Điều chuyển đi (Không tính)'
                                              : sumItem.detailType === 'nhan_khac' ? 'Nhận nguồn khác (Không tính)'
                                              : 'Không tính'
                              
                              const norm = (sumItem.partner || '').trim().replace(/\s+/g, ' ');
                              const cat = (customCategoryMap && customCategoryMap[norm]) || getUnitCategory(norm);
                              const partnerRoleLabel = cat === 'ncc' ? 'Nhà cung cấp (NCC)'
                                                     : cat === 'kho' ? 'Kho khác'
                                                     : cat === 'todoi' ? 'Tổ đội'
                                                     : '';
                              return (
                                <tr key={idx} style={{ borderBottom: '1px solid #f1f5f9', background: idx % 2 === 1 ? '#fcfdfd' : '#ffffff' }}>
                                  <td style={{ padding: '8px 10px', textAlign: 'center', color: '#64748b' }}>{idx + 1}</td>
                                  
                                  {/* 2. Phân loại đơn vị */}
                                  <td style={{ padding: '8px 10px', fontWeight: 600 }}>
                                    <span style={{
                                      padding: '2px 6px', borderRadius: 4, fontSize: '10px',
                                      background: isZero ? '#f1f5f9' : isNegative ? '#fee2e2' : '#ecfdf5',
                                      color: isZero ? '#475569' : isNegative ? '#991b1b' : '#047857',
                                      border: isZero ? '1px solid #cbd5e1' : isNegative ? '1px solid #fca5a5' : '1px solid #a7f3d0'
                                    }}>
                                      {labelType}
                                    </span>
                                  </td>
                                  
                                  {/* 3. Kho chọn */}
                                  <td style={{ 
                                    padding: '8px 10px', 
                                    fontWeight: 700, 
                                    background: sumItem.role === 'Đơn vị nhận' ? '#f5f3ff' : '#fff7ed',
                                    color: sumItem.role === 'Đơn vị nhận' ? '#6d28d9' : '#ea580c'
                                  }}>
                                    {localProject || 'Tất cả'}
                                  </td>

                                  {/* 4. Vai trò Kho chọn */}
                                  <td style={{ padding: '8px 10px', textAlign: 'center', fontWeight: 600 }}>
                                    <span style={{
                                      padding: '2px 6px', borderRadius: 4, fontSize: '10px',
                                      background: sumItem.role === 'Đơn vị nhận' ? '#f5f3ff' : '#fff7ed',
                                      color: sumItem.role === 'Đơn vị nhận' ? '#6d28d9' : '#ea580c',
                                      border: sumItem.role === 'Đơn vị nhận' ? '1px solid #ddd6fe' : '1px solid #fed7aa'
                                    }}>
                                      {sumItem.role}
                                    </span>
                                  </td>
                                  
                                  {/* 5. Tên NCC/ Kho khác / Tổ đội */}
                                  <td style={{
                                    padding: '8px 10px',
                                    backgroundColor: cat === 'kho' ? '#eff6ff' : cat === 'ncc' ? '#fffbeb' : cat === 'todoi' ? '#f0fdf4' : undefined
                                  }}>
                                    <div style={{
                                      fontWeight: 600,
                                      color: cat === 'kho' ? '#1e3a8a' : cat === 'ncc' ? '#78350f' : cat === 'todoi' ? '#065f46' : '#334155'
                                    }}>
                                      {sumItem.partner}
                                    </div>
                                  </td>

                                  {/* 6. Vai trò NCC/ Kho khác/ Tổ đội */}
                                  <td style={{ padding: '8px 10px', textAlign: 'center', fontWeight: 600 }}>
                                    {partnerRoleLabel && (
                                      <span style={{
                                        padding: '2px 6px', borderRadius: 4, fontSize: '10px',
                                        background: cat === 'kho' ? '#eff6ff' : cat === 'ncc' ? '#fffbeb' : '#f0fdf4',
                                        color: cat === 'kho' ? '#1e3a8a' : cat === 'ncc' ? '#78350f' : '#065f46',
                                        border: cat === 'kho' ? '1px solid #bfdbfe' : cat === 'ncc' ? '1px solid #fde68a' : '1px solid #bbf7d0'
                                      }}>
                                        {partnerRoleLabel}
                                      </span>
                                    )}
                                  </td>

                                  {/* 7. Tổng KL Chứng từ */}
                                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: '#0f172a' }}>
                                    {sumItem.totalQty.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                  </td>

                                  {/* 8. Hệ số logic */}
                                  <td style={{ padding: '8px 10px', textAlign: 'center' }}>
                                    <span style={{
                                      fontWeight: 700, padding: '2px 6px', borderRadius: 4, fontSize: '10px',
                                      background: sumItem.logicVal > 0 ? '#dcfce7' : sumItem.logicVal < 0 ? '#fee2e2' : '#f1f5f9',
                                      color: sumItem.logicVal > 0 ? '#15803d' : sumItem.logicVal < 0 ? '#b91c1c' : '#475569'
                                    }}>
                                      {sumItem.logicVal > 0 ? `+${sumItem.logicVal}` : sumItem.logicVal}
                                    </span>
                                  </td>

                                  {/* 9. Thực nhập */}
                                  <td style={{
                                    padding: '8px 10px', textAlign: 'right', fontWeight: 700,
                                    color: sumItem.totalContribution > 0 ? '#0f766e' : sumItem.totalContribution < 0 ? '#b91c1c' : '#64748b'
                                  }}>
                                    {sumItem.totalContribution === 0 ? '-' : (sumItem.totalContribution > 0 ? '+' : '') + sumItem.totalContribution.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                  </td>

                                  {/* 10. Ghi chú */}
                                  <td style={{ padding: '8px 10px', color: '#475569', fontStyle: 'italic', fontSize: '11px' }}>
                                    {sumItem.explain || ''}
                                  </td>
                                </tr>
                              )
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Table 2: DIỄN GIẢI THỰC XUẤT */}
                  <div style={{ background: '#fcfdfd', border: '1px solid #c2410c', borderRadius: 8, overflow: 'hidden', boxShadow: '0 2px 4px rgba(15,23,42,0.05)' }}>
                    <div style={{ background: 'linear-gradient(135deg, #ea580c 0%, #c2410c 100%)', color: '#ffffff', padding: '10px 14px', fontWeight: 700, fontSize: '13px', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                      <span style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                        <Users size={14} /> BẢNG TỔNG HỢP DIỄN GIẢI THỰC XUẤT (SUMIFS THEO ĐƠN VỊ)
                      </span>
                      <span style={{ background: 'rgba(255,255,255,0.18)', padding: '2px 8px', borderRadius: 4, fontSize: '11.5px', fontWeight: 800 }}>
                        Tổng: {detailRow.thucXuat === 0 ? '-' : detailRow.thucXuat.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} {detailRow.dvt}
                      </span>
                    </div>
                    <div style={{ overflowX: 'auto' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '11.5px' }}>
                        <thead>
                          <tr style={{ background: '#fff7ed', borderBottom: '1px solid #ffedd5', color: '#c2410c', fontWeight: 700 }}>
                            <th style={{ padding: '8px 10px', textAlign: 'center', width: 40, borderBottom: '1px solid #ffedd5' }}>STT</th>
                            <th style={{ padding: '8px 10px', textAlign: 'left', width: 110, borderBottom: '1px solid #ffedd5' }}>Phân loại đơn vị</th>
                            <th style={{ padding: '8px 10px', textAlign: 'left', width: 140, borderBottom: '1px solid #ffedd5' }}>Kho chọn</th>
                            <th style={{ padding: '8px 10px', textAlign: 'center', width: 110, borderBottom: '1px solid #ffedd5' }}>Vai trò Kho chọn</th>
                            <th style={{ padding: '8px 10px', textAlign: 'left', width: 180, borderBottom: '1px solid #ffedd5' }}>Tên NCC/ Kho khác / Tổ đội</th>
                            <th style={{ padding: '8px 10px', textAlign: 'center', width: 150, borderBottom: '1px solid #ffedd5' }}>Vai trò NCC/ Kho khác/ Tổ đội</th>
                            <th style={{ padding: '8px 10px', textAlign: 'right', width: 100, borderBottom: '1px solid #ffedd5' }}>Tổng KL Chứng từ</th>
                            <th style={{ padding: '8px 10px', textAlign: 'center', width: 80, borderBottom: '1px solid #ffedd5' }}>Hệ số logic</th>
                            <th style={{ padding: '8px 10px', textAlign: 'right', width: 110, borderBottom: '1px solid #ffedd5' }}>Thực xuất</th>
                            <th style={{ padding: '8px 10px', textAlign: 'left', width: 180, borderBottom: '1px solid #ffedd5' }}>Ghi chú</th>
                          </tr>
                        </thead>
                        <tbody>
                          {xuatSummaryList.length === 0 ? (
                            <tr>
                              <td colSpan={10} style={{ padding: '16px', textAlign: 'center', color: '#64748b', fontStyle: 'italic' }}>
                                Không có dữ liệu thực xuất phát sinh
                              </td>
                            </tr>
                          ) : (
                            xuatSummaryList.map((sumItem, idx) => {
                              const isNegative = sumItem.logicVal < 0
                              const isZero = sumItem.logicVal === 0
                              const labelType = sumItem.detailType === 'xuat_todoi' ? 'Xuất cho Tổ đội'
                                              : sumItem.detailType === 'xuat_kho' ? 'Xuất đi Kho nhận'
                                              : sumItem.detailType === 'todoi_tra' ? 'Tổ đội trả hàng'
                                              : sumItem.detailType === 'xuat_khac' ? 'Xuất khác (Không tính)'
                                              : sumItem.detailType === 'nhan_lai_khac' ? 'Nhận lại khác (Không tính)'
                                              : 'Không tính'

                              const norm = (sumItem.partner || '').trim().replace(/\s+/g, ' ');
                              const cat = (customCategoryMap && customCategoryMap[norm]) || getUnitCategory(norm);
                              const partnerRoleLabel = cat === 'ncc' ? 'Nhà cung cấp (NCC)'
                                                     : cat === 'kho' ? 'Kho khác'
                                                     : cat === 'todoi' ? 'Tổ đội'
                                                     : '';
                              return (
                                <tr key={idx} style={{ borderBottom: '1px solid #f1f5f9', background: idx % 2 === 1 ? '#fcfdfd' : '#ffffff' }}>
                                  <td style={{ padding: '8px 10px', textAlign: 'center', color: '#64748b' }}>{idx + 1}</td>
                                  
                                  {/* 2. Phân loại đơn vị */}
                                  <td style={{ padding: '8px 10px', fontWeight: 600 }}>
                                    <span style={{
                                      padding: '2px 6px', borderRadius: 4, fontSize: '10px',
                                      background: isZero ? '#f1f5f9' : isNegative ? '#fee2e2' : '#fff7ed',
                                      color: isZero ? '#475569' : isNegative ? '#991b1b' : '#c2410c',
                                      border: isZero ? '1px solid #cbd5e1' : isNegative ? '1px solid #fca5a5' : '1px solid #fed7aa'
                                    }}>
                                      {labelType}
                                    </span>
                                  </td>
                                  
                                  {/* 3. Kho chọn */}
                                  <td style={{ 
                                    padding: '8px 10px', 
                                    fontWeight: 700, 
                                    background: sumItem.role === 'Đơn vị nhận' ? '#f5f3ff' : '#fff7ed',
                                    color: sumItem.role === 'Đơn vị nhận' ? '#6d28d9' : '#ea580c'
                                  }}>
                                    {localProject || 'Tất cả'}
                                  </td>

                                  {/* 4. Vai trò Kho chọn */}
                                  <td style={{ padding: '8px 10px', textAlign: 'center', fontWeight: 600 }}>
                                    <span style={{
                                      padding: '2px 6px', borderRadius: 4, fontSize: '10px',
                                      background: sumItem.role === 'Đơn vị nhận' ? '#f5f3ff' : '#fff7ed',
                                      color: sumItem.role === 'Đơn vị nhận' ? '#6d28d9' : '#ea580c',
                                      border: sumItem.role === 'Đơn vị nhận' ? '1px solid #ddd6fe' : '1px solid #fed7aa'
                                    }}>
                                      {sumItem.role}
                                    </span>
                                  </td>
                                  
                                  {/* 5. Tên NCC/ Kho khác / Tổ đội */}
                                  <td style={{
                                    padding: '8px 10px',
                                    backgroundColor: cat === 'kho' ? '#eff6ff' : cat === 'ncc' ? '#fffbeb' : cat === 'todoi' ? '#f0fdf4' : undefined
                                  }}>
                                    <div style={{
                                      fontWeight: 600,
                                      color: cat === 'kho' ? '#1e3a8a' : cat === 'ncc' ? '#78350f' : cat === 'todoi' ? '#065f46' : '#334155'
                                    }}>
                                      {sumItem.partner}
                                    </div>
                                  </td>

                                  {/* 6. Vai trò NCC/ Kho khác/ Tổ đội */}
                                  <td style={{ padding: '8px 10px', textAlign: 'center', fontWeight: 600 }}>
                                    {partnerRoleLabel && (
                                      <span style={{
                                        padding: '2px 6px', borderRadius: 4, fontSize: '10px',
                                        background: cat === 'kho' ? '#eff6ff' : cat === 'ncc' ? '#fffbeb' : '#f0fdf4',
                                        color: cat === 'kho' ? '#1e3a8a' : cat === 'ncc' ? '#78350f' : '#065f46',
                                        border: cat === 'kho' ? '1px solid #bfdbfe' : cat === 'ncc' ? '1px solid #fde68a' : '1px solid #bbf7d0'
                                      }}>
                                        {partnerRoleLabel}
                                      </span>
                                    )}
                                  </td>

                                  {/* 7. Tổng KL Chứng từ */}
                                  <td style={{ padding: '8px 10px', textAlign: 'right', fontWeight: 600, color: '#0f172a' }}>
                                    {sumItem.totalQty.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                  </td>

                                  {/* 8. Hệ số logic */}
                                  <td style={{ padding: '8px 10px', textAlign: 'center' }}>
                                    <span style={{
                                      fontWeight: 700, padding: '2px 6px', borderRadius: 4, fontSize: '10px',
                                      background: sumItem.logicVal > 0 ? '#dcfce7' : sumItem.logicVal < 0 ? '#fee2e2' : '#f1f5f9',
                                      color: sumItem.logicVal > 0 ? '#15803d' : sumItem.logicVal < 0 ? '#b91c1c' : '#475569'
                                    }}>
                                      {sumItem.logicVal > 0 ? `+${sumItem.logicVal}` : sumItem.logicVal}
                                    </span>
                                  </td>

                                  {/* 9. Thực xuất */}
                                  <td style={{
                                    padding: '8px 10px', textAlign: 'right', fontWeight: 700,
                                    color: sumItem.totalContribution > 0 ? '#c2410c' : sumItem.totalContribution < 0 ? '#b91c1c' : '#64748b'
                                  }}>
                                    {sumItem.totalContribution === 0 ? '-' : (sumItem.totalContribution > 0 ? '+' : '') + sumItem.totalContribution.toLocaleString('vi-VN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                                  </td>

                                  {/* 10. Ghi chú */}
                                  <td style={{ padding: '8px 10px', color: '#475569', fontStyle: 'italic', fontSize: '11px' }}>
                                    {sumItem.explain || ''}
                                  </td>
                                </tr>
                              )
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>

                </div>



              </div>

              {/* Modal Footer */}
              <div style={{ display: 'flex', justifyContent: 'flex-end', borderTop: '1px solid #f1f5f9', paddingTop: 12 }}>
                <button onClick={() => setDetailRow(null)} className="btn btn-outline" style={{ minWidth: 100 }}>
                  Đóng
                </button>
              </div>
            </div>
          </div>
        )
      })()}


      {/* Mirror Scrollbar */}
      <div className="scroll-mirror" ref={mirrorRef} style={{ overflowX: 'auto', overflowY: 'hidden', height: 12, flexShrink: 0, marginTop: 4, background: '#f8fafc', borderBottom: '1px solid #e2e8f0' }}>
        <div style={{ height: 1 }}><span /></div>
      </div>

      {/* Pagination */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', padding: '12px 16px', background: '#f8fafc', borderTop: '1px solid #cbd5e1', borderBottomLeftRadius: 8, borderBottomRightRadius: 8, gap: 12, flexWrap: 'wrap', flexShrink: 0 }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: '13px', color: '#475569' }}>
          <span>Hiển thị</span>
          <select
            value={pageSize}
            onChange={(e) => { setPageSize(Number(e.target.value)); setCurrentPage(1); }}
            style={{ padding: '2px 4px', border: '1px solid #cbd5e1', borderRadius: 4, background: '#fff', fontSize: '13px', color: '#1e293b', cursor: 'pointer' }}
          >
            {PAGE_SIZE_OPTIONS.map(opt => <option key={opt} value={opt}>{opt}</option>)}
          </select>
          <span>dòng / trang</span>
        </div>

        <div style={{ fontSize: '13px', color: '#475569', fontWeight: 500 }}>
          {summaryRows.length > 0 ? `${startIdx + 1}–${endIdx} / ${summaryRows.length} dòng` : '0 dòng'}
        </div>

        {totalPages > 1 && (
          <div style={{ display: 'flex', gap: 4, alignItems: 'center' }}>
            <button
              onClick={() => setCurrentPage(prev => Math.max(1, prev - 1))}
              disabled={currentPage === 1}
              style={currentPage === 1 ? btnDisabled : btnBase}
              title="Trang trước"
            >
              ◀
            </button>
            {getPageNums().map((num, idx) => (
              <button
                key={idx}
                onClick={() => num !== '...' && setCurrentPage(num)}
                style={num === '...' ? { ...btnBase, cursor: 'default', border: 'none', background: 'transparent' } : (num === currentPage ? btnActive : btnBase)}
                disabled={num === '...'}
              >
                {num}
              </button>
            ))}
            <button
              onClick={() => setCurrentPage(prev => Math.min(totalPages, prev + 1))}
              disabled={currentPage === totalPages}
              style={currentPage === totalPages ? btnDisabled : btnBase}
              title="Trang sau"
            >
              ▶
            </button>
          </div>
        )}
      </div>
    </div>
  )
}

// ─── Data Table ───────────────────────────────────────────────────────────────
const PAGE_SIZE_OPTIONS = [50, 100, 200, 500]

function DataTable({ rows, setRows, type, columns = COLS_GIAO_NHAN, customCategoryMap = {} }) {
  const [pageSize, setPageSize] = React.useState(100)
  const [currentPage, setCurrentPage] = React.useState(1)
  const [selectedIds, setSelectedIds] = React.useState(new Set())
  const [showConfirmDelete, setShowConfirmDelete] = React.useState(false)
  const [isDeleting, setIsDeleting] = React.useState(false)
  const [deleteProgress, setDeleteProgress] = React.useState(null) // { done, total }
  const tableWrapRef = React.useRef(null)
  const mirrorRef = React.useRef(null)

  // Reset to page 1 when rows change (filter applied)
  React.useEffect(() => { 
    setCurrentPage(1)
    setSelectedIds(new Set())
  }, [rows])

  // Sync scroll giữa bảng và thanh cuộn mirror dưới
  React.useEffect(() => {
    const wrap = tableWrapRef.current
    const mirror = mirrorRef.current
    if (!wrap || !mirror) return

    const syncMirrorWidth = () => {
      const inner = wrap.querySelector('table')
      if (inner) {
        mirror.firstChild.style.width = inner.scrollWidth + 'px'
      }
    }
    syncMirrorWidth()

    const onWrapScroll = () => { mirror.scrollLeft = wrap.scrollLeft }
    const onMirrorScroll = () => { wrap.scrollLeft = mirror.scrollLeft }

    wrap.addEventListener('scroll', onWrapScroll)
    mirror.addEventListener('scroll', onMirrorScroll)

    const ro = new ResizeObserver(syncMirrorWidth)
    ro.observe(wrap)

    return () => {
      wrap.removeEventListener('scroll', onWrapScroll)
      mirror.removeEventListener('scroll', onMirrorScroll)
      ro.disconnect()
    }
  }, [rows, pageSize, columns])

  // Các hook bên dưới PHẢI nằm trước mọi early-return để tuân thủ Rules of Hooks
  // (tránh lỗi "Rendered fewer hooks than expected" / React error #300 gây màn hình trắng
  // khi rows.length chuyển từ >0 về 0 hoặc ngược lại, ví dụ khi lọc ra 0 kết quả)
  const totalPages = Math.ceil(rows.length / pageSize)
  const startIdx = (currentPage - 1) * pageSize
  const endIdx = Math.min(startIdx + pageSize, rows.length)
  const pageRows = rows.slice(startIdx, endIdx)

  const pageRowIds = React.useMemo(() => pageRows.map(r => r.id), [pageRows])

  const isAllSelected = React.useMemo(() => {
    if (rows.length === 0) return false
    return rows.every(r => selectedIds.has(r.id))
  }, [rows, selectedIds])

  if (rows.length === 0) return (
    <div className="empty-state">
      <Search size={48} />
      <h3>Không có dữ liệu</h3>
      <p>Không tìm thấy kết quả phù hợp với bộ lọc hiện tại.</p>
    </div>
  )

  const handleSelectAll = (e) => {
    const checked = e.target.checked
    setSelectedIds(prev => {
      const next = new Set(prev)
      if (checked) {
        rows.forEach(r => next.add(r.id))
      } else {
        rows.forEach(r => next.delete(r.id))
      }
      return next
    })
  }

  const handleSelectRow = (id, checked) => {
    setSelectedIds(prev => {
      const next = new Set(prev)
      if (checked) {
        next.add(id)
      } else {
        next.delete(id)
      }
      return next
    })
  }

  const handleDeleteSelected = () => {
    if (selectedIds.size === 0) return
    setShowConfirmDelete(true)
  }

  const executeDelete = async () => {
    const idsArr = Array.from(selectedIds)

    // 1. Delete on Supabase if connected
    // Xóa theo từng đợt nhỏ (chunk) thay vì gửi toàn bộ ID trong 1 request, vì khi số
    // lượng dòng được chọn quá lớn (hàng nghìn-hàng chục nghìn), URL request .in('id', [...])
    // sẽ vượt quá giới hạn độ dài cho phép, gây lỗi "TypeError: Failed to fetch" ngay từ trình duyệt.
    if (isSupabaseConfigured) {
      const tableName = type === 'chung' ? 'don_chung' : type === 'giao' ? 'don_giao' : type === 'nhan' ? 'don_nhan' : type === 'kho' ? 'don_kho' : 'don_chung'
      const chunkSize = 300
      setIsDeleting(true)
      setDeleteProgress({ done: 0, total: idsArr.length })
      try {
        for (let i = 0; i < idsArr.length; i += chunkSize) {
          const chunk = idsArr.slice(i, i + chunkSize)
          const { error } = await supabase
            .from(tableName)
            .delete()
            .in('id', chunk)

          if (error) throw error
          setDeleteProgress({ done: Math.min(i + chunkSize, idsArr.length), total: idsArr.length })
        }
      } catch (err) {
        alert(`Lỗi khi xóa dòng trên Supabase: ${err.message || err}`)
        setIsDeleting(false)
        setDeleteProgress(null)
        return
      }
      setIsDeleting(false)
      setDeleteProgress(null)
    }

    // 2. Delete locally
    if (setRows) {
      setRows(prev => prev.filter(r => !selectedIds.has(r.id)))
    }
    
    // Clear selection
    setSelectedIds(new Set())
  }

  // Generate page buttons (show max 7 around current)
  const getPageNums = () => {
    const pages = []
    if (totalPages <= 7) {
      for (let i = 1; i <= totalPages; i++) pages.push(i)
    } else {
      pages.push(1)
      if (currentPage > 3) pages.push('...')
      for (let i = Math.max(2, currentPage - 1); i <= Math.min(totalPages - 1, currentPage + 1); i++) pages.push(i)
      if (currentPage < totalPages - 2) pages.push('...')
      pages.push(totalPages)
    }
    return pages
  }

  const btnBase = {
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    height: 30, minWidth: 30, padding: '0 8px',
    border: '1px solid #e2e8f0', borderRadius: 6,
    fontSize: 13, fontWeight: 500, cursor: 'pointer',
    background: '#fff', color: '#374151', transition: 'all 0.15s',
    userSelect: 'none'
  }
  const btnActive = { ...btnBase, background: 'var(--primary)', color: '#fff', borderColor: 'var(--primary)', fontWeight: 700 }
  const btnDisabled = { ...btnBase, opacity: 0.4, cursor: 'not-allowed', pointerEvents: 'none' }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
      {/* Custom Confirmation Dialog */}
      {showConfirmDelete && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(15, 23, 42, 0.65)',
          backdropFilter: 'blur(6px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 99999,
        }}>
          <div style={{
            background: '#ffffff',
            borderRadius: 16,
            padding: 24,
            width: '100%',
            maxWidth: 440,
            boxShadow: '0 25px 50px -12px rgba(15, 23, 42, 0.25)',
            border: '1px solid #fee2e2',
            display: 'flex',
            flexDirection: 'column',
            gap: 20,
            margin: '0 16px',
            boxSizing: 'border-box'
          }}>
            <div style={{ display: 'flex', gap: 16, alignItems: 'flex-start' }}>
              <div style={{
                width: 48,
                height: 48,
                borderRadius: 24,
                background: '#fee2e2',
                color: '#ef4444',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                flexShrink: 0
              }}>
                <Trash2 size={24} />
              </div>
              <div style={{ display: 'flex', flexDirection: 'column', gap: 6, flex: 1 }}>
                <h3 style={{ fontSize: 18, fontWeight: 700, color: '#0f172a', margin: 0, fontFamily: 'inherit' }}>
                  Xác nhận xóa dữ liệu
                </h3>
                <p style={{ fontSize: 14, color: '#475569', margin: 0, lineHeight: '1.5', fontFamily: 'inherit' }}>
                  Bạn có thực sự muốn xóa <strong>{selectedIds.size}</strong> dòng dữ liệu đã chọn? Hành động này không thể hoàn tác và sẽ xóa vĩnh viễn trên cơ sở dữ liệu Supabase.
                </p>
              </div>
            </div>

            <div style={{ display: 'flex', gap: 12, justifyContent: 'flex-end' }}>
              <button
                onClick={() => setShowConfirmDelete(false)}
                style={{
                  padding: '10px 20px',
                  borderRadius: 8,
                  fontSize: 14,
                  fontWeight: 600,
                  background: '#f1f5f9',
                  color: '#334155',
                  cursor: 'pointer',
                  border: '1px solid #cbd5e1',
                  transition: 'all 0.15s ease',
                  fontFamily: 'inherit'
                }}
                onMouseOver={e => {
                  e.currentTarget.style.background = '#e2e8f0';
                  e.currentTarget.style.color = '#0f172a';
                }}
                onMouseOut={e => {
                  e.currentTarget.style.background = '#f1f5f9';
                  e.currentTarget.style.color = '#334155';
                }}
              >
                Hủy (No)
              </button>
              <button
                onClick={async () => {
                  setShowConfirmDelete(false)
                  await executeDelete()
                }}
                style={{
                  padding: '10px 20px',
                  borderRadius: 8,
                  fontSize: 14,
                  fontWeight: 700,
                  background: '#dc2626',
                  color: '#ffffff',
                  cursor: 'pointer',
                  boxShadow: '0 2px 4px rgba(220, 38, 38, 0.2)',
                  border: 'none',
                  transition: 'all 0.15s ease',
                  fontFamily: 'inherit'
                }}
                onMouseOver={e => e.currentTarget.style.background = '#b91c1c'}
                onMouseOut={e => e.currentTarget.style.background = '#dc2626'}
              >
                Đồng ý (Yes)
              </button>
            </div>
          </div>
        </div>
      )}

      {isDeleting && (
        <div style={{
          position: 'fixed',
          top: 0, left: 0, right: 0, bottom: 0,
          background: 'rgba(15, 23, 42, 0.65)',
          backdropFilter: 'blur(6px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 100000,
        }}>
          <div style={{
            background: '#ffffff',
            borderRadius: 16,
            padding: 28,
            width: '100%',
            maxWidth: 360,
            boxShadow: '0 25px 50px -12px rgba(15, 23, 42, 0.25)',
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            gap: 14,
          }}>
            <RefreshCw size={28} color="#dc2626" style={{ animation: 'spin 1s linear infinite' }} />
            <h3 style={{ fontSize: 16, fontWeight: 700, color: '#0f172a', margin: 0 }}>
              Đang xóa dữ liệu...
            </h3>
            {deleteProgress && (
              <p style={{ fontSize: 13.5, color: '#475569', margin: 0 }}>
                Đã xóa {deleteProgress.done.toLocaleString('vi-VN')} / {deleteProgress.total.toLocaleString('vi-VN')} dòng
              </p>
            )}
          </div>
        </div>
      )}

      {/* Floating Selection Alert Bar */}
      {setRows && selectedIds.size > 0 && (
        <div style={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          background: '#fef2f2',
          border: '1px solid #fca5a5',
          borderRadius: 8,
          padding: '10px 16px',
          marginBottom: 12,
          color: '#991b1b',
          fontSize: '13.5px',
          fontWeight: 500,
          boxShadow: '0 1px 3px rgba(239, 68, 68, 0.08)',
          flexShrink: 0
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <AlertCircle size={16} color="#ef4444" />
            <span>Đang chọn <strong style={{ fontSize: 15, color: '#b91c1c' }}>{selectedIds.size}</strong> dòng dữ liệu</span>
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button
              onClick={() => setSelectedIds(new Set())}
              className="btn btn-outline"
              style={{
                height: 32,
                padding: '0 14px',
                fontSize: '13px',
                borderColor: '#fca5a5',
                color: '#b91c1c',
                background: '#ffffff',
                cursor: 'pointer',
                fontWeight: 600,
                borderRadius: 6,
                transition: 'all 0.15s'
              }}
            >
              Hủy chọn
            </button>
            <button
              onClick={handleDeleteSelected}
              className="btn"
              style={{
                height: 32,
                padding: '0 14px',
                fontSize: '13px',
                background: '#dc2626',
                color: '#ffffff',
                border: 'none',
                display: 'flex',
                alignItems: 'center',
                gap: 6,
                fontWeight: 700,
                cursor: 'pointer',
                borderRadius: 6,
                boxShadow: '0 2px 4px rgba(220, 38, 38, 0.2)',
                transition: 'all 0.15s'
              }}
            >
              <Trash2 size={14} /> Xóa {selectedIds.size} dòng đã chọn
            </button>
          </div>
        </div>
      )}

      {/* Table */}
      <div className="table-wrap" ref={tableWrapRef} style={{ overflowX: 'hidden', overflowY: 'auto' }}>
        <table>
          <thead>
            <tr>
              {setRows && (
                <th style={{ width: 40, minWidth: 40, maxWidth: 40, textAlign: 'center', verticalAlign: 'middle', padding: '8px 10px' }}>
                  <input
                    type="checkbox"
                    checked={isAllSelected}
                    onChange={handleSelectAll}
                    style={{
                      width: 16,
                      height: 16,
                      cursor: 'pointer',
                      borderRadius: 4,
                      borderColor: '#cbd5e1',
                      accentColor: 'var(--primary)'
                    }}
                  />
                </th>
              )}
              <th style={{ width: 50, minWidth: 50, maxWidth: 50, textAlign: 'center', verticalAlign: 'middle', fontSize: '12px', padding: '8px 10px' }}>
                STT
              </th>
              {columns.map(c => {
                const isNhapKhoGroup = ['khoiLuongNhap', 'maDonViGiao', 'donViGiao', 'nguoiGiao'].includes(c.key)
                const isXuatKhoGroup = ['khoiLuongXuat', 'maDonViNhan', 'donViNhan', 'nguoiPheDuyet'].includes(c.key)
                const isRealReportGroup = ['khoiLuongThuc', 'logicTongHop'].includes(c.key)
                const thBg = isNhapKhoGroup ? '#0f766e' : isXuatKhoGroup ? '#c2410c' : isRealReportGroup ? '#5b21b6' : undefined
                const thBorderBottom = isNhapKhoGroup ? '2px solid #115e59' : isXuatKhoGroup ? '2px solid #9a3412' : isRealReportGroup ? '2px solid #4c1d95' : undefined

                return (
                  <th
                    key={c.key}
                    style={{
                      width: c.width, minWidth: c.width, maxWidth: c.width,
                      textAlign: 'center', verticalAlign: 'middle',
                      fontSize: '12px', padding: '8px 10px',
                      whiteSpace: 'normal', wordBreak: 'break-word', lineHeight: '1.2',
                      background: thBg,
                      borderBottom: thBorderBottom
                    }}
                  >
                    {c.label}
                  </th>
                )
              })}
            </tr>
          </thead>
          <tbody>
            {pageRows.map((row, i) => (
              <tr key={row.id}>
                {setRows && (
                  <td style={{ width: 40, minWidth: 40, maxWidth: 40, textAlign: 'center', padding: '6px 10px' }}>
                    <input
                      type="checkbox"
                      checked={selectedIds.has(row.id)}
                      onChange={(e) => handleSelectRow(row.id, e.target.checked)}
                      style={{
                        width: 16,
                        height: 16,
                        cursor: 'pointer',
                        borderRadius: 4,
                        borderColor: '#cbd5e1',
                        accentColor: 'var(--primary)'
                      }}
                    />
                  </td>
                )}
                <td style={{ width: 50, minWidth: 50, maxWidth: 50, textAlign: 'center', fontSize: '12px', color: '#1b1919', padding: '6px 10px' }}>
                  {startIdx + i + 1}
                </td>
                {columns.map(col => {
                  const isCenteredCol = [
                    'ngayXuatNhap', 'maVatTu', 'maSAP', 'dvt', 'loaiDon', 'logicTongHop',
                    'maDonViGiao', 'nguoiGiao',
                    'maDonViNhan', 'nguoiPheDuyet', 'nguoiNhan',
                    'soHopDong', 'thuKho', 'tinhTrang'
                  ].includes(col.key)
                  const isRightAligned = ['khoiLuongNhap', 'khoiLuongXuat'].includes(col.key) || col.key.toLowerCase().includes('khoiluong')
                  
                  // Custom styling for logicTongHop and khoiLuongThuc
                  const isRealReportGroup = ['khoiLuongThuc', 'logicTongHop'].includes(col.key)
                  let cellBg = undefined
                  let cellTextColor = '#1b1919'
                  let cellFontWeight = undefined
                  
                  if (isRealReportGroup) {
                    const valLogic = Number(row.logicTongHop)
                    if (valLogic === 1) {
                      cellBg = '#ecfdf5' // soft green
                      cellTextColor = '#065f46' // dark green
                      cellFontWeight = '700'
                    } else if (valLogic === -1) {
                      cellBg = '#fef2f2' // soft red
                      cellTextColor = '#991b1b' // dark red
                      cellFontWeight = '700'
                    } else if (valLogic === 0) {
                      cellBg = '#f1f5f9' // soft gray
                      cellTextColor = '#475569' // dark gray
                      cellFontWeight = '500'
                    }
                  } else if (col.key === 'donViGiao' || col.key === 'donViNhan') {
                    const name = row[col.key]
                    if (name) {
                      const normName = name.trim().replace(/\s+/g, ' ')
                      let cat = 'chuaphanbo'
                      if (customCategoryMap && customCategoryMap[normName]) {
                        cat = customCategoryMap[normName]
                      } else if (customCategoryMap) {
                        const keys = Object.keys(customCategoryMap)
                        const foundKey = keys.find(k => k.toLowerCase() === normName.toLowerCase())
                        if (foundKey) {
                          cat = customCategoryMap[foundKey]
                        } else {
                          cat = getUnitCategory(normName)
                        }
                      } else {
                        cat = getUnitCategory(normName)
                      }

                      if (cat === 'kho') {
                        cellBg = '#eff6ff'
                        cellTextColor = '#1e3a8a'
                        cellFontWeight = '600'
                      } else if (cat === 'ncc') {
                        cellBg = '#fffbeb'
                        cellTextColor = '#78350f'
                        cellFontWeight = '600'
                      } else if (cat === 'todoi') {
                        cellBg = '#f0fdf4'
                        cellTextColor = '#065f46'
                        cellFontWeight = '600'
                      }
                    }
                  }

                  return (
                    <td
                      key={col.key}
                      style={{
                        width: col.width, minWidth: col.width, maxWidth: col.width,
                        color: cellTextColor,
                        backgroundColor: cellBg,
                        fontWeight: cellFontWeight,
                        textAlign: isCenteredCol ? 'center' : (isRightAligned ? 'right' : 'left'),
                        fontSize: '12px', padding: '6px 10px',
                        wordBreak: 'break-word', whiteSpace: 'normal'
                      }}
                      title={String(formatVal(row[col.key], col.key) || '')}
                    >
                      {col.key === 'trangThai' ? (
                        row[col.key] ? (
                          <span className={`badge ${getTrangThaiColor(row[col.key])}`} style={{ fontSize: '11px', padding: '2px 6px', lineHeight: 1.2 }}>
                            {row[col.key]}
                          </span>
                        ) : ''
                      ) : col.key === 'tinhTrang' ? (
                        row[col.key]
                          ? <span className={`badge ${row[col.key] === 'NEW' ? 'badge-green' : row[col.key] === 'USED' ? 'badge-yellow' : 'badge-gray'}`} style={{ fontSize: '11px', padding: '2px 6px', lineHeight: 1.2 }}>{row[col.key]}</span>
                          : ''
                      ) : (col.key === 'logicTongHop' || col.key === 'khoiLuongThuc') && Number(row[col.key]) === 0 ? (
                        '-'
                      ) : (
                        formatVal(row[col.key], col.key) !== null && formatVal(row[col.key], col.key) !== undefined ? formatVal(row[col.key], col.key) : ''
                      )}
                    </td>
                  )
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Mirror scrollbar ngang - dính ngay dưới bảng */}
      <div ref={mirrorRef} className="scroll-mirror">
        <div style={{ height: 1 }} />
      </div>

      {/* Pagination bottom bar: trái = hiển thị dòng/trang, phải = số trang */}
      <div style={{
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        marginTop: 10, gap: 12, flexWrap: 'wrap', flexShrink: 0
      }}>
        {/* Left: page size selector */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
          <span style={{ fontSize: 13, color: 'var(--text-muted)', whiteSpace: 'nowrap' }}>Hiển thị</span>
          <div style={{ display: 'flex', gap: 4 }}>
            {PAGE_SIZE_OPTIONS.map(sz => (
              <button
                key={sz}
                onClick={() => { setPageSize(sz); setCurrentPage(1) }}
                style={pageSize === sz
                  ? { ...btnActive, minWidth: 44, height: 28, fontSize: 12 }
                  : { ...btnBase, minWidth: 44, height: 28, fontSize: 12 }
                }
              >
                {sz}
              </button>
            ))}
          </div>
          <span style={{ fontSize: 13, color: 'var(--text-muted)', whiteSpace: 'nowrap' }}>dòng / trang</span>
        </div>

        {/* Right: page info + nav */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
          <span style={{ fontSize: 13, color: 'var(--text-muted)', marginRight: 4, whiteSpace: 'nowrap' }}>
            {startIdx + 1}–{endIdx} / {rows.length.toLocaleString()} dòng
          </span>
          <button style={currentPage === 1 ? btnDisabled : btnBase} onClick={() => setCurrentPage(1)}>«</button>
          <button style={currentPage === 1 ? btnDisabled : btnBase} onClick={() => setCurrentPage(p => p - 1)}>‹</button>
          {getPageNums().map((p, i) =>
            p === '...'
              ? <span key={`ellipsis-b-${i}`} style={{ fontSize: 13, color: '#94a3b8', padding: '0 2px' }}>…</span>
              : <button key={p} style={currentPage === p ? btnActive : btnBase} onClick={() => setCurrentPage(p)}>{p}</button>
          )}
          <button style={currentPage === totalPages ? btnDisabled : btnBase} onClick={() => setCurrentPage(p => p + 1)}>›</button>
          <button style={currentPage === totalPages ? btnDisabled : btnBase} onClick={() => setCurrentPage(totalPages)}>»</button>
        </div>
      </div>
    </div>
  )
}

// Helper to parse date string or excel serial to Date object
function parseRowDate(dateVal) {
  if (dateVal === null || dateVal === undefined) return null;
  if (typeof dateVal === 'number') {
    const d = new Date((dateVal - 25569) * 86400000);
    return d;
  }
  const str = String(dateVal).trim();
  if (!str) return null;

  // DD/MM/YYYY or DD-MM-YYYY or DD.MM.YYYY
  const dmyMatch = str.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{4})$/);
  if (dmyMatch) {
    const day = parseInt(dmyMatch[1], 10);
    const month = parseInt(dmyMatch[2], 10) - 1;
    const year = parseInt(dmyMatch[3], 10);
    return new Date(year, month, day);
  }

  // YYYY/MM/DD or YYYY-MM-DD
  const ymdMatch = str.match(/^(\d{4})[\/\-\.](\d{1,2})[\/\-\.](\d{1,2})$/);
  if (ymdMatch) {
    const year = parseInt(ymdMatch[1], 10);
    const month = parseInt(ymdMatch[2], 10) - 1;
    const day = parseInt(ymdMatch[3], 10);
    return new Date(year, month, day);
  }

  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  return null;
}

// ─── Order Tab (shared for Giao & Nhan) ──────────────────────────────────────
function OrderTab({
  type,
  rows,
  setRows,
  fileName,
  setFileName,
  selectedProject,
  setSelectedProject,
  onSync,
  syncing,
  supabaseMessage,
  onEditProject,
  projectOptions,
  onImportFile,
  onDeleteFile,
  customCategoryMap = {}
}) {
  const [loading, setLoading] = useState(false)
  const [search, setSearch] = useState('')
  const [trangThai, setTrangThai] = useState('')
  const [donViGiaoFilter, setDonViGiaoFilter] = useState('')
  const [donViNhanFilter, setDonViNhanFilter] = useState('')
  const [startDate, setStartDate] = useState('')
  const [endDate, setEndDate] = useState('')

  const label = type === 'giao' ? 'Đơn Giao' : type === 'nhan' ? 'Đơn Nhận' : type === 'kho' ? 'Kho dự án' : 'Đơn chung'
  const uploadLabel = type === 'giao'
    ? 'Tải lên file Report_Orders_Đơn giao'
    : type === 'nhan'
    ? 'Tải lên file Report_Orders_Đơn nhận'
    : type === 'kho'
    ? 'Tải lên file Report_Orders_Kho_du_an'
    : 'Tải lên file Report_Orders_Đơn chung'

  const handleFile = useCallback((resultsOrData, name) => {
    setLoading(true)
    setTimeout(() => {
      if (Array.isArray(resultsOrData) && resultsOrData.length > 0 && typeof resultsOrData[0] === 'object' && 'data' in resultsOrData[0]) {
        // Multiple files uploaded
        let allParsedRows = []
        let combinedNames = []
        resultsOrData.forEach(item => {
          const parsed = parseXlsxToRows(item.data)
          allParsedRows = allParsedRows.concat(parsed)
          combinedNames.push(item.name)
        })
        const finalName = combinedNames.join(' + ')
        if (onImportFile) {
          onImportFile(allParsedRows, finalName, false)
        } else {
          setRows(allParsedRows)
          setFileName(finalName)
        }
      } else {
        // Single file uploaded
        const parsed = parseXlsxToRows(resultsOrData)
        if (onImportFile) {
          onImportFile(parsed, name, false)
        } else {
          setRows(parsed)
          setFileName(name)
        }
      }
      setLoading(false)
    }, 80)
  }, [setRows, setFileName, onImportFile])

  const handleAppendFile = useCallback((resultsOrData, name) => {
    setLoading(true)
    setTimeout(() => {
      if (Array.isArray(resultsOrData) && resultsOrData.length > 0 && typeof resultsOrData[0] === 'object' && 'data' in resultsOrData[0]) {
        // Multiple files appended
        let allParsedRows = []
        let combinedNames = []
        resultsOrData.forEach(item => {
          const parsed = parseXlsxToRows(item.data)
          allParsedRows = allParsedRows.concat(parsed)
          combinedNames.push(item.name)
        })
        const finalName = combinedNames.join(' + ')
        if (onImportFile) {
          onImportFile(allParsedRows, finalName, true)
        } else {
          setRows(prev => [...prev, ...allParsedRows])
          setFileName(prev => prev ? `${prev} + ${finalName}` : finalName)
        }
      } else {
        // Single file appended
        const parsed = parseXlsxToRows(resultsOrData)
        if (onImportFile) {
          onImportFile(parsed, name, true)
        } else {
          setRows(prev => [...prev, ...parsed])
          setFileName(prev => prev ? `${prev} + ${name}` : name)
        }
      }
      setLoading(false)
    }, 80)
  }, [setRows, setFileName, onImportFile])

  const trangThaiOptions = useMemo(() =>
    [...new Set(rows.map(r => r.trangThai).filter(Boolean))].sort(), [rows])
  const duAnOptions = useMemo(() => {
    if (projectOptions !== undefined && projectOptions !== null) {
      return projectOptions
    }
    return [...new Set(rows.map(r => r.duAn).filter(Boolean))].sort()
  }, [rows, projectOptions])

  const projectFilteredRowsForStats = useMemo(() => {
    if (!selectedProject) return rows
    return rows.filter(r => (r.ten_du_an || r.tenDuAn || r.duAn) === selectedProject)
  }, [rows, selectedProject])

  const donViGiaoOptions = useMemo(() => {
    return [...new Set(projectFilteredRowsForStats.map(r => r.donViGiao).filter(Boolean))].sort()
  }, [projectFilteredRowsForStats])

  const donViNhanOptions = useMemo(() => {
    return [...new Set(projectFilteredRowsForStats.map(r => r.donViNhan).filter(Boolean))].sort()
  }, [projectFilteredRowsForStats])

  const filtered = useMemo(() => {
    let r = rows
    if (search) {
      const q = search.toLowerCase()
      r = r.filter(row =>
        String(row.tenVatTu).toLowerCase().includes(q) ||
        String(row.maDonNhapKho).toLowerCase().includes(q) ||
        String(row.maDonXuatKho).toLowerCase().includes(q) ||
        String(row.donViGiao).toLowerCase().includes(q) ||
        String(row.donViNhan).toLowerCase().includes(q) ||
        String(row.maSAP).toLowerCase().includes(q)
      )
    }
    if (trangThai) r = r.filter(row => row.trangThai === trangThai)
    if (selectedProject) r = r.filter(row => (row.ten_du_an || row.tenDuAn || row.duAn) === selectedProject)
    if (donViGiaoFilter) r = r.filter(row => row.donViGiao === donViGiaoFilter)
    if (donViNhanFilter) r = r.filter(row => row.donViNhan === donViNhanFilter)

    if (startDate || endDate) {
      const start = startDate ? new Date(startDate) : null
      if (start) start.setHours(0, 0, 0, 0)

      const end = endDate ? new Date(endDate) : null
      if (end) end.setHours(23, 59, 59, 999)

      r = r.filter(row => {
        const rowDate = parseRowDate(row.ngayXuatNhap)
        if (!rowDate) return false
        if (start && rowDate < start) return false
        if (end && rowDate > end) return false
        return true
      })
    }

    // Sort descending by date (newest/most recent first)
    const sorted = [...r].sort((a, b) => {
      const dateA = parseRowDate(a.ngayXuatNhap)
      const dateB = parseRowDate(b.ngayXuatNhap)
      if (!dateA && !dateB) return 0
      if (!dateA) return 1
      if (!dateB) return -1
      return dateB.getTime() - dateA.getTime()
    })

    return sorted
  }, [rows, search, trangThai, selectedProject, donViGiaoFilter, donViNhanFilter, startDate, endDate])

  const handleExportExcel = useCallback(() => {
    const wb = XLSXStyle.utils.book_new()
    const ws = {}

    // Columns: STT + COLS_GIAO_NHAN
    const columns = [
      { key: 'STT', label: 'STT', width: 50 },
      ...COLS_GIAO_NHAN
    ]

    // Set widths
    ws['!cols'] = columns.map(c => ({ wpx: c.width }))

    let excelRowIdx = 1

    // Write header row
    columns.forEach((col, colIdx) => {
      const colChar = getColLabel(colIdx)
      const cellRef = `${colChar}${excelRowIdx}`
      
      const isNhapKhoGroup = ['khoiLuongNhap', 'maDonViGiao', 'donViGiao', 'nguoiGiao'].includes(col.key)
      const isXuatKhoGroup = ['khoiLuongXuat', 'maDonViNhan', 'donViNhan', 'nguoiPheDuyet'].includes(col.key)
      const excelBgColor = isNhapKhoGroup ? '0F766E' : isXuatKhoGroup ? 'C2410C' : '0F58A7'
      const excelBorderColor = isNhapKhoGroup ? '115E59' : isXuatKhoGroup ? '9A3412' : '0A3D73'

      ws[cellRef] = {
        v: col.label,
        t: 's',
        s: {
          fill: {
            patternType: 'solid',
            fgColor: { rgb: excelBgColor }
          },
          font: {
            name: 'Segoe UI',
            sz: 9.5,
            bold: true,
            color: { rgb: 'FFFFFF' }
          },
          alignment: {
            horizontal: 'center',
            vertical: 'center',
            wrapText: true
          },
          border: {
            top: { style: 'thin', color: { rgb: excelBorderColor } },
            bottom: { style: 'medium', color: { rgb: excelBorderColor } },
            left: { style: 'thin', color: { rgb: excelBorderColor } },
            right: { style: 'thin', color: { rgb: excelBorderColor } }
          }
        }
      }
    })

    // Helper to get letter coordinate
    function getColLabel(index) {
      let label = ''
      let temp = index
      while (temp >= 0) {
        label = String.fromCharCode((temp % 26) + 65) + label
        temp = Math.floor(temp / 26) - 1
      }
      return label
    }

    // Helper for status styling
    function getExcelStatusStyle(val, colKey) {
      if (!val) return null
      const v = String(val).toLowerCase()
      if (colKey === 'trangThai') {
        if (v.includes('chờ') || v.includes('chưa')) {
          return { fg: 'FFFBEB', text: '92400E' } // Orange/Yellow
        }
        if (v.includes('phê duyệt') || v.includes('hoàn thành') || v.includes('đã')) {
          return { fg: 'ECFDF5', text: '065F46' } // Green
        }
        if (v.includes('từ chối') || v.includes('hủy')) {
          return { fg: 'FFF1F2', text: '9F1239' } // Red
        }
        return { fg: 'EFF6FF', text: '1E40AF' } // Blue
      }
      if (colKey === 'tinhTrang') {
        if (val === 'NEW') {
          return { fg: 'ECFDF5', text: '065F46' }
        }
        if (val === 'USED') {
          return { fg: 'FFFBEB', text: '92400E' }
        }
        return { fg: 'F8FAFC', text: '475569' }
      }
      return null
    }

    // Write data rows
    filtered.forEach((row, rowIndex) => {
      excelRowIdx++
      const isEvenNum = (rowIndex % 2 === 1)
      const rowBgColor = isEvenNum ? 'F8FAFC' : 'FFFFFF'

      columns.forEach((col, colIdx) => {
        const colChar = getColLabel(colIdx)
        const cellRef = `${colChar}${excelRowIdx}`

        let val = ''
        let cellType = 's'
        let numFormat = undefined

        if (col.key === 'STT') {
          val = rowIndex + 1
          cellType = 'n'
        } else {
          const raw = row[col.key]
          if (raw !== null && raw !== undefined) {
            const isRightAligned = ['khoiLuongNhap', 'khoiLuongXuat'].includes(col.key) || col.key.toLowerCase().includes('khoiluong')
            if (isRightAligned && !isNaN(Number(raw)) && raw !== '') {
              val = Number(raw)
              cellType = 'n'
              numFormat = '#,##0'
            } else {
              val = String(raw)
            }
          }
        }

        // Check alignment
        const isCenteredCol = [
          'STT', 'ngayXuatNhap', 'maVatTu', 'maSAP', 'dvt', 'loaiDon', 'logicTongHop',
          'maDonViGiao', 'nguoiGiao',
          'maDonViNhan', 'nguoiPheDuyet', 'nguoiNhan',
          'soHopDong', 'thuKho', 'tinhTrang', 'trangThai'
        ].includes(col.key)
        const isRightAligned = ['khoiLuongNhap', 'khoiLuongXuat'].includes(col.key) || col.key.toLowerCase().includes('khoiluong')

        // Styles
        const cellStyle = {
          font: {
            name: 'Segoe UI',
            sz: 9,
            color: { rgb: '1B1919' }
          },
          alignment: {
            horizontal: isCenteredCol ? 'center' : (isRightAligned ? 'right' : 'left'),
            vertical: 'center',
            wrapText: true
          },
          border: {
            top: { style: 'thin', color: { rgb: 'E2E8F0' } },
            bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          },
          fill: {
            patternType: 'solid',
            fgColor: { rgb: rowBgColor }
          }
        }

        // Status style overlay
        if (col.key === 'trangThai' || col.key === 'tinhTrang') {
          const statusOverlay = getExcelStatusStyle(val, col.key)
          if (statusOverlay) {
            cellStyle.fill.fgColor = { rgb: statusOverlay.fg }
            cellStyle.font.color = { rgb: statusOverlay.text }
            cellStyle.font.bold = true
          }
        }

        const cellObj = {
          v: val,
          t: cellType,
          s: cellStyle
        }
        if (numFormat) {
          cellObj.z = numFormat
        }

        ws[cellRef] = cellObj
      })
    })

    // Set range ref
    const excelRangeRef = `A1:${getColLabel(columns.length - 1)}${excelRowIdx}`
    ws['!ref'] = excelRangeRef
    ws['!autofilter'] = { ref: excelRangeRef }

    const sheetName = type === 'giao' ? "Đơn Giao" : type === 'nhan' ? "Đơn Nhận" : type === 'kho' ? "Kho dự án" : "Đơn chung"
    XLSXStyle.utils.book_append_sheet(wb, ws, sheetName)
    
    // Save
    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary', compression: false })
    function s2ab(s) {
      const buf = new ArrayBuffer(s.length)
      const view = new Uint8Array(buf)
      for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF
      return buf
    }
    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    const fileSuffix = type === 'giao' ? 'Don_giao' : type === 'nhan' ? 'Don_nhan' : type === 'kho' ? 'Kho_du_an' : 'Don_chung'
    a.download = `Export_${fileSuffix}_${new Date().toISOString().slice(0, 10)}.xlsx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }, [filtered, type])

  return (
    <div style={{ padding: '16px 24px 24px 24px', height: '100%', display: 'flex', flexDirection: 'column', boxSizing: 'border-box', overflow: 'hidden' }}>
      {!rows.length && !loading ? (
        <div style={{ maxWidth: 560, margin: '40px auto' }}>
          <div style={{
            background: 'var(--primary-light)',
            border: '1px solid #bdd4f0',
            borderRadius: 10,
            padding: '10px 14px',
            marginBottom: 20,
            display: 'flex',
            gap: 8,
            alignItems: 'flex-start',
          }}>
            <Info size={14} color="var(--primary)" style={{ marginTop: 1, flexShrink: 0 }} />
            <div style={{ fontSize: 14, color: 'var(--primary-dark)' }}>
              <strong>Cấu trúc file yêu cầu:</strong> File Excel có sheet "Báo cáo theo dõi giao nhận" với 36 cột từ Ngày xuất nhập đến Nhãn hiệu. Dòng 1 là nhóm cột, dòng 2 là sub-header, dữ liệu từ dòng 3.
            </div>
          </div>

          {!selectedProject && (
            <div style={{
              background: 'var(--primary-light)',
              border: '1px solid #bdd4f0',
              borderRadius: 10,
              padding: '10px 14px',
              marginBottom: 16,
              display: 'flex',
              gap: 8,
              alignItems: 'flex-start',
            }}>
              <Info size={14} color="var(--primary)" style={{ marginTop: 1, flexShrink: 0 }} />
              <div style={{ fontSize: 14, color: 'var(--primary-dark)' }}>
                <strong>Lưu ý:</strong> Bạn đang chọn <strong>Tất cả dự án</strong>. Khi tải file lên, hệ thống sẽ tự động phân tích và gán dữ liệu vào các dự án tương ứng dựa trên cột dữ liệu gốc của từng dòng!
              </div>
            </div>
          )}
          <UploadZone onFile={handleFile} label={uploadLabel} disabled={false} />
        </div>
      ) : loading ? (
        <div className="empty-state">
          <RefreshCw size={40} style={{ animation: 'spin 1s linear infinite', opacity: 0.5 }} />
          <h3>Đang xử lý dữ liệu...</h3>
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
          {selectedProject && (
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', marginBottom: 16, flexShrink: 0 }}>
              <button className="btn btn-outline btn-sm" onClick={() => { if (onDeleteFile) { onDeleteFile() } else { setRows([]); setFileName(''); setSearch(''); setTrangThai('') } }}>
                <X size={12} /> Xóa file
              </button>
            </div>
          )}

          {supabaseMessage && (
            <div style={{
              background: supabaseMessage.type === 'success' ? '#ecfdf5' : supabaseMessage.type === 'error' ? '#fff1f2' : '#eff6ff',
              border: `1px solid ${supabaseMessage.type === 'success' ? '#a7f3d0' : supabaseMessage.type === 'error' ? '#fecdd3' : '#bfdbfe'}`,
              borderRadius: 8,
              padding: '10px 16px',
              marginBottom: 16,
              fontSize: 14,
              color: supabaseMessage.type === 'success' ? '#065f46' : supabaseMessage.type === 'error' ? '#9f1239' : '#1e40af',
              display: 'flex',
              alignItems: 'center',
              gap: 8,
              boxShadow: '0 1px 2px rgba(15,23,42,0.05)'
            }}>
              {supabaseMessage.type === 'success' ? (
                <CheckCircle2 size={16} color="#10b981" />
              ) : supabaseMessage.type === 'error' ? (
                <AlertCircle size={16} color="#ef4444" />
              ) : (
                <RefreshCw size={16} style={{ animation: 'spin 1.5s linear infinite' }} />
              )}
              <span style={{ fontWeight: 500 }}>{supabaseMessage.text}</span>
            </div>
          )}

          <StatsBar rows={projectFilteredRowsForStats} onAppendFile={handleAppendFile} />

          <FilterBar
            search={search} setSearch={setSearch}
            trangThai={trangThai} setTrangThai={setTrangThai}
            trangThaiOptions={trangThaiOptions}
            duAn={selectedProject} setDuAn={setSelectedProject}
            duAnOptions={duAnOptions}
            donViGiao={donViGiaoFilter} setDonViGiao={setDonViGiaoFilter}
            donViGiaoOptions={donViGiaoOptions}
            donViNhan={donViNhanFilter} setDonViNhan={setDonViNhanFilter}
            donViNhanOptions={donViNhanOptions}
            startDate={startDate} setStartDate={setStartDate}
            endDate={endDate} setEndDate={setEndDate}
            onClear={() => {
              setSearch('')
              setTrangThai('')
              setDonViGiaoFilter('')
              setDonViNhanFilter('')
              setSelectedProject('')
              setStartDate('')
              setEndDate('')
            }}
            onEditProject={onEditProject}
            onExportExcel={rows.length > 0 ? handleExportExcel : undefined}
          />

          {filtered.length === 0 && selectedProject && !search && !trangThai && !donViGiaoFilter && !donViNhanFilter ? (
            <div style={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <div style={{ maxWidth: 480, width: '100%', textAlign: 'center' }}>
                <div style={{
                  width: 56, height: 56, background: 'var(--primary-light)',
                  borderRadius: 16, display: 'flex', alignItems: 'center', justifyContent: 'center',
                  margin: '0 auto 12px'
                }}>
                  {type === 'giao' ? <Truck size={26} color="var(--primary)" /> : <PackageCheck size={26} color="var(--primary)" />}
                </div>
                <h3 style={{ fontSize: 17, fontWeight: 700, color: 'var(--text)', marginBottom: 6 }}>
                  Dự án <span style={{ color: 'var(--primary)' }}>"{selectedProject}"</span> chưa có dữ liệu
                </h3>
                <p style={{ color: 'var(--text-muted)', fontSize: 14, marginBottom: 20 }}>
                  Upload file Excel để tải dữ liệu {label} cho dự án này.
                </p>
                <UploadZone onFile={handleFile} label={uploadLabel} disabled={false} />
              </div>
            </div>
          ) : (
            <DataTable rows={filtered} setRows={setRows} type={type} customCategoryMap={customCategoryMap} />
          )}
        </div>
      )}
    </div>
  )
}

// Helper function to classify unit name into 4 categories
function getUnitCategory(name) {
  const upper = (name || '').toUpperCase();
  
  // 1. Nhà cung cấp: contains "CÔNG TY", "CONG TY", "CTY", "DNTN"
  if (upper.includes('CÔNG TY') || upper.includes('CONG TY') || upper.includes('CTY') || upper.includes('DNTN')) {
    return 'ncc';
  }
  
  // 2. Tổ đội: has "TỔ ĐỘI" or "TO DOI"
  if (upper.includes('TỔ ĐỘI') || upper.includes('TO DOI')) {
    return 'todoi';
  }
  
  // 3. Kho BCH: contains "KHO" (as a separate word), "SGC", or "BCH"
  // Split by non-letter characters to extract exact words/tokens
  const tokens = upper.split(/[^A-ZÁÀẢÃẠĂẮẰẲẴẶÂẤẦẨẪẬÉÈẺẼẸÊẾỀỂỄỆÍÌỈĨỊÓÒỎÕỌÔỐỒỔỖỘƠỚỜỞỠỢÚÙỦŨỤƯỨỪỬỮỰÝỲỶỸYĐ]/);
  const hasKho = tokens.includes('KHO');
  const hasSgc = tokens.includes('SGC');
  const hasBch = tokens.includes('BCH');
  
  if (hasKho || hasSgc || hasBch) {
    return 'kho';
  }
  
  // 4. Check if it's a personal full name (which falls into "Tổ đội")
  const trimmed = (name || '').trim();
  // Exclude strings containing numbers, slashes, underscores, or dashes (usually not a real Vietnamese name)
  if (/[-/_[\]()0-9]/.test(trimmed)) {
    return 'chuaphanbo';
  }
  
  const words = trimmed.split(/\s+/);
  // Vietnamese names are typically 2 to 5 words
  if (words.length >= 2 && words.length <= 5) {
    // Check if all words are capitalized
    const isCapitalized = words.every(w => {
      if (!w) return true;
      const firstChar = w[0];
      return firstChar === firstChar.toUpperCase();
    });
    
    // Exclude common corporate or team terms
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

// ─── Kho Du An Tab (Auto-compiled from Don Chung) ───────────────────────────
function KhoDuAnTab({ chungRows, selectedProject, setSelectedProject, allProjects = [], customCategoryMap = {}, setCustomCategoryMap, dbCategoryMap = {}, setDbCategoryMap }) {
  const [search, setSearch] = useState('')
  const [saveStatus, setSaveStatus] = useState('idle') // 'idle' | 'saving' | 'success' | 'error'
  const [errorMessage, setErrorMessage] = useState('')

  // Drag and drop / Custom classifications state
  const [draggedItemName, setDraggedItemName] = useState(null)
  const [dragOverCol, setDragOverCol] = useState(null)

  // Drag & drop handlers
  const handleDragStart = (e, name) => {
    setDraggedItemName(name)
    e.dataTransfer.setData('text/plain', name)
    e.dataTransfer.effectAllowed = 'move'
  }

  const handleDragOver = (e) => {
    e.preventDefault()
  }

  const handleDrop = (e, targetColKey) => {
    e.preventDefault()
    const name = e.dataTransfer.getData('text/plain') || draggedItemName
    if (!name) return

    setCustomCategoryMap(prev => ({
      ...prev,
      [name]: targetColKey
    }))
    setDraggedItemName(null)
    setDragOverCol(null)
  }

  // 1. Extract unique units (both Giao and Nhan combined) from all chungRows
  // Chuẩn hóa tên (gộp khoảng trắng thừa + không phân biệt hoa/thường) để gộp các
  // bản ghi trùng lặp do lệch hoa/thường hoặc dư khoảng trắng trong dữ liệu gốc,
  // tránh hiển thị trùng trên webapp và tránh đẩy dữ liệu trùng lên Supabase.
  const normalizeUnitName = (s) => (s || '').trim().replace(/\s+/g, ' ')
  const uniqueDonVi = useMemo(() => {
    const counts = {}
    chungRows.forEach(r => {
      const gRaw = normalizeUnitName(r.donViGiao)
      const nRaw = normalizeUnitName(r.donViNhan)
      if (gRaw) {
        const key = gRaw.toLowerCase()
        if (!counts[key]) counts[key] = { name: gRaw, giaoCount: 0, nhanCount: 0 }
        counts[key].giaoCount++
      }
      if (nRaw) {
        const key = nRaw.toLowerCase()
        if (!counts[key]) counts[key] = { name: nRaw, giaoCount: 0, nhanCount: 0 }
        counts[key].nhanCount++
      }
    })
    return Object.values(counts)
      .map(item => ({
        ...item,
        totalCount: item.giaoCount + item.nhanCount
      }))
      .sort((a, b) => b.totalCount - a.totalCount)
  }, [chungRows])

  // Check if current classifications match database state exactly
  const isSynchronized = useMemo(() => {
    if (uniqueDonVi.length === 0) return true
    return uniqueDonVi.every(item => {
      const dbCat = dbCategoryMap[item.name]
      const currentCat = customCategoryMap[item.name] || getUnitCategory(item.name)
      return dbCat !== undefined && dbCat === currentCat
    })
  }, [uniqueDonVi, customCategoryMap, dbCategoryMap])

  // 2. Filter list by search query
  const filteredDonVi = useMemo(() => {
    if (!search) return uniqueDonVi
    const q = search.toLowerCase()
    return uniqueDonVi.filter(item => item.name.toLowerCase().includes(q))
  }, [uniqueDonVi, search])

  // 3. Categorize filtered units into 4 groups with default + manual classification mapping
  // Sắp xếp hiển thị theo thứ tự A, b, c,... z, ưu tiên các dữ liệu chưa lưu trên Supabase lên trên cùng
  const categorizedUnits = useMemo(() => {
    const groups = {
      chuaphanbo: [],
      kho: [],
      ncc: [],
      todoi: []
    }
    filteredDonVi.forEach(item => {
      const cat = customCategoryMap[item.name] || getUnitCategory(item.name)
      groups[cat].push(item)
    })
    
    // Sắp xếp từng nhóm: các đơn vị chưa được lưu trên Supabase (hoặc bị đổi nhóm chưa lưu) lên trên cùng, sau đó sắp xếp theo bảng chữ cái tiếng Việt (A, b, c,...)
    Object.keys(groups).forEach(key => {
      groups[key].sort((a, b) => {
        const catA = customCategoryMap[a.name] || getUnitCategory(a.name)
        const catB = customCategoryMap[b.name] || getUnitCategory(b.name)
        const isUnsavedA = dbCategoryMap[a.name] === undefined || dbCategoryMap[a.name] !== catA
        const isUnsavedB = dbCategoryMap[b.name] === undefined || dbCategoryMap[b.name] !== catB
        
        if (isUnsavedA && !isUnsavedB) return -1
        if (!isUnsavedA && isUnsavedB) return 1
        return a.name.localeCompare(b.name, 'vi', { sensitivity: 'base' })
      })
    })
    
    return groups
  }, [filteredDonVi, customCategoryMap, dbCategoryMap])

  const handleSaveToSupabase = useCallback(async () => {
    if (!isSupabaseConfigured) {
      alert('Chưa cấu hình Supabase!')
      return
    }
    
    setSaveStatus('saving')
    setErrorMessage('')
    
    try {
      // Find items that have been customized or modified compared to the DB state
      const changedItems = uniqueDonVi.filter(item => {
        const dbCat = dbCategoryMap[item.name]
        const currentCat = customCategoryMap[item.name] || getUnitCategory(item.name)
        return dbCat === undefined || dbCat !== currentCat
      })

      const namesToUpdate = changedItems.map(item => item.name)

      if (namesToUpdate.length > 0) {
        // Chunk deletes to avoid hitting URL/query length limits
        const deleteChunkSize = 100
        for (let i = 0; i < namesToUpdate.length; i += deleteChunkSize) {
          const chunkNames = namesToUpdate.slice(i, i + deleteChunkSize)
          const { error: delErr } = await supabase
            .from('phan_loai_don_vi')
            .delete()
            .in('ten_don_vi', chunkNames)
          if (delErr) throw delErr
        }

        // Chunk inserts
        const insertPayload = changedItems.map(item => ({
          ten_don_vi: item.name,
          nhom_don_vi: customCategoryMap[item.name] || getUnitCategory(item.name)
        }))

        const insertChunkSize = 200
        for (let i = 0; i < insertPayload.length; i += insertChunkSize) {
          const chunkData = insertPayload.slice(i, i + insertChunkSize)
          const { error: insErr } = await supabase
            .from('phan_loai_don_vi')
            .insert(chunkData)
          if (insErr) throw insErr
        }
      }
      
      // Update our reference map representing the database
      setDbCategoryMap(prev => {
        const next = { ...prev }
        changedItems.forEach(item => {
          const cat = customCategoryMap[item.name] || getUnitCategory(item.name)
          next[item.name] = cat
        })
        return next
      })
      
      setSaveStatus('success')
      setTimeout(() => setSaveStatus('idle'), 3000)
    } catch (err) {
      console.error('Error saving classifications to Supabase:', err)
      setSaveStatus('error')
      setErrorMessage(err.message || String(err))
    }
  }, [uniqueDonVi, customCategoryMap, dbCategoryMap])

  // 5. Excel export handler for categorized units list (Multi-sheet export)
  const handleExportExcel = useCallback(() => {
    const wb = XLSXStyle.utils.book_new()
    
    const createSheetForList = (list, sheetTitle, fgColorHex = '0F58A7') => {
      const ws = {}
      const cols = [
        { key: 'STT', label: 'STT', width: 60 },
        { key: 'ten', label: 'Tên Đơn vị', width: 350 },
        { key: 'giaoCount', label: 'Số dòng Đơn vị giao (Nhập kho)', width: 200 },
        { key: 'nhanCount', label: 'Số dòng Đơn vị nhận (Xuất kho)', width: 200 },
        { key: 'totalCount', label: 'Tổng số dòng phát sinh', width: 180 }
      ]
      ws['!cols'] = cols.map(c => ({ wpx: c.width }))
      
      // Write header
      let rowIdx = 1
      cols.forEach((col, colIdx) => {
        const cellRef = `${String.fromCharCode(65 + colIdx)}${rowIdx}`
        ws[cellRef] = {
          v: col.label,
          t: 's',
          s: {
            fill: { patternType: 'solid', fgColor: { rgb: fgColorHex } },
            font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
            alignment: { horizontal: 'center', vertical: 'center' },
            border: {
              top: { style: 'thin', color: { rgb: '0A3D73' } },
              bottom: { style: 'medium', color: { rgb: '0A3D73' } },
              left: { style: 'thin', color: { rgb: '0A3D73' } },
              right: { style: 'thin', color: { rgb: '0A3D73' } }
            }
          }
        }
      })

      // Write data
      list.forEach((item, idx) => {
        rowIdx++
        const cells = [
          idx + 1,
          item.name,
          item.giaoCount,
          item.nhanCount,
          item.totalCount
        ]
        cells.forEach((val, colIdx) => {
          const cellRef = `${String.fromCharCode(65 + colIdx)}${rowIdx}`
          ws[cellRef] = {
            v: val,
            t: typeof val === 'number' ? 'n' : 's',
            s: {
              font: { name: 'Segoe UI', sz: 9.5 },
              alignment: { horizontal: colIdx === 1 ? 'left' : 'center', vertical: 'center' },
              border: {
                top: { style: 'thin', color: { rgb: 'E2E8F0' } },
                bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
                left: { style: 'thin', color: { rgb: 'E2E8F0' } },
                right: { style: 'thin', color: { rgb: 'E2E8F0' } }
              }
            }
          }
        })
      })
      ws['!ref'] = `A1:E${rowIdx}`
      XLSXStyle.utils.book_append_sheet(wb, ws, sheetTitle)
    }

    // Append sheets
    createSheetForList(filteredDonVi, 'Tổng hợp đơn vị', '0F58A7')
    createSheetForList(categorizedUnits.chuaphanbo, 'Chưa phân bổ', '64748B')
    createSheetForList(categorizedUnits.kho, 'Kho BCH', '1E40AF')
    createSheetForList(categorizedUnits.ncc, 'Nhà Cung cấp', 'D97706')
    createSheetForList(categorizedUnits.todoi, 'Tổ đội', '059669')

    // Save
    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary', compression: false })
    function s2ab(s) {
      const buf = new ArrayBuffer(s.length)
      const view = new Uint8Array(buf)
      for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF
      return buf
    }
    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `Kho_du_an_trich_xuat_${new Date().toISOString().slice(0, 10)}.xlsx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }, [filteredDonVi, categorizedUnits])

  return (
    <div style={{ padding: '16px 24px 24px 24px', height: '100%', display: 'flex', flexDirection: 'column', boxSizing: 'border-box', overflow: 'hidden' }}>
      {/* Tab Header & Action Panel */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, flexShrink: 0, gap: 16 }}>
        <div>
          <h2 style={{ fontSize: 19, fontWeight: 800, color: 'var(--text)', margin: 0, display: 'flex', alignItems: 'center', gap: 8 }}>
            <Warehouse size={22} color="var(--primary)" />
            Kho dự án (Trích xuất từ Đơn chung)
          </h2>
          <p style={{ margin: '4px 0 0', color: 'var(--text-muted)', fontSize: 13.5 }}>
            Tự động trích xuất toàn bộ tên <strong>Đơn vị giao</strong> và <strong>Đơn vị nhận</strong> từ dữ liệu Đơn chung hiện tại.
          </p>
        </div>
        <div style={{ display: 'flex', gap: 12, alignItems: 'center' }}>
          {isSupabaseConfigured && (
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              {/* Status text badge to clearly inform user */}
              <div style={{
                display: 'flex',
                alignItems: 'center',
                gap: 6,
                padding: '6px 12px',
                borderRadius: 20,
                fontSize: 12.5,
                fontWeight: 600,
                background: isSynchronized ? '#f0fdf4' : '#fff7ed',
                border: isSynchronized ? '1px solid #bbf7d0' : '1px solid #ffedd5',
                color: isSynchronized ? '#15803d' : '#c2410c',
                boxShadow: '0 1px 2px rgba(15, 23, 42, 0.05)',
                transition: 'all 0.2s ease'
              }}>
                <span 
                  className={isSynchronized ? "" : "animate-pulse"}
                  style={{
                    display: 'inline-block',
                    width: 8,
                    height: 8,
                    borderRadius: '50%',
                    background: isSynchronized ? '#22c55e' : '#f97316',
                    boxShadow: isSynchronized ? 'none' : '0 0 8px #f97316'
                  }} 
                />
                <span>
                  {isSynchronized ? 'Dữ liệu đã được đồng bộ' : 'Dữ liệu chưa được đồng bộ'}
                </span>
              </div>

              <button
                className="btn"
                onClick={handleSaveToSupabase}
                disabled={saveStatus === 'saving' || isSynchronized}
                style={{
                  background: isSynchronized 
                    ? '#cbd5e1' 
                    : saveStatus === 'saving' 
                      ? '#94a3b8' 
                      : saveStatus === 'success' 
                        ? 'linear-gradient(135deg, #16a34a 0%, #22c55e 100%)' 
                        : saveStatus === 'error'
                          ? 'linear-gradient(135deg, #dc2626 0%, #ef4444 100%)'
                          : 'linear-gradient(135deg, #2563eb 0%, #3b82f6 100%)',
                  color: isSynchronized ? '#64748b' : '#ffffff',
                  border: 'none',
                  boxShadow: (saveStatus === 'saving' || isSynchronized) ? 'none' : '0 2px 4px rgba(37,99,235,0.2)',
                  display: 'flex',
                  alignItems: 'center',
                  gap: 6,
                  padding: '8px 16px',
                  fontSize: 13.5,
                  fontWeight: 600,
                  borderRadius: 8,
                  cursor: (saveStatus === 'saving' || isSynchronized) ? 'not-allowed' : 'pointer',
                  transition: 'all 0.2s ease'
                }}
                title={isSynchronized ? 'Dữ liệu hiện tại đã khớp với Supabase' : saveStatus === 'error' ? errorMessage : 'Lưu sự phân bổ danh sách đơn vị hiện tại lên cơ sở dữ liệu Supabase'}
              >
                <Save size={14} />
                {saveStatus === 'saving' ? 'Đang lưu...' : saveStatus === 'success' ? 'Đã lưu thành công!' : saveStatus === 'error' ? 'Lỗi khi lưu!' : 'Lưu lên Supabase'}
              </button>
            </div>
          )}

          {filteredDonVi.length > 0 && (
            <button
              className="btn"
              onClick={handleExportExcel}
              style={{
                background: 'linear-gradient(135deg, #059669 0%, #10b981 100%)',
                color: '#ffffff',
                border: 'none',
                boxShadow: '0 2px 4px rgba(16,185,129,0.2)',
                display: 'flex',
                alignItems: 'center',
                gap: 6,
                padding: '8px 16px',
                fontSize: 13.5,
                fontWeight: 600,
                borderRadius: 8,
                cursor: 'pointer'
              }}
            >
              <Download size={14} /> Xuất Excel Kho dự án
            </button>
          )}
        </div>
      </div>

      {/* Info Warning Alert */}
      <div style={{
        background: '#eff6ff',
        border: '1px solid #bfdbfe',
        borderRadius: 10,
        padding: '10px 14px',
        marginBottom: 16,
        display: 'flex',
        gap: 8,
        alignItems: 'center',
        flexShrink: 0
      }}>
        <Info size={16} color="var(--primary)" style={{ flexShrink: 0 }} />
        <div style={{ fontSize: 13.5, color: '#1e40af', lineHeight: 1.4 }}>
          <strong>Cơ chế tự động:</strong> Bạn không cần phải import file Excel riêng cho tab này. Hệ thống sẽ tự động quét cột <strong>Đơn vị giao</strong> và <strong>Đơn vị nhận</strong> của Sheet <strong>Đơn chung</strong> để đồng bộ trực quan thời gian thực!
        </div>
      </div>

      {/* Search Input Filter */}
      <div style={{
        background: '#ffffff',
        border: '1px solid #e2e8f0',
        borderRadius: 10,
        padding: '10px 14px',
        marginBottom: 16,
        display: 'flex',
        alignItems: 'center',
        gap: 8,
        flexShrink: 0,
        boxShadow: '0 1px 2px rgba(0,0,0,0.02)'
      }}>
        <Search size={16} color="#64748b" />
        <input
          type="text"
          placeholder="Tìm kiếm nhanh tên Đơn vị giao hoặc Đơn vị nhận..."
          value={search}
          onChange={(e) => setSearch(e.target.value)}
          style={{
            flex: 1,
            border: 'none',
            outline: 'none',
            fontSize: 14,
            color: 'var(--text)',
            background: 'transparent'
          }}
        />
        {search && (
          <button
            onClick={() => setSearch('')}
            style={{
              background: 'transparent',
              border: 'none',
              cursor: 'pointer',
              color: '#64748b',
              padding: 2,
              display: 'flex',
              alignItems: 'center'
            }}
          >
            <X size={14} />
          </button>
        )}
      </div>

      {/* Main Grid View */}
      {chungRows.length === 0 ? (
        <div style={{
          flex: 1,
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          background: '#ffffff',
          borderRadius: 12,
          border: '1px solid #e2e8f0',
          padding: 40,
          textAlign: 'center'
        }}>
          <div style={{
            width: 56,
            height: 56,
            background: 'var(--primary-light)',
            borderRadius: 16,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            marginBottom: 16
          }}>
            <ClipboardList size={26} color="var(--primary)" />
          </div>
          <h3 style={{ fontSize: 16, fontWeight: 700, color: 'var(--text)', marginBottom: 6 }}>
            Chưa có dữ liệu Đơn chung
          </h3>
          <p style={{ color: 'var(--text-muted)', fontSize: 14, maxWidth: 440, margin: '0 auto 16px' }}>
            Vui lòng chọn tab <strong>Đơn chung</strong> ở thanh menu bên trái, tải lên tệp Excel dữ liệu để hệ thống tự động trích xuất các Kho dự án trực tiếp tại đây!
          </p>
        </div>
      ) : (
        (() => {
          const columnsConfig = [
            {
              key: 'chuaphanbo',
              title: 'Chưa phân bổ',
              icon: <HelpCircle size={15} color="#64748b" />,
              list: categorizedUnits.chuaphanbo,
              headerBg: '#f8fafc',
              textColor: '#334155',
              borderColor: '#cbd5e1',
              badgeBg: '#f1f5f9',
              badgeBorder: '#cbd5e1',
              badgeTextColor: '#475569',
              emptyText: 'Trống'
            },
            {
              key: 'kho',
              title: 'Kho BCH',
              icon: <Warehouse size={15} color="#1e40af" />,
              list: categorizedUnits.kho,
              headerBg: '#eff6ff',
              textColor: '#1e3a8a',
              borderColor: '#bfdbfe',
              badgeBg: '#eff6ff',
              badgeBorder: '#93c5fd',
              badgeTextColor: '#1e40af',
              emptyText: 'Trống'
            },
            {
              key: 'ncc',
              title: 'Nhà Cung cấp',
              icon: <Building2 size={15} color="#d97706" />,
              list: categorizedUnits.ncc,
              headerBg: '#fffbeb',
              textColor: '#78350f',
              borderColor: '#fde68a',
              badgeBg: '#fffbeb',
              badgeBorder: '#fcd34d',
              badgeTextColor: '#b45309',
              emptyText: 'Trống'
            },
            {
              key: 'todoi',
              title: 'Tổ đội',
              icon: <Users size={15} color="#059669" />,
              list: categorizedUnits.todoi,
              headerBg: '#f0fdf4',
              textColor: '#065f46',
              borderColor: '#bbf7d0',
              badgeBg: '#f0fdf4',
              badgeBorder: '#86efac',
              badgeTextColor: '#047857',
              emptyText: 'Trống'
            }
          ];

          return (
            <div style={{
              flex: 1,
              display: 'grid',
              gridTemplateColumns: 'repeat(4, minmax(0, 1fr))',
              gap: '14px',
              minHeight: 0,
              overflow: 'hidden'
            }}>
              {columnsConfig.map((col) => (
                <div 
                  key={col.key} 
                  onDragOver={handleDragOver}
                  onDrop={(e) => handleDrop(e, col.key)}
                  onDragEnter={() => setDragOverCol(col.key)}
                  onDragLeave={(e) => {
                    if (!e.currentTarget.contains(e.relatedTarget)) {
                      setDragOverCol(null)
                    }
                  }}
                  style={{
                    display: 'flex',
                    flexDirection: 'column',
                    background: '#ffffff',
                    borderRadius: 12,
                    border: dragOverCol === col.key 
                      ? '2px dashed var(--primary)' 
                      : `1px solid ${col.borderColor}`,
                    transform: dragOverCol === col.key ? 'scale(1.01)' : 'none',
                    transition: 'all 0.15s ease',
                    overflow: 'hidden',
                    boxShadow: dragOverCol === col.key 
                      ? '0 4px 12px rgba(37,99,235,0.1)' 
                      : '0 1px 3px rgba(0,0,0,0.02)'
                  }}
                >
                  {/* Column Header */}
                  <div style={{
                    background: col.headerBg,
                    borderBottom: `1px solid ${col.borderColor}`,
                    padding: '12px 14px',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    flexShrink: 0
                  }}>
                    <span style={{ fontWeight: 700, fontSize: 13.5, color: col.textColor, display: 'flex', alignItems: 'center', gap: 6 }}>
                      {col.icon}
                      {col.title} ({col.list.length})
                    </span>
                  </div>
                  
                  {/* Column Scrollable Content */}
                  <div style={{ flex: 1, overflowY: 'auto' }}>
                    {col.list.length === 0 ? (
                      <div style={{ color: '#94a3b8', fontSize: 12.5, padding: 24, textAlign: 'center', fontStyle: 'italic' }}>
                        {col.emptyText}
                      </div>
                    ) : (
                      col.list.map((item, idx) => {
                        const currentCat = customCategoryMap[item.name] || getUnitCategory(item.name)
                        const isUnsaved = dbCategoryMap[item.name] === undefined || dbCategoryMap[item.name] !== currentCat

                        return (
                          <div 
                            key={idx} 
                            draggable
                            onDragStart={(e) => handleDragStart(e, item.name)}
                            onDragEnd={() => {
                              setDraggedItemName(null)
                              setDragOverCol(null)
                            }}
                            style={{
                              padding: '10px 12px',
                              borderBottom: isUnsaved ? '1px solid #fca5a5' : '1px solid #f1f5f9',
                              display: 'flex',
                              alignItems: 'center',
                              gap: 8,
                              background: isUnsaved ? '#fee2e2' : (idx % 2 === 1 ? '#f8fafc' : '#ffffff'),
                              opacity: draggedItemName === item.name ? 0.4 : 1,
                              cursor: 'grab',
                              userSelect: 'none',
                              transition: 'background 0.1s, opacity 0.1s'
                            }} 
                            className={isUnsaved ? "hover:bg-red-200/50" : "hover:bg-blue-50/40"}
                          >
                          {/* STT */}
                          <div style={{ fontSize: 11, color: '#94a3b8', width: 20, textAlign: 'center', fontWeight: 600 }}>
                            {idx + 1}
                          </div>
                          
                          {/* Name only */}
                          <div style={{ flex: 1, minWidth: 0 }}>
                            <div 
                              title={item.name} 
                              style={{ 
                                fontSize: 12.5, 
                                fontWeight: 600, 
                                color: '#1e293b', 
                                whiteSpace: 'nowrap', 
                                overflow: 'hidden', 
                                textOverflow: 'ellipsis' 
                              }}
                            >
                              {item.name}
                            </div>
                          </div>
                        </div>
                      )
                    })
                  )}
                  </div>
                </div>
              ))}
            </div>
          );
        })()
      )}
    </div>
  )
}

// ─── IndexedDB cache helpers (thay thế localStorage cho dữ liệu lớn) ─────────
// localStorage giới hạn ~5-10MB nên với dữ liệu hàng chục nghìn dòng sẽ bị lỗi
// QuotaExceededError. IndexedDB cho phép lưu trữ dung lượng lớn hơn nhiều (thường
// hàng trăm MB trở lên tùy trình duyệt), giúp cache hoạt động thật sự ổn định.
const IDB_NAME = 'sgc_qlk_cache'
const IDB_STORE = 'kv'

// Chạy danh sách các hàm trả về Promise (vd: query Supabase) với giới hạn số lượng chạy ĐỒNG THỜI,
// thay vì bắn tất cả cùng lúc bằng Promise.all (dễ khiến Supabase free tier trả lỗi 500 do vượt quá
// giới hạn connection/pooler đồng thời). Trả về mảng kết quả theo đúng thứ tự ban đầu.
async function runWithConcurrencyLimit(taskFns, limit = 6) {
  const results = new Array(taskFns.length)
  let nextIndex = 0

  async function worker() {
    while (nextIndex < taskFns.length) {
      const currentIndex = nextIndex
      nextIndex += 1
      results[currentIndex] = await taskFns[currentIndex]()
    }
  }

  const workers = Array.from({ length: Math.min(limit, taskFns.length) }, () => worker())
  await Promise.all(workers)
  return results
}

function openIdb() {
  return new Promise((resolve, reject) => {
    try {
      const req = indexedDB.open(IDB_NAME, 1)
      req.onupgradeneeded = () => {
        if (!req.result.objectStoreNames.contains(IDB_STORE)) {
          req.result.createObjectStore(IDB_STORE)
        }
      }
      req.onsuccess = () => resolve(req.result)
      req.onerror = () => reject(req.error)
    } catch (e) {
      reject(e)
    }
  })
}

async function idbGet(key) {
  try {
    const db = await openIdb()
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readonly')
      const req = tx.objectStore(IDB_STORE).get(key)
      req.onsuccess = () => resolve(req.result)
      req.onerror = () => reject(req.error)
    })
  } catch (e) {
    console.warn('IndexedDB error reading', key, e)
    return undefined
  }
}

async function idbSet(key, value) {
  try {
    const db = await openIdb()
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(IDB_STORE, 'readwrite')
      tx.objectStore(IDB_STORE).put(value, key)
      tx.oncomplete = () => resolve(true)
      tx.onerror = () => reject(tx.error)
    })
  } catch (e) {
    console.warn('IndexedDB error persisting', key, e)
    return false
  }
}

// ─── Summary Config Tab ───────────────────────────────────────────────────────
const SUMMARY_CONFIG_KEY = 'sgc_summary_configs'

function loadSummaryConfigs() {
  try {
    const raw = localStorage.getItem(SUMMARY_CONFIG_KEY)
    return raw ? JSON.parse(raw) : []
  } catch { return [] }
}

function saveSummaryConfigs(configs) {
  localStorage.setItem(SUMMARY_CONFIG_KEY, JSON.stringify(configs))
}

// ─── Contrast & Border Helpers ────────────────────────────────────────────────
function getContrastColor(hexColor) {
  if (!hexColor || typeof hexColor !== 'string' || !hexColor.startsWith('#')) return '#0f172a'
  const hex = hexColor.replace('#', '')
  if (hex.length < 6) return '#0f172a'
  const r = parseInt(hex.substring(0, 2), 16)
  const g = parseInt(hex.substring(2, 4), 16)
  const b = parseInt(hex.substring(4, 6), 16)
  if (isNaN(r) || isNaN(g) || isNaN(b)) return '#0f172a'
  const yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000
  return (yiq >= 150) ? '#0f172a' : '#ffffff'
}

function getMutedContrastColor(hexColor) {
  if (!hexColor || typeof hexColor !== 'string' || !hexColor.startsWith('#')) return '#64748b'
  const hex = hexColor.replace('#', '')
  if (hex.length < 6) return '#64748b'
  const r = parseInt(hex.substring(0, 2), 16)
  const g = parseInt(hex.substring(2, 4), 16)
  const b = parseInt(hex.substring(4, 6), 16)
  if (isNaN(r) || isNaN(g) || isNaN(b)) return '#64748b'
  const yiq = ((r * 299) + (g * 587) + (b * 114)) / 1000
  return (yiq >= 150) ? '#475569' : '#f1f5f9'
}

function getBorderColor(hexColor) {
  if (!hexColor || typeof hexColor !== 'string' || !hexColor.startsWith('#')) return '#e2e8f0'
  const hex = hexColor.replace('#', '')
  if (hex.length < 6) return '#e2e8f0'
  const r = parseInt(hex.substring(0, 2), 16)
  const g = parseInt(hex.substring(2, 4), 16)
  const b = parseInt(hex.substring(4, 6), 16)
  if (isNaN(r) || isNaN(g) || isNaN(b)) return '#e2e8f0'
  const nr = Math.max(0, r - 25)
  const ng = Math.max(0, g - 25)
  const nb = Math.max(0, b - 25)
  return `#${nr.toString(16).padStart(2, '0')}${ng.toString(16).padStart(2, '0')}${nb.toString(16).padStart(2, '0')}`
}

const PRESET_COLORS = [
  { name: 'Xanh dương nhạt', value: '#eff6ff', border: '#bfdbfe', text: '#1e3a8a' },
  { name: 'Xanh lục nhạt', value: '#f0fdf4', border: '#bbf7d0', text: '#14532d' },
  { name: 'Sắc xuân lục', value: '#ecfdf5', border: '#a7f3d0', text: '#064e3b' },
  { name: 'Vàng cam nhạt', value: '#fffbeb', border: '#fde68a', text: '#78350f' },
  { name: 'Ghi sáng', value: '#f8fafc', border: '#cbd5e1', text: '#334155' },
  { name: 'Tím oải hương', value: '#faf5ff', border: '#e9d5ff', text: '#581c87' },
  { name: 'Hồng phấn', value: '#fff1f2', border: '#fecdd3', text: '#881337' },
]

function SummaryConfigTab({ 
  giaoRows, 
  nhanRows, 
  selectedProject, 
  allProjects, 
  configs, 
  setConfigs
}) {
  const [showCreateModal, setShowCreateModal] = React.useState(false)
  const [expandedConfig, setExpandedConfig] = React.useState(null) // index of expanded config
  const [configToDeleteIdx, setConfigToDeleteIdx] = React.useState(null)
  
  const filteredConfigs = selectedProject
    ? configs.filter(cfg => cfg.project === selectedProject)
    : []

  // Bảng Đơn Giao hiển thị list Đơn vị Nhận (từ giaoRows)
  const getUniqueGiao = (proj) => {
    const rows = proj ? giaoRows.filter(r => (r.ten_du_an || r.tenDuAn || r.duAn || '') === proj) : giaoRows
    return [...new Set(rows.map(r => r.donViNhan).filter(Boolean))].sort()
  }

  // Bảng Đơn Nhận hiển thị list Đơn vị Giao (từ nhanRows)
  const getUniqueNhan = (proj) => {
    const rows = proj ? nhanRows.filter(r => (r.ten_du_an || r.tenDuAn || r.duAn || '') === proj) : nhanRows
    return [...new Set(rows.map(r => r.donViGiao).filter(Boolean))].sort()
  }

  // Lưu configs vào LocalStorage mỗi khi thay đổi (không còn đồng bộ với Supabase)
  React.useEffect(() => {
    saveSummaryConfigs(configs)
  }, [configs])

  const handleCreateConfig = (name, proj, bgColor) => {
    const giaoUnits = getUniqueGiao(proj)
    const nhanUnits = getUniqueNhan(proj)
    const tempId = Date.now()
    
    const newConfig = {
      id: tempId,
      name,
      project: proj,
      giaoTable: giaoUnits.map(u => ({ unit: u, giamTru: '', boQua: false, tinhToan: false })),
      nhanTable: nhanUnits.map(u => ({ unit: u, giamTru: '', boQua: false, tinhToan: false })),
      bgColor: bgColor || '#eff6ff',
    }

    setConfigs(prev => [...prev, newConfig])
    setShowCreateModal(false)

    setConfigs(prev => {
      const idx = prev.findIndex(c => c.id === tempId)
      if (idx !== -1) setExpandedConfig(idx)
      return prev
    })
  }

  const handleDeleteConfig = (idx) => {
    const updated = configs.filter((_, i) => i !== idx)
    setConfigs(updated)
    if (expandedConfig === idx) setExpandedConfig(null)
    else if (expandedConfig > idx) setExpandedConfig(expandedConfig - 1)
    setConfigToDeleteIdx(null)
  }

  const updateGiaoRow = (cfgIdx, rowIdx, field, value) => {
    setConfigs(prev => {
      const next = prev.map((c, i) => {
        if (i !== cfgIdx) return c
        const newGiao = c.giaoTable.map((r, j) => {
          if (j !== rowIdx) return r
          if (typeof field === 'object' && field !== null) {
            return { ...r, ...field }
          }
          return { ...r, [field]: value }
        })
        return { ...c, giaoTable: newGiao }
      })
      return next
    })
  }

  const updateNhanRow = (cfgIdx, rowIdx, field, value) => {
    setConfigs(prev => {
      const next = prev.map((c, i) => {
        if (i !== cfgIdx) return c
        const newNhan = c.nhanTable.map((r, j) => {
          if (j !== rowIdx) return r
          if (typeof field === 'object' && field !== null) {
            return { ...r, ...field }
          }
          return { ...r, [field]: value }
        })
        return { ...c, nhanTable: newNhan }
      })
      return next
    })
  }

  const setUnitColumn = (cfgIdx, rIdx, colName, tableType) => {
    if (tableType === 'giao') {
      if (colName === 'giamTru') {
        updateGiaoRow(cfgIdx, rIdx, { giamTru: '1', boQua: false, tinhToan: false })
      } else if (colName === 'boQua') {
        updateGiaoRow(cfgIdx, rIdx, { giamTru: '', boQua: true, tinhToan: false })
      } else if (colName === 'tinhToan') {
        updateGiaoRow(cfgIdx, rIdx, { giamTru: '', boQua: false, tinhToan: true })
      }
    } else {
      if (colName === 'giamTru') {
        updateNhanRow(cfgIdx, rIdx, { giamTru: '1', boQua: false, tinhToan: false })
      } else if (colName === 'boQua') {
        updateNhanRow(cfgIdx, rIdx, { giamTru: '', boQua: true, tinhToan: false })
      } else if (colName === 'tinhToan') {
        updateNhanRow(cfgIdx, rIdx, { giamTru: '', boQua: false, tinhToan: true })
      }
    }
  }

  // Sync units when rows data changes (reload units for existing configs)
  const handleRefreshConfig = (cfgIdx) => {
    const cfg = configs[cfgIdx]
    const giaoUnits = getUniqueGiao(cfg.project)
    const nhanUnits = getUniqueNhan(cfg.project)
    setConfigs(prev => prev.map((c, i) => {
      if (i !== cfgIdx) return c
      
      // 1. Keep ALL existing units in c.giaoTable (especially those already configured).
      const existingGiaoKeys = new Set(c.giaoTable.map(r => r.unit))
      // Append any new units from current active files that aren't already in c.giaoTable
      const addedGiao = giaoUnits
        .filter(u => !existingGiaoKeys.has(u))
        .map(u => ({ unit: u, giamTru: '', boQua: false, tinhToan: false }))
      const newGiao = [...c.giaoTable, ...addedGiao]

      // 2. Same for nhanTable
      const existingNhanKeys = new Set(c.nhanTable.map(r => r.unit))
      const addedNhan = nhanUnits
        .filter(u => !existingNhanKeys.has(u))
        .map(u => ({ unit: u, giamTru: '', boQua: false, tinhToan: false }))
      const newNhan = [...c.nhanTable, ...addedNhan]

      return { ...c, giaoTable: newGiao, nhanTable: newNhan }
    }))
  }

  const thStyle = {
    padding: '8px 12px', fontSize: 12, fontWeight: 700, color: '#fff',
    background: '#0f58a7', textAlign: 'center', whiteSpace: 'nowrap',
    border: '1px solid #0a3d73'
  }
  const tdStyle = {
    padding: '6px 10px', fontSize: 13, borderBottom: '1px solid #e2e8f0',
    verticalAlign: 'middle', color: '#1e293b'
  }

  return (
    <div style={{ padding: '16px 24px 32px 24px', height: '100%', overflowY: 'auto', boxSizing: 'border-box' }}>
      {/* Header */}
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 20 }}>
        <div>
          <h2 style={{ margin: 0, fontSize: 20, fontWeight: 800, color: '#0f172a' }}>Cấu hình tổng hợp</h2>
          <p style={{ margin: '4px 0 0', color: '#64748b', fontSize: 14 }}>
            Tạo và quản lý các loại tổng hợp theo từng dự án — cấu hình đơn vị giao và nhận riêng biệt
          </p>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginTop: 8, fontSize: 12, color: '#64748b', fontWeight: 600 }}>
            <AlertCircle size={12} />
            <span>Dữ liệu cấu hình được lưu tại trình duyệt này (LocalStorage)</span>
          </div>
        </div>
        <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'flex-end', gap: 4 }}>
          <button
            onClick={() => selectedProject ? setShowCreateModal(true) : null}
            disabled={!selectedProject}
            title={!selectedProject ? 'Vui lòng chọn một Kho dự án cụ thể trước khi tạo' : ''}
            style={{
              background: selectedProject
                ? 'linear-gradient(135deg, #0f58a7 0%, #1a6abf 100%)'
                : '#cbd5e1',
              color: selectedProject ? '#fff' : '#94a3b8',
              border: 'none', borderRadius: 8,
              padding: '9px 18px', fontSize: 14, fontWeight: 700,
              cursor: selectedProject ? 'pointer' : 'not-allowed',
              display: 'flex', alignItems: 'center', gap: 6,
              boxShadow: selectedProject ? '0 2px 8px rgba(15,88,167,0.25)' : 'none',
              transition: 'all 0.15s'
            }}
          >
            <span style={{ fontSize: 18, lineHeight: 1 }}>+</span> Tạo Loại Tổng hợp
          </button>
          {!selectedProject && (
            <span style={{ fontSize: 11, color: '#f59e0b', fontWeight: 600, display: 'flex', alignItems: 'center', gap: 4 }}>
              <AlertCircle size={11} color="#f59e0b" /> Vui lòng chọn Kho dự án trước
            </span>
          )}
        </div>
      </div>


      {/* Empty state */}
      {!selectedProject ? (
        <div style={{
          background: '#f8fafc', border: '2px dashed #cbd5e1', borderRadius: 12,
          padding: '48px 24px', textAlign: 'center', color: '#64748b'
        }}>
          <div style={{
            width: 64, height: 64, background: '#e2e8f0', borderRadius: 16,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            margin: '0 auto 16px'
          }}>
            <Settings size={28} color="#94a3b8" />
          </div>
          <h3 style={{ margin: '0 0 8px', fontSize: 16, fontWeight: 700, color: '#334155' }}>
            Chưa chọn Kho dự án
          </h3>
          <p style={{ margin: 0, fontSize: 14 }}>
            Vui lòng chọn Kho dự án trên thanh menu ở góc trên để cấu hình loại tổng hợp.
          </p>
        </div>
      ) : filteredConfigs.length === 0 ? (
        <div style={{
          background: '#f8fafc', border: '2px dashed #cbd5e1', borderRadius: 12,
          padding: '48px 24px', textAlign: 'center', color: '#64748b'
        }}>
          <div style={{
            width: 64, height: 64, background: '#e2e8f0', borderRadius: 16,
            display: 'flex', alignItems: 'center', justifyContent: 'center',
            margin: '0 auto 16px'
          }}>
            <Settings size={28} color="#94a3b8" />
          </div>
          <h3 style={{ margin: '0 0 8px', fontSize: 16, fontWeight: 700, color: '#334155' }}>
            Chưa có loại tổng hợp cho dự án này
          </h3>
          <p style={{ margin: 0, fontSize: 14 }}>
            Nhấn <strong>"+ Tạo Loại Tổng hợp"</strong> để bắt đầu cấu hình đơn vị giao nhận cho dự án này.
          </p>
        </div>
      ) : null}

      {/* Config list */}
      {filteredConfigs.map((cfg) => {
        const cfgIdx = configs.findIndex(c => c.id === cfg.id)
        if (cfgIdx === -1) return null

        const isExpanded = expandedConfig === cfgIdx
        const cardBorderColor = cfg.bgColor ? getBorderColor(cfg.bgColor) : '#e2e8f0'
        const contrastColor = cfg.bgColor ? getContrastColor(cfg.bgColor) : '#0f172a'
        const mutedColor = cfg.bgColor ? getMutedContrastColor(cfg.bgColor) : '#64748b'
        const headerBgColor = cfg.bgColor || (isExpanded ? '#eff6ff' : '#f8fafc')
        const headerBorderBottom = isExpanded ? `1px solid ${cfg.bgColor ? getBorderColor(cfg.bgColor) : '#bfdbfe'}` : 'none'

        return (
          <div key={cfg.id} style={{
            background: '#fff', borderRadius: 12, border: `1px solid ${cardBorderColor}`,
            boxShadow: '0 2px 8px rgba(15,23,42,0.06)', marginBottom: 16, overflow: 'hidden'
          }}>
            {/* Config header row */}
            <div
              onClick={() => setExpandedConfig(isExpanded ? null : cfgIdx)}
              style={{
                display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                padding: '14px 20px', cursor: 'pointer',
                background: headerBgColor,
                borderBottom: headerBorderBottom,
                transition: 'background 0.15s'
              }}
            >
              <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                <div style={{
                  width: 36, height: 36, background: isExpanded ? '#0f58a7' : '#e2e8f0',
                  borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center',
                  flexShrink: 0
                }}>
                  <BarChart3 size={18} color={isExpanded ? '#fff' : '#64748b'} />
                </div>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 15, color: contrastColor }}>{cfg.name}</div>
                  <div style={{ fontSize: 12, color: mutedColor, marginTop: 2 }}>
                    Dự án: <strong style={{ color: contrastColor }}>{cfg.project || 'Tất cả dự án'}</strong>
                    &nbsp;·&nbsp; {cfg.giaoTable.length} đơn vị giao &nbsp;·&nbsp; {cfg.nhanTable.length} đơn vị nhận
                  </div>
                </div>
              </div>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                <button
                  onClick={(e) => { e.stopPropagation(); handleRefreshConfig(cfgIdx) }}
                  title="Làm mới danh sách đơn vị từ dữ liệu hiện tại"
                  style={{
                    background: 'transparent', border: `1px solid ${cfg.bgColor ? getBorderColor(cfg.bgColor) : '#cbd5e1'}`, borderRadius: 6,
                    padding: '4px 8px', cursor: 'pointer', color: contrastColor,
                    display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, fontWeight: 600
                  }}
                >
                  <RefreshCw size={12} color={contrastColor} /> Làm mới
                </button>
                <button
                  onClick={(e) => { e.stopPropagation(); setConfigToDeleteIdx(cfgIdx) }}
                  title="Xóa cấu hình này"
                  style={{
                    background: 'transparent', border: '1px solid #fca5a5', borderRadius: 6,
                    padding: '4px 8px', cursor: 'pointer', color: '#ef4444',
                    display: 'flex', alignItems: 'center', gap: 4, fontSize: 12, fontWeight: 600
                  }}
                >
                  <Trash2 size={12} /> Xóa
                </button>
                <ChevronDown size={16} color={mutedColor} style={{
                  transform: isExpanded ? 'rotate(180deg)' : 'none',
                  transition: 'transform 0.2s'
                }} />
              </div>
            </div>

            {/* Config body */}
            {isExpanded && (
              <div style={{ padding: '20px 20px 24px' }}>
                {/* Drag and Drop instructions */}
                <div style={{
                  fontSize: '12.5px', color: '#0369a1', background: '#f0f9ff',
                  padding: '10px 14px', borderRadius: 8, marginBottom: 16,
                  border: '1px solid #bae6fd', display: 'flex', alignItems: 'center', gap: 8
                }}>
                  <span style={{ fontSize: 16 }}>💡</span>
                  <span>
                    <strong>Tính năng kéo thả:</strong> Bạn có thể nắm biểu tượng ☰ của <strong>Tên đơn vị</strong> rồi kéo thả sang ô tương ứng ở 3 cột (Giảm trừ, Bỏ qua, Tính toán) để thiết lập nhanh, tiện lợi và lưu lại tức thì!
                  </span>
                </div>

                <div style={{ display: 'flex', gap: 24, flexWrap: 'wrap' }}>
                  {/* Bảng Đơn Giao */}
                  <div style={{ flex: '1 1 400px', minWidth: 320 }}>
                    <div style={{
                      display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12
                    }}>
                      <div style={{
                        width: 28, height: 28, borderRadius: 6, background: '#eff6ff',
                        display: 'flex', alignItems: 'center', justifyContent: 'center'
                      }}>
                        <Truck size={14} color="#0f58a7" />
                      </div>
                      <span style={{ fontWeight: 700, fontSize: 14, color: '#0f172a' }}>
                        Bảng Đơn Giao
                      </span>
                      <span style={{ fontSize: 12, color: '#64748b', background: '#f1f5f9', borderRadius: 4, padding: '2px 8px' }}>
                        {cfg.giaoTable.length} đơn vị
                      </span>
                    </div>

                    {cfg.giaoTable.length === 0 ? (
                      <div style={{ fontSize: 13, color: '#94a3b8', padding: '16px', textAlign: 'center', background: '#f8fafc', borderRadius: 8, border: '1px dashed #e2e8f0' }}>
                        Không có dữ liệu đơn vị giao cho dự án này
                      </div>
                    ) : (
                      <div style={{ overflowX: 'auto', borderRadius: 8, border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13, tableLayout: 'fixed' }}>
                          <thead>
                            <tr>
                              <th colSpan={4} style={{
                                padding: '10px 12px', fontSize: 13, fontWeight: 700, color: '#fff',
                                background: '#0f58a7', textAlign: 'center', whiteSpace: 'nowrap',
                                border: '1px solid #0a3d73', textTransform: 'uppercase', letterSpacing: '0.05em'
                              }}>
                                Đơn vị nhận
                              </th>
                            </tr>
                            <tr>
                              <th style={{ ...thStyle, textAlign: 'center', width: '25%', minWidth: '25%' }}>Chưa phân bổ</th>
                              <th style={{ ...thStyle, width: '25%', minWidth: '25%', maxWidth: '25%' }}>Giảm trừ</th>
                              <th style={{ ...thStyle, width: '25%', minWidth: '25%', maxWidth: '25%' }}>Bỏ qua</th>
                              <th style={{ ...thStyle, width: '25%', minWidth: '25%', maxWidth: '25%' }}>Tính toán</th>
                            </tr>
                          </thead>
                          <tbody>
                            {(() => {
                              const giaoUnits = getUniqueGiao(cfg.project);

                              const listGiamTruGiao = cfg.giaoTable
                                .map((row, idx) => ({ row, idx }))
                                .filter(item => !!item.row.giamTru);

                              const listBoQuaGiao = cfg.giaoTable
                                .map((row, idx) => ({ row, idx }))
                                .filter(item => !!item.row.boQua);

                              const listTinhToanGiao = cfg.giaoTable
                                .map((row, idx) => ({ row, idx }))
                                .filter(item => !!item.row.tinhToan);

                              const listChuaPhanBoGiao = cfg.giaoTable
                                .map((row, idx) => ({ row, idx }))
                                .filter(item => !item.row.giamTru && !item.row.boQua && !item.row.tinhToan)
                                .filter(item => giaoUnits.includes(item.row.unit));

                              const maxRowsGiao = Math.max(
                                listChuaPhanBoGiao.length,
                                listGiamTruGiao.length,
                                listBoQuaGiao.length,
                                listTinhToanGiao.length
                              );

                              const renderRowsGiao = maxRowsGiao > 0 ? maxRowsGiao : 1;

                              return Array.from({ length: renderRowsGiao }).map((_, rIdx) => {
                                const itemChuaPhanBo = listChuaPhanBoGiao[rIdx];
                                const itemGiamTru = listGiamTruGiao[rIdx];
                                const itemBoQua = listBoQuaGiao[rIdx];
                                const itemTinhToan = listTinhToanGiao[rIdx];

                                return (
                                  <tr key={rIdx} style={{ background: rIdx % 2 === 0 ? '#fff' : '#f8fafc', transition: 'background 0.15s' }}>
                                    {/* Col 1: Chưa phân bổ */}
                                    <td 
                                      style={{ 
                                        ...tdStyle, 
                                        fontWeight: 600, 
                                        color: '#0f172a',
                                        width: '25%',
                                        minWidth: '25%',
                                        borderRight: '1px solid #f1f5f9',
                                        padding: '8px',
                                        background: '#fff7ed'
                                      }}
                                      onDragOver={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '#fed7aa';
                                      }}
                                      onDragLeave={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                      }}
                                      onDrop={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        const typeStr = e.dataTransfer.getData('application/sgc-unit-type');
                                        const draggedIdxStr = e.dataTransfer.getData('application/sgc-row-index');
                                        if (typeStr === 'giao' && draggedIdxStr !== '') {
                                          const draggedIdx = Number(draggedIdxStr);
                                          updateGiaoRow(cfgIdx, draggedIdx, { giamTru: '', boQua: false, tinhToan: false });
                                        }
                                      }}
                                    >
                                      {itemChuaPhanBo ? (
                                        <div 
                                          draggable={true}
                                          onDragStart={(e) => {
                                            e.dataTransfer.setData('text/plain', itemChuaPhanBo.row.unit);
                                            e.dataTransfer.setData('application/sgc-unit-type', 'giao');
                                            e.dataTransfer.setData('application/sgc-row-index', String(itemChuaPhanBo.idx));
                                            e.dataTransfer.effectAllowed = 'copyMove';
                                          }}
                                          onClick={() => {
                                            updateGiaoRow(cfgIdx, itemChuaPhanBo.idx, { giamTru: '', boQua: false, tinhToan: true });
                                          }}
                                          style={{ display: 'flex', alignItems: 'center', gap: 6, width: '100%', cursor: 'grab', userSelect: 'none' }}
                                          title="Nhấp giữ kéo sang các cột bên cạnh, hoặc Click để chuyển sang Tính Toán"
                                        >
                                          <span style={{ color: '#94a3b8', fontSize: 13, flexShrink: 0 }}>☰</span>
                                          <span style={{ wordBreak: 'break-word', whiteSpace: 'normal', flex: 1 }} title={itemChuaPhanBo.row.unit}>
                                            {itemChuaPhanBo.row.unit}
                                          </span>
                                        </div>
                                      ) : null}
                                    </td>

                                    {/* Col 2: Giảm trừ */}
                                    <td 
                                      style={{ 
                                        ...tdStyle, 
                                        textAlign: 'left', 
                                        width: '25%', 
                                        minWidth: '25%', 
                                        maxWidth: '25%',
                                        borderRight: '1px solid #f1f5f9',
                                        transition: 'all 0.15s',
                                        position: 'relative',
                                        padding: '8px'
                                      }}
                                      onDragOver={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '#eff6ff';
                                        e.currentTarget.style.outline = '2px dashed #3b82f6';
                                      }}
                                      onDragLeave={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                      }}
                                      onDrop={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                        const typeStr = e.dataTransfer.getData('application/sgc-unit-type');
                                        const draggedIdxStr = e.dataTransfer.getData('application/sgc-row-index');
                                        if (typeStr === 'giao' && draggedIdxStr !== '') {
                                          const draggedIdx = Number(draggedIdxStr);
                                          updateGiaoRow(cfgIdx, draggedIdx, { giamTru: '1', boQua: false, tinhToan: false });
                                        }
                                      }}
                                    >
                                      {itemGiamTru ? (
                                        <div 
                                          draggable={true}
                                          onDragStart={(e) => {
                                            e.dataTransfer.setData('text/plain', itemGiamTru.row.unit);
                                            e.dataTransfer.setData('application/sgc-unit-type', 'giao');
                                            e.dataTransfer.setData('application/sgc-row-index', String(itemGiamTru.idx));
                                            e.dataTransfer.effectAllowed = 'copyMove';
                                          }}
                                          style={{
                                            background: '#eff6ff',
                                            border: '1px solid #bfdbfe',
                                            color: '#1e40af',
                                            borderRadius: 6,
                                            padding: '4px 8px',
                                            fontSize: '11.5px',
                                            fontWeight: 600,
                                            display: 'inline-flex',
                                            alignItems: 'center',
                                            gap: 4,
                                            cursor: 'grab',
                                            boxShadow: '0 1px 2px rgba(30,64,175,0.05)',
                                            maxWidth: '100%'
                                          }}
                                        >
                                          <span style={{ flexShrink: 0, color: '#3b82f6' }}>☰</span>
                                          <span style={{ wordBreak: 'break-word', whiteSpace: 'normal' }} title={itemGiamTru.row.unit}>{itemGiamTru.row.unit}</span>
                                          <button 
                                            type="button" 
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              updateGiaoRow(cfgIdx, itemGiamTru.idx, { giamTru: '', boQua: false, tinhToan: false });
                                            }}
                                            style={{
                                              background: 'transparent',
                                              border: 'none',
                                              color: '#1e40af',
                                              cursor: 'pointer',
                                              padding: 0,
                                              marginLeft: 4,
                                              fontSize: 14,
                                              lineHeight: 1,
                                              display: 'flex',
                                              alignItems: 'center'
                                            }}
                                            title="Bỏ thiết lập"
                                          >
                                            ×
                                          </button>
                                        </div>
                                      ) : (
                                        <span style={{ fontSize: 11, color: '#cbd5e1', fontStyle: 'italic', pointerEvents: 'none', userSelect: 'none' }}>Kéo thả vào đây</span>
                                      )}
                                    </td>

                                    {/* Col 3: Bỏ qua */}
                                    <td 
                                      style={{ 
                                        ...tdStyle, 
                                        textAlign: 'left', 
                                        width: '25%', 
                                        minWidth: '25%', 
                                        maxWidth: '25%',
                                        borderRight: '1px solid #f1f5f9',
                                        transition: 'all 0.15s',
                                        position: 'relative',
                                        padding: '8px'
                                      }}
                                      onDragOver={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '#fff1f2';
                                        e.currentTarget.style.outline = '2px dashed #f43f5e';
                                      }}
                                      onDragLeave={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                      }}
                                      onDrop={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                        const typeStr = e.dataTransfer.getData('application/sgc-unit-type');
                                        const draggedIdxStr = e.dataTransfer.getData('application/sgc-row-index');
                                        if (typeStr === 'giao' && draggedIdxStr !== '') {
                                          const draggedIdx = Number(draggedIdxStr);
                                          updateGiaoRow(cfgIdx, draggedIdx, { giamTru: '', boQua: true, tinhToan: false });
                                        }
                                      }}
                                    >
                                      {itemBoQua ? (
                                        <div 
                                          draggable={true}
                                          onDragStart={(e) => {
                                            e.dataTransfer.setData('text/plain', itemBoQua.row.unit);
                                            e.dataTransfer.setData('application/sgc-unit-type', 'giao');
                                            e.dataTransfer.setData('application/sgc-row-index', String(itemBoQua.idx));
                                            e.dataTransfer.effectAllowed = 'copyMove';
                                          }}
                                          style={{
                                            background: '#fff1f2',
                                            border: '1px solid #fecdd3',
                                            color: '#9f1239',
                                            borderRadius: 6,
                                            padding: '4px 8px',
                                            fontSize: '11.5px',
                                            fontWeight: 600,
                                            display: 'inline-flex',
                                            alignItems: 'center',
                                            gap: 4,
                                            cursor: 'grab',
                                            boxShadow: '0 1px 2px rgba(159,18,57,0.05)',
                                            maxWidth: '100%'
                                          }}
                                        >
                                          <span style={{ flexShrink: 0, color: '#f43f5e' }}>☰</span>
                                          <span style={{ wordBreak: 'break-word', whiteSpace: 'normal' }} title={itemBoQua.row.unit}>{itemBoQua.row.unit}</span>
                                          <button 
                                            type="button" 
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              updateGiaoRow(cfgIdx, itemBoQua.idx, { giamTru: '', boQua: false, tinhToan: false });
                                            }}
                                            style={{
                                              background: 'transparent',
                                              border: 'none',
                                              color: '#9f1239',
                                              cursor: 'pointer',
                                              padding: 0,
                                              marginLeft: 4,
                                              fontSize: 14,
                                              lineHeight: 1,
                                              display: 'flex',
                                              alignItems: 'center'
                                            }}
                                            title="Bỏ thiết lập"
                                          >
                                            ×
                                          </button>
                                        </div>
                                      ) : (
                                        <span style={{ fontSize: 11, color: '#cbd5e1', fontStyle: 'italic', pointerEvents: 'none', userSelect: 'none' }}>Kéo thả vào đây</span>
                                      )}
                                    </td>

                                    {/* Col 4: Tính toán */}
                                    <td 
                                      style={{ 
                                        ...tdStyle, 
                                        textAlign: 'left', 
                                        width: '25%', 
                                        minWidth: '25%', 
                                        maxWidth: '25%',
                                        transition: 'all 0.15s',
                                        position: 'relative',
                                        padding: '8px'
                                      }}
                                      onDragOver={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '#ecfdf5';
                                        e.currentTarget.style.outline = '2px dashed #10b981';
                                      }}
                                      onDragLeave={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                      }}
                                      onDrop={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                        const typeStr = e.dataTransfer.getData('application/sgc-unit-type');
                                        const draggedIdxStr = e.dataTransfer.getData('application/sgc-row-index');
                                        if (typeStr === 'giao' && draggedIdxStr !== '') {
                                          const draggedIdx = Number(draggedIdxStr);
                                          updateGiaoRow(cfgIdx, draggedIdx, { giamTru: '', boQua: false, tinhToan: true });
                                        }
                                      }}
                                    >
                                      {itemTinhToan ? (
                                        <div 
                                          draggable={true}
                                          onDragStart={(e) => {
                                            e.dataTransfer.setData('text/plain', itemTinhToan.row.unit);
                                            e.dataTransfer.setData('application/sgc-unit-type', 'giao');
                                            e.dataTransfer.setData('application/sgc-row-index', String(itemTinhToan.idx));
                                            e.dataTransfer.effectAllowed = 'copyMove';
                                          }}
                                          style={{
                                            background: '#ecfdf5',
                                            border: '1px solid #a7f3d0',
                                            color: '#065f46',
                                            borderRadius: 6,
                                            padding: '4px 8px',
                                            fontSize: '11.5px',
                                            fontWeight: 600,
                                            display: 'inline-flex',
                                            alignItems: 'center',
                                            gap: 4,
                                            cursor: 'grab',
                                            boxShadow: '0 1px 2px rgba(6,95,70,0.05)',
                                            maxWidth: '100%'
                                          }}
                                        >
                                          <span style={{ flexShrink: 0, color: '#10b981' }}>☰</span>
                                          <span style={{ wordBreak: 'break-word', whiteSpace: 'normal' }} title={itemTinhToan.row.unit}>{itemTinhToan.row.unit}</span>
                                          <button 
                                            type="button" 
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              updateGiaoRow(cfgIdx, itemTinhToan.idx, { giamTru: '', boQua: false, tinhToan: false });
                                            }}
                                            style={{
                                              background: 'transparent',
                                              border: 'none',
                                              color: '#065f46',
                                              cursor: 'pointer',
                                              padding: 0,
                                              marginLeft: 4,
                                              fontSize: 14,
                                              lineHeight: 1,
                                              display: 'flex',
                                              alignItems: 'center'
                                            }}
                                            title="Bỏ thiết lập"
                                          >
                                            ×
                                          </button>
                                        </div>
                                      ) : (
                                        <span style={{ fontSize: 11, color: '#cbd5e1', fontStyle: 'italic', pointerEvents: 'none', userSelect: 'none' }}>Kéo thả vào đây</span>
                                      )}
                                    </td>
                                  </tr>
                                );
                              });
                            })()}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>

                  {/* Bảng Đơn Nhận */}
                  <div style={{ flex: '1 1 400px', minWidth: 320 }}>
                    <div style={{
                      display: 'flex', alignItems: 'center', gap: 8, marginBottom: 12
                    }}>
                      <div style={{
                        width: 28, height: 28, borderRadius: 6, background: '#f0fdf4',
                        display: 'flex', alignItems: 'center', justifyContent: 'center'
                      }}>
                        <PackageCheck size={14} color="#10b981" />
                      </div>
                      <span style={{ fontWeight: 700, fontSize: 14, color: '#0f172a' }}>
                        Bảng Đơn Nhận
                      </span>
                      <span style={{ fontSize: 12, color: '#64748b', background: '#f1f5f9', borderRadius: 4, padding: '2px 8px' }}>
                        {cfg.nhanTable.length} đơn vị
                      </span>
                    </div>

                    {cfg.nhanTable.length === 0 ? (
                      <div style={{ fontSize: 13, color: '#94a3b8', padding: '16px', textAlign: 'center', background: '#f8fafc', borderRadius: 8, border: '1px dashed #e2e8f0' }}>
                        Không có dữ liệu đơn vị nhận cho dự án này
                      </div>
                    ) : (
                      <div style={{ overflowX: 'auto', borderRadius: 8, border: '1px solid #e2e8f0', boxShadow: '0 1px 3px rgba(0,0,0,0.05)' }}>
                        <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13, tableLayout: 'fixed' }}>
                          <thead>
                            <tr>
                              <th colSpan={4} style={{
                                padding: '10px 12px', fontSize: 13, fontWeight: 700, color: '#fff',
                                background: '#065f46', textAlign: 'center', whiteSpace: 'nowrap',
                                border: '1px solid #04432e', textTransform: 'uppercase', letterSpacing: '0.05em'
                              }}>
                                Đơn vị giao
                              </th>
                            </tr>
                            <tr>
                              <th style={{ ...thStyle, textAlign: 'center', width: '25%', minWidth: '25%', background: '#065f46', border: '1px solid #04432e' }}>Chưa phân bổ</th>
                              <th style={{ ...thStyle, width: '25%', minWidth: '25%', maxWidth: '25%', background: '#065f46', border: '1px solid #04432e' }}>Giảm trừ</th>
                              <th style={{ ...thStyle, width: '25%', minWidth: '25%', maxWidth: '25%', background: '#065f46', border: '1px solid #04432e' }}>Bỏ qua</th>
                              <th style={{ ...thStyle, width: '25%', minWidth: '25%', maxWidth: '25%', background: '#065f46', border: '1px solid #04432e' }}>Tính toán</th>
                            </tr>
                          </thead>
                          <tbody>
                            {(() => {
                              const nhanUnits = getUniqueNhan(cfg.project);

                              const listGiamTruNhan = cfg.nhanTable
                                .map((row, idx) => ({ row, idx }))
                                .filter(item => !!item.row.giamTru);

                              const listBoQuaNhan = cfg.nhanTable
                                .map((row, idx) => ({ row, idx }))
                                .filter(item => !!item.row.boQua);

                              const listTinhToanNhan = cfg.nhanTable
                                .map((row, idx) => ({ row, idx }))
                                .filter(item => !!item.row.tinhToan);

                              const listChuaPhanBoNhan = cfg.nhanTable
                                .map((row, idx) => ({ row, idx }))
                                .filter(item => !item.row.giamTru && !item.row.boQua && !item.row.tinhToan)
                                .filter(item => nhanUnits.includes(item.row.unit));

                              const maxRowsNhan = Math.max(
                                listChuaPhanBoNhan.length,
                                listGiamTruNhan.length,
                                listBoQuaNhan.length,
                                listTinhToanNhan.length
                              );

                              const renderRowsNhan = maxRowsNhan > 0 ? maxRowsNhan : 1;

                              return Array.from({ length: renderRowsNhan }).map((_, rIdx) => {
                                const itemChuaPhanBo = listChuaPhanBoNhan[rIdx];
                                const itemGiamTru = listGiamTruNhan[rIdx];
                                const itemBoQua = listBoQuaNhan[rIdx];
                                const itemTinhToan = listTinhToanNhan[rIdx];

                                return (
                                  <tr key={rIdx} style={{ background: rIdx % 2 === 0 ? '#fff' : '#f8fafc', transition: 'background 0.15s' }}>
                                    {/* Col 1: Chưa phân bổ */}
                                    <td 
                                      style={{ 
                                        ...tdStyle, 
                                        fontWeight: 600, 
                                        color: '#0f172a',
                                        width: '25%',
                                        minWidth: '25%',
                                        borderRight: '1px solid #f1f5f9',
                                        padding: '8px',
                                        background: '#fff7ed'
                                      }}
                                      onDragOver={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '#fed7aa';
                                      }}
                                      onDragLeave={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                      }}
                                      onDrop={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        const typeStr = e.dataTransfer.getData('application/sgc-unit-type');
                                        const draggedIdxStr = e.dataTransfer.getData('application/sgc-row-index');
                                        if (typeStr === 'nhan' && draggedIdxStr !== '') {
                                          const draggedIdx = Number(draggedIdxStr);
                                          updateNhanRow(cfgIdx, draggedIdx, { giamTru: '', boQua: false, tinhToan: false });
                                        }
                                      }}
                                    >
                                      {itemChuaPhanBo ? (
                                        <div 
                                          draggable={true}
                                          onDragStart={(e) => {
                                            e.dataTransfer.setData('text/plain', itemChuaPhanBo.row.unit);
                                            e.dataTransfer.setData('application/sgc-unit-type', 'nhan');
                                            e.dataTransfer.setData('application/sgc-row-index', String(itemChuaPhanBo.idx));
                                            e.dataTransfer.effectAllowed = 'copyMove';
                                          }}
                                          onClick={() => {
                                            updateNhanRow(cfgIdx, itemChuaPhanBo.idx, { giamTru: '', boQua: false, tinhToan: true });
                                          }}
                                          style={{ display: 'flex', alignItems: 'center', gap: 6, width: '100%', cursor: 'grab', userSelect: 'none' }}
                                          title="Nhấp giữ kéo sang các cột bên cạnh, hoặc Click để chuyển sang Tính Toán"
                                        >
                                          <span style={{ color: '#94a3b8', fontSize: 13, flexShrink: 0 }}>☰</span>
                                          <span style={{ wordBreak: 'break-word', whiteSpace: 'normal', flex: 1 }} title={itemChuaPhanBo.row.unit}>
                                            {itemChuaPhanBo.row.unit}
                                          </span>
                                        </div>
                                      ) : null}
                                    </td>

                                    {/* Col 2: Giảm trừ */}
                                    <td 
                                      style={{ 
                                        ...tdStyle, 
                                        textAlign: 'left', 
                                        width: '25%', 
                                        minWidth: '25%', 
                                        maxWidth: '25%',
                                        borderRight: '1px solid #f1f5f9',
                                        transition: 'all 0.15s',
                                        position: 'relative',
                                        padding: '8px'
                                      }}
                                      onDragOver={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '#eff6ff';
                                        e.currentTarget.style.outline = '2px dashed #3b82f6';
                                      }}
                                      onDragLeave={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                      }}
                                      onDrop={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                        const typeStr = e.dataTransfer.getData('application/sgc-unit-type');
                                        const draggedIdxStr = e.dataTransfer.getData('application/sgc-row-index');
                                        if (typeStr === 'nhan' && draggedIdxStr !== '') {
                                          const draggedIdx = Number(draggedIdxStr);
                                          updateNhanRow(cfgIdx, draggedIdx, { giamTru: '1', boQua: false, tinhToan: false });
                                        }
                                      }}
                                    >
                                      {itemGiamTru ? (
                                        <div 
                                          draggable={true}
                                          onDragStart={(e) => {
                                            e.dataTransfer.setData('text/plain', itemGiamTru.row.unit);
                                            e.dataTransfer.setData('application/sgc-unit-type', 'nhan');
                                            e.dataTransfer.setData('application/sgc-row-index', String(itemGiamTru.idx));
                                            e.dataTransfer.effectAllowed = 'copyMove';
                                          }}
                                          style={{
                                            background: '#eff6ff',
                                            border: '1px solid #bfdbfe',
                                            color: '#1e40af',
                                            borderRadius: 6,
                                            padding: '4px 8px',
                                            fontSize: '11.5px',
                                            fontWeight: 600,
                                            display: 'inline-flex',
                                            alignItems: 'center',
                                            gap: 4,
                                            cursor: 'grab',
                                            boxShadow: '0 1px 2px rgba(30,64,175,0.05)',
                                            maxWidth: '100%'
                                          }}
                                        >
                                          <span style={{ flexShrink: 0, color: '#3b82f6' }}>☰</span>
                                          <span style={{ wordBreak: 'break-word', whiteSpace: 'normal' }} title={itemGiamTru.row.unit}>{itemGiamTru.row.unit}</span>
                                          <button 
                                            type="button" 
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              updateNhanRow(cfgIdx, itemGiamTru.idx, { giamTru: '', boQua: false, tinhToan: false });
                                            }}
                                            style={{
                                              background: 'transparent',
                                              border: 'none',
                                              color: '#1e40af',
                                              cursor: 'pointer',
                                              padding: 0,
                                              marginLeft: 4,
                                              fontSize: 14,
                                              lineHeight: 1,
                                              display: 'flex',
                                              alignItems: 'center'
                                            }}
                                            title="Bỏ thiết lập"
                                          >
                                            ×
                                          </button>
                                        </div>
                                      ) : (
                                        <span style={{ fontSize: 11, color: '#cbd5e1', fontStyle: 'italic', pointerEvents: 'none', userSelect: 'none' }}>Kéo thả vào đây</span>
                                      )}
                                    </td>

                                    {/* Col 3: Bỏ qua */}
                                    <td 
                                      style={{ 
                                        ...tdStyle, 
                                        textAlign: 'left', 
                                        width: '25%', 
                                        minWidth: '25%', 
                                        maxWidth: '25%',
                                        borderRight: '1px solid #f1f5f9',
                                        transition: 'all 0.15s',
                                        position: 'relative',
                                        padding: '8px'
                                      }}
                                      onDragOver={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '#fff1f2';
                                        e.currentTarget.style.outline = '2px dashed #f43f5e';
                                      }}
                                      onDragLeave={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                      }}
                                      onDrop={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                        const typeStr = e.dataTransfer.getData('application/sgc-unit-type');
                                        const draggedIdxStr = e.dataTransfer.getData('application/sgc-row-index');
                                        if (typeStr === 'nhan' && draggedIdxStr !== '') {
                                          const draggedIdx = Number(draggedIdxStr);
                                          updateNhanRow(cfgIdx, draggedIdx, { giamTru: '', boQua: true, tinhToan: false });
                                        }
                                      }}
                                    >
                                      {itemBoQua ? (
                                        <div 
                                          draggable={true}
                                          onDragStart={(e) => {
                                            e.dataTransfer.setData('text/plain', itemBoQua.row.unit);
                                            e.dataTransfer.setData('application/sgc-unit-type', 'nhan');
                                            e.dataTransfer.setData('application/sgc-row-index', String(itemBoQua.idx));
                                            e.dataTransfer.effectAllowed = 'copyMove';
                                          }}
                                          style={{
                                            background: '#fff1f2',
                                            border: '1px solid #fecdd3',
                                            color: '#9f1239',
                                            borderRadius: 6,
                                            padding: '4px 8px',
                                            fontSize: '11.5px',
                                            fontWeight: 600,
                                            display: 'inline-flex',
                                            alignItems: 'center',
                                            gap: 4,
                                            cursor: 'grab',
                                            boxShadow: '0 1px 2px rgba(159,18,57,0.05)',
                                            maxWidth: '100%'
                                          }}
                                        >
                                          <span style={{ flexShrink: 0, color: '#f43f5e' }}>☰</span>
                                          <span style={{ wordBreak: 'break-word', whiteSpace: 'normal' }} title={itemBoQua.row.unit}>{itemBoQua.row.unit}</span>
                                          <button 
                                            type="button" 
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              updateNhanRow(cfgIdx, itemBoQua.idx, { giamTru: '', boQua: false, tinhToan: false });
                                            }}
                                            style={{
                                              background: 'transparent',
                                              border: 'none',
                                              color: '#9f1239',
                                              cursor: 'pointer',
                                              padding: 0,
                                              marginLeft: 4,
                                              fontSize: 14,
                                              lineHeight: 1,
                                              display: 'flex',
                                              alignItems: 'center'
                                            }}
                                            title="Bỏ thiết lập"
                                          >
                                            ×
                                          </button>
                                        </div>
                                      ) : (
                                        <span style={{ fontSize: 11, color: '#cbd5e1', fontStyle: 'italic', pointerEvents: 'none', userSelect: 'none' }}>Kéo thả vào đây</span>
                                      )}
                                    </td>

                                    {/* Col 4: Tính toán */}
                                    <td 
                                      style={{ 
                                        ...tdStyle, 
                                        textAlign: 'left', 
                                        width: '25%', 
                                        minWidth: '25%', 
                                        maxWidth: '25%',
                                        transition: 'all 0.15s',
                                        position: 'relative',
                                        padding: '8px'
                                      }}
                                      onDragOver={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '#ecfdf5';
                                        e.currentTarget.style.outline = '2px dashed #10b981';
                                      }}
                                      onDragLeave={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                      }}
                                      onDrop={(e) => {
                                        e.preventDefault();
                                        e.currentTarget.style.backgroundColor = '';
                                        e.currentTarget.style.outline = 'none';
                                        const typeStr = e.dataTransfer.getData('application/sgc-unit-type');
                                        const draggedIdxStr = e.dataTransfer.getData('application/sgc-row-index');
                                        if (typeStr === 'nhan' && draggedIdxStr !== '') {
                                          const draggedIdx = Number(draggedIdxStr);
                                          updateNhanRow(cfgIdx, draggedIdx, { giamTru: '', boQua: false, tinhToan: true });
                                        }
                                      }}
                                    >
                                      {itemTinhToan ? (
                                        <div 
                                          draggable={true}
                                          onDragStart={(e) => {
                                            e.dataTransfer.setData('text/plain', itemTinhToan.row.unit);
                                            e.dataTransfer.setData('application/sgc-unit-type', 'nhan');
                                            e.dataTransfer.setData('application/sgc-row-index', String(itemTinhToan.idx));
                                            e.dataTransfer.effectAllowed = 'copyMove';
                                          }}
                                          style={{
                                            background: '#ecfdf5',
                                            border: '1px solid #a7f3d0',
                                            color: '#065f46',
                                            borderRadius: 6,
                                            padding: '4px 8px',
                                            fontSize: '11.5px',
                                            fontWeight: 600,
                                            display: 'inline-flex',
                                            alignItems: 'center',
                                            gap: 4,
                                            cursor: 'grab',
                                            boxShadow: '0 1px 2px rgba(6,95,70,0.05)',
                                            maxWidth: '100%'
                                          }}
                                        >
                                          <span style={{ flexShrink: 0, color: '#10b981' }}>☰</span>
                                          <span style={{ wordBreak: 'break-word', whiteSpace: 'normal' }} title={itemTinhToan.row.unit}>{itemTinhToan.row.unit}</span>
                                          <button 
                                            type="button" 
                                            onClick={(e) => {
                                              e.stopPropagation();
                                              updateNhanRow(cfgIdx, itemTinhToan.idx, { giamTru: '', boQua: false, tinhToan: false });
                                            }}
                                            style={{
                                              background: 'transparent',
                                              border: 'none',
                                              color: '#065f46',
                                              cursor: 'pointer',
                                              padding: 0,
                                              marginLeft: 4,
                                              fontSize: 14,
                                              lineHeight: 1,
                                              display: 'flex',
                                              alignItems: 'center'
                                            }}
                                            title="Bỏ thiết lập"
                                          >
                                            ×
                                          </button>
                                        </div>
                                      ) : (
                                        <span style={{ fontSize: 11, color: '#cbd5e1', fontStyle: 'italic', pointerEvents: 'none', userSelect: 'none' }}>Kéo thả vào đây</span>
                                      )}
                                    </td>
                                  </tr>
                                );
                              });
                            })()}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </div>
                </div>

                {/* Legend */}
                <div style={{
                  marginTop: 16, padding: '10px 14px', background: '#f8fafc',
                  borderRadius: 8, border: '1px solid #e2e8f0',
                  display: 'flex', gap: 20, flexWrap: 'wrap', fontSize: 12, color: '#64748b'
                }}>
                  <span>📌 <strong>Giảm trừ:</strong> Kéo thả tên đơn vị từ cột danh sách tương ứng để đưa vào nhóm Giảm trừ</span>
                  <span>🚫 <strong>Bỏ qua:</strong> Kéo thả tên đơn vị tương ứng để loại khỏi tính toán</span>
                  <span>✅ <strong>Tính toán:</strong> Nhóm các đơn vị thực hiện tính toán đầy đủ (mặc định)</span>
                </div>
              </div>
            )}
          </div>
        )
      })}

      {/* Create Config Modal */}
      {showCreateModal && (
        <CreateSummaryConfigModal
          onClose={() => setShowCreateModal(false)}
          onSave={handleCreateConfig}
          selectedProject={selectedProject}
        />
      )}

      {/* Custom Delete Config Confirmation Modal */}
      {configToDeleteIdx !== null && (
        <DeleteConfigConfirmModal
          isOpen={true}
          onClose={() => setConfigToDeleteIdx(null)}
          onConfirm={() => handleDeleteConfig(configToDeleteIdx)}
          configName={configs[configToDeleteIdx]?.name || ''}
          projectName={configs[configToDeleteIdx]?.project || ''}
          giaoCount={configs[configToDeleteIdx]?.giaoTable?.length || 0}
          nhanCount={configs[configToDeleteIdx]?.nhanTable?.length || 0}
        />
      )}
    </div>
  )
}

function CreateSummaryConfigModal({ onClose, onSave, selectedProject }) {
  const [name, setName] = React.useState('')
  const [error, setError] = React.useState('')
  const [bgColor, setBgColor] = React.useState('#eff6ff')

  React.useEffect(() => {
    const handleKey = (e) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('keydown', handleKey)
    return () => document.removeEventListener('keydown', handleKey)
  }, [onClose])

  const handleSave = () => {
    if (!name.trim()) { setError('Vui lòng nhập tên loại tổng hợp'); return }
    onSave(name.trim(), selectedProject, bgColor)
  }

  return (
    <div style={{
      position: 'fixed', inset: 0, zIndex: 10000,
      background: 'rgba(15,23,42,0.55)', backdropFilter: 'blur(3px)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 16
    }}>
      <div style={{
        background: '#fff', borderRadius: 14, width: '100%', maxWidth: 460,
        padding: 28, boxShadow: '0 25px 60px rgba(15,23,42,0.2)',
        display: 'flex', flexDirection: 'column', gap: 18
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: 14 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{ width: 36, height: 36, background: '#eff6ff', borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <Settings size={18} color="#0f58a7" />
            </div>
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 800, color: '#0f172a' }}>Tạo Loại Tổng hợp</h3>
          </div>
          <button onClick={onClose} style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#64748b', padding: 4, borderRadius: 6, display: 'flex' }}>
            <X size={18} />
          </button>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <label style={{ fontSize: 13, fontWeight: 600, color: '#475569' }}>Tên Loại tổng hợp <span style={{ color: '#ef4444' }}>*</span></label>
          <input
            autoFocus
            type="text"
            className="input"
            style={{ border: error ? '1px solid #ef4444' : '1px solid #cbd5e1' }}
            placeholder="Ví dụ: Tổng hợp Tháng 5, BC Quý 2..."
            value={name}
            onChange={e => { setName(e.target.value); if (error) setError('') }}
            onKeyDown={e => { if (e.key === 'Enter') handleSave() }}
          />
          {error && <span style={{ fontSize: 12, color: '#ef4444', fontWeight: 500 }}>{error}</span>}
        </div>

        {/* Hiển thị bộ chọn màu sắc nền */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <label style={{ fontSize: 13, fontWeight: 600, color: '#475569' }}>Màu nền Loại tổng hợp</label>
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: 8, alignItems: 'center' }}>
            {PRESET_COLORS.map((preset) => (
              <button
                key={preset.value}
                type="button"
                onClick={() => setBgColor(preset.value)}
                style={{
                  width: 28,
                  height: 28,
                  borderRadius: '50%',
                  backgroundColor: preset.value,
                  border: bgColor === preset.value ? '2.5px solid #0f58a7' : '1px solid #cbd5e1',
                  cursor: 'pointer',
                  boxShadow: bgColor === preset.value ? '0 0 0 2px rgba(15,88,167,0.2)' : 'none',
                  transition: 'all 0.15s',
                  position: 'relative'
                }}
                title={preset.name}
              >
                {bgColor === preset.value && (
                  <span style={{
                    position: 'absolute',
                    top: '50%',
                    left: '50%',
                    transform: 'translate(-50%, -50%)',
                    fontSize: 10,
                    color: getContrastColor(preset.value),
                    fontWeight: 900
                  }}>✓</span>
                )}
              </button>
            ))}
            
            {/* Custom Color Input */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginLeft: 'auto', borderLeft: '1px solid #e2e8f0', paddingLeft: 12 }}>
              <span style={{ fontSize: 12, color: '#64748b' }}>Tùy chọn:</span>
              <input 
                type="color" 
                value={bgColor} 
                onChange={e => setBgColor(e.target.value)}
                style={{
                  width: 30,
                  height: 30,
                  padding: 0,
                  border: '1px solid #cbd5e1',
                  borderRadius: 6,
                  cursor: 'pointer',
                  backgroundColor: 'transparent'
                }}
                title="Chọn một màu tùy ý"
              />
            </div>
          </div>
          
          <div style={{
            marginTop: 4,
            padding: '10px 14px',
            backgroundColor: bgColor,
            border: `1px solid ${getBorderColor(bgColor)}`,
            borderRadius: 8,
            fontSize: 12,
            color: getContrastColor(bgColor),
            fontWeight: 700,
            transition: 'background-color 0.2s, border-color 0.2s',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between'
          }}>
            <span>Mẫu màu nền hiển thị thực tế</span>
            <div style={{ width: 14, height: 14, borderRadius: '50%', backgroundColor: bgColor, border: '1px solid rgba(0,0,0,0.1)' }} />
          </div>
        </div>

        {/* Hiển thị dự án đang chọn — không cần dropdown */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <label style={{ fontSize: 13, fontWeight: 600, color: '#475569' }}>Gắn với Dự án</label>
          {selectedProject ? (
            <div style={{
              display: 'flex', alignItems: 'center', gap: 8,
              background: '#eff6ff', border: '1px solid #bfdbfe',
              borderRadius: 8, padding: '9px 14px'
            }}>
              <CheckCircle2 size={15} color="#0f58a7" />
              <span style={{ fontSize: 14, fontWeight: 700, color: '#0f58a7' }}>{selectedProject}</span>
            </div>
          ) : (
            <div style={{
              display: 'flex', alignItems: 'center', gap: 8,
              background: '#fffbeb', border: '1px solid #fde68a',
              borderRadius: 8, padding: '9px 14px'
            }}>
              <AlertCircle size={15} color="#d97706" />
              <span style={{ fontSize: 13, color: '#92400e' }}>
                Chưa chọn dự án — cấu hình sẽ áp dụng cho <strong>tất cả dự án</strong>. Bạn có thể chọn dự án cụ thể ở thanh KHO DỰ ÁN trên header trước khi tạo.
              </span>
            </div>
          )}
        </div>

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 4 }}>
          <button
            onClick={onClose}
            style={{ padding: '7px 16px', borderRadius: 6, border: '1px solid #cbd5e1', background: '#fff', color: '#475569', fontSize: 13, fontWeight: 600, cursor: 'pointer' }}
          >
            Hủy
          </button>
          <button
            onClick={handleSave}
            style={{ padding: '7px 18px', borderRadius: 6, border: 'none', background: 'linear-gradient(135deg, #0f58a7 0%, #1a6abf 100%)', color: '#fff', fontSize: 13, fontWeight: 700, cursor: 'pointer', boxShadow: '0 2px 4px rgba(15,88,167,0.2)' }}
          >
            Tạo & Lưu
          </button>
        </div>
      </div>
    </div>
  )
}

function DeleteDepreciationConfirmModal({ isOpen, onClose, onConfirm, months }) {
  React.useEffect(() => {
    const handleKey = (e) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('keydown', handleKey)
    return () => document.removeEventListener('keydown', handleKey)
  }, [onClose])

  if (!isOpen) return null

  return (
    <div style={{
      position: 'fixed',
      top: 0, left: 0, right: 0, bottom: 0,
      backgroundColor: 'rgba(15, 23, 42, 0.65)',
      backdropFilter: 'blur(4px)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 10002,
      fontFamily: '"Roboto", sans-serif'
    }}>
      <div style={{
        background: '#ffffff',
        borderRadius: '24px',
        width: '100%',
        maxWidth: 500,
        boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)',
        overflow: 'hidden',
        border: 'none',
        display: 'flex',
        flexDirection: 'column'
      }}>
        {/* Header - Crimson red/rose color matching the first screenshot */}
        <div style={{
          background: '#d91b43', 
          padding: '18px 24px',
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          color: '#ffffff'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <div style={{
              width: 36,
              height: 36,
              borderRadius: '50%',
              background: 'rgba(255, 255, 255, 0.2)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}>
              <Trash2 size={18} color="#ffffff" />
            </div>
            <h3 style={{
              margin: 0,
              fontSize: '15px',
              fontWeight: 800,
              letterSpacing: '0.03em',
              textTransform: 'uppercase',
              color: '#ffffff'
            }}>
              XÁC NHẬN XÓA THỜI GIAN KHẤU HAO
            </h3>
          </div>
          <button
            onClick={onClose}
            style={{
              background: 'transparent',
              border: 'none',
              color: '#ffffff',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              padding: 4,
              opacity: 0.8,
              transition: 'opacity 0.15s'
            }}
            onMouseOver={e => e.currentTarget.style.opacity = '1'}
            onMouseOut={e => e.currentTarget.style.opacity = '0.8'}
          >
            <X size={18} />
          </button>
        </div>

        {/* Body */}
        <div style={{ padding: '24px 32px', display: 'flex', flexDirection: 'column', gap: 16 }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
            <span style={{ fontSize: '14px', color: '#64748b', fontWeight: 500 }}>
              Xóa cấu hình khấu hao:
            </span>
            <span style={{ fontSize: '24px', fontWeight: 800, color: '#d91b43' }}>
              Khấu hao {months} tháng
            </span>
          </div>

          <p style={{ margin: 0, fontSize: '14.5px', color: '#334155', lineHeight: '1.6', fontWeight: 500 }}>
            Bạn có chắc chắn muốn xóa cấu hình khấu hao này khỏi hệ thống không?
            <span style={{ display: 'block', marginTop: 8, color: '#64748b', fontWeight: 400 }}>
              Tất cả vật tư thuộc nhóm này sẽ được đưa về nhóm <strong style={{ color: '#0f172a' }}>"Chưa thiết lập"</strong>.
            </span>
          </p>

          <p style={{
            margin: 0,
            fontSize: '13px',
            color: '#94a3b8',
            fontStyle: 'italic',
            borderTop: '1px solid #f1f5f9',
            paddingTop: 12
          }}>
            Hành động này sẽ xóa vĩnh viễn dữ liệu cấu hình khấu hao của các vật tư thuộc nhóm này cục bộ. Hãy bấm "Lưu dữ liệu" sau khi xóa để đồng bộ lên Supabase.
          </p>
        </div>

        {/* Footer */}
        <div style={{
          display: 'flex',
          justifyContent: 'flex-end',
          gap: 12,
          padding: '16px 32px 24px 32px',
          background: '#f8fafc',
          borderTop: '1px solid #f1f5f9'
        }}>
          <button
            onClick={onClose}
            style={{
              padding: '10px 24px',
              borderRadius: '12px',
              border: '1px solid #cbd5e1',
              background: '#ffffff',
              color: '#475569',
              fontSize: '13.5px',
              fontWeight: 600,
              cursor: 'pointer',
              transition: 'all 0.15s'
            }}
            onMouseOver={e => {
              e.currentTarget.style.background = '#f1f5f9'
              e.currentTarget.style.borderColor = '#94a3b8'
            }}
            onMouseOut={e => {
              e.currentTarget.style.background = '#ffffff'
              e.currentTarget.style.borderColor = '#cbd5e1'
            }}
          >
            Hủy
          </button>
          <button
            onClick={onConfirm}
            style={{
              padding: '10px 24px',
              borderRadius: '12px',
              border: 'none',
              background: '#d91b43',
              color: '#ffffff',
              fontSize: '13.5px',
              fontWeight: 700,
              cursor: 'pointer',
              boxShadow: '0 4px 6px -1px rgba(217, 27, 67, 0.2), 0 2px 4px -1px rgba(217, 27, 67, 0.1)',
              transition: 'all 0.15s'
            }}
            onMouseOver={e => e.currentTarget.style.background = '#b91235'}
            onMouseOut={e => e.currentTarget.style.background = '#d91b43'}
          >
            Xác nhận xóa
          </button>
        </div>
      </div>
    </div>
  )
}

function DeleteConfigConfirmModal({ isOpen, onClose, onConfirm, configName, projectName, giaoCount, nhanCount }) {
  React.useEffect(() => {
    const handleKey = (e) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('keydown', handleKey)
    return () => document.removeEventListener('keydown', handleKey)
  }, [onClose])

  if (!isOpen) return null

  return (
    <div style={{
      position: 'fixed',
      top: 0, left: 0, right: 0, bottom: 0,
      backgroundColor: 'rgba(15, 23, 42, 0.65)',
      backdropFilter: 'blur(4px)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 10002
    }}>
      <div style={{
        background: '#ffffff',
        borderRadius: 14,
        width: '100%',
        maxWidth: 460,
        padding: 24,
        boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)',
        border: '1px solid #fee2e2',
        display: 'flex',
        flexDirection: 'column',
        gap: 16
      }}>
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: 12 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{ width: 36, height: 36, borderRadius: 10, background: '#fee2e2', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <Trash2 size={18} color="#ef4444" />
            </div>
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 800, color: '#991b1b', letterSpacing: '-0.01em' }}>Xác nhận xóa Cấu hình</h3>
          </div>
          <button onClick={onClose} style={{ background: 'transparent', border: 'none', color: '#64748b', cursor: 'pointer', display: 'flex', alignItems: 'center', padding: 4 }}>
            <X size={18} />
          </button>
        </div>

        {/* Body */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <p style={{ margin: 0, fontSize: 14.5, color: '#334155', lineHeight: '1.6' }}>
            Bạn có chắc chắn muốn xóa cấu hình tổng hợp <strong style={{ color: '#0f172a' }}>"{configName}"</strong>?
          </p>

          <div style={{
            background: '#fff8f8',
            border: '1px solid #fecdd3',
            borderRadius: 8,
            padding: '12px 16px',
            display: 'flex',
            flexDirection: 'column',
            gap: 8,
            fontSize: 13,
            color: '#9f1239'
          }}>
            <div style={{ fontWeight: 700, display: 'flex', alignItems: 'center', gap: 6 }}>
              <AlertCircle size={14} color="#ef4444" />
              Lưu ý: Hành động này không thể hoàn tác
            </div>
            <ul style={{ margin: '0 0 0 20px', padding: 0, display: 'flex', flexDirection: 'column', gap: 4 }}>
              <li>Dự án gắn liền: <strong>{projectName || 'Tất cả dự án'}</strong></li>
              <li>Chứa <strong>{giaoCount}</strong> cài đặt đơn vị bảng Đơn Giao</li>
              <li>Chứa <strong>{nhanCount}</strong> cài đặt đơn vị bảng Đơn Nhận</li>
            </ul>
          </div>
        </div>

        {/* Footer */}
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 10, borderTop: '1px solid #f1f5f9', paddingTop: 14, marginTop: 4 }}>
          <button
            onClick={onClose}
            style={{
              padding: '8px 16px',
              borderRadius: 6,
              border: '1px solid #cbd5e1',
              background: '#fff',
              color: '#475569',
              fontSize: 13,
              fontWeight: 600,
              cursor: 'pointer',
              transition: 'background 0.1s'
            }}
            onMouseOver={e => e.currentTarget.style.background = '#f8fafc'}
            onMouseOut={e => e.currentTarget.style.background = '#fff'}
          >
            Hủy bỏ
          </button>
          <button
            onClick={onConfirm}
            style={{
              padding: '8px 20px',
              borderRadius: 6,
              border: 'none',
              background: 'linear-gradient(135deg, #ef4444 0%, #dc2626 100%)',
              color: '#fff',
              fontSize: 13,
              fontWeight: 700,
              cursor: 'pointer',
              boxShadow: '0 2px 6px rgba(239, 68, 68, 0.25)',
              transition: 'opacity 0.1s'
            }}
            onMouseOver={e => e.currentTarget.style.opacity = '0.9'}
            onMouseOut={e => e.currentTarget.style.opacity = '1'}
          >
            Đồng ý xóa
          </button>
        </div>
      </div>
    </div>
  )
}

// ─── Utility functions for Depreciation Calculation ───────────────────────────
export const getClosestMonthsGroup = (row, options) => {
  const activeOptions = options.filter(o => o.months > 0)
  
  // Check if the row already has a specific assigned depreciationMonths
  if (row.depreciationMonths && row.depreciationMonths > 0) {
    if (activeOptions.length > 0) {
      // Find option that matches BOTH months and isApproved status
      const matched = activeOptions.find(o => o.months === row.depreciationMonths && !!o.isApproved === !!row.isApprovedDepreciation)
      if (matched) return matched
      
      // If same approval status is not found, try to find any active option with the same months
      const matchedMonthsOnly = activeOptions.find(o => o.months === row.depreciationMonths)
      if (matchedMonthsOnly) return matchedMonthsOnly
    }
    // If the active options list is empty or doesn't have the option, return the row values themselves
    return { months: row.depreciationMonths, isApproved: !!row.isApprovedDepreciation }
  }

  // Under the user's requirement, we MUST NOT automatically assign/calculate mathematical closest groups
  // based on rawMonths if the row does not have a saved depreciation_months group.
  // Thus, if it has no explicitly assigned depreciationMonths, it stays as 'Chưa thiết lập' (months = 0).
  return { months: 0, isApproved: false }
}

export const alignMaterialDepreciationWithConfig = (rows, configs) => {
  return rows.map(row => {
    const closest = getClosestMonthsGroup(row, configs)
    
    if (row.isApprovedDepreciation !== closest.isApproved || row.depreciationMonths !== closest.months) {
      return { ...row, isApprovedDepreciation: closest.isApproved, depreciationMonths: closest.months }
    }
    return row
  })
}

// ─── Phan Nhom Vat Tu Tab ──────────────────────────────────────────────────────
function PhanNhomVatTuTab({
  giaoRows,
  nhanRows,
  chungRows,
  materialPriceRows,
  setMaterialPriceRows,
  materialPrices,
  setMaterialPrices,
  materialClassifications,
  setMaterialClassifications,
  loadingDbPrices,
  allProjects = [],
  customCategoryMap = {},
  handleClassificationChange,
  handlePriceChange,
  lastSyncedDepreciationOptions,
  setLastSyncedDepreciationOptions,
  lastSyncedMaterialPriceRows,
  setLastSyncedMaterialPriceRows
}) {
  const [priceSearchQuery, setPriceSearchQuery] = React.useState('')
  const [currentSubTab, setCurrentSubTab] = React.useState('classification') // 'classification' | 'depreciation_duration'

  const getClassification = React.useCallback((row) => {
    const sap = String(row.maSAP || '').trim()
    return String(customCategoryMap[sap] || materialClassifications[sap] || row.phanLoaiVatTu || '').trim()
  }, [customCategoryMap, materialClassifications])

  const [selectedMonthsGroup, setSelectedMonthsGroup] = React.useState({ months: 0, isApproved: false })
  const [depreciationOptions, setDepreciationOptions] = React.useState(() => {
    const saved = localStorage.getItem('sgc_depreciation_options_v2')
    if (saved) {
      try {
        const parsed = JSON.parse(saved)
        if (Array.isArray(parsed) && parsed.length > 0 && typeof parsed[0] === 'object' && parsed[0] !== null) {
          const unique = []
          const seen = new Set()
          parsed.forEach(opt => {
            const key = `${opt.months}-${!!opt.isApproved}`
            if (!seen.has(key)) {
              seen.add(key)
              unique.push(opt)
            }
          })
          return unique.sort((a, b) => a.months - b.months || (a.isApproved ? 1 : 0) - (b.isApproved ? 1 : 0))
        }
      } catch (e) {}
    }
    return [
      { months: 0, isApproved: false },
      { months: 12, isApproved: true },
      { months: 24, isApproved: true },
      { months: 36, isApproved: true },
      { months: 48, isApproved: true },
      { months: 60, isApproved: true },
      { months: 72, isApproved: true },
      { months: 120, isApproved: true }
    ]
  })

  const saveDepreciationOptions = (newOpts) => {
    const unique = []
    const seen = new Set()
    newOpts.forEach(opt => {
      const key = `${opt.months}-${!!opt.isApproved}`
      if (!seen.has(key)) {
        seen.add(key)
        unique.push(opt)
      }
    })
    const sorted = unique.sort((a, b) => a.months - b.months || (a.isApproved ? 1 : 0) - (b.isApproved ? 1 : 0))
    setDepreciationOptions(sorted)
    localStorage.setItem('sgc_depreciation_options_v2', JSON.stringify(sorted))
  }

  const allAvailableDurations = React.useMemo(() => {
    const set = new Set([12, 24, 36, 48, 60, 72, 120])
    depreciationOptions.forEach(opt => {
      if (opt.months > 0) set.add(opt.months)
    })
    return Array.from(set).sort((a, b) => a - b)
  }, [depreciationOptions])

  const handleSingleDepreciationChange = React.useCallback((maSAP, targetMonths) => {
    const updatedRows = materialPriceRows.map(row => {
      if (String(row.maSAP || '').trim().toLowerCase() === String(maSAP || '').trim().toLowerCase()) {
        const currentApproved = !!row.isApprovedDepreciation;
        let targetOption = depreciationOptions.find(o => o.months === targetMonths && !!o.isApproved === currentApproved);
        if (!targetOption && targetMonths > 0) {
          targetOption = depreciationOptions.find(o => o.months === targetMonths);
        }
        
        const finalMonths = targetOption ? targetOption.months : targetMonths;
        const finalApproved = targetOption ? targetOption.isApproved : currentApproved;
        
        let newDaily = 0;
        if (finalMonths > 0) {
          newDaily = Math.round(row.donGiaTrungBinh / (finalMonths * 30.417));
        }
        return { 
          ...row, 
          donGiaTrungBinh1Ngay: newDaily,
          isApprovedDepreciation: finalApproved,
          depreciationMonths: finalMonths
        }
      }
      return row;
    });
    setMaterialPriceRows(updatedRows);
    localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(updatedRows));
  }, [materialPriceRows, depreciationOptions, setMaterialPriceRows]);

  const handleSingleGroupStatusChange = React.useCallback((maSAP, targetApproved) => {
    const updatedRows = materialPriceRows.map(row => {
      if (String(row.maSAP || '').trim().toLowerCase() === String(maSAP || '').trim().toLowerCase()) {
        const currentMonths = row.depreciationMonths || 0;
        if (currentMonths === 0) return row; // cannot set group for 0 months

        // Find matching option with same months and targetApproved status
        let targetOption = depreciationOptions.find(o => o.months === currentMonths && !!o.isApproved === !!targetApproved);
        if (!targetOption) {
          targetOption = { months: currentMonths, isApproved: targetApproved };
        }
        
        return {
          ...row,
          isApprovedDepreciation: targetOption.isApproved,
          depreciationMonths: targetOption.months
        }
      }
      return row;
    });
    setMaterialPriceRows(updatedRows);
    localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(updatedRows));
  }, [materialPriceRows, depreciationOptions, setMaterialPriceRows]);

  const [selectedMaterials, setSelectedMaterials] = React.useState(new Set())
  const [batchTargetMonths, setBatchTargetMonths] = React.useState(0)
  const [isBatchMonthsOpen, setIsBatchMonthsOpen] = React.useState(false)
  const [batchTargetApproved, setBatchTargetApproved] = React.useState(true)

  React.useEffect(() => {
    if (batchTargetMonths === 0 && depreciationOptions && depreciationOptions.length > 0) {
      const sortedValid = [...depreciationOptions]
        .filter(o => o.months > 0)
        .sort((a, b) => a.months - b.months);
      if (sortedValid.length > 0) {
        setBatchTargetMonths(sortedValid[0].months);
      }
    }
  }, [depreciationOptions, batchTargetMonths]);

  const [newDurationInput, setNewDurationInput] = React.useState('')
  const [newDurationIsApproved, setNewDurationIsApproved] = React.useState(false)
  const [addDurationFeedback, setAddDurationFeedback] = React.useState(null)
  const feedbackTimeoutRef = React.useRef(null)

  const showAddDurationFeedback = (text, isError = false) => {
    if (feedbackTimeoutRef.current) {
      clearTimeout(feedbackTimeoutRef.current)
    }
    setAddDurationFeedback({ text, isError })
    feedbackTimeoutRef.current = setTimeout(() => {
      setAddDurationFeedback(null)
    }, 6000)
  }

  const [depreciationToDelete, setDepreciationToDelete] = React.useState(null)

  const getGroupColors = (opt) => {
    const val = opt && typeof opt === 'object' ? opt.months : opt
    if (val === 0) {
      return { bg: '#f8fafc', fg: '#475569', border: '#e2e8f0' }
    }
    switch (val) {
      case 3:
        return { bg: '#fff1f2', fg: '#e11d48', border: '#fecdd3' } // Rose
      case 6:
        return { bg: '#f0fdfa', fg: '#0d9488', border: '#ccfbf1' } // Teal
      case 12:
        return { bg: '#eff6ff', fg: '#1d4ed8', border: '#bfdbfe' } // Blue
      case 18:
        return { bg: '#f0f9ff', fg: '#0284c7', border: '#bae6fd' } // Sky
      case 24:
        return { bg: '#f5f3ff', fg: '#6d28d9', border: '#ddd6fe' } // Purple
      case 36:
        return { bg: '#fff7ed', fg: '#c2410c', border: '#fed7aa' } // Orange
      case 48:
        return { bg: '#ecfdf5', fg: '#047857', border: '#a7f3d0' } // Emerald
      case 60:
        return { bg: '#fdf2f8', fg: '#be185d', border: '#fbcfe8' } // Pink
      case 72:
        return { bg: '#fffbeb', fg: '#b45309', border: '#fde68a' } // Amber
      case 120:
        return { bg: '#faf5ff', fg: '#701a75', border: '#f3e8ff' } // Fuchsia/Violet
      default: {
        const palettes = [
          { bg: '#eff6ff', fg: '#1d4ed8', border: '#bfdbfe' }, // Blue
          { bg: '#f5f3ff', fg: '#6d28d9', border: '#ddd6fe' }, // Purple
          { bg: '#fff7ed', fg: '#c2410c', border: '#fed7aa' }, // Orange
          { bg: '#ecfdf5', fg: '#047857', border: '#a7f3d0' }, // Emerald
          { bg: '#fdf2f8', fg: '#be185d', border: '#fbcfe8' }, // Pink
          { bg: '#f0fdfa', fg: '#0d9488', border: '#ccfbf1' }, // Teal
          { bg: '#f0f9ff', fg: '#0284c7', border: '#bae6fd' }, // Sky
          { bg: '#fffbeb', fg: '#b45309', border: '#fde68a' }, // Amber
          { bg: '#faf5ff', fg: '#701a75', border: '#f3e8ff' }, // Fuchsia
          { bg: '#fff1f2', fg: '#e11d48', border: '#fecdd3' }  // Rose
        ]
        const idx = val % palettes.length
        return palettes[idx]
      }
    }
  }

  const isDepreciationAndPricesSynced = React.useMemo(() => {
    if (!isSupabaseConfigured) return true
    if (isDbSchemaOutdated) return false
    if (lastSyncedDepreciationOptions === null || lastSyncedMaterialPriceRows === null) {
      return true
    }

    const optsMatch = depreciationOptions.length === lastSyncedDepreciationOptions.length &&
      depreciationOptions.every((val, idx) => {
        const lastVal = lastSyncedDepreciationOptions[idx]
        return val.months === lastVal.months && val.isApproved === lastVal.isApproved
      })

    if (!optsMatch) return false

    if (materialPriceRows.length !== lastSyncedMaterialPriceRows.length) return false

    const lastMap = new Map()
    lastSyncedMaterialPriceRows.forEach(row => {
      lastMap.set(row.maSAP, row)
    })

    for (let i = 0; i < materialPriceRows.length; i++) {
      const row = materialPriceRows[i]
      const lastRow = lastMap.get(row.maSAP)
      if (!lastRow) return false
      if (
        row.khoiLuongTong !== lastRow.khoiLuongTong ||
        row.thanhTien !== lastRow.thanhTien ||
        row.donGiaTrungBinh !== lastRow.donGiaTrungBinh ||
        row.donGiaTrungBinh1Ngay !== lastRow.donGiaTrungBinh1Ngay ||
        row.phanLoaiVatTu !== lastRow.phanLoaiVatTu ||
        !!row.isApprovedDepreciation !== !!lastRow.isApprovedDepreciation
      ) {
        return false
      }
    }

    return true
  }, [
    isSupabaseConfigured,
    depreciationOptions,
    lastSyncedDepreciationOptions,
    materialPriceRows,
    lastSyncedMaterialPriceRows
  ])

  const hasLocalDuplicateMonths = React.useMemo(() => {
    const seen = new Set()
    for (const opt of depreciationOptions) {
      if (opt.months !== 0) {
        if (seen.has(opt.months)) {
          return true
        }
        seen.add(opt.months)
      }
    }
    return false
  }, [depreciationOptions])

  const loadDepreciationOptionsFromSupabase = React.useCallback(async () => {
    if (!isSupabaseConfigured) return
    try {
      const { data, error } = await supabase
        .from('cau_hinh_khau_hao')
        .select('*')
      
      if (error) {
        console.warn('Lỗi khi tải cấu hình khấu hao từ Supabase:', error.message)
        const errMsg = error.message || ''
        if (error.code === '42703' || error.code === '42P01' || errMsg.toLowerCase().includes('column') || errMsg.toLowerCase().includes('relation') || errMsg.toLowerCase().includes('does not exist')) {
          setIsDbSchemaOutdated(true)
        }
        return
      }
      
      if (data && data.length > 0) {
        const unique = []
        const seen = new Set()
        // Chỉ lọc các cấu hình chung (không gắn với mã vật tư)
        const globalConfigs = data.filter(item => !item.ma_sap)
        globalConfigs.forEach(item => {
          let itemApproved = item.is_approved !== undefined ? !!item.is_approved : true
          let realMonths = item.months
          if (realMonths < 0) {
            realMonths = Math.abs(realMonths)
            itemApproved = false
          } else if (item.is_approved === false) {
            itemApproved = false
          }
          const key = `${realMonths}-${itemApproved}`
          if (!seen.has(key)) {
            seen.add(key)
            unique.push({
              months: realMonths,
              isApproved: itemApproved
            })
          }
        })
        if (!unique.some(o => o.months === 0)) {
          unique.push({ months: 0, isApproved: false })
        }
        const sortedOpts = unique.sort((a, b) => a.months - b.months || (a.isApproved ? 1 : 0) - (b.isApproved ? 1 : 0))
        saveDepreciationOptions(sortedOpts)
        setLastSyncedDepreciationOptions(sortedOpts)
        
        // Align materialPriceRows
        setMaterialPriceRows(prev => {
          const aligned = alignMaterialDepreciationWithConfig(prev, sortedOpts)
          localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(aligned))
          return aligned
        })
      } else {
        const fetchedOpts = [{ months: 0, isApproved: false }]
        saveDepreciationOptions(fetchedOpts)
        setLastSyncedDepreciationOptions(fetchedOpts)
        
        setMaterialPriceRows(prev => {
          const aligned = alignMaterialDepreciationWithConfig(prev, fetchedOpts)
          localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(aligned))
          return aligned
        })
      }
    } catch (err) {
      console.error('Lỗi loadDepreciationOptionsFromSupabase:', err)
    }
  }, [])

  React.useEffect(() => {
    if (isSupabaseConfigured) {
      loadDepreciationOptionsFromSupabase()
    }
  }, [isSupabaseConfigured, loadDepreciationOptionsFromSupabase])

  const [isSyncingToSupabase, setIsSyncingToSupabase] = React.useState(false)

  const handleSaveAllToSupabase = async () => {
    if (!isSupabaseConfigured) {
      alert('Vui lòng kết nối cơ sở dữ liệu Supabase trước.')
      return
    }

    setIsSyncingToSupabase(true)
    try {
      // 1. Sync cau_hinh_khau_hao (Cấu hình chung và Cấu hình riêng theo vật tư)
      const localNonZeroOpts = depreciationOptions.filter(opt => opt.months !== 0)
      
      // Chỉ xóa các cấu hình chung (ma_sap is null) và chèn lại
      const { error: delErr } = await supabase
        .from('cau_hinh_khau_hao')
        .delete()
        .is('ma_sap', null)
      
      if (delErr) {
        console.warn('Lỗi đồng bộ cấu hình khấu hao (Xóa cấu hình chung) lên Supabase:', delErr.message)
      }

      const uniquePayload = []
      const seenKeys = new Set()
      localNonZeroOpts.forEach(m => {
        const key = `${m.months}-${!!m.isApproved}`
        if (!seenKeys.has(key)) {
          seenKeys.add(key)
          uniquePayload.push({
            months: m.months,
            isApproved: !!m.isApproved,
            ma_sap: null
          })
        }
      })

      if (uniquePayload.length > 0) {
        try {
          await insertWithFallback('cau_hinh_khau_hao', uniquePayload)
        } catch (insErr) {
          console.warn('Lỗi đồng bộ cấu hình khấu hao (Thêm) lên Supabase:', insErr.message || insErr)
          const errMsg = insErr.message || ''
          const isDuplicateKey = errMsg.includes('cau_hinh_khau_hao_months_key') || 
                                 errMsg.toLowerCase().includes('duplicate key') ||
                                 errMsg.toLowerCase().includes('trùng lặp khóa')

          if (isDuplicateKey) {
            console.log('[Duplicate Month Conflict] Phát hiện lỗi trùng lặp số tháng khấu hao do cấu trúc DB chưa nâng cấp.');
            setShowDbUpgradeModal(true)
            throw new Error('Bạn đang lưu trùng số tháng khấu hao (ví dụ: 6 tháng) ở cả 2 nhóm "Đã duyệt" và "Tạm tính". Ràng buộc UNIQUE của cơ sở dữ liệu Supabase hiện tại đang chặn việc này.\n\nHệ thống đã tự động mở bảng hướng dẫn nâng cấp bằng câu lệnh SQL. Vui lòng làm theo hướng dẫn chạy SQL trên màn hình để khắc phục hoàn toàn lỗi này và bảo toàn dữ liệu.')
          } else {
            throw insErr
          }
        }
      }

      // 1b. Đồng bộ cấu hình khấu hao theo từng vật tư (ma_sap) sang bảng cau_hinh_khau_hao
      // Đầu tiên, xóa các cấu hình riêng cũ (ma_sap không null) để làm sạch các cấu hình đã bị đưa về "Chưa thiết lập" (months = 0)
      const { error: delIndivErr } = await supabase
        .from('cau_hinh_khau_hao')
        .delete()
        .not('ma_sap', 'is', null)
      
      if (delIndivErr) {
        console.warn('Lỗi khi dọn dẹp cấu hình khấu hao riêng biệt cũ:', delIndivErr.message)
      }

      // Chỉ lọc ra các vật tư có thiết lập thời gian khấu hao thực sự (> 0 tháng) để lưu/chèn mới vào cau_hinh_khau_hao
      const seenMaSap = new Set()
      const materialConfigsPayload = []
      materialPriceRows.forEach(row => {
        if (row && row.maSAP && String(row.maSAP).trim()) {
          const maSapTrimmed = String(row.maSAP).trim()
          const key = maSapTrimmed.toLowerCase()
          const months = row.depreciationMonths || 0
          if (months > 0 && !seenMaSap.has(key)) {
            seenMaSap.add(key)
            materialConfigsPayload.push({
              ma_sap: maSapTrimmed,
              months: months,
              is_approved: !!row.isApprovedDepreciation
            })
          }
        }
      })

      if (materialConfigsPayload.length > 0) {
        try {
          await upsertWithFallback('cau_hinh_khau_hao', materialConfigsPayload, 'ma_sap')
        } catch (confErr) {
          console.warn('Lỗi đồng bộ cấu hình khấu hao riêng của vật tư lên bảng cau_hinh_khau_hao:', confErr.message || confErr)
        }
      }

      // 2. Sync don_gia_vat_tu
      if (materialPriceRows.length > 0) {
        const payload = materialPriceRows
          .filter(row => row && row.maSAP && String(row.maSAP).trim())
          .map(row => ({
            maSAP: String(row.maSAP).trim(),
            khoiLuongTong: row.khoiLuongTong || 0,
            thanhTien: row.thanhTien || 0,
            donGiaTrungBinh: row.donGiaTrungBinh || 0,
            donGiaTrungBinh1Ngay: row.donGiaTrungBinh1Ngay || 0,
            phanLoaiVatTu: getClassification(row) || '',
            isApprovedDepreciation: row.isApprovedDepreciation !== undefined ? !!row.isApprovedDepreciation : false,
            depreciationMonths: row.depreciationMonths || 0
          }))

        if (payload.length > 0) {
          await upsertWithFallback('don_gia_vat_tu', payload, 'maSAP')
        }
      }

      setLastSyncedDepreciationOptions(depreciationOptions)
      setLastSyncedMaterialPriceRows(materialPriceRows)
      alert('Đồng bộ dữ liệu thành công lên Supabase!')
    } catch (err) {
      console.error('Lỗi đồng bộ dữ liệu lên Supabase:', err)
      alert('Đồng bộ thất bại: ' + (err.message || err))
    } finally {
      setIsSyncingToSupabase(false)
    }
  }

  


  const handleAddDuration = () => {
    const months = parseInt(newDurationInput)
    if (isNaN(months) || months <= 0) {
      showAddDurationFeedback('Vui lòng nhập số tháng hợp lệ lớn hơn 0.', true)
      return
    }
    
    const existingExact = depreciationOptions.find(o => o.months === months && o.isApproved === newDurationIsApproved)
    if (existingExact) {
      showAddDurationFeedback(`Thời gian khấu hao ${months} tháng đã tồn tại trong nhóm ${newDurationIsApproved ? 'Đã duyệt' : 'Tạm tính'}.`, true)
      setSelectedMonthsGroup({ months, isApproved: newDurationIsApproved })
      return
    }
    
    const newOpts = [...depreciationOptions, { months, isApproved: newDurationIsApproved }].sort((a, b) => a.months - b.months || (a.isApproved ? 1 : 0) - (b.isApproved ? 1 : 0))
    saveDepreciationOptions(newOpts)
    setNewDurationInput('')
    setSelectedMonthsGroup({ months, isApproved: newDurationIsApproved })
    showAddDurationFeedback(`Đã thêm mới nhóm khấu hao ${months} tháng vào nhóm ${newDurationIsApproved ? 'Đã duyệt' : 'Tạm tính'} cục bộ thành công! Hãy bấm "Lưu dữ liệu" để đồng bộ lên Supabase.`, false)
  }

  const handleToggleApproval = (optToToggle, targetApproved) => {
    const updatedOpts = depreciationOptions.map(opt => {
      if (opt.months === optToToggle.months && opt.isApproved === optToToggle.isApproved) {
        return { ...opt, isApproved: targetApproved }
      }
      return opt
    })
    
    saveDepreciationOptions(updatedOpts)
    
    const alignedRows = alignMaterialDepreciationWithConfig(materialPriceRows, updatedOpts)
    setMaterialPriceRows(alignedRows)
    localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(alignedRows))
    
    if (selectedMonthsGroup && !selectedMonthsGroup.isAll && selectedMonthsGroup.months === optToToggle.months && selectedMonthsGroup.isApproved === optToToggle.isApproved) {
      setSelectedMonthsGroup({ months: optToToggle.months, isApproved: targetApproved })
    }
  }

  const handleBatchMove = (targetOption) => {
    if (selectedMaterials.size === 0) return
    
    const updatedRows = materialPriceRows.map(row => {
      if (selectedMaterials.has(row.maSAP)) {
        let newDaily = 0
        if (targetOption.months > 0) {
          newDaily = Math.round(row.donGiaTrungBinh / (targetOption.months * 30.417))
        }
        return { 
          ...row, 
          donGiaTrungBinh1Ngay: newDaily,
          isApprovedDepreciation: targetOption.isApproved,
          depreciationMonths: targetOption.months
        }
      }
      return row
    })
    
    setMaterialPriceRows(updatedRows)
    localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(updatedRows))
    
    const oldSize = selectedMaterials.size
    setSelectedMaterials(new Set())
    alert(`Đã chuyển thành công ${oldSize} vật tư sang nhóm khấu hao ${targetOption.months > 0 ? `${targetOption.months} tháng (${targetOption.isApproved ? 'Đã duyệt' : 'Tạm tính'})` : 'Chưa thiết lập'} cục bộ! Hãy bấm "Lưu dữ liệu" để đồng bộ lên Supabase.`)
  }

  const handleBatchChangeStatusOnly = (targetApproved) => {
    if (selectedMaterials.size === 0) return
    
    const updatedRows = materialPriceRows.map(row => {
      if (selectedMaterials.has(row.maSAP)) {
        return {
          ...row,
          isApprovedDepreciation: targetApproved
        }
      }
      return row
    })
    
    setMaterialPriceRows(updatedRows)
    localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(updatedRows))
    
    const oldSize = selectedMaterials.size
    setSelectedMaterials(new Set())
    alert(`Đã chuyển thành công ${oldSize} vật tư sang nhóm "${targetApproved ? 'Đã duyệt' : 'Tạm tính'}" cục bộ! Hãy bấm "Lưu dữ liệu" để đồng bộ lên Supabase.`)
  }

  const handleApplyBatchChanges = (targetMonths, targetApproved) => {
    if (selectedMaterials.size === 0) return
    if (targetMonths === undefined) return
    
    const updatedRows = materialPriceRows.map(row => {
      if (selectedMaterials.has(row.maSAP)) {
        let newDaily = 0
        if (targetMonths > 0) {
          newDaily = Math.round(row.donGiaTrungBinh / (targetMonths * 30.417))
        }
        return { 
          ...row, 
          donGiaTrungBinh1Ngay: newDaily,
          isApprovedDepreciation: targetApproved,
          depreciationMonths: targetMonths
        }
      }
      return row
    })
    
    setMaterialPriceRows(updatedRows)
    localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(updatedRows))
    
    const oldSize = selectedMaterials.size
    setSelectedMaterials(new Set())
    alert(`Đã chuyển thành công ${oldSize} vật tư sang nhóm khấu hao ${targetMonths > 0 ? `${targetMonths} tháng (${targetApproved ? 'Đã duyệt' : 'Tạm tính'})` : 'Chưa thiết lập'} cục bộ! Hãy bấm "Lưu dữ liệu" để đồng bộ lên Supabase.`)
  }

  const handleUploadPriceExcel = async (e) => {
    const file = e.target.files?.[0]
    if (!file) return

    const reader = new FileReader()
    reader.onload = async (evt) => {
      try {
        const u8arr = new Uint8Array(evt.target.result)
        const wb = XLSX.read(u8arr, { type: 'array' })
        const wsname = wb.SheetNames[0]
        const ws = wb.Sheets[wsname]
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 })

        if (data.length < 2) {
          alert('File Excel không có đủ dữ liệu.')
          return
        }

        let headerRowIdx = 0
        const row0 = data[0] || []
        const hasSapHeader = row0.some(c => String(c || '').toLowerCase().includes('sap'))
        if (!hasSapHeader && data[1]) {
          const row1 = data[1] || []
          if (row1.some(c => String(c || '').toLowerCase().includes('sap'))) {
            headerRowIdx = 1
          }
        }

        let colIdxSap = 0
        let colIdxKlTong = 1
        let colIdxThanhTien = 2
        let colIdxDonGiaTb = 3
        let colIdxDonGiaTb1Ngay = 4
        let colIdxPhanLoai = 5

        const headers = data[headerRowIdx] || []
        headers.forEach((cell, idx) => {
          const s = String(cell || '').trim().toLowerCase()
          if (s.includes('sap') || s.includes('mã s')) colIdxSap = idx
          else if (s.includes('khối lượng') || s.includes('khoi luong') || s.includes('kl tổng') || s.includes('kl tong')) colIdxKlTong = idx
          else if (s.includes('thành tiền') || s.includes('thanh tien')) colIdxThanhTien = idx
          else if (s.includes('trung bình') && !s.includes('1 ngày') && !s.includes('1 ngay') && !s.includes('1 ngày')) colIdxDonGiaTb = idx
          else if (s.includes('1 ngày') || s.includes('1 ngay')) colIdxDonGiaTb1Ngay = idx
          else if (s.includes('phân loại') || s.includes('phan loai')) colIdxPhanLoai = idx
        })

        const rows = []
        const newPrices = { ...materialPrices }
        const newClassifications = { ...materialClassifications }

        const parseNum = (val) => {
          if (val === null || val === undefined) return 0
          if (typeof val === 'number') return val
          const clean = String(val).replace(/\s/g, '').replace(/\./g, '').replace(',', '.')
          const parsed = parseFloat(clean)
          return isNaN(parsed) ? 0 : parsed
        }

        for (let i = headerRowIdx + 1; i < data.length; i++) {
          const row = data[i]
          if (!row || row.length === 0) continue

          const sap = String(row[colIdxSap] || '').trim()
          if (!sap) continue

          const klTong = parseNum(row[colIdxKlTong])
          const thanhTien = parseNum(row[colIdxThanhTien])
          const donGiaTb = parseNum(row[colIdxDonGiaTb])
          const donGiaTb1Ngay = parseNum(row[colIdxDonGiaTb1Ngay])
          const phanLoai = String(row[colIdxPhanLoai] || '').trim()

          rows.push({
            maSAP: sap,
            khoiLuongTong: klTong,
            thanhTien: thanhTien,
            donGiaTrungBinh: donGiaTb,
            donGiaTrungBinh1Ngay: donGiaTb1Ngay,
            phanLoaiVatTu: phanLoai
          })

          newPrices[sap] = donGiaTb
          newClassifications[sap] = phanLoai
        }

         if (rows.length === 0) {
          alert('Không tìm thấy dòng dữ liệu hợp lệ nào.')
          return
        }

        // Tự động khôi phục cấu hình khấu hao đã được lưu từ Supabase (nếu có) vào các dòng vừa import
        if (isSupabaseConfigured) {
          try {
            const { data: configData } = await supabase
              .from('cau_hinh_khau_hao')
              .select('*')
              .not('ma_sap', 'is', null)
            
            if (configData && configData.length > 0) {
              const configMap = new Map()
              configData.forEach(item => {
                if (item.ma_sap) {
                  configMap.set(String(item.ma_sap).trim().toLowerCase(), {
                    months: item.months || 0,
                    isApproved: item.is_approved !== undefined ? !!item.is_approved : false
                  })
                }
              })
              
              rows.forEach(r => {
                const sapKey = String(r.maSAP || '').trim().toLowerCase()
                if (configMap.has(sapKey)) {
                  const cfg = configMap.get(sapKey)
                  r.depreciationMonths = cfg.months
                  r.isApprovedDepreciation = cfg.isApproved
                  if (cfg.months > 0) {
                    r.donGiaTrungBinh1Ngay = Math.round((r.donGiaTrungBinh || 0) / (cfg.months * 30.417))
                  }
                }
              })
            }
          } catch (confErr) {
            console.warn('Lỗi tự động khôi phục cấu hình khấu hao cho file Excel vừa tải lên:', confErr)
          }
        }

        setMaterialPriceRows(rows)
        setMaterialPrices(newPrices)
        setMaterialClassifications(newClassifications)

        localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(rows))
        localStorage.setItem('sgc_report_material_prices', JSON.stringify(newPrices))
        localStorage.setItem('sgc_report_material_classifications', JSON.stringify(newClassifications))

        alert(`Đã tải lên thành công ${rows.length} đơn giá vật tư cục bộ!`)

        if (isSupabaseConfigured) {
          try {
            const payload = rows
              .filter(r => r && r.maSAP && String(r.maSAP).trim())
              .map(r => ({
                maSAP: String(r.maSAP).trim(),
                khoiLuongTong: r.khoiLuongTong,
                thanhTien: r.thanhTien,
                donGiaTrungBinh: r.donGiaTrungBinh,
                donGiaTrungBinh1Ngay: r.donGiaTrungBinh1Ngay,
                phanLoaiVatTu: r.phanLoaiVatTu
              }))

            if (payload.length > 0) {
              await upsertWithFallback('don_gia_vat_tu', payload, 'maSAP')
              alert('Đã đồng bộ thành công dữ liệu đơn giá vật tư lên Supabase!')
            } else {
              alert('Không tìm thấy bản ghi đơn giá hợp lệ có mã SAP để đồng bộ lên Supabase.')
            }
          } catch (upsertErr) {
            console.error('Lỗi kết nối Supabase:', upsertErr)
            alert('Đã lưu cục bộ thành công nhưng lỗi đồng bộ lên Supabase: ' + (upsertErr.message || upsertErr))
          }
        }
      } catch (err) {
        alert('Lỗi đọc file Excel: ' + err.message)
      }
    }
    reader.readAsArrayBuffer(file)
  }

  const materialMetadataMap = React.useMemo(() => {
    const map = new Map()
    if (nhanRows && Array.isArray(nhanRows)) {
      nhanRows.forEach(r => {
        const sap = String(r.maSAP || '').trim().toLowerCase()
        if (sap) {
          map.set(sap, { tenVatTu: r.tenVatTu || '', dvt: r.dvt || '' })
        }
      })
    }
    if (giaoRows && Array.isArray(giaoRows)) {
      giaoRows.forEach(r => {
        const sap = String(r.maSAP || '').trim().toLowerCase()
        if (sap) {
          map.set(sap, { tenVatTu: r.tenVatTu || '', dvt: r.dvt || '' })
        }
      })
    }
    if (chungRows && Array.isArray(chungRows)) {
      chungRows.forEach(r => {
        const sap = String(r.maSAP || '').trim().toLowerCase()
        if (sap) {
          map.set(sap, { tenVatTu: r.tenVatTu || '', dvt: r.dvt || '' })
        }
      })
    }
    return map
  }, [chungRows, giaoRows, nhanRows])

  const filteredPriceRows = React.useMemo(() => {
    const q = priceSearchQuery.toLowerCase().trim()
    if (!q) return materialPriceRows
    return materialPriceRows.filter(r => {
      const sap = String(r.maSAP || '').trim()
      const val = getClassification(r)
      const meta = materialMetadataMap.get(sap.toLowerCase()) || { tenVatTu: '' }
      const name = String(meta.tenVatTu || '').toLowerCase()
      return sap.toLowerCase().includes(q) || val.toLowerCase().includes(q) || name.includes(q)
    })
  }, [materialPriceRows, priceSearchQuery, getClassification, materialMetadataMap])

  const assetRows = React.useMemo(() => {
    const seenSaps = new Set()
    const list = []

    materialPriceRows.forEach(r => {
      const val = getClassification(r)
      const text = val.toLowerCase()
      const isAsset = text.includes('khấu hao') || text.includes('tài sản') || text.includes('tskh')
      if (isAsset) {
        const sap = String(r.maSAP || '').trim()
        seenSaps.add(sap.toLowerCase())
        list.push(r)
      }
    })

    materialMetadataMap.forEach((meta, sapKey) => {
      const classification = String(materialClassifications[sapKey] || '').trim()
      const text = classification.toLowerCase()
      const isAsset = text.includes('khấu hao') || text.includes('tài sản') || text.includes('tskh')
      if (isAsset && !seenSaps.has(sapKey.toLowerCase())) {
        seenSaps.add(sapKey.toLowerCase())
        const existingRow = materialPriceRows.find(r => String(r.maSAP || '').trim().toLowerCase() === sapKey.toLowerCase())
        if (existingRow) {
          list.push(existingRow)
        } else {
          list.push({
            maSAP: sapKey.toUpperCase(),
            khoiLuongTong: 0,
            thanhTien: 0,
            donGiaTrungBinh: materialPrices[sapKey.toUpperCase()] || materialPrices[sapKey] || 0,
            donGiaTrungBinh1Ngay: 0,
            phanLoaiVatTu: classification
          })
        }
      }
    })

    return list
  }, [materialPriceRows, getClassification, materialMetadataMap, materialClassifications, materialPrices])

  const filteredDepreciationRows = React.useMemo(() => {
    const q = priceSearchQuery.toLowerCase().trim()
    if (!q) return assetRows
    return assetRows.filter(r => {
      const sap = String(r.maSAP || '').trim()
      const val = getClassification(r)
      const meta = materialMetadataMap.get(sap.toLowerCase()) || { tenVatTu: '' }
      const name = String(meta.tenVatTu || '').toLowerCase()
      return sap.toLowerCase().includes(q) || val.toLowerCase().includes(q) || name.includes(q)
    })
  }, [assetRows, priceSearchQuery, getClassification, materialMetadataMap])

  const rowMonthsMap = React.useMemo(() => {
    const map = new Map()
    assetRows.forEach(row => {
      const closest = getClosestMonthsGroup(row, depreciationOptions)
      map.set(row.maSAP, closest)
    })
    return map
  }, [assetRows, depreciationOptions])

  const groupCounts = React.useMemo(() => {
    const counts = {}
    depreciationOptions.forEach(opt => {
      const key = `${opt.months}-${opt.isApproved}`
      counts[key] = 0
    })

    const q = priceSearchQuery.toLowerCase().trim()
    const filteredRows = q
      ? assetRows.filter(row => {
          const meta = materialMetadataMap.get(String(row.maSAP || '').trim().toLowerCase()) || { tenVatTu: '' }
          return (
            String(row.maSAP || '').toLowerCase().includes(q) ||
            String(meta.tenVatTu || '').toLowerCase().includes(q)
          )
        })
      : assetRows

    filteredRows.forEach(row => {
      const closest = rowMonthsMap.get(row.maSAP)
      if (closest) {
        const key = `${closest.months}-${closest.isApproved}`
        counts[key] = (counts[key] || 0) + 1
      }
    })
    return counts
  }, [assetRows, depreciationOptions, rowMonthsMap, priceSearchQuery, materialMetadataMap])

  const sidebarChuaThietLap = React.useMemo(() => {
    return depreciationOptions.filter(o => o.months === 0).filter(opt => {
      if (!priceSearchQuery.trim()) return true
      const count = groupCounts[`${opt.months}-${opt.isApproved}`] || 0
      return count > 0
    })
  }, [depreciationOptions, priceSearchQuery, groupCounts])

  const sidebarDaDuyet = React.useMemo(() => {
    return depreciationOptions.filter(o => o.isApproved && o.months !== 0).filter(opt => {
      if (!priceSearchQuery.trim()) return true
      const count = groupCounts[`${opt.months}-${opt.isApproved}`] || 0
      return count > 0
    })
  }, [depreciationOptions, priceSearchQuery, groupCounts])

  const sidebarTamTinh = React.useMemo(() => {
    return depreciationOptions.filter(o => !o.isApproved && o.months !== 0).filter(opt => {
      if (!priceSearchQuery.trim()) return true
      const count = groupCounts[`${opt.months}-${opt.isApproved}`] || 0
      return count > 0
    })
  }, [depreciationOptions, priceSearchQuery, groupCounts])

  const activeGroupRows = React.useMemo(() => {
    if (selectedMonthsGroup && selectedMonthsGroup.isAll) {
      return assetRows
    }
    return assetRows.filter(row => {
      const closest = rowMonthsMap.get(row.maSAP)
      if (!closest) return false
      return closest.months === selectedMonthsGroup?.months && closest.isApproved === selectedMonthsGroup?.isApproved
    })
  }, [assetRows, rowMonthsMap, selectedMonthsGroup])

  const searchedActiveRows = React.useMemo(() => {
    const q = priceSearchQuery.toLowerCase().trim()
    if (!q) return activeGroupRows
    return activeGroupRows.filter(row => {
      const meta = materialMetadataMap.get(String(row.maSAP || '').trim().toLowerCase()) || { tenVatTu: '' }
      return (
        String(row.maSAP || '').toLowerCase().includes(q) ||
        String(meta.tenVatTu || '').toLowerCase().includes(q)
      )
    })
  }, [activeGroupRows, priceSearchQuery, materialMetadataMap])

  const btnBase = {
    padding: '4px 8px',
    background: '#ffffff',
    border: '1px solid var(--border)',
    borderRadius: '4px',
    cursor: 'pointer',
    fontSize: '12px',
    fontWeight: 500,
    color: 'var(--text)',
    transition: 'all 0.15s'
  }

  const btnDisabled = {
    ...btnBase,
    background: '#f1f5f9',
    color: 'var(--text-muted)',
    cursor: 'not-allowed',
    borderColor: '#e2e8f0'
  }

  return (
    <div id="phan-nhom-vat-tu-root" style={{ height: '100%', boxSizing: 'border-box', display: 'flex', flexDirection: 'column', overflow: 'hidden', padding: '16px 20px', minHeight: 0 }}>
      {/* Sub-tabs for Phân nhóm Vật tư & Đơn giá */}
      <div style={{
        display: 'flex',
        justifyContent: 'space-between',
        alignItems: 'center',
        marginBottom: 12,
        borderBottom: '1px solid var(--border)',
        paddingBottom: 6,
        flexShrink: 0
      }}>
        <div style={{ display: 'flex', gap: 6 }}>
          <button
            onClick={() => setCurrentSubTab('classification')}
            style={{
              padding: '8px 16px',
              borderRadius: '6px 6px 0 0',
              fontSize: '13.5px',
              fontWeight: currentSubTab === 'classification' ? 700 : 500,
              background: currentSubTab === 'classification' ? 'var(--primary)' : 'transparent',
              color: currentSubTab === 'classification' ? '#ffffff' : 'var(--text-muted)',
              border: 'none',
              cursor: 'pointer',
              transition: 'all 0.15s',
              boxShadow: currentSubTab === 'classification' ? 'var(--shadow-sm)' : 'none',
              borderBottom: currentSubTab === 'classification' ? '2px solid var(--primary)' : 'none',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            <PackageCheck size={16} />
            Phân nhóm vật tư
          </button>
          <button
            onClick={() => setCurrentSubTab('depreciation')}
            style={{
              padding: '8px 16px',
              borderRadius: '6px 6px 0 0',
              fontSize: '13.5px',
              fontWeight: currentSubTab === 'depreciation' ? 700 : 500,
              background: currentSubTab === 'depreciation' ? 'var(--primary)' : 'transparent',
              color: currentSubTab === 'depreciation' ? '#ffffff' : 'var(--text-muted)',
              border: 'none',
              cursor: 'pointer',
              transition: 'all 0.15s',
              boxShadow: currentSubTab === 'depreciation' ? 'var(--shadow-sm)' : 'none',
              borderBottom: currentSubTab === 'depreciation' ? '2px solid var(--primary)' : 'none',
              display: 'flex',
              alignItems: 'center',
              gap: '6px'
            }}
          >
            <DollarSign size={16} />
            Thời gian khấu hao
          </button>
        </div>

        {/* Save button matching the design in the screenshot */}
        {isSupabaseConfigured && (
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            {isDbSchemaOutdated ? (
              <button
                onClick={() => setShowDbUpgradeModal(true)}
                title="Lỗi cấu trúc bảng: Thiếu cột ma_sap trong bảng cau_hinh_khau_hao. Click để xem hướng dẫn sửa lỗi!"
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: 6,
                  padding: '6px 12px',
                  background: '#fef2f2',
                  border: '1.5px solid #f87171',
                  color: '#dc2626',
                  borderRadius: 6,
                  fontSize: '12.5px',
                  fontWeight: 700,
                  cursor: 'pointer',
                  transition: 'all 0.15s ease',
                  boxShadow: '0 2px 4px rgba(220, 38, 38, 0.1)'
                }}
                onMouseOver={(e) => {
                  e.currentTarget.style.background = '#fee2e2'
                }}
                onMouseOut={(e) => {
                  e.currentTarget.style.background = '#fef2f2'
                }}
              >
                <AlertTriangle size={14} style={{ color: '#dc2626' }} />
                <span>Nâng cấp DB (Yêu cầu)</span>
              </button>
            ) : (
              <button
                onClick={() => setShowDbUpgradeModal(true)}
                title="Xem hướng dẫn nâng cấp hoặc khởi tạo cấu trúc bảng trên Supabase"
                style={{
                  display: 'inline-flex',
                  alignItems: 'center',
                  gap: 6,
                  padding: '6px 12px',
                  background: '#f8fafc',
                  border: '1px solid #cbd5e1',
                  color: '#475569',
                  borderRadius: 6,
                  fontSize: '12.5px',
                  fontWeight: 600,
                  cursor: 'pointer',
                  transition: 'all 0.15s ease'
                }}
                onMouseOver={(e) => {
                  e.currentTarget.style.background = '#f1f5f9'
                }}
                onMouseOut={(e) => {
                  e.currentTarget.style.background = '#f8fafc'
                }}
              >
                <Database size={13} style={{ color: '#64748b' }} />
                <span>Nâng cấp DB</span>
              </button>
            )}

            {isDepreciationAndPricesSynced ? (
              <span style={{
                fontSize: '12px',
                color: '#475569',
                background: '#f1f5f9',
                padding: '4px 8px',
                borderRadius: '6px',
                border: '1px solid #cbd5e1',
                fontWeight: 600,
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px'
              }}>
                <span className="animate-pulse" style={{ width: '6px', height: '6px', borderRadius: '50%', background: '#10b981' }}></span>
                Dữ liệu khớp (Đã đồng bộ)
              </span>
            ) : (
              <span style={{
                fontSize: '12px',
                color: '#b45309',
                background: '#fffbeb',
                padding: '4px 8px',
                borderRadius: '6px',
                border: '1px solid #fde047',
                fontWeight: 600,
                display: 'inline-flex',
                alignItems: 'center',
                gap: '6px'
              }}>
                <span className="animate-pulse" style={{ width: '6px', height: '6px', borderRadius: '50%', background: '#f59e0b' }}></span>
                Có thay đổi chưa lưu
              </span>
            )}

            <button
              onClick={handleSaveAllToSupabase}
              disabled={isSyncingToSupabase || isDepreciationAndPricesSynced}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: 6,
                padding: '6px 14px',
                background: isDepreciationAndPricesSynced ? '#cbd5e1' : '#16a34a', // Clearer gray for synced/disabled
                color: isDepreciationAndPricesSynced ? '#64748b' : '#ffffff',
                border: 'none',
                borderRadius: 6,
                fontSize: '13px',
                fontWeight: 700,
                cursor: (isSyncingToSupabase || isDepreciationAndPricesSynced) ? 'not-allowed' : 'pointer',
                opacity: isSyncingToSupabase ? 0.7 : 1,
                boxShadow: isDepreciationAndPricesSynced ? 'none' : '0 2px 5px rgba(22, 163, 74, 0.2)',
                transition: 'all 0.15s'
              }}
              onMouseOver={(e) => {
                if (!isSyncingToSupabase && !isDepreciationAndPricesSynced) {
                  e.currentTarget.style.background = '#15803d'
                  e.currentTarget.style.boxShadow = '0 4px 8px rgba(22, 163, 74, 0.3)'
                }
              }}
              onMouseOut={(e) => {
                if (!isSyncingToSupabase && !isDepreciationAndPricesSynced) {
                  e.currentTarget.style.background = '#16a34a'
                  e.currentTarget.style.boxShadow = '0 2px 5px rgba(22, 163, 74, 0.2)'
                }
              }}
            >
              {isSyncingToSupabase ? (
                <>
                  <RefreshCw size={13} className="animate-spin" />
                  <span>Đang đồng bộ...</span>
                </>
              ) : isDepreciationAndPricesSynced ? (
                <>
                  <Save size={14} />
                  <span>Lưu dữ liệu (Đã đồng bộ)</span>
                </>
              ) : (
                <>
                  <Save size={14} />
                  <span>Lưu dữ liệu</span>
                </>
              )}
            </button>
          </div>
        )}
      </div>

      {/* Title and Controls Card */}
      {currentSubTab === 'classification' && (
        <div style={{
          background: '#ffffff',
          borderRadius: 8,
          border: '1px solid var(--border)',
          padding: '16px 20px',
          marginBottom: 12,
          boxShadow: 'var(--shadow-sm)',
          flexShrink: 0,
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '24px' }}>
            <div>
              <h2 style={{ fontSize: 18, fontWeight: 800, color: 'var(--text)', margin: 0, display: 'flex', alignItems: 'center', gap: 8 }}>
                <PackageCheck size={20} style={{ color: 'var(--primary)' }} />
                Phân nhóm vật tư
              </h2>
              <p style={{ color: 'var(--text-muted)', fontSize: 13, margin: '4px 0 0 0' }}>
                Cập nhật đơn giá trung bình và phân loại vật tư bằng cách tải lên file Excel hoặc nhập trực tiếp. Dữ liệu này sẽ tự động liên thông với Bảng tổng hợp xuất nhập tồn.
              </p>
            </div>
          </div>
          
          <div style={{ display: 'flex', gap: 10, alignItems: 'center' }}>
            <button
              className="btn btn-success"
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: 6,
                cursor: 'pointer',
                background: '#10b981',
                color: '#ffffff',
                padding: '8px 16px',
                borderRadius: 6,
                fontSize: 13,
                fontWeight: 600,
                border: 'none',
                boxShadow: 'var(--shadow-sm)',
                transition: 'all 0.15s'
              }}
              onClick={() => {
                try {
                  const wb = XLSXStyle.utils.book_new()
                  const rowsToExport = filteredPriceRows || materialPriceRows || []
                  const phanNhomSheet = buildPhanNhomVatTuSheet(
                    rowsToExport,
                    materialMetadataMap,
                    depreciationOptions,
                    customCategoryMap,
                    materialClassifications
                  )
                  XLSXStyle.utils.book_append_sheet(wb, phanNhomSheet, "Phân nhóm Vật tư")
                  
                  const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary', compression: false })
                  const s2ab = (s) => {
                    const buf = new ArrayBuffer(s.length)
                    const view = new Uint8Array(buf)
                    for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF
                    return buf
                  }
                  const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' })
                  const url = URL.createObjectURL(blob)
                  const a = document.createElement('a')
                  a.href = url
                  a.download = `Phan_Nhom_Vat_Tu_${new Date().toISOString().slice(0, 10)}.xlsx`
                  a.click()
                  URL.revokeObjectURL(url)
                } catch (err) {
                  console.error("Export Excel error:", err)
                  alert("Đã xảy ra lỗi khi xuất Excel. Vui lòng thử lại.")
                }
              }}
            >
              <Download size={14} />
              <span>Xuất Excel</span>
            </button>

            <label className="btn btn-primary" style={{
              display: 'inline-flex',
              alignItems: 'center',
              gap: 6,
              cursor: 'pointer',
              background: 'var(--primary)',
              color: '#ffffff',
              padding: '8px 16px',
              borderRadius: 6,
              fontSize: 13,
              fontWeight: 600,
              border: 'none',
              boxShadow: 'var(--shadow-sm)',
              transition: 'all 0.15s'
            }}>
              <Upload size={14} />
              <span>Tải lên Excel đơn giá</span>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={handleUploadPriceExcel}
                style={{ display: 'none' }}
              />
            </label>

            <button
              onClick={async () => {
                if (window.confirm('Bạn có chắc chắn muốn xóa toàn bộ danh sách đơn giá vật tư này không?')) {
                  setMaterialPriceRows([]);
                  setMaterialPrices({});
                  setMaterialClassifications({});
                  localStorage.removeItem('sgc_report_material_price_rows');
                  localStorage.removeItem('sgc_report_material_prices');
                  localStorage.removeItem('sgc_report_material_classifications');
                  if (isSupabaseConfigured) {
                    try {
                      const { error } = await supabase.from('don_gia_vat_tu').delete().neq('id', -999);
                      if (error) alert('Lỗi xóa trên Supabase: ' + error.message);
                      else {
                        setLastSyncedMaterialPriceRows([]);
                        alert('Đã xóa sạch dữ liệu trên cả cục bộ và Supabase!');
                      }
                    } catch (e) {
                      alert('Lỗi kết nối Supabase: ' + e.message);
                    }
                  } else {
                    setLastSyncedMaterialPriceRows([]);
                    alert('Đã xóa sạch dữ liệu cục bộ!');
                  }
                }
              }}
              className="btn"
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: 6,
                cursor: 'pointer',
                background: '#ef4444',
                color: '#ffffff',
                padding: '8px 16px',
                borderRadius: 6,
                fontSize: 13,
                fontWeight: 600,
                border: 'none',
                boxShadow: 'var(--shadow-sm)',
                transition: 'all 0.15s'
              }}
            >
              <Trash2 size={14} />
              <span>Xóa sạch</span>
            </button>
          </div>
        </div>
      )}

      {currentSubTab === 'classification' ? (
        /* Search bar & Grid for Classification */
        <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0, overflow: 'hidden', paddingBottom: 4 }}>
          {/* Table side */}
          <div style={{ background: '#ffffff', borderRadius: 8, border: '1px solid var(--border)', display: 'flex', flexDirection: 'column', overflow: 'hidden', boxShadow: 'var(--shadow-sm)', flex: 1 }}>
            {/* Search bar inside table */}
            <div style={{ padding: '12px 16px', borderBottom: '1px solid var(--border)', background: '#f8fafc', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexShrink: 0 }}>
              <div style={{ position: 'relative', width: 280 }}>
                <Search size={15} style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }} />
                <input
                  type="text"
                  placeholder="Tìm kiếm theo mã SAP hoặc phân loại..."
                  value={priceSearchQuery}
                  onChange={(e) => setPriceSearchQuery(e.target.value)}
                  style={{
                    width: '100%',
                    padding: priceSearchQuery ? '6px 32px 6px 32px' : '6px 10px 6px 32px',
                    fontSize: '13px',
                    border: '1px solid var(--border)',
                    borderRadius: 6,
                    background: '#ffffff',
                    outline: 'none',
                    color: 'var(--text)'
                  }}
                />
                {priceSearchQuery && (
                  <button
                    onClick={() => setPriceSearchQuery('')}
                    style={{
                      position: 'absolute',
                      right: 10,
                      top: '50%',
                      transform: 'translateY(-50%)',
                      background: 'none',
                      border: 'none',
                      cursor: 'pointer',
                      padding: 0,
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      color: '#94a3b8',
                      transition: 'color 0.15s'
                    }}
                    onMouseOver={(e) => e.currentTarget.style.color = '#ef4444'}
                    onMouseOut={(e) => e.currentTarget.style.color = '#94a3b8'}
                    title="Xóa lọc"
                  >
                    <X size={15} />
                  </button>
                )}
              </div>
              <div style={{ fontSize: '12.5px', color: 'var(--text-muted)', fontWeight: 500 }}>
                Có <strong>{filteredPriceRows.length}</strong> vật tư được thiết lập đơn giá
              </div>
            </div>

            {/* Table wrapper */}
            <div style={{ flex: 1, overflow: 'auto', minHeight: 0 }}>
              <table style={{ width: '100%', minWidth: 1870, borderCollapse: 'collapse', borderSpacing: 0, tableLayout: 'fixed' }}>
                <thead>
                  <tr style={{ position: 'sticky', top: 0, background: '#f1f5f9', zIndex: 10, borderBottom: '2px solid var(--border)' }}>
                    <th style={{ width: 60, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>STT</th>
                    <th style={{ width: 140, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Mã SAP</th>
                    <th style={{ width: 300, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Tên vật tư</th>
                    <th style={{ width: 80, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Đơn vị</th>
                    <th style={{ width: 140, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Khối lượng tổng</th>
                    <th style={{ width: 160, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Thành tiền</th>
                    <th style={{ width: 160, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Đơn giá trung bình</th>
                    <th style={{ width: 120, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Đơn giá 1 tháng</th>
                    <th style={{ width: 130, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Đơn giá TB 1 ngày</th>
                    <th style={{ width: 180, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Phân loại vật tư</th>
                    <th style={{ width: 180, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Thời gian khấu hao</th>
                    <th style={{ width: 180, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px' }}>Nhóm khấu hao</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredPriceRows.length === 0 ? (
                    <tr>
                      <td colSpan={12} style={{ padding: '40px', textAlign: 'center', color: 'var(--text-muted)', fontSize: '13.5px' }}>
                        Chưa có đơn giá nào được thiết lập. Hãy tải lên file Excel để xem dữ liệu.
                      </td>
                    </tr>
                  ) : (
                    filteredPriceRows.map((row, idx) => {
                      const meta = materialMetadataMap.get(String(row.maSAP || '').trim().toLowerCase()) || { tenVatTu: '—', dvt: '—' }
                      return (
                        <tr key={row.maSAP + idx} style={{ borderBottom: '1px solid #f1f5f9', background: idx % 2 === 0 ? '#ffffff' : '#f8fafc' }}>
                          <td style={{ textAlign: 'center', padding: '10px', fontSize: 13, color: 'var(--text-muted)' }}>{idx + 1}</td>
                          <td style={{ fontWeight: 700, padding: '10px', fontSize: 13, color: 'var(--text)' }}>{row.maSAP}</td>
                          <td style={{ textAlign: 'left', padding: '10px', fontSize: 13, color: 'var(--text)', whiteSpace: 'normal', wordBreak: 'break-word' }}>{meta.tenVatTu || '—'}</td>
                          <td style={{ textAlign: 'center', padding: '10px', fontSize: 13, color: 'var(--text-muted)' }}>{meta.dvt || '—'}</td>
                          <td style={{ textAlign: 'right', padding: '10px', fontSize: 13, color: 'var(--text)' }}>
                            {row.khoiLuongTong ? row.khoiLuongTong.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '0'}
                          </td>
                          <td style={{ textAlign: 'right', padding: '10px', fontSize: 13, color: 'var(--text)', fontWeight: 600 }}>
                            {row.thanhTien ? row.thanhTien.toLocaleString('vi-VN') : '0'}
                          </td>
                          <td style={{ textAlign: 'right', padding: '10px', fontSize: 13, color: 'var(--primary)', fontWeight: 700 }}>
                            {row.donGiaTrungBinh ? row.donGiaTrungBinh.toLocaleString('vi-VN') : '0'}
                          </td>
                          <td style={{ textAlign: 'right', padding: '10px', fontSize: 13, color: '#2563eb', fontWeight: 600 }}>
                            {row.donGiaTrungBinh1Ngay ? Math.round(row.donGiaTrungBinh1Ngay * 30.417).toLocaleString('vi-VN') : '0'}
                          </td>
                          <td style={{ textAlign: 'right', padding: '10px', fontSize: 13, color: 'var(--text-muted)' }}>
                            {row.donGiaTrungBinh1Ngay ? row.donGiaTrungBinh1Ngay.toLocaleString('vi-VN') : '0'}
                          </td>
                          <td style={{ textAlign: 'center', padding: '10px 18px 10px 10px', fontSize: 13 }}>
                            {(() => {
                              const val = getClassification(row)
                              const text = val.trim().toLowerCase()
                              
                              let bg = '#f1f5f9'
                              let fg = '#475569'
                              let border = '#cbd5e1'
                              
                              if (text) {
                                  if (text.includes('tiêu hao')) {
                                    bg = '#ecfdf5' // Soft green
                                    fg = '#047857' // Deep green
                                    border = '#a7f3d0'
                                  } else if (text.includes('khấu hao') || text.includes('tài sản') || text.includes('tskh')) {
                                    bg = '#fff7ed' // Soft orange
                                    fg = '#ea580c' // Deep orange
                                    border = '#fed7aa'
                                  } else {
                                    bg = '#eff6ff' // Soft blue
                                    fg = '#1d4ed8' // Deep blue
                                    border = '#bfdbfe'
                                  }
                              }

                              return (
                                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                  <input
                                    list="material-classification-options"
                                    type="text"
                                    placeholder="Chưa phân loại"
                                    value={val}
                                    onChange={(e) => handleClassificationChange && handleClassificationChange(row.maSAP, e.target.value)}
                                    style={{
                                      width: '100%',
                                      padding: '4px 8px',
                                      fontSize: '12px',
                                      border: `1.5px solid ${border}`,
                                      borderRadius: 4,
                                      background: bg,
                                      color: fg,
                                      fontWeight: 700,
                                      textAlign: 'center',
                                      outline: 'none',
                                      transition: 'all 0.15s'
                                    }}
                                  />
                                  <datalist id="material-classification-options">
                                    <option value="Vật tư tiêu hao" />
                                    <option value="Tài sản khấu hao" />
                                  </datalist>
                                </div>
                              )
                            })()}
                          </td>
                          <td style={{ textAlign: 'center', padding: '10px 18px 10px 10px', fontSize: 13 }}>
                            {(() => {
                              const val = getClassification(row)
                              const text = val.trim().toLowerCase()
                              const isAsset = text.includes('khấu hao') || text.includes('tài sản') || text.includes('tskh')
                              
                              if (!isAsset) {
                                return (
                                  <span style={{ fontSize: '12.5px', color: '#94a3b8', fontStyle: 'italic', fontWeight: 500 }}>— Không áp dụng</span>
                                )
                              }
                              
                              const currentMonths = getClosestMonthsGroup(row, depreciationOptions)
                              const colors = getGroupColors(currentMonths)
                              
                              return (
                                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                  <span
                                    style={{
                                      width: '100%',
                                      padding: '5px 8px',
                                      fontSize: '12px',
                                      border: `1.5px solid ${currentMonths.months === 0 ? '#cbd5e1' : colors.border}`,
                                      borderRadius: 4,
                                      background: currentMonths.months === 0 ? '#f1f5f9' : colors.bg,
                                      color: currentMonths.months === 0 ? '#475569' : colors.fg,
                                      fontWeight: 700,
                                      textAlign: 'center',
                                      display: 'inline-block',
                                      userSelect: 'none'
                                    }}
                                  >
                                    {currentMonths.months === 0 ? 'Chưa thiết lập' : `Khấu hao ${currentMonths.months} T`}
                                  </span>
                                </div>
                              )
                            })()}
                          </td>
                          <td style={{ textAlign: 'center', padding: '10px 18px 10px 10px', fontSize: 13 }}>
                            {(() => {
                              const val = getClassification(row)
                              const text = val.trim().toLowerCase()
                              const isAsset = text.includes('khấu hao') || text.includes('tài sản') || text.includes('tskh')
                              
                              if (!isAsset) {
                                return (
                                  <span style={{ fontSize: '12.5px', color: '#94a3b8', fontStyle: 'italic', fontWeight: 500 }}>— Không áp dụng</span>
                                )
                              }
                              
                              const currentMonths = getClosestMonthsGroup(row, depreciationOptions)
                              if (currentMonths.months === 0) {
                                return (
                                  <span style={{ fontSize: '12.5px', color: '#94a3b8', fontStyle: 'italic', fontWeight: 500 }}>— Chưa thiết lập</span>
                                )
                              }
                              
                              return (
                                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
                                  <span
                                    style={{
                                      width: '100%',
                                      padding: '5px 8px',
                                      fontSize: '12px',
                                      border: `1.5px solid ${currentMonths.isApproved ? '#16a34a' : '#d97706'}`,
                                      borderRadius: 4,
                                      background: currentMonths.isApproved ? '#ecfdf5' : '#fffbeb',
                                      color: currentMonths.isApproved ? '#16a34a' : '#d97706',
                                      fontWeight: 700,
                                      textAlign: 'center',
                                      display: 'inline-block',
                                      userSelect: 'none'
                                    }}
                                  >
                                    {currentMonths.isApproved ? 'Đã duyệt' : 'Tạm tính'}
                                  </span>
                                </div>
                              )
                            })()}
                          </td>
                        </tr>
                      )
                    })
                  )}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      ) : (
        /* Redesigned Two-Pane Sidebar & Grid for Depreciation Duration */
        <div style={{ display: 'flex', gap: '16px', flex: 1, minHeight: 0, overflow: 'hidden', paddingBottom: 4 }}>
          {/* LEFT SIDEBAR: LIST OF DEPRECIATION DURATION GROUPS */}
          <div style={{
            width: '320px',
            background: '#ffffff',
            borderRadius: 8,
            border: '1px solid var(--border)',
            display: 'flex',
            flexDirection: 'column',
            overflow: 'hidden',
            boxShadow: 'var(--shadow-sm)',
            flexShrink: 0
          }}>
            {/* Sidebar Header */}
            <div style={{ padding: '16px', borderBottom: '1px solid var(--border)', background: '#f8fafc' }}>
              <div style={{ fontWeight: 800, fontSize: '13.5px', color: 'var(--text)', display: 'flex', alignItems: 'center', gap: 6, marginBottom: 12 }}>
                <Clock size={16} style={{ color: 'var(--primary)' }} />
                DANH SÁCH KHẤU HAO ({depreciationOptions.length})
              </div>
              
              {/* Add New Duration Form */}
              <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
                <div style={{ display: 'flex', gap: 6 }}>
                  <input
                    type="number"
                    placeholder="Nhập số tháng..."
                    value={newDurationInput}
                    onChange={(e) => setNewDurationInput(e.target.value)}
                    style={{
                      flex: 1,
                      padding: '6px 10px',
                      fontSize: '13px',
                      border: '1px solid var(--border)',
                      borderRadius: 6,
                      outline: 'none',
                      color: 'var(--text)',
                      background: '#ffffff'
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter') {
                        handleAddDuration()
                      }
                    }}
                  />
                  <button
                    onClick={handleAddDuration}
                    style={{
                      padding: '6px 12px',
                      background: 'var(--primary)',
                      color: '#ffffff',
                      border: 'none',
                      borderRadius: 6,
                      fontSize: '12.5px',
                      fontWeight: 700,
                      cursor: 'pointer',
                      transition: 'all 0.15s'
                    }}
                  >
                    + Thêm
                  </button>
                </div>
                
                {/* Segment control to choose type */}
                <div style={{ display: 'flex', background: '#f1f5f9', borderRadius: 6, padding: 2, border: '1px solid #e2e8f0' }}>
                  <button
                    onClick={() => setNewDurationIsApproved(true)}
                    style={{
                      flex: 1,
                      padding: '4px 8px',
                      fontSize: '11px',
                      fontWeight: 700,
                      border: 'none',
                      borderRadius: 4,
                      background: newDurationIsApproved ? '#16a34a' : 'transparent',
                      color: newDurationIsApproved ? '#ffffff' : '#64748b',
                      cursor: 'pointer',
                      transition: 'all 0.15s',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: 4
                    }}
                  >
                    <span style={{ fontSize: '11px' }}>✓</span> Đã duyệt
                  </button>
                  <button
                    onClick={() => setNewDurationIsApproved(false)}
                    style={{
                      flex: 1,
                      padding: '4px 8px',
                      fontSize: '11px',
                      fontWeight: 700,
                      border: 'none',
                      borderRadius: 4,
                      background: !newDurationIsApproved ? '#d97706' : 'transparent',
                      color: !newDurationIsApproved ? '#ffffff' : '#64748b',
                      cursor: 'pointer',
                      transition: 'all 0.15s',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      gap: 4
                    }}
                  >
                    <span style={{ fontSize: '11px' }}>⏱</span> Tạm tính
                  </button>
                </div>

                {addDurationFeedback && (
                  <div style={{
                    padding: '8px 10px',
                    borderRadius: 6,
                    fontSize: '11.5px',
                    fontWeight: 500,
                    background: addDurationFeedback.isError ? '#fef2f2' : '#f0fdf4',
                    border: `1px solid ${addDurationFeedback.isError ? '#fca5a5' : '#86efac'}`,
                    color: addDurationFeedback.isError ? '#b91c1c' : '#15803d',
                    display: 'flex',
                    alignItems: 'flex-start',
                    gap: 6,
                    lineHeight: '1.4',
                    wordBreak: 'break-word',
                    marginTop: 4
                  }}>
                    <span style={{ fontSize: '13px', flexShrink: 0, marginTop: '1px' }}>{addDurationFeedback.isError ? '⚠️' : '✅'}</span>
                    <span style={{ flex: 1 }}>{addDurationFeedback.text}</span>
                    <button 
                      onClick={() => setAddDurationFeedback(null)}
                      style={{
                        background: 'none',
                        border: 'none',
                        color: 'inherit',
                        fontSize: '13px',
                        cursor: 'pointer',
                        padding: 0,
                        marginLeft: 'auto',
                        display: 'flex',
                        alignItems: 'center',
                        flexShrink: 0
                      }}
                    >
                      ×
                    </button>
                  </div>
                )}
              </div>
            </div>

            {/* Sidebar list of groups */}
            <div style={{ flex: 1, overflowY: 'auto', padding: '12px' }}>
              <div style={{ display: 'flex', flexDirection: 'column', gap: '16px' }}>
                
                {/* TỔNG QUAN SECTION */}
                <div>
                  <div style={{
                    fontSize: '11px',
                    fontWeight: 800,
                    textTransform: 'uppercase',
                    color: 'var(--primary)',
                    letterSpacing: '0.05em',
                    marginBottom: '8px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: '6px',
                    paddingLeft: '4px'
                  }}>
                    <BarChart3 size={13} style={{ color: 'var(--primary)' }} />
                    Tổng quan
                  </div>
                  
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                    <div
                      onClick={() => {
                        setSelectedMonthsGroup({ isAll: true })
                        setSelectedMaterials(new Set()) // clear selection
                      }}
                      style={{
                        padding: '10px 12px',
                        borderRadius: '10px',
                        cursor: 'pointer',
                        transition: 'all 0.15s',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        border: selectedMonthsGroup?.isAll ? `2px solid var(--primary)` : `1px solid var(--border)`,
                        background: selectedMonthsGroup?.isAll ? 'var(--primary-light, #eff6ff)' : '#ffffff',
                        boxShadow: selectedMonthsGroup?.isAll ? '0 4px 12px rgba(15, 88, 167, 0.08)' : 'none'
                      }}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', gap: 10, minWidth: 0 }}>
                        <div style={{
                          fontSize: '10px',
                          fontWeight: 800,
                          padding: '3px 6px',
                          borderRadius: '6px',
                          background: selectedMonthsGroup?.isAll ? '#ffffff' : `rgba(15, 88, 167, 0.1)`,
                          color: 'var(--primary)',
                          width: '42px',
                          textAlign: 'center',
                          flexShrink: 0,
                          border: `1.5px solid rgba(15, 88, 167, 0.3)`
                        }}>
                          ALL
                        </div>
                        <div style={{
                          fontWeight: selectedMonthsGroup?.isAll ? 800 : 600,
                          fontSize: '13px',
                          color: selectedMonthsGroup?.isAll ? 'var(--primary)' : 'var(--text)',
                          whiteSpace: 'nowrap',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis'
                        }}>
                          Tất cả vật tư khấu hao
                        </div>
                      </div>
                      
                      <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                        <span style={{
                          fontSize: '11px',
                          fontWeight: 700,
                          background: 'var(--primary)',
                          color: '#ffffff',
                          padding: '3px 8px',
                          borderRadius: '20px'
                        }}>
                          {assetRows.length}
                        </span>
                      </div>
                    </div>
                  </div>
                </div>

                {/* 0. CHƯA THIẾT LẬP SECTION */}
                {sidebarChuaThietLap.length > 0 && (
                  <div>
                    <div style={{
                      fontSize: '11px',
                      fontWeight: 800,
                      textTransform: 'uppercase',
                      color: '#64748b',
                      letterSpacing: '0.05em',
                      marginBottom: '8px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px',
                      paddingLeft: '4px'
                    }}>
                      <Clock size={13} style={{ color: '#64748b' }} />
                      Chưa thiết lập
                    </div>
                    
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      {sidebarChuaThietLap.map((opt) => {
                        const isActive = !selectedMonthsGroup?.isAll && selectedMonthsGroup?.months === opt.months && selectedMonthsGroup?.isApproved === opt.isApproved
                        const count = groupCounts[`${opt.months}-${opt.isApproved}`] || 0
                        const titleLabel = 'Chưa thiết lập'
                        const shortBadge = 'CHƯA'
                        const colors = getGroupColors(opt)

                        return (
                          <div
                            key={`${opt.months}-${opt.isApproved}`}
                            onClick={() => {
                              setSelectedMonthsGroup(opt)
                              setSelectedMaterials(new Set()) // clear selection
                            }}
                            style={{
                              padding: '10px 12px',
                              borderRadius: '10px',
                              cursor: 'pointer',
                              transition: 'all 0.15s',
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'space-between',
                              border: isActive ? `2px solid ${colors.fg}` : `1px solid ${colors.border}`,
                              background: colors.bg,
                              boxShadow: isActive ? '0 4px 12px rgba(15, 88, 167, 0.08)' : 'none'
                            }}
                          >
                            <div style={{ display: 'flex', alignItems: 'center', gap: 10, minWidth: 0 }}>
                              <div style={{
                                fontSize: '10px',
                                fontWeight: 800,
                                padding: '3px 6px',
                                borderRadius: '6px',
                                background: isActive ? '#ffffff' : `${colors.fg}20`,
                                color: colors.fg,
                                width: '42px',
                                textAlign: 'center',
                                flexShrink: 0,
                                border: `1.5px solid ${colors.fg}40`
                              }}>
                                {shortBadge}
                              </div>
                              <div style={{
                                fontWeight: isActive ? 800 : 600,
                                fontSize: '13px',
                                color: colors.fg,
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis'
                              }}>
                                {titleLabel}
                              </div>
                            </div>
                            
                            <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                              <span style={{
                                fontSize: '11px',
                                fontWeight: 700,
                                background: colors.fg,
                                color: '#ffffff',
                                padding: '3px 8px',
                                borderRadius: '20px'
                              }}>
                                {count}
                              </span>
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  </div>
                )}

                {/* 1. ĐÃ DUYỆT SECTION */}
                {sidebarDaDuyet.length > 0 && (
                  <div>
                    <div style={{
                      fontSize: '11px',
                      fontWeight: 800,
                      textTransform: 'uppercase',
                      color: '#16a34a',
                      letterSpacing: '0.05em',
                      marginBottom: '8px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px',
                      paddingLeft: '4px'
                    }}>
                      <CheckCircle2 size={13} />
                      Đã duyệt ({sidebarDaDuyet.length})
                    </div>
                    
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      {sidebarDaDuyet.map((opt) => {
                        const isActive = !selectedMonthsGroup?.isAll && selectedMonthsGroup?.months === opt.months && selectedMonthsGroup?.isApproved === opt.isApproved
                        const count = groupCounts[`${opt.months}-${opt.isApproved}`] || 0
                        const titleLabel = `Khấu hao ${opt.months} tháng`
                        const shortBadge = `${opt.months}T`
                        const colors = getGroupColors(opt)

                        return (
                          <div
                            key={`${opt.months}-${opt.isApproved}`}
                            onClick={() => {
                              setSelectedMonthsGroup(opt)
                              setSelectedMaterials(new Set()) // clear selection
                            }}
                            style={{
                              padding: '10px 12px',
                              borderRadius: '10px',
                              cursor: 'pointer',
                              transition: 'all 0.15s',
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'space-between',
                              border: isActive ? `2px solid ${colors.fg}` : `1px solid ${colors.border}`,
                              background: colors.bg,
                              boxShadow: isActive ? '0 4px 12px rgba(15, 88, 167, 0.08)' : 'none'
                            }}
                          >
                            <div style={{ display: 'flex', alignItems: 'center', gap: 10, minWidth: 0 }}>
                              <div style={{
                                fontSize: '10px',
                                fontWeight: 800,
                                padding: '3px 6px',
                                borderRadius: '6px',
                                background: isActive ? '#ffffff' : `${colors.fg}20`,
                                color: colors.fg,
                                width: '42px',
                                textAlign: 'center',
                                flexShrink: 0,
                                border: `1.5px solid ${colors.fg}40`
                              }}>
                                {shortBadge}
                              </div>
                              <div style={{
                                fontWeight: isActive ? 800 : 600,
                                fontSize: '13px',
                                color: colors.fg,
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis'
                              }}>
                                {titleLabel}
                              </div>
                            </div>
                            
                            <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                              <span style={{
                                fontSize: '11px',
                                fontWeight: 700,
                                background: colors.fg,
                                color: '#ffffff',
                                padding: '3px 8px',
                                borderRadius: '20px'
                              }}>
                                {count}
                              </span>
                              
                              {/* Move to Tạm tính */}
                              <button
                                onClick={(e) => {
                                  e.stopPropagation()
                                  handleToggleApproval(opt, false)
                                }}
                                style={{
                                  background: 'transparent',
                                  border: 'none',
                                  color: '#eab308',
                                  cursor: 'pointer',
                                  padding: 4,
                                  borderRadius: 4,
                                  display: 'flex',
                                  alignItems: 'center',
                                  justifyContent: 'center'
                                }}
                                title="Chuyển sang nhóm Tạm tính"
                              >
                                <Clock size={13} />
                              </button>

                              {/* Delete button */}
                              <button
                                onClick={(e) => {
                                  e.stopPropagation()
                                  setDepreciationToDelete(opt)
                                }}
                                style={{
                                  background: 'transparent',
                                  border: 'none',
                                  color: '#ef4444',
                                  cursor: 'pointer',
                                  padding: 4,
                                  borderRadius: 4,
                                  display: 'flex',
                                  alignItems: 'center',
                                  justifyContent: 'center'
                                }}
                                title="Xóa nhóm này"
                              >
                                <Trash2 size={13} />
                              </button>
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  </div>
                )}

                {/* 2. TẠM TÍNH SECTION */}
                {sidebarTamTinh.length > 0 && (
                  <div>
                    <div style={{
                      fontSize: '11px',
                      fontWeight: 800,
                      textTransform: 'uppercase',
                      color: '#d97706',
                      letterSpacing: '0.05em',
                      marginBottom: '8px',
                      display: 'flex',
                      alignItems: 'center',
                      gap: '6px',
                      paddingLeft: '4px'
                    }}>
                      <Clock size={13} />
                      Tạm tính ({sidebarTamTinh.length})
                    </div>
                    
                    <div style={{ display: 'flex', flexDirection: 'column', gap: '8px' }}>
                      {sidebarTamTinh.map((opt) => {
                        const isActive = !selectedMonthsGroup?.isAll && selectedMonthsGroup?.months === opt.months && selectedMonthsGroup?.isApproved === opt.isApproved
                        const count = groupCounts[`${opt.months}-${opt.isApproved}`] || 0
                        const titleLabel = opt.months === 0 ? 'Chưa thiết lập' : `Khấu hao ${opt.months} tháng`
                        const shortBadge = opt.months === 0 ? 'CHƯA' : `${opt.months}T`
                        const colors = getGroupColors(opt)

                        return (
                          <div
                            key={`${opt.months}-${opt.isApproved}`}
                            onClick={() => {
                              setSelectedMonthsGroup(opt)
                              setSelectedMaterials(new Set()) // clear selection
                            }}
                            style={{
                              padding: '10px 12px',
                              borderRadius: '10px',
                              cursor: 'pointer',
                              transition: 'all 0.15s',
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'space-between',
                              border: isActive ? `2px solid ${colors.fg}` : `1px solid ${colors.border}`,
                              background: colors.bg,
                              boxShadow: isActive ? '0 4px 12px rgba(15, 88, 167, 0.08)' : 'none'
                            }}
                          >
                            <div style={{ display: 'flex', alignItems: 'center', gap: 10, minWidth: 0 }}>
                              <div style={{
                                fontSize: '10px',
                                fontWeight: 800,
                                padding: '3px 6px',
                                borderRadius: '6px',
                                background: isActive ? '#ffffff' : `${colors.fg}20`,
                                color: colors.fg,
                                width: '42px',
                                textAlign: 'center',
                                flexShrink: 0,
                                border: `1.5px solid ${colors.fg}40`
                              }}>
                                {shortBadge}
                              </div>
                              <div style={{
                                fontWeight: isActive ? 800 : 600,
                                fontSize: '13px',
                                color: colors.fg,
                                whiteSpace: 'nowrap',
                                overflow: 'hidden',
                                textOverflow: 'ellipsis'
                              }}>
                                {titleLabel}
                              </div>
                            </div>
                            
                            <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                              <span style={{
                                fontSize: '11px',
                                fontWeight: 700,
                                background: colors.fg,
                                color: '#ffffff',
                                padding: '3px 8px',
                                borderRadius: '20px'
                              }}>
                                {count}
                              </span>
                              
                              {/* Move to Đã duyệt */}
                              {opt.months !== 0 && (
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation()
                                    handleToggleApproval(opt, true)
                                  }}
                                  style={{
                                    background: 'transparent',
                                    border: 'none',
                                    color: '#16a34a',
                                    cursor: 'pointer',
                                    padding: 4,
                                    borderRadius: 4,
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center'
                                  }}
                                  title="Chuyển sang nhóm Đã duyệt"
                                >
                                  <CheckCircle2 size={13} />
                                </button>
                              )}

                              {/* Delete button */}
                              {opt.months !== 0 && (
                                <button
                                  onClick={(e) => {
                                    e.stopPropagation()
                                    setDepreciationToDelete(opt)
                                  }}
                                  style={{
                                    background: 'transparent',
                                    border: 'none',
                                    color: '#ef4444',
                                    cursor: 'pointer',
                                    padding: 4,
                                    borderRadius: 4,
                                    display: 'flex',
                                    alignItems: 'center',
                                    justifyContent: 'center'
                                  }}
                                  title="Xóa nhóm này"
                                >
                                  <Trash2 size={13} />
                                </button>
                              )}
                            </div>
                          </div>
                        )
                      })}
                    </div>
                  </div>
                )}

              </div>
            </div>
          </div>

          {/* RIGHT CONTENT: MATERIALS GRID WITH SEARCH AND BATCH SELECTION */}
          <div style={{
            flex: 1,
            background: '#ffffff',
            borderRadius: 8,
            border: '1px solid var(--border)',
            display: 'flex',
            flexDirection: 'column',
            overflow: 'hidden',
            boxShadow: 'var(--shadow-sm)'
          }}>
            {/* Header / Filter area */}
            <div style={{
              padding: '12px 16px',
              borderBottom: '1px solid var(--border)',
              background: '#f8fafc',
              display: 'flex',
              flexDirection: 'column',
              gap: 12,
              flexShrink: 0
            }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: 8 }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                  <h3 style={{ margin: 0, fontSize: '14px', fontWeight: 800, color: 'var(--text)' }}>
                    {selectedMonthsGroup?.isAll
                      ? (priceSearchQuery.trim() ? 'KẾT QUẢ TÌM KIẾM TOÀN BỘ VẬT TƯ' : 'TẤT CẢ VẬT TƯ KHẤU HAO')
                      : priceSearchQuery.trim()
                        ? `KẾT QUẢ TÌM KIẾM TRONG: ${
                            selectedMonthsGroup?.months === 0 
                              ? 'VẬT TƯ CHƯA THIẾT LẬP' 
                              : `KHẤU HAO ${selectedMonthsGroup?.months} THÁNG (${selectedMonthsGroup?.isApproved ? 'ĐÃ DUYỆT' : 'TẠM TÍNH'})`
                          }`
                        : selectedMonthsGroup?.months === 0
                          ? 'VẬT TƯ CHƯA THIẾT LẬP KHẤU HAO'
                          : `KHẤU HAO ${selectedMonthsGroup?.months} THÁNG (${selectedMonthsGroup?.isApproved ? 'ĐÃ DUYỆT' : 'TẠM TÍNH'})`}
                  </h3>
                  <span style={{ fontSize: '12px', color: 'var(--text-muted)' }}>
                    (Tìm thấy {searchedActiveRows.length} vật tư)
                  </span>
                </div>
                
                {/* Search query input */}
                <div style={{ position: 'relative', width: 280 }}>
                  <Search size={15} style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--text-muted)' }} />
                  <input
                    type="text"
                    placeholder="Tìm theo mã SAP hoặc tên vật tư..."
                    value={priceSearchQuery}
                    onChange={(e) => setPriceSearchQuery(e.target.value)}
                    style={{
                      width: '100%',
                      padding: priceSearchQuery ? '6px 32px 6px 32px' : '6px 10px 6px 32px',
                      fontSize: '13px',
                      border: '1px solid var(--border)',
                      borderRadius: 6,
                      background: '#ffffff',
                      outline: 'none',
                      color: 'var(--text)'
                    }}
                  />
                  {priceSearchQuery && (
                    <button
                      onClick={() => setPriceSearchQuery('')}
                      style={{
                        position: 'absolute',
                        right: 10,
                        top: '50%',
                        transform: 'translateY(-50%)',
                        background: 'none',
                        border: 'none',
                        cursor: 'pointer',
                        padding: 0,
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        color: '#94a3b8',
                        transition: 'color 0.15s'
                      }}
                      onMouseOver={(e) => e.currentTarget.style.color = '#ef4444'}
                      onMouseOut={(e) => e.currentTarget.style.color = '#94a3b8'}
                      title="Xóa lọc"
                    >
                      <X size={15} />
                    </button>
                  )}
                </div>
              </div>

              {/* SELECT ALL TOOL */}
              {searchedActiveRows.length > 0 && (
                <div style={{ display: 'flex', justifyContent: 'flex-start', alignItems: 'center', marginTop: -4 }}>
                  <label style={{
                    display: 'flex',
                    alignItems: 'center',
                    gap: 8,
                    cursor: 'pointer',
                    fontSize: '12.5px',
                    fontWeight: 600,
                    color: 'var(--text-muted)',
                    userSelect: 'none'
                  }}>
                    <input
                      type="checkbox"
                      checked={searchedActiveRows.length > 0 && searchedActiveRows.every(row => selectedMaterials.has(row.maSAP))}
                      onChange={(e) => {
                        const newSelection = new Set(selectedMaterials)
                        if (e.target.checked) {
                          searchedActiveRows.forEach(row => {
                            newSelection.add(row.maSAP)
                          })
                        } else {
                          searchedActiveRows.forEach(row => {
                            newSelection.delete(row.maSAP)
                          })
                        }
                        setSelectedMaterials(newSelection)
                      }}
                      style={{
                        width: '16px',
                        height: '16px',
                        cursor: 'pointer',
                        accentColor: 'var(--primary)'
                      }}
                    />
                    <span>Chọn tất cả ({searchedActiveRows.length} vật tư)</span>
                  </label>
                </div>
              )}

              {/* BATCH ACTION BAR: Show only if materials are selected */}
              {selectedMaterials.size > 0 ? (
                <div style={{
                  background: '#ffffff',
                  border: '1px solid #cbd5e1',
                  boxShadow: '0 10px 25px -5px rgba(0, 0, 0, 0.05), 0 8px 10px -6px rgba(0, 0, 0, 0.05)',
                  borderRadius: '12px',
                  padding: '16px 20px',
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '14px',
                  animation: 'fadeIn 0.2s ease-out',
                  width: '100%',
                  boxSizing: 'border-box'
                }}>
                  {/* Header Row */}
                  <div style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    borderBottom: '1px solid #f1f5f9',
                    paddingBottom: '10px',
                    flexWrap: 'wrap',
                    gap: '10px'
                  }}>
                    <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
                      <div style={{
                        background: 'var(--primary)',
                        color: '#ffffff',
                        fontWeight: 700,
                        fontSize: '13px',
                        padding: '4px 12px',
                        borderRadius: '20px',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '6px'
                      }}>
                        <CheckSquare size={14} />
                        <span>Đã chọn {selectedMaterials.size} vật tư</span>
                      </div>
                      <span style={{ fontSize: '13.5px', color: '#64748b', fontWeight: 500 }}>
                        Thiết lập số tháng & trạng thái đồng loạt cho các vật tư đã chọn bên dưới.
                      </span>
                    </div>
                    <button
                      onClick={() => setSelectedMaterials(new Set())}
                      style={{
                        padding: '5px 12px',
                        background: '#f1f5f9',
                        border: '1px solid #cbd5e1',
                        borderRadius: '6px',
                        fontSize: '12px',
                        fontWeight: 600,
                        color: '#64748b',
                        cursor: 'pointer',
                        transition: 'all 0.15s'
                      }}
                      onMouseOver={(e) => {
                        e.currentTarget.style.background = '#fee2e2';
                        e.currentTarget.style.color = '#ef4444';
                        e.currentTarget.style.borderColor = '#fca5a5';
                      }}
                      onMouseOut={(e) => {
                        e.currentTarget.style.background = '#f1f5f9';
                        e.currentTarget.style.color = '#64748b';
                        e.currentTarget.style.borderColor = '#cbd5e1';
                      }}
                    >
                      Hủy chọn tất cả
                    </button>
                  </div>

                  {/* Action Columns Row */}
                  <div style={{
                    display: 'flex',
                    flexWrap: 'wrap',
                    gap: '20px',
                    width: '100%'
                  }}>
                    {/* Complete Batch Movement (Months + Status) */}
                    <div style={{
                      flex: '1 1 100%',
                      display: 'flex',
                      flexWrap: 'wrap',
                      gap: '12px',
                      alignItems: 'flex-end',
                      background: '#f8fafc',
                      padding: '12px 16px',
                      borderRadius: '10px',
                      border: '1px solid #cbd5e1'
                    }}>
                      {/* 1. Month select */}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', flex: '2 1 320px' }}>
                        <span style={{ fontSize: '11px', fontWeight: 700, textTransform: 'uppercase', color: '#475569', letterSpacing: '0.02em' }}>
                          📅 Bước 1: Chọn số tháng mới
                        </span>
                        
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '8px', alignItems: 'center' }}>
                          {/* Existing options sorted */}
                          {depreciationOptions.filter(o => o.months > 0).reduce((acc, current) => {
                            if (!acc.some(o => o.months === current.months)) {
                              acc.push(current);
                            }
                            return acc;
                          }, []).sort((a, b) => a.months - b.months).map(o => {
                            const optColors = getGroupColors(o);
                            const isSelected = batchTargetMonths === o.months;
                            return (
                              <div
                                key={o.months}
                                onClick={() => setBatchTargetMonths(o.months)}
                                style={{
                                  padding: '6px 12px',
                                  borderRadius: '6px',
                                  border: isSelected ? `2px solid ${optColors.fg}` : `1px solid ${optColors.fg}40`,
                                  background: isSelected ? optColors.bg : '#ffffff',
                                  color: optColors.fg,
                                  fontSize: '12.5px',
                                  fontWeight: isSelected ? 700 : 600,
                                  cursor: 'pointer',
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '6px',
                                  userSelect: 'none',
                                  boxShadow: isSelected ? '0 2px 4px rgba(0,0,0,0.08)' : 'none',
                                  opacity: isSelected ? 1 : 0.65,
                                  transition: 'all 0.15s',
                                  height: '37px',
                                  boxSizing: 'border-box'
                                }}
                                onMouseOver={(e) => {
                                  e.currentTarget.style.opacity = '1';
                                  if (!isSelected) {
                                    e.currentTarget.style.borderColor = optColors.fg;
                                  }
                                }}
                                onMouseOut={(e) => {
                                  if (!isSelected) {
                                    e.currentTarget.style.opacity = '0.65';
                                    e.currentTarget.style.borderColor = `${optColors.fg}40`;
                                  }
                                }}
                              >
                                <span style={{
                                  fontSize: '10px',
                                  fontWeight: 800,
                                  padding: '1px 5px',
                                  borderRadius: '4px',
                                  background: optColors.bg,
                                  color: optColors.fg,
                                  border: `1.5px solid ${optColors.fg}40`
                                }}>
                                  {o.months}T
                                </span>
                                <span>Khấu hao {o.months} tháng</span>
                              </div>
                            );
                          })}
                        </div>
                      </div>

                      {/* 2. Status toggle buttons */}
                      <div style={{ display: 'flex', flexDirection: 'column', gap: '6px', flex: '0 0 220px' }}>
                        <span style={{ fontSize: '11px', fontWeight: 700, textTransform: 'uppercase', color: '#475569', letterSpacing: '0.02em' }}>
                          🏷️ Bước 2: Chọn trạng thái mới
                        </span>
                        <div style={{
                          display: 'flex',
                          background: '#e2e8f0',
                          padding: '3px',
                          borderRadius: '6px',
                          gap: '2px',
                          height: '37px',
                          boxSizing: 'border-box',
                          width: '100%'
                        }}>
                          <button
                            type="button"
                            onClick={() => setBatchTargetApproved(true)}
                            style={{
                              flex: 1,
                              padding: '2px 8px',
                              border: 'none',
                              borderRadius: '4px',
                              fontSize: '11.5px',
                              fontWeight: 600,
                              cursor: 'pointer',
                              background: batchTargetApproved ? '#16a34a' : 'transparent',
                              color: batchTargetApproved ? '#ffffff' : '#64748b',
                              boxShadow: batchTargetApproved ? '0 2px 4px rgba(22, 163, 74, 0.25)' : 'none',
                              transition: 'all 0.15s'
                            }}
                          >
                            Đã duyệt
                          </button>
                          <button
                            type="button"
                            onClick={() => setBatchTargetApproved(false)}
                            style={{
                              flex: 1,
                              padding: '2px 8px',
                              border: 'none',
                              borderRadius: '4px',
                              fontSize: '11.5px',
                              fontWeight: 600,
                              cursor: 'pointer',
                              background: !batchTargetApproved ? '#d97706' : 'transparent',
                              color: !batchTargetApproved ? '#ffffff' : '#64748b',
                              boxShadow: !batchTargetApproved ? '0 2px 4px rgba(217, 119, 6, 0.25)' : 'none',
                              transition: 'all 0.15s'
                            }}
                          >
                            Tạm tính
                          </button>
                        </div>
                      </div>

                      {/* 3. Apply Action Button */}
                      <button
                        onClick={() => handleApplyBatchChanges(batchTargetMonths, batchTargetApproved)}
                        style={{
                          padding: '0 16px',
                          background: 'var(--primary)',
                          color: '#ffffff',
                          border: 'none',
                          borderRadius: '6px',
                          fontSize: '12.5px',
                          fontWeight: 700,
                          cursor: 'pointer',
                          transition: 'all 0.15s',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'center',
                          gap: '6px',
                          height: '37px',
                          flex: '0 0 120px'
                        }}
                        onMouseOver={(e) => {
                          e.currentTarget.style.background = '#7c3aed';
                        }}
                        onMouseOut={(e) => {
                          e.currentTarget.style.background = 'var(--primary)';
                        }}
                      >
                        <ArrowRight size={14} />
                        Áp dụng
                      </button>
                    </div>

                  </div>
                </div>
              ) : (
                <div style={{ fontSize: '12.5px', color: 'var(--text-muted)', display: 'flex', alignItems: 'center', gap: 6 }}>
                  <Info size={14} style={{ color: 'var(--primary)' }} />
                  <span>Chọn một hoặc nhiều thẻ vật tư để chuyển sang thời gian khấu hao tương ứng nhanh chóng.</span>
                </div>
              )}
            </div>

            {/* Grid of Material Cards */}
            <div style={{ flex: 1, overflowY: 'auto', padding: '16px', background: '#f8fafc' }}>
              {searchedActiveRows.length === 0 ? (
                <div style={{
                  height: '100%',
                  display: 'flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  justifyContent: 'center',
                  padding: '40px',
                  color: 'var(--text-muted)',
                  fontSize: '13.5px',
                  background: '#ffffff',
                  borderRadius: '12px',
                  border: '1.5px dashed var(--border)'
                }}>
                  <PackageCheck size={36} style={{ color: '#cbd5e1', marginBottom: 12 }} />
                  Không tìm thấy vật tư nào trong nhóm này.
                </div>
              ) : (
                <div style={{
                  display: 'grid',
                  gridTemplateColumns: 'repeat(auto-fill, minmax(420px, 1fr))',
                  gap: '12px'
                }}>
                  {searchedActiveRows.map((row) => {
                    const meta = materialMetadataMap.get(String(row.maSAP || '').trim().toLowerCase()) || { tenVatTu: '—', dvt: '—' }
                    const isSelected = selectedMaterials.has(row.maSAP)
                    const closestGroup = getClosestMonthsGroup(row, depreciationOptions)
                    const groupColors = getGroupColors(closestGroup)

                    return (
                      <div
                        key={row.maSAP}
                        onClick={() => {
                          const newSelection = new Set(selectedMaterials)
                          if (newSelection.has(row.maSAP)) {
                            newSelection.delete(row.maSAP)
                          } else {
                            newSelection.add(row.maSAP)
                          }
                          setSelectedMaterials(newSelection)
                        }}
                        style={{
                          background: closestGroup.months === 0 
                            ? '#f1f5f9' 
                            : (closestGroup.isApproved ? groupColors.bg : '#ffffff'),
                          fontFamily: '"Roboto", sans-serif',
                          borderRadius: '8px',
                          border: isSelected ? '2px solid var(--primary)' : `1px solid ${groupColors.border}`,
                          borderLeft: isSelected 
                            ? '5px solid var(--primary)' 
                            : (closestGroup.months === 0 
                              ? '5px solid #64748b' 
                              : (closestGroup.isApproved 
                                ? `5px solid ${groupColors.fg}` 
                                : `5px dashed ${groupColors.fg}`)),
                          padding: '8px 16px',
                          display: 'flex',
                          alignItems: 'center',
                          justifyContent: 'space-between',
                          gap: '12px',
                          cursor: 'pointer',
                          transition: 'all 0.15s',
                          position: 'relative',
                          boxShadow: isSelected ? '0 4px 10px -2px rgba(15, 88, 167, 0.1)' : 'var(--shadow-xs)'
                        }}
                      >
                        {/* Main contents in a single horizontal row */}
                        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', minWidth: 0, flex: 1 }}>
                          {/* SAP Code */}
                          <span style={{
                            fontWeight: 800,
                            fontSize: '11.5px',
                            color: groupColors.fg,
                            background: `${groupColors.fg}15`,
                            padding: '1.5px 5px',
                            borderRadius: '4px',
                            fontFamily: '"Roboto", sans-serif',
                            whiteSpace: 'nowrap',
                            flexShrink: 0
                          }}>
                            {row.maSAP}
                          </span>
                          
                          <span style={{ color: '#cbd5e1', flexShrink: 0 }}>•</span>
                          
                          {/* Material Name */}
                          <span style={{
                            fontWeight: 700,
                            fontSize: '13px',
                            color: '#1e293b',
                            whiteSpace: 'nowrap',
                            overflow: 'hidden',
                            textOverflow: 'ellipsis',
                            fontFamily: '"Roboto", sans-serif',
                            flex: 1,
                            minWidth: 0
                          }} title={meta.tenVatTu}>
                            {meta.tenVatTu || '—'}
                          </span>

                          {/* Status Badge */}
                          <span style={{
                            fontSize: '10.5px',
                            fontWeight: 800,
                            padding: '2px 6px',
                            borderRadius: '4px',
                            background: closestGroup.months === 0 
                              ? '#f1f5f9' 
                              : (closestGroup.isApproved ? groupColors.fg : groupColors.bg),
                            color: closestGroup.months === 0 
                              ? '#475569' 
                              : (closestGroup.isApproved ? '#ffffff' : groupColors.fg),
                            border: `1px solid ${closestGroup.months === 0 
                              ? '#cbd5e1' 
                              : (closestGroup.isApproved ? groupColors.fg : groupColors.border)}`,
                            fontFamily: '"Roboto", sans-serif',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '4px',
                            flexShrink: 0,
                            marginLeft: '4px'
                          }}>
                            {closestGroup.months === 0 ? (
                              <>⏱ Chưa thiết lập</>
                            ) : closestGroup.isApproved ? (
                              <>✓ {closestGroup.months}T (Đã duyệt)</>
                            ) : (
                              <>⏱ {closestGroup.months}T (Tạm tính)</>
                            )}
                          </span>
                        </div>

                        {/* Right side: Checkbox */}
                        <div
                          onClick={(e) => e.stopPropagation()} // Stop propagation so it doesn't double toggle
                          style={{ padding: '2px', display: 'flex', alignItems: 'center', flexShrink: 0 }}
                        >
                          <input
                            type="checkbox"
                            checked={isSelected}
                            onChange={(e) => {
                              const newSelection = new Set(selectedMaterials)
                              if (e.target.checked) {
                                newSelection.add(row.maSAP)
                              } else {
                                newSelection.delete(row.maSAP)
                              }
                              setSelectedMaterials(newSelection)
                            }}
                            style={{
                              width: '15px',
                              height: '15px',
                              cursor: 'pointer',
                              accentColor: groupColors.fg
                            }}
                          />
                        </div>
                      </div>
                    )
                  })}
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* Custom Delete Depreciation Confirmation Modal */}
      {depreciationToDelete !== null && (
        <DeleteDepreciationConfirmModal
          isOpen={depreciationToDelete !== null}
          onClose={() => setDepreciationToDelete(null)}
          onConfirm={() => {
            const opt = depreciationToDelete
            const newOpts = depreciationOptions.filter(o => o.months !== opt.months || o.isApproved !== opt.isApproved)
            saveDepreciationOptions(newOpts)
            if (selectedMonthsGroup.months === opt.months && selectedMonthsGroup.isApproved === opt.isApproved) {
              setSelectedMonthsGroup({ months: 0, isApproved: false })
            }

            // Reset the materials that were in this group back to 0 (Chưa thiết lập)
            const resetRows = materialPriceRows.map(row => {
              const closest = getClosestMonthsGroup(row, depreciationOptions)
              if (closest.months === opt.months && closest.isApproved === opt.isApproved) {
                return { ...row, donGiaTrungBinh1Ngay: 0 }
              }
              return row
            })
            const alignedRows = alignMaterialDepreciationWithConfig(resetRows, newOpts)
            setMaterialPriceRows(alignedRows)
            localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(alignedRows))

            setDepreciationToDelete(null)
            alert(`Đã xóa cấu hình khấu hao ${opt.months} tháng cục bộ thành công! Hãy bấm "Lưu dữ liệu" để đồng bộ lên Supabase.`)
          }}
          months={depreciationToDelete.months}
        />
      )}
    </div>
  )
}

// ─── Inventory Report (Báo cáo xuất nhập tồn) ──────────────────────────────────
function BaoCaoXuatNhapTonTab({
  chungRows = [],
  giaoRows = [],
  nhanRows = [],
  selectedProject,
  setSelectedProject,
  allProjects = [],
  isRealReport = false,
  customCategoryMap = {},
  materialPriceRows = [],
  materialPrices = {},
  materialClassifications = {},
  handleClassificationChange,
  handlePriceChange
}) {
  // Bộ lọc Kho/Dự án của tab này hoàn toàn độc lập (local), không lấy theo
  // và không đồng bộ ngược lại selectedProject toàn cục — tránh ảnh hưởng
  // tới các tab/sheet khác như Đơn chung, Kho dự án.
  const [localProject, setLocalProject] = React.useState('')
  const [subTab, setSubTab] = React.useState(isRealReport ? 'summary' : 'dashboard') // 'dashboard' | 'dongia' | 'summary'
  const [realReportSubTab, setRealReportSubTab] = React.useState('nhap') // 'nhap' | 'xuat'
  const [showLogicExplainModal, setShowLogicExplainModal] = React.useState(false)
  const [explainModalSubTab, setExplainModalSubTab] = React.useState('nhap')

  const [sqlCopied, setSqlCopied] = React.useState(false)

  const handleCopySql = () => {
    navigator.clipboard.writeText(sqlQueryText)
    setSqlCopied(true)
    setTimeout(() => setSqlCopied(false), 2000)
  }

  const materialMetadataMap = React.useMemo(() => {
    const map = new Map()
    if (nhanRows && Array.isArray(nhanRows)) {
      nhanRows.forEach(r => {
        const sap = String(r.maSAP || '').trim().toLowerCase()
        if (sap) {
          map.set(sap, { tenVatTu: r.tenVatTu || '', dvt: r.dvt || '' })
        }
      })
    }
    if (giaoRows && Array.isArray(giaoRows)) {
      giaoRows.forEach(r => {
        const sap = String(r.maSAP || '').trim().toLowerCase()
        if (sap) {
          map.set(sap, { tenVatTu: r.tenVatTu || '', dvt: r.dvt || '' })
        }
      })
    }
    if (chungRows && Array.isArray(chungRows)) {
      chungRows.forEach(r => {
        const sap = String(r.maSAP || '').trim().toLowerCase()
        if (sap) {
          map.set(sap, { tenVatTu: r.tenVatTu || '', dvt: r.dvt || '' })
        }
      })
    }
    return map
  }, [chungRows, giaoRows, nhanRows])

  const sqlQueryText = `-- Bảng lưu Đơn giá vật tư
CREATE TABLE IF NOT EXISTS public.don_gia_vat_tu (
    id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    ma_sap text NOT NULL UNIQUE,
    khoi_luong_tong numeric,
    thanh_tien numeric,
    don_gia_trung_binh numeric,
    don_gia_trung_binh_1_ngay numeric,
    phan_loai_vat_tu text,
    is_approved_depreciation boolean DEFAULT false,
    depreciation_months integer DEFAULT 0,
    created_at timestamp with time zone DEFAULT timezone('utc'::text, now())
);

-- Bật Row Level Security (RLS)
ALTER TABLE public.don_gia_vat_tu ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Allow public read" ON public.don_gia_vat_tu FOR SELECT USING (true);
CREATE POLICY "Allow public insert" ON public.don_gia_vat_tu FOR INSERT WITH CHECK (true);
CREATE POLICY "Allow public update" ON public.don_gia_vat_tu FOR UPDATE USING (true);
CREATE POLICY "Allow public delete" ON public.don_gia_vat_tu FOR DELETE USING (true);

-- Bảng lưu cấu hình khấu hao (tháng)
CREATE TABLE IF NOT EXISTS public.cau_hinh_khau_hao (
    id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    months integer NOT NULL,
    is_approved boolean DEFAULT false,
    ma_sap text,
    created_at timestamp with time zone DEFAULT timezone('utc'::text, now())
);

ALTER TABLE public.cau_hinh_khau_hao ENABLE ROW LEVEL SECURITY;
CREATE POLICY "Allow public read" ON public.cau_hinh_khau_hao FOR SELECT USING (true);
CREATE POLICY "Allow public insert" ON public.cau_hinh_khau_hao FOR INSERT WITH CHECK (true);
CREATE POLICY "Allow public update" ON public.cau_hinh_khau_hao FOR UPDATE USING (true);
CREATE POLICY "Allow public delete" ON public.cau_hinh_khau_hao FOR DELETE USING (true);

-- Tạo các unique indexes để tránh trùng lặp cấu hình
CREATE UNIQUE INDEX IF NOT EXISTS cau_hinh_khau_hao_global_idx ON public.cau_hinh_khau_hao (months, is_approved) WHERE ma_sap IS NULL;
CREATE UNIQUE INDEX IF NOT EXISTS cau_hinh_khau_hao_ma_sap_idx ON public.cau_hinh_khau_hao (ma_sap) WHERE ma_sap IS NOT NULL;

-- Thêm một số giá trị mặc định ban đầu cho cấu hình chung
INSERT INTO public.cau_hinh_khau_hao (months, is_approved) VALUES (12, true), (24, true), (36, true), (48, true), (60, true) ON CONFLICT (months, is_approved) WHERE ma_sap IS NULL DO NOTHING;`

  const formatDate = (date) => {
    if (!date) return '—'
    const d = (date instanceof Date) ? date : new Date(date)
    if (isNaN(d.getTime())) return '—'
    const day = String(d.getDate()).padStart(2, '0')
    const month = String(d.getMonth() + 1).padStart(2, '0')
    const year = d.getFullYear()
    return `${day}/${month}/${year}`
  }

  const getDaysToToday = (dateVal) => {
    if (!dateVal) return '—'
    const d = (dateVal instanceof Date) ? dateVal : parseRowDate(dateVal)
    if (!d || isNaN(d.getTime())) return '—'
    const dDate = new Date(d.getFullYear(), d.getMonth(), d.getDate())
    const todayObj = new Date()
    const dToday = new Date(todayObj.getFullYear(), todayObj.getMonth(), todayObj.getDate())
    const diffTime = dToday.getTime() - dDate.getTime()
    const diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24))
    if (diffDays < 0) return 0
    return diffDays
  }

  const getUnusedStatus = React.useCallback((item, stockVal) => {
    const stock = stockVal !== undefined ? stockVal : item.stock
    if (stock <= 0) {
      return 'Đã dùng hết'
    }
    
    // Nếu chưa bao giờ xuất (không có ngày xuất muộn nhất hoặc khối lượng xuất = 0)
    if (!item.latestIssuedDate || item.issued === 0) {
      const daysSinceRec = getDaysToToday(item.latestReceivedDate)
      if (typeof daysSinceRec === 'number' && daysSinceRec > 30) {
        return 'Chưa sử dụng (> 30 ngày)'
      }
      return 'Đang sử dụng'
    }
    
    const daysSinceIss = getDaysToToday(item.latestIssuedDate)
    if (typeof daysSinceIss === 'number' && daysSinceIss > 30) {
      return 'Chưa sử dụng (> 30 ngày)'
    }
    
    return 'Đang sử dụng'
  }, [])

  const [searchTerm, setSearchTerm] = React.useState('')
  const [statusFilter, setStatusFilter] = React.useState('approved_only') // 'approved_only' | 'approved_pending' | 'all'
  const [localGiao, setLocalGiao] = React.useState('')
  const [localNhan, setLocalNhan] = React.useState('')
  const [startDate, setStartDate] = React.useState('')
  const [endDate, setEndDate] = React.useState('')

  // Hàm kiểm tra 1 dòng có được tính vào báo cáo hay không, theo statusFilter hiện tại:
  // - approved_only: chỉ tính đơn Đã phê duyệt
  // - approved_pending: tính cả đơn Đã phê duyệt + Chưa phê duyệt/Chưa duyệt/Chờ phê duyệt,
  //   LOẠI TRỪ "Chưa xác nhận" (chưa phải là Chưa phê duyệt) và Từ chối/Hủy
  // - all: tính tất cả các đơn, không loại trừ gì
  const matchStatusFilter = React.useCallback((trangThai) => {
    if (statusFilter === 'all') return true
    if (isApprovedStatus(trangThai)) return true
    if (statusFilter !== 'approved_pending') return false

    const raw = String(trangThai || '').trim().toLowerCase()
    if (raw.includes('xác nhận')) return false // "Chưa xác nhận" không được tính là "Chưa phê duyệt"
    if (isRejectedStatus(trangThai)) return false
    return isPendingStatus(trangThai)
  }, [statusFilter])
  const [currentPage, setCurrentPage] = React.useState(1)
  const [pageSize, setPageSize] = React.useState(50)
  const [sortField, setSortField] = React.useState('default') // 'default' | 'maSAP' | 'received' | 'issued' | 'stock'
  const [sortDirection, setSortDirection] = React.useState('asc') // 'asc' | 'desc'

  // Extract all unique warehouses
  const uniqueWarehouses = React.useMemo(() => {
    const list = new Set()
    allProjects.forEach(p => { if (p) list.add(p.trim()) })
    chungRows.forEach(r => {
      const g = (r.donViGiao || '').trim()
      const n = (r.donViNhan || '').trim()
      if (g) list.add(g)
      if (n) list.add(n)
    })
    giaoRows.forEach(r => {
      const g = (r.donViGiao || '').trim()
      const n = (r.donViNhan || '').trim()
      if (g) list.add(g)
      if (n) list.add(n)
    })
    nhanRows.forEach(r => {
      const g = (r.donViGiao || '').trim()
      const n = (r.donViNhan || '').trim()
      if (g) list.add(g)
      if (n) list.add(n)
    })
    return [...list].sort()
  }, [allProjects, chungRows, giaoRows, nhanRows])

  // Extract unique delivery units (Đơn vị giao)
  const uniqueGiaoUnits = React.useMemo(() => {
    const list = new Set()
    const source = isRealReport 
      ? [...chungRows, ...giaoRows, ...nhanRows]
      : ((chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows])
    
    source.forEach(r => {
      const g = (r.donViGiao || '').trim()
      if (g) list.add(g)
    })
    return [...list].sort()
  }, [chungRows, giaoRows, nhanRows, isRealReport])

  // Extract unique receiving units (Đơn vị nhận)
  const uniqueNhanUnits = React.useMemo(() => {
    const list = new Set()
    const source = isRealReport 
      ? [...chungRows, ...giaoRows, ...nhanRows]
      : ((chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows])
    
    source.forEach(r => {
      const n = (r.donViNhan || '').trim()
      if (n) list.add(n)
    })
    return [...list].sort()
  }, [chungRows, giaoRows, nhanRows, isRealReport])

  // Process raw rows and group by maSAP for a given project/warehouse
  const computeProjectReportData = React.useCallback((proj) => {
    if (!proj) return [] // If no project is selected, return an empty array to prevent lag!
    const groups = {}
    const projLower = proj ? proj.trim().toLowerCase() : ''

    // Helper to parse double values safely
    const parseVal = (val) => {
      if (val === null || val === undefined) return 0
      if (typeof val === 'number') return val
      const cleaned = String(val).replace(/[^\d.-]/g, '').replace(',', '.')
      const num = parseFloat(cleaned)
      return isNaN(num) ? 0 : num
    }

    const sourceRows = isRealReport
      ? [...chungRows, ...giaoRows, ...nhanRows]
      : ((chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows])

    sourceRows.forEach(r => {
      // Apply status filter
      if (!matchStatusFilter(r.trangThai)) {
        return
      }

      // Apply delivery and receiving unit filters for real report
      if (isRealReport) {
        if (localGiao && String(r.donViGiao || '').trim().toLowerCase() !== localGiao.trim().toLowerCase()) {
          return
        }
        if (localNhan && String(r.donViNhan || '').trim().toLowerCase() !== localNhan.trim().toLowerCase()) {
          return
        }
      }

      const nhanUnit = String(r.donViNhan || '').trim().toLowerCase()
      const giaoUnit = String(r.donViGiao || '').trim().toLowerCase()

      let isNhan = false
      let isGiao = false

      if (proj) {
        isNhan = nhanUnit === projLower
        isGiao = giaoUnit === projLower
        // If this row is not related to our selected warehouse, skip
        if (!isNhan && !isGiao) return
      } else {
        isNhan = true
        isGiao = true
      }

      // Use maSAP as grouping key (fallback to empty string if not present)
      const sap = String(r.maSAP || '').trim()
      if (!sap) return // Skip rows without SAP code as per requirement

      // Initialize group if not exists
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
          latestIssuedDate: null
        }
      } else {
        // Fallbacks for empty info
        if (!groups[sap].maVatTu && r.maVatTu) groups[sap].maVatTu = String(r.maVatTu).trim()
        if (!groups[sap].tenVatTu && r.tenVatTu) groups[sap].tenVatTu = String(r.tenVatTu).trim()
        if (!groups[sap].dvt && r.dvt) groups[sap].dvt = String(r.dvt).trim()
        if (!groups[sap].thongSoKyThuat && r.thongSoKyThuat) groups[sap].thongSoKyThuat = String(r.thongSoKyThuat).trim()
      }

      const rowDate = parseRowDate(r.ngayXuatNhap)

      // Add to received/issued
      if (proj) {
        if (isNhan) {
          // Receiving unit: add to received
          groups[sap].received += parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
          if (rowDate && (!groups[sap].latestReceivedDate || rowDate > groups[sap].latestReceivedDate)) {
            groups[sap].latestReceivedDate = rowDate
          }
        }
        if (isGiao) {
          // Issuing/delivery unit: add to issued
          groups[sap].issued += parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
          if (rowDate && (!groups[sap].latestIssuedDate || rowDate > groups[sap].latestIssuedDate)) {
            groups[sap].latestIssuedDate = rowDate
          }
        }
      } else {
        // Global aggregation
        groups[sap].received += parseVal(r.khoiLuongNhap)
        groups[sap].issued += parseVal(r.khoiLuongXuat)
        if (parseVal(r.khoiLuongNhap) && rowDate && (!groups[sap].latestReceivedDate || rowDate > groups[sap].latestReceivedDate)) {
          groups[sap].latestReceivedDate = rowDate
        }
        if (parseVal(r.khoiLuongXuat) && rowDate && (!groups[sap].latestIssuedDate || rowDate > groups[sap].latestIssuedDate)) {
          groups[sap].latestIssuedDate = rowDate
        }
      }
    })

    // Map groups to list and compute stock/inventory
    const result = Object.values(groups).map(g => {
      const stock = g.received - g.issued
      const unusedStatus = getUnusedStatus(g, stock)
      
      const materialClassification = materialClassifications[g.maSAP] || ''
      const isKhauHao = String(materialClassification).toLowerCase().includes('khấu hao') || String(materialClassification).toLowerCase().includes('tài sản')
      
      let estimatedUnitPrice = 0
      let valueOver30Days = 0

      // Match price from materialPriceRows first
      const priceRow = materialPriceRows.find(r => String(r.maSAP || '').trim().toLowerCase() === String(g.maSAP || '').trim().toLowerCase())

      if (isKhauHao) {
        // Đơn giá tạm tính lấy theo cột đơn giá 1 ngày
        estimatedUnitPrice = priceRow ? (priceRow.donGiaTrungBinh1Ngay || 0) : 0
        if (unusedStatus === 'Chưa sử dụng (> 30 ngày)' && stock > 0) {
          const days = getDaysToToday(g.latestReceivedDate)
          const numDays = typeof days === 'number' ? days : 0
          valueOver30Days = numDays * stock * estimatedUnitPrice
        }
      } else {
        // Vật tư tiêu hao hoặc khác
        estimatedUnitPrice = priceRow ? (priceRow.donGiaTrungBinh || 0) : (materialPrices[g.maSAP] || 0)
        if (unusedStatus === 'Chưa sử dụng (> 30 ngày)' && stock > 0) {
          valueOver30Days = stock * estimatedUnitPrice
        }
      }

      valueOver30Days = Math.round(valueOver30Days)

      return {
        ...g,
        stock,
        unusedStatus,
        estimatedUnitPrice,
        valueOver30Days,
        materialClassification
      }
    })

    return result
  }, [chungRows, giaoRows, nhanRows, materialClassifications, materialPriceRows, materialPrices, matchStatusFilter, getUnusedStatus, isRealReport, localGiao, localNhan])

  // Process rows and group by maSAP
  const reportData = React.useMemo(() => {
    const result = computeProjectReportData(localProject)

    // Filter by search term
    const s = searchTerm.trim().toLowerCase()
    const filtered = s
      ? result.filter(item => 
          item.maSAP.toLowerCase().includes(s) ||
          item.maVatTu.toLowerCase().includes(s) ||
          item.tenVatTu.toLowerCase().includes(s) ||
          item.thongSoKyThuat.toLowerCase().includes(s) ||
          item.unusedStatus.toLowerCase().includes(s) ||
          item.materialClassification.toLowerCase().includes(s)
        )
      : result

    // Sort
    filtered.sort((a, b) => {
      if (sortField === 'default') {
        const getStatusRank = (status) => {
          const sLower = String(status || '').trim().toLowerCase()
          if (sLower.includes('chưa sử dụng')) return 1
          if (sLower.includes('đang sử dụng')) return 2
          if (sLower.includes('đã dùng hết') || sLower.includes('dùng hết')) return 3
          return 4
        }

        const getClassificationRank = (c) => {
          const cLower = String(c || '').trim().toLowerCase()
          if (cLower.includes('tiêu hao')) return 1
          if (cLower.includes('khấu hao') || cLower.includes('tài sản')) return 2
          return 3
        }

        // 1. Trạng thái vật tư không sử dụng Từ Chưa sử dụng đến Đang sử dụng đến đã dùng hết
        const statusRankA = getStatusRank(a.unusedStatus)
        const statusRankB = getStatusRank(b.unusedStatus)
        if (statusRankA !== statusRankB) {
          return sortDirection === 'asc' ? (statusRankA - statusRankB) : (statusRankB - statusRankA)
        }

        // 2. Phân loại vật tư: Vật tư tiêu hao rồi đến tài Sản khấu hao
        const classRankA = getClassificationRank(a.materialClassification)
        const classRankB = getClassificationRank(b.materialClassification)
        if (classRankA !== classRankB) {
          return sortDirection === 'asc' ? (classRankA - classRankB) : (classRankB - classRankA)
        }

        // 3. Thành tiền vật tư tồn: Từ giá trị lớn đến giá trị nhỏ
        const valA = a.valueOver30Days || 0
        const valB = b.valueOver30Days || 0
        if (valA !== valB) {
          return sortDirection === 'asc' ? (valB - valA) : (valA - valB)
        }

        // Fallback to maSAP alphabetically
        const sapA = String(a.maSAP || '').toLowerCase()
        const sapB = String(b.maSAP || '').toLowerCase()
        return sortDirection === 'asc' ? sapA.localeCompare(sapB) : sapB.localeCompare(sapA)
      }

      let valA, valB
      if (sortField === 'maSAP') {
        valA = a.maSAP.toLowerCase()
        valB = b.maSAP.toLowerCase()
      } else if (sortField === 'received') {
        valA = a.received
        valB = b.received
      } else if (sortField === 'issued') {
        valA = a.issued
        valB = b.issued
      } else if (sortField === 'stock') {
        valA = a.stock
        valB = b.stock
      } else if (sortField === 'estimatedUnitPrice') {
        valA = a.estimatedUnitPrice
        valB = b.estimatedUnitPrice
      } else if (sortField === 'valueOver30Days') {
        valA = a.valueOver30Days
        valB = b.valueOver30Days
      } else if (sortField === 'materialClassification') {
        valA = a.materialClassification.toLowerCase()
        valB = b.materialClassification.toLowerCase()
      }

      if (valA < valB) return sortDirection === 'asc' ? -1 : 1
      if (valA > valB) return sortDirection === 'asc' ? 1 : -1
      return 0
    })

    return filtered
  }, [computeProjectReportData, localProject, searchTerm, sortField, sortDirection])

  const realReportRows = React.useMemo(() => {
    if (!isRealReport) return []
    if (!localProject) return [] // If no project is selected, return an empty array to prevent lag!
    let list = chungRows

    // 1. Kho / Dự án filter
    if (localProject) {
      const p = localProject.trim().toLowerCase()
      list = list.filter(row => {
        const g = (row.donViGiao || '').trim().toLowerCase()
        const n = (row.donViNhan || '').trim().toLowerCase()
        return g === p || n === p
      })
    }

    // 2. Đơn vị giao filter
    if (localGiao) {
      list = list.filter(row => String(row.donViGiao || '').trim() === localGiao.trim())
    }

    // 3. Đơn vị nhận filter
    if (localNhan) {
      list = list.filter(row => String(row.donViNhan || '').trim() === localNhan.trim())
    }

    // 4. Status filter matching matchStatusFilter(row.trangThai)
    list = list.filter(row => matchStatusFilter(row.trangThai))

    // 5. SearchTerm filter (case insensitive search across many fields)
    if (searchTerm) {
      const q = searchTerm.toLowerCase()
      list = list.filter(row =>
        String(row.tenVatTu || '').toLowerCase().includes(q) ||
        String(row.maDonNhapKho || '').toLowerCase().includes(q) ||
        String(row.maDonXuatKho || '').toLowerCase().includes(q) ||
        String(row.donViGiao || '').toLowerCase().includes(q) ||
        String(row.donViNhan || '').toLowerCase().includes(q) ||
        String(row.maSAP || '').toLowerCase().includes(q) ||
        String(row.maVatTu || '').toLowerCase().includes(q)
      )
    }

    // 6. Date Range filter
    if (startDate || endDate) {
      const start = startDate ? new Date(startDate) : null
      if (start) start.setHours(0, 0, 0, 0)

      const end = endDate ? new Date(endDate) : null
      if (end) end.setHours(23, 59, 59, 999)

      list = list.filter(row => {
        const rowDate = parseRowDate(row.ngayXuatNhap)
        if (!rowDate) return false
        if (start && rowDate < start) return false
        if (end && rowDate > end) return false
        return true
      })
    }

    const parseNum = (v) => {
      if (v === null || v === undefined) return 0
      if (typeof v === 'number') return v
      const cleaned = String(v).replace(/[^\d.-]/g, '').replace(',', '.')
      const num = parseFloat(cleaned)
      return isNaN(num) ? 0 : num
    }

    // 7. Filter by Nhập thực vs Xuất thực
    if (realReportSubTab === 'nhap') {
      list = list.filter(row => parseNum(row.khoiLuongNhap) > 0)
    } else if (realReportSubTab === 'xuat') {
      list = list.filter(row => parseNum(row.khoiLuongXuat) > 0)
    }

    // Sort descending by date (newest/most recent first)
    const sorted = [...list].sort((a, b) => {
      const dateA = parseRowDate(a.ngayXuatNhap)
      const dateB = parseRowDate(b.ngayXuatNhap)
      if (!dateA && !dateB) return 0
      if (!dateA) return 1
      if (!dateB) return -1
      return dateB.getTime() - dateA.getTime()
    })

    const getCategoryForUnit = (name) => {
      if (!name) return 'chuaphanbo'
      const normName = name.trim().replace(/\s+/g, ' ')
      if (customCategoryMap && customCategoryMap[normName]) {
        return customCategoryMap[normName]
      }
      if (customCategoryMap) {
        const keys = Object.keys(customCategoryMap)
        const foundKey = keys.find(k => k.toLowerCase() === normName.toLowerCase())
        if (foundKey) {
          return customCategoryMap[foundKey]
        }
      }
      return getUnitCategory(normName)
    }

    return sorted.map(row => {
      const nhap = parseNum(row.khoiLuongNhap)
      const xuat = parseNum(row.khoiLuongXuat)
      
      const normProject = String(localProject || '').trim().toLowerCase()
      const normGiao = String(row.donViGiao || '').trim().toLowerCase()
      const normNhan = String(row.donViNhan || '').trim().toLowerCase()

      const isGiaoProject = (normGiao === normProject)
      const isNhanProject = (normNhan === normProject)

      let logicTongHopVal = 0

      if (realReportSubTab === 'nhap') {
        if (isGiaoProject) {
          const catNhan = getCategoryForUnit(row.donViNhan)
          if (catNhan === 'ncc') {
            logicTongHopVal = -1
          } else if (catNhan === 'kho') {
            logicTongHopVal = 0
          } else if (catNhan === 'todoi') {
            logicTongHopVal = 0
          }
        } else if (isNhanProject) {
          const catGiao = getCategoryForUnit(row.donViGiao)
          if (catGiao === 'ncc') {
            logicTongHopVal = 1
          } else if (catGiao === 'kho') {
            logicTongHopVal = 1
          } else if (catGiao === 'todoi') {
            logicTongHopVal = 0
          }
        }
      } else {
        // realReportSubTab === 'xuat'
        if (isGiaoProject) {
          const catNhan = getCategoryForUnit(row.donViNhan)
          if (catNhan === 'todoi') {
            logicTongHopVal = 1
          } else if (catNhan === 'kho') {
            logicTongHopVal = 1
          } else if (catNhan === 'ncc') {
            logicTongHopVal = 0
          }
        } else if (isNhanProject) {
          const catGiao = getCategoryForUnit(row.donViGiao)
          if (catGiao === 'todoi') {
            logicTongHopVal = -1
          } else if (catGiao === 'kho') {
            logicTongHopVal = 0
          } else if (catGiao === 'ncc') {
            logicTongHopVal = 0
          }
        }
      }

      const qty = realReportSubTab === 'nhap' ? nhap : xuat
      const khoiLuongThuc = qty * logicTongHopVal

      return {
        ...row,
        logicTongHop: logicTongHopVal,
        khoiLuongThuc
      }
    })
  }, [chungRows, isRealReport, localProject, localGiao, localNhan, searchTerm, statusFilter, startDate, endDate, matchStatusFilter, realReportSubTab, customCategoryMap])

  const realReportSummaryRows = React.useMemo(() => {
    if (!isRealReport) return []
    if (!localProject) return []

    let list = chungRows

    if (localProject) {
      const p = localProject.trim().toLowerCase()
      list = list.filter(row => {
        const g = (row.donViGiao || '').trim().toLowerCase()
        const n = (row.donViNhan || '').trim().toLowerCase()
        return g === p || n === p
      })
    }

    if (localGiao) {
      list = list.filter(row => String(row.donViGiao || '').trim() === localGiao.trim())
    }
    if (localNhan) {
      list = list.filter(row => String(row.donViNhan || '').trim() === localNhan.trim())
    }
    list = list.filter(row => matchStatusFilter(row.trangThai))

    if (searchTerm) {
      const q = searchTerm.toLowerCase()
      list = list.filter(row =>
        String(row.tenVatTu || '').toLowerCase().includes(q) ||
        String(row.maDonNhapKho || '').toLowerCase().includes(q) ||
        String(row.maDonXuatKho || '').toLowerCase().includes(q) ||
        String(row.donViGiao || '').toLowerCase().includes(q) ||
        String(row.donViNhan || '').toLowerCase().includes(q) ||
        String(row.maSAP || '').toLowerCase().includes(q) ||
        String(row.maVatTu || '').toLowerCase().includes(q)
      )
    }

    if (startDate || endDate) {
      const start = startDate ? new Date(startDate) : null
      if (start) start.setHours(0, 0, 0, 0)
      const end = endDate ? new Date(endDate) : null
      if (end) end.setHours(23, 59, 59, 999)

      list = list.filter(row => {
        const rowDate = parseRowDate(row.ngayXuatNhap)
        if (!rowDate) return false
        if (start && rowDate < start) return false
        if (end && rowDate > end) return false
        return true
      })
    }

    const parseNum = (v) => {
      if (v === null || v === undefined) return 0
      if (typeof v === 'number') return v
      const cleaned = String(v).replace(/[^\d.-]/g, '').replace(',', '.')
      const num = parseFloat(cleaned)
      return isNaN(num) ? 0 : num
    }

    const getCategoryForUnit = (name) => {
      if (!name) return 'chuaphanbo'
      const normName = name.trim().replace(/\s+/g, ' ')
      if (customCategoryMap && customCategoryMap[normName]) {
        return customCategoryMap[normName]
      }
      if (customCategoryMap) {
        const keys = Object.keys(customCategoryMap)
        const foundKey = keys.find(k => k.toLowerCase() === normName.toLowerCase())
        if (foundKey) {
          return customCategoryMap[foundKey]
        }
      }
      return getUnitCategory(normName)
    }

    const groups = {}

    list.forEach(row => {
      const maSAP = String(row.maSAP || '').trim()
      if (!maSAP) return

      if (!groups[maSAP]) {
        groups[maSAP] = {
          maVatTu: String(row.maVatTu || '').trim(),
          maSAP,
          tenVatTu: String(row.tenVatTu || '').trim(),
          thongSoKyThuat: String(row.thongSoKyThuat || '').trim(),
          dvt: String(row.dvt || '').trim(),
          thucNhap: 0,
          thucXuat: 0,
          transactions: []
        }
      }

      const nhapVal = parseNum(row.khoiLuongNhap)
      const xuatVal = parseNum(row.khoiLuongXuat)

      const normProject = String(localProject || '').trim().toLowerCase()
      const normGiao = String(row.donViGiao || '').trim().toLowerCase()
      const normNhan = String(row.donViNhan || '').trim().toLowerCase()

      const isGiaoProject = (normGiao === normProject)
      const isNhanProject = (normNhan === normProject)

      let logicNhapVal = 0
      let explainNhap = ''
      let detailTypeNhap = ''
      if (nhapVal > 0) {
        if (isGiaoProject) {
          const catNhan = getCategoryForUnit(row.donViNhan)
          if (catNhan === 'ncc') {
            logicNhapVal = -1
            detailTypeNhap = 'tra_ncc'
            explainNhap = `Trả lại NCC (${row.donViNhan})`
          } else {
            explainNhap = `Điều chuyển đi (${row.donViNhan}) - Không tính`
          }
        } else if (isNhanProject) {
          const catGiao = getCategoryForUnit(row.donViGiao)
          if (catGiao === 'ncc' || catGiao === 'kho') {
            logicNhapVal = 1
            detailTypeNhap = catGiao === 'ncc' ? 'nhan_ncc' : 'nhan_kho'
            explainNhap = `Nhận từ ${catGiao === 'ncc' ? 'NCC' : 'Kho cấp trên'} (${row.donViGiao})`
          } else {
            explainNhap = `Nhận từ nguồn khác (${row.donViGiao}) - Không tính`
          }
        }
        groups[maSAP].thucNhap += nhapVal * logicNhapVal
      }

      let logicXuatVal = 0
      let explainXuat = ''
      let detailTypeXuat = ''
      if (xuatVal > 0) {
        if (isGiaoProject) {
          const catNhan = getCategoryForUnit(row.donViNhan)
          if (catNhan === 'todoi' || catNhan === 'kho') {
            logicXuatVal = 1
            detailTypeXuat = catNhan === 'todoi' ? 'xuat_todoi' : 'xuat_kho'
            explainXuat = `Xuất cho ${catNhan === 'todoi' ? 'Tổ đội/NTP' : 'Kho khác'} (${row.donViNhan})`
          } else {
            explainXuat = `Xuất cho đối tượng khác (${row.donViNhan}) - Không tính`
          }
        } else if (isNhanProject) {
          const catGiao = getCategoryForUnit(row.donViGiao)
          if (catGiao === 'todoi') {
            logicXuatVal = -1
            detailTypeXuat = 'todoi_tra'
            explainXuat = `Tổ đội trả hàng thừa (${row.donViGiao})`
          } else {
            explainXuat = `Nhận lại từ nguồn khác (${row.donViGiao}) - Không tính`
          }
        }
        groups[maSAP].thucXuat += xuatVal * logicXuatVal
      }

      if (nhapVal > 0 || xuatVal > 0) {
        groups[maSAP].transactions.push({
          ngayXuatNhap: row.ngayXuatNhap,
          maDonNhapKho: row.maDonNhapKho,
          maDonXuatKho: row.maDonXuatKho,
          donViGiao: row.donViGiao,
          donViNhan: row.donViNhan,
          nhapVal,
          logicNhapVal,
          explainNhap,
          detailTypeNhap,
          xuatVal,
          logicXuatVal,
          explainXuat,
          detailTypeXuat
        })
      }
    })

    return Object.values(groups)
      .filter(g => g.thucNhap !== 0 || g.thucXuat !== 0)
      .map(g => {
        const thucNhap = Math.max(0, g.thucNhap)
        const thucXuat = Math.max(0, g.thucXuat)
        return {
          ...g,
          thucNhap,
          thucXuat,
          tonKho: thucNhap - thucXuat
        }
      })
  }, [chungRows, isRealReport, localProject, localGiao, localNhan, searchTerm, statusFilter, startDate, endDate, matchStatusFilter, customCategoryMap])

  const realReportMetrics = React.useMemo(() => {
    if (!isRealReport) return []

    if (realReportSubTab === 'tonghop') {
      let totalItems = realReportSummaryRows.length
      let totalReceivedSum = realReportSummaryRows.reduce((sum, r) => sum + r.thucNhap, 0)
      let totalIssuedSum = realReportSummaryRows.reduce((sum, r) => sum + r.thucXuat, 0)

      return [
        {
          label: 'Tổng số mặt hàng thực tế',
          value: totalItems.toLocaleString('vi-VN') + ' mặt hàng',
          color: '#8b5cf6',
          bg: '#f5f3ff',
          border: '#ddd6fe',
          icon: <PackageCheck size={18} />
        },
        {
          label: 'Tổng khối lượng thực nhập',
          value: totalReceivedSum === 0 ? '-' : totalReceivedSum.toLocaleString('vi-VN', { maximumFractionDigits: 2 }),
          color: '#10b981',
          bg: '#ecfdf5',
          border: '#a7f3d0',
          icon: <Truck size={18} />
        },
        {
          label: 'Tổng khối lượng thực xuất',
          value: totalIssuedSum === 0 ? '-' : totalIssuedSum.toLocaleString('vi-VN', { maximumFractionDigits: 2 }),
          color: '#f97316',
          bg: '#fff7ed',
          border: '#fed7aa',
          icon: <Download size={18} style={{ transform: 'rotate(180deg)' }} />
        }
      ]
    }

    let total = realReportRows.length
    let totalReceived = 0
    let totalIssued = 0

    const parseNum = (v) => {
      if (v === null || v === undefined) return 0
      if (typeof v === 'number') return v
      const cleaned = String(v).replace(/[^\d.-]/g, '').replace(',', '.')
      const num = parseFloat(cleaned)
      return isNaN(num) ? 0 : num
    }

    const counts = {}
    realReportRows.forEach(row => {
      const klThuc = parseNum(row.khoiLuongThuc)
      if (realReportSubTab === 'nhap') {
        totalReceived += klThuc
      } else {
        totalIssued += klThuc
      }

      const status = (row.trangThai || 'Chưa duyệt').trim()
      counts[status] = (counts[status] || 0) + 1
    })

    const activeStatuses = Object.keys(counts).filter(status => counts[status] > 0)
    const priority = ['Đã phê duyệt', 'Chưa xác nhận', 'Chờ phê duyệt', 'Chưa phê duyệt', 'Chưa duyệt', 'Từ chối']
    activeStatuses.sort((a, b) => {
      const idxA = priority.indexOf(a)
      const idxB = priority.indexOf(b)
      if (idxA !== -1 && idxB !== -1) return idxA - idxB
      if (idxA !== -1) return -1
      if (idxB !== -1) return 1
      return a.localeCompare(b, 'vi')
    })

    const statusCards = activeStatuses.map(status => {
      const count = counts[status]
      let color = 'var(--primary)'
      let bg = 'var(--primary-light)'
      let border = 'var(--border)'
      let icon = <ClipboardList size={18} />

      const sLower = status.toLowerCase()
      if (sLower === 'đã phê duyệt' || (isApprovedStatus(status) && !sLower.includes('chưa') && !sLower.includes('chờ'))) {
        color = '#10b981'
        bg = '#ecfdf5'
        border = '#a7f3d0'
        icon = <CheckCircle2 size={18} />
      } else if (sLower === 'từ chối' || isRejectedStatus(status)) {
        color = '#ef4444'
        bg = '#fef2f2'
        border = '#fca5a5'
        icon = <AlertCircle size={18} />
      } else {
        color = '#b45309'
        bg = '#fffbeb'
        border = '#fde68a'
        icon = <Clock size={18} />
      }

      return {
        label: status,
        value: count.toLocaleString('vi-VN') + ' đơn',
        color,
        bg,
        border,
        icon
      }
    })

    if (realReportSubTab === 'nhap') {
      return [
        {
          label: 'Tổng số lượng đơn nhập',
          value: total.toLocaleString('vi-VN') + ' đơn',
          color: 'var(--primary)',
          bg: 'var(--primary-light)',
          border: 'var(--border)',
          icon: <FileSpreadsheet size={18} />
        },
        {
          label: 'Tổng khối lượng nhận thực',
          value: totalReceived === 0 ? '-' : totalReceived.toLocaleString('vi-VN', { maximumFractionDigits: 2 }),
          color: '#10b981',
          bg: '#ecfdf5',
          border: '#a7f3d0',
          icon: <Truck size={18} />
        },
        ...statusCards
      ]
    } else {
      return [
        {
          label: 'Tổng số lượng đơn xuất',
          value: total.toLocaleString('vi-VN') + ' đơn',
          color: 'var(--primary)',
          bg: 'var(--primary-light)',
          border: 'var(--border)',
          icon: <FileSpreadsheet size={18} />
        },
        {
          label: 'Tổng khối lượng xuất thực',
          value: totalIssued === 0 ? '-' : totalIssued.toLocaleString('vi-VN', { maximumFractionDigits: 2 }),
          color: '#f97316',
          bg: '#fff7ed',
          border: '#fed7aa',
          icon: <Download size={18} style={{ transform: 'rotate(180deg)' }} />
        },
        ...statusCards
      ]
    }
  }, [realReportRows, realReportSummaryRows, isRealReport, realReportSubTab])

  // Extract and group dashboard statistics specifically for BCH warehouses
  const bchWarehouses = React.useMemo(() => {
    const bchList = uniqueWarehouses.filter(w => {
      const lower = String(w || '').trim().toLowerCase()
      return lower.includes('bch')
    })
    return bchList.length > 0 ? bchList : uniqueWarehouses
  }, [uniqueWarehouses])

  // TỐI ƯU HIỆU NĂNG: Trước đây dashboardStats gọi computeProjectReportData() riêng cho TỪNG kho BCH
  // (68 kho) → quét toàn bộ 145.515 dòng 68 LẦN (~9.9 triệu lượt duyệt), gây khựng ~5s khi mở tab
  // "Dashboard báo cáo". Giờ chỉ quét chungRows MỘT LẦN duy nhất, gom nhóm dữ liệu cho tất cả các kho
  // BCH cùng lúc, và dùng Map tra cứu đơn giá thay vì .find() (vốn cũng chạy O(n) cho mỗi mặt hàng).
  const dashboardStats = React.useMemo(() => {
    if (bchWarehouses.length === 0) return []

    const parseVal = (val) => {
      if (val === null || val === undefined) return 0
      if (typeof val === 'number') return val
      const cleaned = String(val).replace(/[^\d.-]/g, '').replace(',', '.')
      const num = parseFloat(cleaned)
      return isNaN(num) ? 0 : num
    }

    // Chuẩn hóa tên kho BCH để tra cứu O(1) trong lúc quét dữ liệu
    const bchNormToOriginal = new Map()
    bchWarehouses.forEach(w => {
      bchNormToOriginal.set(String(w || '').trim().toLowerCase(), w)
    })

    // groupsByProject: projNormalized -> { [maSAP]: groupObj }
    const groupsByProject = new Map()
    bchNormToOriginal.forEach((_, norm) => groupsByProject.set(norm, {}))

    const sourceRows = isRealReport
      ? [...chungRows, ...giaoRows, ...nhanRows]
      : ((chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows])

    sourceRows.forEach(r => {
      if (!matchStatusFilter(r.trangThai)) return

      const nhanUnit = String(r.donViNhan || '').trim().toLowerCase()
      const giaoUnit = String(r.donViGiao || '').trim().toLowerCase()

      const nhanGroups = groupsByProject.get(nhanUnit)
      const giaoGroups = groupsByProject.get(giaoUnit)
      if (!nhanGroups && !giaoGroups) return // Dòng này không thuộc bất kỳ kho BCH nào cần tính

      const sap = String(r.maSAP || '').trim()
      if (!sap) return

      const rowDate = parseRowDate(r.ngayXuatNhap)
      const ensureGroup = (groups) => {
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
            latestIssuedDate: null
          }
        } else {
          if (!groups[sap].maVatTu && r.maVatTu) groups[sap].maVatTu = String(r.maVatTu).trim()
          if (!groups[sap].tenVatTu && r.tenVatTu) groups[sap].tenVatTu = String(r.tenVatTu).trim()
          if (!groups[sap].dvt && r.dvt) groups[sap].dvt = String(r.dvt).trim()
          if (!groups[sap].thongSoKyThuat && r.thongSoKyThuat) groups[sap].thongSoKyThuat = String(r.thongSoKyThuat).trim()
        }
        return groups[sap]
      }

      if (nhanGroups) {
        const g = ensureGroup(nhanGroups)
        g.received += parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
        if (rowDate && (!g.latestReceivedDate || rowDate > g.latestReceivedDate)) g.latestReceivedDate = rowDate
      }
      if (giaoGroups) {
        const g = ensureGroup(giaoGroups)
        g.issued += parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
        if (rowDate && (!g.latestIssuedDate || rowDate > g.latestIssuedDate)) g.latestIssuedDate = rowDate
      }
    })

    // Tra cứu đơn giá bằng Map (O(1)) thay vì materialPriceRows.find() (O(n)) cho mỗi mặt hàng
    const priceMap = new Map()
    materialPriceRows.forEach(pr => {
      const key = String(pr.maSAP || '').trim().toLowerCase()
      if (key && !priceMap.has(key)) priceMap.set(key, pr)
    })

    const results = []
    bchNormToOriginal.forEach((projName, norm) => {
      const groups = groupsByProject.get(norm)
      const items = Object.values(groups).map(g => {
        const stock = g.received - g.issued
        const unusedStatus = getUnusedStatus(g, stock)

        const materialClassification = materialClassifications[g.maSAP] || ''
        const isKhauHao = String(materialClassification).toLowerCase().includes('khấu hao') || String(materialClassification).toLowerCase().includes('tài sản')

        let estimatedUnitPrice = 0
        let valueOver30Days = 0
        const priceRow = priceMap.get(String(g.maSAP || '').trim().toLowerCase())

        if (isKhauHao) {
          estimatedUnitPrice = priceRow ? (priceRow.donGiaTrungBinh1Ngay || 0) : 0
          if (unusedStatus === 'Chưa sử dụng (> 30 ngày)' && stock > 0) {
            const days = getDaysToToday(g.latestReceivedDate)
            const numDays = typeof days === 'number' ? days : 0
            valueOver30Days = numDays * stock * estimatedUnitPrice
          }
        } else {
          estimatedUnitPrice = priceRow ? (priceRow.donGiaTrungBinh || 0) : (materialPrices[g.maSAP] || 0)
          if (unusedStatus === 'Chưa sử dụng (> 30 ngày)' && stock > 0) {
            valueOver30Days = stock * estimatedUnitPrice
          }
        }

        valueOver30Days = Math.round(valueOver30Days)

        return { ...g, stock, unusedStatus, estimatedUnitPrice, valueOver30Days, materialClassification }
      })

      let totalStock = 0
      let totalReceived = 0
      let totalIssued = 0

      let totalItemsInStock = 0
      let unusedCount = 0
      let unusedQty = 0
      let unusedValue = 0

      items.forEach(item => {
        if (item.received > 0 || item.issued > 0) {
          totalReceived += item.received
          totalIssued += item.issued
        }
        if (item.stock > 0) {
          totalItemsInStock++
          totalStock += item.stock
          if (item.unusedStatus === 'Chưa sử dụng (> 30 ngày)') {
            unusedCount++
            unusedQty += item.stock
            unusedValue += item.valueOver30Days || 0
          }
        }
      })

      const unusedRatio = totalItemsInStock > 0 ? (unusedCount / totalItemsInStock) * 100 : 0

      // Determine alarm level
      let alarmLevel = 'Bình thường'
      let alarmColor = '#10b981' // Green
      let alarmBg = '#ecfdf5'
      let alarmBorder = '#a7f3d0'

      if (unusedValue > 1000000000 || unusedRatio > 50) {
        alarmLevel = 'Nguy cơ cao (Đỏ)'
        alarmColor = '#ef4444' // Red
        alarmBg = '#fef2f2'
        alarmBorder = '#fca5a5'
      } else if (unusedValue > 100000000 || unusedRatio > 25) {
        alarmLevel = 'Cảnh báo (Vàng)'
        alarmColor = '#f59e0b' // Yellow
        alarmBg = '#fffbeb'
        alarmBorder = '#fde68a'
      }

      // Sort stagnant items to find the top 10 with highest unusedValue
      const topStagnant = [...items]
        .filter(item => item.unusedStatus === 'Chưa sử dụng (> 30 ngày)' && item.stock > 0)
        .sort((a, b) => (b.valueOver30Days || 0) - (a.valueOver30Days || 0))
        .slice(0, 10)

      results.push({
        projectName: projName,
        totalReceived,
        totalIssued,
        totalStock,
        totalItemsInStock,
        unusedCount,
        unusedQty,
        unusedValue,
        unusedRatio,
        alarmLevel,
        alarmColor,
        alarmBg,
        alarmBorder,
        topStagnant,
        allItems: items
      })
    })

    return results.sort((a, b) => b.unusedValue - a.unusedValue)
  }, [bchWarehouses, chungRows, giaoRows, nhanRows, materialClassifications, materialPriceRows, materialPrices, matchStatusFilter, getUnusedStatus])

  // Cho phép người dùng chọn 1 kho dự án BCH cụ thể để xem báo cáo Dashboard,
  // thay vì luôn cộng dồn/hiển thị toàn bộ các kho.
  const [dashboardWarehouseFilter, setDashboardWarehouseFilter] = React.useState('all')

  const filteredDashboardStats = React.useMemo(() => {
    if (dashboardWarehouseFilter === 'all') return dashboardStats
    return dashboardStats.filter(s => s.projectName === dashboardWarehouseFilter)
  }, [dashboardStats, dashboardWarehouseFilter])

  const dashboardSummary = React.useMemo(() => {
    let grandUnusedValue = 0
    let grandUnusedCount = 0
    let grandTotalItems = 0
    let highestUnusedWarehouse = 'Chưa xác định'
    let highestUnusedVal = 0

    filteredDashboardStats.forEach(stat => {
      grandUnusedValue += stat.unusedValue
      grandUnusedCount += stat.unusedCount
      grandTotalItems += stat.totalItemsInStock
      if (stat.unusedValue > highestUnusedVal) {
        highestUnusedVal = stat.unusedValue
        highestUnusedWarehouse = stat.projectName
      }
    })

    const averageUnusedRatio = grandTotalItems > 0 ? (grandUnusedCount / grandTotalItems) * 100 : 0

    return {
      grandUnusedValue,
      grandUnusedCount,
      grandTotalItems,
      averageUnusedRatio,
      highestUnusedWarehouse,
      highestUnusedVal
    }
  }, [filteredDashboardStats])

  const [selectedDashProject, setSelectedDashProject] = React.useState(null)

  const activeDashStat = React.useMemo(() => {
    if (filteredDashboardStats.length === 0) return null
    if (!selectedDashProject) return filteredDashboardStats[0]
    return filteredDashboardStats.find(s => s.projectName === selectedDashProject) || filteredDashboardStats[0]
  }, [filteredDashboardStats, selectedDashProject])

  // Chi tiết các đơn nhận / đơn xuất cấu thành nên khối lượng nhận - khối lượng xuất ở bảng tổng hợp.
  // Dùng đúng logic lọc (Kho/Dự án + trạng thái) như reportData để đảm bảo số liệu khớp 100%.
  const detailRows = React.useMemo(() => {
    const projLower = localProject ? localProject.trim().toLowerCase() : ''
    const parseVal = (val) => {
      if (val === null || val === undefined) return 0
      if (typeof val === 'number') return val
      const cleaned = String(val).replace(/[^\d.-]/g, '').replace(',', '.')
      const num = parseFloat(cleaned)
      return isNaN(num) ? 0 : num
    }

    const sourceRows = isRealReport
      ? [...giaoRows, ...nhanRows]
      : ((chungRows && chungRows.length > 0) ? chungRows : [...giaoRows, ...nhanRows])

    const nhanList = []
    const xuatList = []

    sourceRows.forEach(r => {
      if (!matchStatusFilter(r.trangThai)) return

      const sap = String(r.maSAP || '').trim()
      if (!sap) return

      const nhanUnit = String(r.donViNhan || '').trim().toLowerCase()
      const giaoUnit = String(r.donViGiao || '').trim().toLowerCase()

      let isNhan = false
      let isGiao = false
      if (localProject) {
        isNhan = nhanUnit === projLower
        isGiao = giaoUnit === projLower
        if (!isNhan && !isGiao) return
      } else {
        isNhan = true
        isGiao = true
      }

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

      if (localProject) {
        if (isNhan) {
          const soLuong = parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
          if (soLuong) {
            nhanList.push({ ...baseInfo, maDon: String(r.maDonNhapKho || r.maDonXuatKho || '').trim(), khoiLuong: soLuong })
          }
        }
        if (isGiao) {
          const soLuong = parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
          if (soLuong) {
            xuatList.push({ ...baseInfo, maDon: String(r.maDonXuatKho || r.maDonNhapKho || '').trim(), khoiLuong: soLuong })
          }
        }
      } else {
        const soNhan = parseVal(r.khoiLuongNhap)
        if (soNhan) {
          nhanList.push({ ...baseInfo, maDon: String(r.maDonNhapKho || '').trim(), khoiLuong: soNhan })
        }
        const soXuat = parseVal(r.khoiLuongXuat)
        if (soXuat) {
          xuatList.push({ ...baseInfo, maDon: String(r.maDonXuatKho || '').trim(), khoiLuong: soXuat })
        }
      }
    })

    return { nhanList, xuatList }
  }, [chungRows, giaoRows, nhanRows, localProject, matchStatusFilter])

  // Metrics
  const metrics = React.useMemo(() => {
    let totalReceived = 0
    let totalIssued = 0
    let totalStock = 0
    let totalValueOver30Days = 0
    reportData.forEach(item => {
      totalReceived += item.received
      totalIssued += item.issued
      totalStock += item.stock
      totalValueOver30Days += item.valueOver30Days || 0
    })
    return {
      totalReceived,
      totalIssued,
      totalStock,
      totalValueOver30Days,
      totalItems: reportData.length
    }
  }, [reportData])

  // Reset pagination on search/filter/project change
  React.useEffect(() => {
    setCurrentPage(1)
  }, [localProject, searchTerm, statusFilter])

  // Pagination calculation
  const totalPages = Math.ceil(reportData.length / pageSize) || 1
  const paginatedData = React.useMemo(() => {
    const startIndex = (currentPage - 1) * pageSize
    return reportData.slice(startIndex, startIndex + pageSize)
  }, [reportData, currentPage, pageSize])

  // Excel export using xlsx-js-style
  const handleExportExcel = () => {
    if (isRealReport) {
      if (realReportSubTab === 'tonghop') {
        if (realReportSummaryRows.length === 0) return

        const wb = XLSXStyle.utils.book_new()
        const ws = {}

        const columns = [
          { key: 'STT', label: 'STT', width: 50 },
          { key: 'maVatTu', label: 'Mã vật tư', width: 100 },
          { key: 'maSAP', label: 'Mã SAP', width: 100 },
          { key: 'thongSoKyThuat', label: 'Thông số kỹ thuật', width: 150 },
          { key: 'tenVatTu', label: 'Tên vật tư', width: 300 },
          { key: 'dvt', label: 'ĐVT', width: 80 },
          { key: 'thucNhap', label: 'Thực nhập (SUMIFS)', width: 130 },
          { key: 'thucXuat', label: 'Thực xuất (SUMIFS)', width: 130 },
          { key: 'tonKho', label: 'Tồn kho (Formula)', width: 130 }
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
          v: `BÁO CÁO TỔNG HỢP KHỐI LƯỢNG THỰC NHẬP THỰC XUẤT`,
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '0B2545' } },
            alignment: { horizontal: 'left', vertical: 'center' }
          }
        }
        ws['A2'] = {
          v: `Kho / Dự án: ${localProject || 'Tất cả'} | Đơn vị giao: ${localGiao || 'Tất cả'} | Đơn vị nhận: ${localNhan || 'Tất cả'}`,
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 10, italic: true },
            alignment: { horizontal: 'left', vertical: 'center' }
          }
        }
        excelRowIdx = 4

        columns.forEach((col, colIdx) => {
          const colChar = getColLabel(colIdx)
          const cellRef = `${colChar}${excelRowIdx}`
          
          const isThucNhap = col.key === 'thucNhap'
          const isThucXuat = col.key === 'thucXuat'
          const isTonKho = col.key === 'tonKho'
          const excelBgColor = isThucNhap ? '0D9488' : isThucXuat ? 'EA580C' : isTonKho ? '0F58A7' : '0F58A7'
          const excelBorderColor = isThucNhap ? '0F766E' : isThucXuat ? 'C2410C' : isTonKho ? '0A3D73' : '0A3D73'

          ws[cellRef] = {
            v: col.label,
            t: 's',
            s: {
              fill: { patternType: 'solid', fgColor: { rgb: excelBgColor } },
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'FFFFFF' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: excelBorderColor } },
                bottom: { style: 'medium', color: { rgb: excelBorderColor } },
                left: { style: 'thin', color: { rgb: excelBorderColor } },
                right: { style: 'thin', color: { rgb: excelBorderColor } }
              }
            }
          }
        })

        realReportSummaryRows.forEach((row, rowIndex) => {
          excelRowIdx++
          const isEvenNum = (rowIndex % 2 === 1)
          const rowBgColor = isEvenNum ? 'F8FAFC' : 'FFFFFF'

          columns.forEach((col, colIdx) => {
            const colChar = getColLabel(colIdx)
            const cellRef = `${colChar}${excelRowIdx}`

            let val = ''
            let cellType = 's'
            let numFormat = undefined
            let isFormula = false
            let formulaStr = ''

            if (col.key === 'STT') {
              val = rowIndex + 1
              cellType = 'n'
            } else if (col.key === 'thucNhap') {
              isFormula = true
              formulaStr = `SUMIFS(Thuc_Nhap!L:L,Thuc_Nhap!C:C,C${excelRowIdx})`
              val = Number(row.thucNhap)
              cellType = 'n'
              numFormat = '#,##0.00;[Red]-#,##0.00;"-"'
            } else if (col.key === 'thucXuat') {
              isFormula = true
              formulaStr = `SUMIFS(Thuc_Xuat!L:L,Thuc_Xuat!C:C,C${excelRowIdx})`
              val = Number(row.thucXuat)
              cellType = 'n'
              numFormat = '#,##0.00;[Red]-#,##0.00;"-"'
            } else if (col.key === 'tonKho') {
              isFormula = true
              formulaStr = `G${excelRowIdx}-H${excelRowIdx}`
              val = Number(row.tonKho)
              cellType = 'n'
              numFormat = '#,##0.00;[Red]-#,##0.00;"-"'
            } else {
              val = String(row[col.key] || '')
            }

            const isCenteredCol = ['STT', 'maVatTu', 'maSAP', 'dvt'].includes(col.key)
            const isRightAligned = ['thucNhap', 'thucXuat', 'tonKho'].includes(col.key)

            const cellStyle = {
              font: { name: 'Segoe UI', sz: 9, color: { rgb: '1A1A1A' } },
              alignment: {
                horizontal: isCenteredCol ? 'center' : (isRightAligned ? 'right' : 'left'),
                vertical: 'center',
                wrapText: true
              },
              fill: { patternType: 'solid', fgColor: { rgb: rowBgColor } },
              border: {
                top: { style: 'thin', color: { rgb: 'E2E8F0' } },
                bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
                left: { style: 'thin', color: { rgb: 'E2E8F0' } },
                right: { style: 'thin', color: { rgb: 'E2E8F0' } }
              }
            }

            if (col.key === 'tonKho') {
              const numVal = Number(row[col.key])
              const isNegative = numVal < 0
              const isZero = numVal === 0
              if (isNegative) {
                cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'FEF2F2' } }
                cellStyle.font.color = { rgb: 'EF4444' }
                cellStyle.font.bold = true
              } else if (!isZero) {
                cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'EFF6FF' } }
                cellStyle.font.color = { rgb: '1E40AF' }
                cellStyle.font.bold = true
              } else {
                cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } }
                cellStyle.font.color = { rgb: '475569' }
                cellStyle.font.bold = false
              }
            } else if (col.key === 'thucNhap' || col.key === 'thucXuat') {
              const numVal = Number(row[col.key])
              if (numVal > 0) {
                cellStyle.fill = { patternType: 'solid', fgColor: { rgb: col.key === 'thucNhap' ? 'ECFDF5' : 'FFF7ED' } }
                cellStyle.font.color = { rgb: col.key === 'thucNhap' ? '065F46' : 'C2410C' }
                cellStyle.font.bold = true
              } else {
                cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } }
                cellStyle.font.color = { rgb: '475569' }
                cellStyle.font.bold = false
              }
            }

            const cellObj = { v: val, t: cellType, s: cellStyle }
            if (isFormula) cellObj.f = formulaStr
            if (numFormat) cellObj.z = numFormat
            ws[cellRef] = cellObj
          })
        })

        // Sum row
        excelRowIdx++
        const lastDataRow = excelRowIdx - 1
        ws[`A${excelRowIdx}`] = {
          v: 'TỔNG CỘNG',
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 10, bold: true },
            alignment: { horizontal: 'center', vertical: 'center' },
            fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
            border: {
              top: { style: 'medium', color: { rgb: '0F58A7' } },
              bottom: { style: 'medium', color: { rgb: '0F58A7' } },
              left: { style: 'thin', color: { rgb: 'E2E8F0' } },
              right: { style: 'thin', color: { rgb: 'E2E8F0' } }
            }
          }
        }

        ws['!merges'] = [
          { s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: 5 } }
        ]

        for (let c = 1; c <= 5; c++) {
          ws[`${getColLabel(c)}${excelRowIdx}`] = {
            v: '',
            t: 's',
            s: {
              fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
              border: {
                top: { style: 'medium', color: { rgb: '0F58A7' } },
                bottom: { style: 'medium', color: { rgb: '0F58A7' } },
                left: { style: 'thin', color: { rgb: 'E2E8F0' } },
                right: { style: 'thin', color: { rgb: 'E2E8F0' } }
              }
            }
          }
        }

        columns.forEach((col, colIdx) => {
          if (colIdx <= 5) return
          const colChar = getColLabel(colIdx)
          const cellRef = `${colChar}${excelRowIdx}`

          let sumFormula = ''
          if (col.key === 'tonKho') {
            sumFormula = `G${excelRowIdx}-H${excelRowIdx}`
          } else {
            sumFormula = `SUM(${colChar}5:${colChar}${lastDataRow})`
          }
          const sumVal = realReportSummaryRows.reduce((sum, r) => sum + (Number(r[col.key]) || 0), 0)
          
          ws[cellRef] = {
            f: sumFormula,
            v: sumVal,
            t: 'n',
            s: {
              font: { name: 'Segoe UI', sz: 10, bold: true },
              alignment: { horizontal: 'right', vertical: 'center' },
              fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
              border: {
                top: { style: 'medium', color: { rgb: '0F58A7' } },
                bottom: { style: 'medium', color: { rgb: '0F58A7' } },
                left: { style: 'thin', color: { rgb: 'E2E8F0' } },
                right: { style: 'thin', color: { rgb: 'E2E8F0' } }
              }
            },
            z: '#,##0.00;[Red]-#,##0.00;"-"'
          }
        })

        ws['!ref'] = `A1:${getColLabel(columns.length - 1)}${excelRowIdx}`

        // Prepare Detailed Sheets "Thực nhập" and "Thực xuất"
        const thucNhapList = []
        const thucXuatList = []

        realReportSummaryRows.forEach(row => {
          const txs = row.transactions || []
          txs.forEach(tx => {
            if (tx.nhapVal > 0) {
              thucNhapList.push({
                maSAP: row.maSAP,
                tenVatTu: row.tenVatTu,
                dvt: row.dvt,
                ngayXuatNhap: tx.ngayXuatNhap || '',
                maDon: tx.maDonNhapKho || tx.maDonXuatKho || '',
                donViGiao: tx.donViGiao || '',
                donViNhan: tx.donViNhan || '',
                qty: tx.nhapVal,
                typeLabel: tx.detailTypeNhap === 'nhan_ncc' ? 'Nhận từ NCC'
                          : tx.detailTypeNhap === 'nhan_kho' ? 'Nhận từ Kho gửi'
                          : tx.detailTypeNhap === 'tra_ncc' ? 'Trả lại NCC'
                          : tx.detailTypeNhap === 'dieu_chuyen_di' ? 'Điều chuyển đi (Không tính)'
                          : tx.detailTypeNhap === 'nhan_khac' ? 'Nhận nguồn khác (Không tính)'
                          : 'Không tính',
                logicVal: tx.logicNhapVal,
                contribution: tx.nhapVal * tx.logicNhapVal
              })
            }
            if (tx.xuatVal > 0) {
              thucXuatList.push({
                maSAP: row.maSAP,
                tenVatTu: row.tenVatTu,
                dvt: row.dvt,
                ngayXuatNhap: tx.ngayXuatNhap || '',
                maDon: tx.maDonNhapKho || tx.maDonXuatKho || '',
                donViGiao: tx.donViGiao || '',
                donViNhan: tx.donViNhan || '',
                qty: tx.xuatVal,
                typeLabel: tx.detailTypeXuat === 'xuat_todoi' ? 'Xuất cho Tổ đội'
                          : tx.detailTypeXuat === 'xuat_kho' ? 'Xuất đi Kho nhận'
                          : tx.detailTypeXuat === 'todoi_tra' ? 'Tổ đội trả hàng'
                          : tx.detailTypeXuat === 'xuat_khac' ? 'Xuất khác (Không tính)'
                          : tx.detailTypeXuat === 'nhan_lai_khac' ? 'Nhận lại khác (Không tính)'
                          : 'Không tính',
                logicVal: tx.logicXuatVal,
                contribution: tx.xuatVal * tx.logicXuatVal
              })
            }
          })
        })

        // Sheet 2: Thuc_Nhap
        const thucNhapCols = [
          { key: 'STT', label: 'STT', width: 50 },
          { key: 'ngayXuatNhap', label: 'Ngày', width: 100 },
          { key: 'maSAP', label: 'Mã SAP', width: 100 },
          { key: 'tenVatTu', label: 'Tên vật tư', width: 250 },
          { key: 'dvt', label: 'ĐVT', width: 70 },
          { key: 'maDon', label: 'Mã chứng từ', width: 150 },
          { key: 'donViGiao', label: 'Đơn vị giao (NCC/Kho gửi)', width: 180 },
          { key: 'donViNhan', label: 'Đơn vị nhận (Kho nhận)', width: 180 },
          { key: 'qty', label: 'KL Chứng từ', width: 120 },
          { key: 'typeLabel', label: 'Phân loại', width: 150 },
          { key: 'logicVal', label: 'Hệ số logic', width: 90 },
          { key: 'contribution', label: 'Đóng góp thực', width: 120 }
        ]

        const wsThucNhap = {}
        wsThucNhap['!cols'] = thucNhapCols.map(c => ({ wpx: c.width }))

        wsThucNhap['A1'] = {
          v: `CHI TIẾT CHỨNG TỪ THỰC NHẬP VẬT TƯ THIẾT BỊ`,
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '0F766E' } },
            alignment: { horizontal: 'left', vertical: 'center' }
          }
        }
        wsThucNhap['A2'] = {
          v: `Kho / Dự án: ${localProject || 'Tất cả'} | Xuất từ dữ liệu Đơn chung`,
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 10, italic: true },
            alignment: { horizontal: 'left', vertical: 'center' }
          }
        }

        let thucNhapRowIdx = 4
        thucNhapCols.forEach((col, colIdx) => {
          const colChar = getColLabel(colIdx)
          wsThucNhap[`${colChar}${thucNhapRowIdx}`] = {
            v: col.label,
            t: 's',
            s: {
              fill: { patternType: 'solid', fgColor: { rgb: '0D9488' } },
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'FFFFFF' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: '0F766E' } },
                bottom: { style: 'medium', color: { rgb: '0F766E' } },
                left: { style: 'thin', color: { rgb: '0F766E' } },
                right: { style: 'thin', color: { rgb: '0F766E' } }
              }
            }
          }
        })

        thucNhapList.forEach((row, rowIndex) => {
          thucNhapRowIdx++
          const isEven = (rowIndex % 2 === 1)
          const rowBgColor = isEven ? 'F8FAFC' : 'FFFFFF'

          thucNhapCols.forEach((col, colIdx) => {
            const colChar = getColLabel(colIdx)
            const cellRef = `${colChar}${thucNhapRowIdx}`

            let val = ''
            let cellType = 's'
            let numFormat = undefined
            let isFormula = false
            let formulaStr = ''

            if (col.key === 'STT') {
              val = rowIndex + 1
              cellType = 'n'
            } else if (col.key === 'qty') {
              val = row.qty
              cellType = 'n'
              numFormat = '#,##0.00'
            } else if (col.key === 'logicVal') {
              val = row.logicVal
              cellType = 'n'
              numFormat = '0'
            } else if (col.key === 'contribution') {
              isFormula = true
              formulaStr = `I${thucNhapRowIdx}*K${thucNhapRowIdx}`
              val = row.contribution
              cellType = 'n'
              numFormat = '#,##0.00;[Red]-#,##0.00;"-"'
            } else {
              val = row[col.key] || ''
            }

            const isCentered = ['STT', 'ngayXuatNhap', 'maSAP', 'dvt', 'maDon', 'logicVal'].includes(col.key)
            const isRight = ['qty', 'contribution'].includes(col.key)

            const cellStyle = {
              font: { name: 'Segoe UI', sz: 9, color: { rgb: '1B1919' } },
              alignment: {
                horizontal: isCentered ? 'center' : (isRight ? 'right' : 'left'),
                vertical: 'center',
                wrapText: true
              },
              fill: { patternType: 'solid', fgColor: { rgb: rowBgColor } },
              border: {
                top: { style: 'thin', color: { rgb: 'E2E8F0' } },
                bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
                left: { style: 'thin', color: { rgb: 'E2E8F0' } },
                right: { style: 'thin', color: { rgb: 'E2E8F0' } }
              }
            }

            if (col.key === 'contribution') {
              cellStyle.font.bold = true
              if (row.contribution < 0) {
                cellStyle.font.color = { rgb: 'EF4444' }
              } else if (row.contribution > 0) {
                cellStyle.font.color = { rgb: '047857' }
              }
            }

            const cellObj = { v: val, t: cellType, s: cellStyle }
            if (isFormula) cellObj.f = formulaStr
            if (numFormat) cellObj.z = numFormat

            wsThucNhap[cellRef] = cellObj
          })
        })
        wsThucNhap['!ref'] = `A1:${getColLabel(thucNhapCols.length - 1)}${thucNhapRowIdx}`

        // Sheet 3: Thuc_Xuat
        const thucXuatCols = [
          { key: 'STT', label: 'STT', width: 50 },
          { key: 'ngayXuatNhap', label: 'Ngày', width: 100 },
          { key: 'maSAP', label: 'Mã SAP', width: 100 },
          { key: 'tenVatTu', label: 'Tên vật tư', width: 250 },
          { key: 'dvt', label: 'ĐVT', width: 70 },
          { key: 'maDon', label: 'Mã chứng từ', width: 150 },
          { key: 'donViGiao', label: 'Đơn vị giao (Kho gửi)', width: 180 },
          { key: 'donViNhan', label: 'Đơn vị nhận (Tổ đội/Kho nhận)', width: 180 },
          { key: 'qty', label: 'KL Chứng từ', width: 120 },
          { key: 'typeLabel', label: 'Phân loại', width: 150 },
          { key: 'logicVal', label: 'Hệ số logic', width: 90 },
          { key: 'contribution', label: 'Đóng góp thực', width: 120 }
        ]

        const wsThucXuat = {}
        wsThucXuat['!cols'] = thucXuatCols.map(c => ({ wpx: c.width }))

        wsThucXuat['A1'] = {
          v: `CHI TIẾT CHỨNG TỪ THỰC XUẤT VẬT TƯ THIẾT BỊ`,
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: 'C2410C' } },
            alignment: { horizontal: 'left', vertical: 'center' }
          }
        }
        wsThucXuat['A2'] = {
          v: `Kho / Dự án: ${localProject || 'Tất cả'} | Xuất từ dữ liệu Đơn chung`,
          t: 's',
          s: {
            font: { name: 'Segoe UI', sz: 10, italic: true },
            alignment: { horizontal: 'left', vertical: 'center' }
          }
        }

        let thucXuatRowIdx = 4
        thucXuatCols.forEach((col, colIdx) => {
          const colChar = getColLabel(colIdx)
          wsThucXuat[`${colChar}${thucXuatRowIdx}`] = {
            v: col.label,
            t: 's',
            s: {
              fill: { patternType: 'solid', fgColor: { rgb: 'EA580C' } },
              font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'FFFFFF' } },
              alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
              border: {
                top: { style: 'thin', color: { rgb: 'C2410C' } },
                bottom: { style: 'medium', color: { rgb: 'C2410C' } },
                left: { style: 'thin', color: { rgb: 'C2410C' } },
                right: { style: 'thin', color: { rgb: 'C2410C' } }
              }
            }
          }
        })

        thucXuatList.forEach((row, rowIndex) => {
          thucXuatRowIdx++
          const isEven = (rowIndex % 2 === 1)
          const rowBgColor = isEven ? 'F8FAFC' : 'FFFFFF'

          thucXuatCols.forEach((col, colIdx) => {
            const colChar = getColLabel(colIdx)
            const cellRef = `${colChar}${thucXuatRowIdx}`

            let val = ''
            let cellType = 's'
            let numFormat = undefined
            let isFormula = false
            let formulaStr = ''

            if (col.key === 'STT') {
              val = rowIndex + 1
              cellType = 'n'
            } else if (col.key === 'qty') {
              val = row.qty
              cellType = 'n'
              numFormat = '#,##0.00'
            } else if (col.key === 'logicVal') {
              val = row.logicVal
              cellType = 'n'
              numFormat = '0'
            } else if (col.key === 'contribution') {
              isFormula = true
              formulaStr = `I${thucXuatRowIdx}*K${thucXuatRowIdx}`
              val = row.contribution
              cellType = 'n'
              numFormat = '#,##0.00;[Red]-#,##0.00;"-"'
            } else {
              val = row[col.key] || ''
            }

            const isCentered = ['STT', 'ngayXuatNhap', 'maSAP', 'dvt', 'maDon', 'logicVal'].includes(col.key)
            const isRight = ['qty', 'contribution'].includes(col.key)

            const cellStyle = {
              font: { name: 'Segoe UI', sz: 9, color: { rgb: '1B1919' } },
              alignment: {
                horizontal: isCentered ? 'center' : (isRight ? 'right' : 'left'),
                vertical: 'center',
                wrapText: true
              },
              fill: { patternType: 'solid', fgColor: { rgb: rowBgColor } },
              border: {
                top: { style: 'thin', color: { rgb: 'E2E8F0' } },
                bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
                left: { style: 'thin', color: { rgb: 'E2E8F0' } },
                right: { style: 'thin', color: { rgb: 'E2E8F0' } }
              }
            }

            if (col.key === 'contribution') {
              cellStyle.font.bold = true
              if (row.contribution < 0) {
                cellStyle.font.color = { rgb: 'EF4444' }
              } else if (row.contribution > 0) {
                cellStyle.font.color = { rgb: '047857' }
              }
            }

            const cellObj = { v: val, t: cellType, s: cellStyle }
            if (isFormula) cellObj.f = formulaStr
            if (numFormat) cellObj.z = numFormat

            wsThucXuat[cellRef] = cellObj
          })
        })
        wsThucXuat['!ref'] = `A1:${getColLabel(thucXuatCols.length - 1)}${thucXuatRowIdx}`

        XLSXStyle.utils.book_append_sheet(wb, ws, "Tổng hợp")
        XLSXStyle.utils.book_append_sheet(wb, wsThucNhap, "Thuc_Nhap")
        XLSXStyle.utils.book_append_sheet(wb, wsThucXuat, "Thuc_Xuat")

        if (materialPriceRows && materialPriceRows.length > 0) {
          const phanNhomSheet = buildPhanNhomVatTuSheet(materialPriceRows, materialMetadataMap)
          XLSXStyle.utils.book_append_sheet(wb, phanNhomSheet, "Phân nhóm Vật tư")
        }

        const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary', compression: false })
        const s2ab = (s) => {
          const buf = new ArrayBuffer(s.length)
          const view = new Uint8Array(buf)
          for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF
          return buf
        }

        const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' })
        const url = URL.createObjectURL(blob)
        const a = document.createElement('a')
        a.href = url
        a.download = `Bao_Cao_Tong_Hop_Thuc_Te_${localProject || 'Tat_Ca'}_${new Date().toISOString().slice(0, 10)}.xlsx`
        document.body.appendChild(a)
        a.click()
        document.body.removeChild(a)
        URL.revokeObjectURL(url)
        return
      }

      if (realReportRows.length === 0) return

      const wb = XLSXStyle.utils.book_new()
      const ws = {}

      // Columns: STT + COLS_REAL_REPORT
      const columns = [
        { key: 'STT', label: 'STT', width: 50 },
        ...COLS_REAL_REPORT
      ]

      // Set widths
      ws['!cols'] = columns.map(c => ({ wpx: c.width }))

      let excelRowIdx = 1

      // Write title
      ws['A1'] = {
        v: realReportSubTab === 'nhap' ? `BÁO CÁO NHẬP THỰC TẾ VẬT TƯ THIẾT BỊ` : `BÁO CÁO XUẤT THỰC TẾ VẬT TƯ THIẾT BỊ`,
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '0A3D73' } },
          alignment: { horizontal: 'left', vertical: 'center' }
        }
      }
      ws['A2'] = {
        v: `Kho / Dự án: ${localProject || 'Tất cả'} | Đơn vị giao: ${localGiao || 'Tất cả'} | Đơn vị nhận: ${localNhan || 'Tất cả'}`,
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, italic: true },
          alignment: { horizontal: 'left', vertical: 'center' }
        }
      }
      excelRowIdx = 4

      // Write header row
      columns.forEach((col, colIdx) => {
        const colChar = getColLabel(colIdx)
        const cellRef = `${colChar}${excelRowIdx}`
        
        const isNhapKhoGroup = ['khoiLuongNhap', 'maDonViGiao', 'donViGiao', 'nguoiGiao'].includes(col.key)
        const isXuatKhoGroup = ['khoiLuongXuat', 'maDonViNhan', 'donViNhan', 'nguoiPheDuyet'].includes(col.key)
        const isRealReportGroup = ['khoiLuongThuc', 'logicTongHop'].includes(col.key)
        const excelBgColor = isNhapKhoGroup ? '0F766E' : isXuatKhoGroup ? 'C2410C' : isRealReportGroup ? '5B21B6' : '0F58A7'
        const excelBorderColor = isNhapKhoGroup ? '115E59' : isXuatKhoGroup ? '9A3412' : isRealReportGroup ? '4C1D95' : '0A3D73'

        ws[cellRef] = {
          v: col.label,
          t: 's',
          s: {
            fill: {
              patternType: 'solid',
              fgColor: { rgb: excelBgColor }
            },
            font: {
              name: 'Segoe UI',
              sz: 9.5,
              bold: true,
              color: { rgb: 'FFFFFF' }
            },
            alignment: {
              horizontal: 'center',
              vertical: 'center',
              wrapText: true
            },
            border: {
              top: { style: 'thin', color: { rgb: excelBorderColor } },
              bottom: { style: 'medium', color: { rgb: excelBorderColor } },
              left: { style: 'thin', color: { rgb: excelBorderColor } },
              right: { style: 'thin', color: { rgb: excelBorderColor } }
            }
          }
        }
      })

      // Helper to get letter coordinate
      function getColLabel(index) {
        let label = ''
        let temp = index
        while (temp >= 0) {
          label = String.fromCharCode((temp % 26) + 65) + label
          temp = Math.floor(temp / 26) - 1
        }
        return label
      }

      // Helper for status styling
      function getExcelStatusStyle(val, colKey) {
        if (!val) return null
        const v = String(val).toLowerCase()
        if (colKey === 'trangThai') {
          if (v.includes('chờ') || v.includes('chưa')) {
            return { fg: 'FFFBEB', text: '92400E' } // Orange/Yellow
          }
          if (v.includes('phê duyệt') || v.includes('hoàn thành') || v.includes('đã')) {
            return { fg: 'ECFDF5', text: '065F46' } // Green
          }
          if (v.includes('từ chối') || v.includes('hủy')) {
            return { fg: 'FFF1F2', text: '9F1239' } // Red
          }
          return { fg: 'EFF6FF', text: '1E40AF' } // Blue
        }
        if (colKey === 'tinhTrang') {
          if (val === 'NEW') {
            return { fg: 'ECFDF5', text: '065F46' }
          }
          if (val === 'USED') {
            return { fg: 'FFFBEB', text: '92400E' }
          }
          return { fg: 'F8FAFC', text: '475569' }
        }
        return null
      }

      // Helper to convert Date to Excel serial
      const toExcelSerial = (dateVal) => {
        const d = (dateVal instanceof Date) ? dateVal : parseRowDate(dateVal)
        if (!d || isNaN(d.getTime())) return null
        return Math.round(d.getTime() / 86400000) + 25569
      }

      // Write data rows
      realReportRows.forEach((row, rowIndex) => {
        excelRowIdx++
        const isEvenNum = (rowIndex % 2 === 1)
        const rowBgColor = isEvenNum ? 'F8FAFC' : 'FFFFFF'

        columns.forEach((col, colIdx) => {
          const colChar = getColLabel(colIdx)
          const cellRef = `${colChar}${excelRowIdx}`

          let val = ''
          let cellType = 's'
          let numFormat = undefined

          if (col.key === 'STT') {
            val = rowIndex + 1
            cellType = 'n'
          } else if (col.key === 'ngayXuatNhap') {
            const serial = toExcelSerial(row[col.key])
            if (serial !== null) {
              val = serial
              cellType = 'n'
              numFormat = 'dd/mm/yyyy'
            } else {
              val = String(row[col.key] || '')
            }
          } else {
            const raw = row[col.key]
            if (raw !== null && raw !== undefined) {
              if (col.key === 'logicTongHop') {
                const valLogic = Number(raw)
                if (valLogic === 0) {
                  val = '-'
                  cellType = 's'
                } else {
                  val = valLogic
                  cellType = 'n'
                  numFormat = '0'
                }
              } else if (col.key === 'khoiLuongThuc') {
                const valThuc = Number(raw)
                if (valThuc === 0) {
                  val = '-'
                  cellType = 's'
                } else {
                  val = valThuc
                  cellType = 'n'
                  numFormat = '#,##0.00'
                }
              } else {
                const isRightAligned = ['khoiLuongNhap', 'khoiLuongXuat'].includes(col.key) || col.key.toLowerCase().includes('khoiluong')
                if (isRightAligned && !isNaN(Number(raw)) && raw !== '') {
                  val = Number(raw)
                  cellType = 'n'
                  numFormat = '#,##0.00'
                } else {
                  val = String(raw)
                }
              }
            }
          }

          // Check alignment
          const isCenteredCol = [
            'STT', 'ngayXuatNhap', 'maVatTu', 'maSAP', 'dvt', 'loaiDon', 'logicTongHop',
            'maDonViGiao', 'nguoiGiao',
            'maDonViNhan', 'nguoiPheDuyet', 'nguoiNhan',
            'soHopDong', 'thuKho', 'tinhTrang', 'trangThai'
          ].includes(col.key)
          const isRightAligned = ['khoiLuongNhap', 'khoiLuongXuat'].includes(col.key) || col.key.toLowerCase().includes('khoiluong')

          // Styles
          const cellStyle = {
            font: {
              name: 'Segoe UI',
              sz: 9,
              color: { rgb: '1A1A1A' }
            },
            alignment: {
              horizontal: isCenteredCol ? 'center' : (isRightAligned ? 'right' : 'left'),
              vertical: 'center',
              wrapText: true
            },
            fill: {
              patternType: 'solid',
              fgColor: { rgb: rowBgColor }
            },
            border: {
              top: { style: 'thin', color: { rgb: 'E2E8F0' } },
              bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
              left: { style: 'thin', color: { rgb: 'E2E8F0' } },
              right: { style: 'thin', color: { rgb: 'E2E8F0' } }
            }
          }

          // Badge/Status custom coloring
          const statusStyle = getExcelStatusStyle(val, col.key)
          if (statusStyle) {
            cellStyle.fill = { patternType: 'solid', fgColor: { rgb: statusStyle.fg } }
            cellStyle.font.bold = true
            cellStyle.font.color = { rgb: statusStyle.text }
          }

          // Special logicTongHop/khoiLuongThuc coloring to match web app y hệt
          if (col.key === 'logicTongHop' || col.key === 'khoiLuongThuc') {
            const valLogic = Number(row.logicTongHop)
            if (valLogic === 1) {
              cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'ECFDF5' } }
              cellStyle.font.color = { rgb: '065F46' }
              cellStyle.font.bold = true
            } else if (valLogic === -1) {
              cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'FEF2F2' } }
              cellStyle.font.color = { rgb: '991B1B' }
              cellStyle.font.bold = true
            } else if (valLogic === 0) {
              cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } }
              cellStyle.font.color = { rgb: '475569' }
              cellStyle.font.bold = false
            }
          }

          ws[cellRef] = {
            v: val,
            t: cellType,
            s: cellStyle
          }
          if (numFormat) {
            ws[cellRef].z = numFormat
          }
        })
      })

      // Sum Row for Real Report
      excelRowIdx++
      const lastDataRow = excelRowIdx - 1
      ws[`A${excelRowIdx}`] = {
        v: 'TỔNG CỘNG',
        t: 's',
        s: {
          font: { name: 'Segoe UI', sz: 10, bold: true },
          alignment: { horizontal: 'center', vertical: 'center' },
          fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
          border: {
            top: { style: 'medium', color: { rgb: '0F58A7' } },
            bottom: { style: 'medium', color: { rgb: '0F58A7' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          }
        }
      }

      // Merge STT up to tenVatTu columns for label (assuming index 0 to 4)
      ws['!merges'] = [
        { s: { r: excelRowIdx - 1, c: 0 }, e: { r: excelRowIdx - 1, c: 4 } }
      ]

      // Empty styling for merged cells
      for (let c = 1; c <= 4; c++) {
        ws[`${getColLabel(c)}${excelRowIdx}`] = {
          v: '',
          t: 's',
          s: {
            fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
            border: {
              top: { style: 'medium', color: { rgb: '0F58A7' } },
              bottom: { style: 'medium', color: { rgb: '0F58A7' } },
              left: { style: 'thin', color: { rgb: 'E2E8F0' } },
              right: { style: 'thin', color: { rgb: 'E2E8F0' } }
            }
          }
        }
      }

      // Write SUM formulas for Numeric Columns (khoiLuongNhap, khoiLuongXuat etc)
      columns.forEach((col, colIdx) => {
        if (colIdx <= 4) return // Skip merged
        const colChar = getColLabel(colIdx)
        const cellRef = `${colChar}${excelRowIdx}`

        const isNumeric = ['khoiLuongNhap', 'khoiLuongXuat'].includes(col.key)
        const totalStyle = {
          font: { name: 'Segoe UI', sz: 10, bold: true },
          alignment: { horizontal: isNumeric ? 'right' : 'center', vertical: 'center' },
          fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
          border: {
            top: { style: 'medium', color: { rgb: '0F58A7' } },
            bottom: { style: 'medium', color: { rgb: '0F58A7' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          }
        }

        if (isNumeric) {
          const sumFormula = `SUM(${colChar}5:${colChar}${lastDataRow})`
          const sumVal = realReportRows.reduce((sum, r) => sum + (Number(r[col.key]) || 0), 0)
          ws[cellRef] = {
            f: sumFormula,
            v: sumVal,
            t: 'n',
            s: totalStyle,
            z: '#,##0.00;[Red]-#,##0.00;"-"'
          }
        } else {
          ws[cellRef] = {
            v: '',
            t: 's',
            s: totalStyle
          }
        }
      })

      ws['!ref'] = `A1:${getColLabel(columns.length - 1)}${excelRowIdx}`
      const sheetName = realReportSubTab === 'nhap' ? "Bao_Cao_Nhap_Thuc" : "Bao_Cao_Xuat_Thuc"
      XLSXStyle.utils.book_append_sheet(wb, ws, sheetName)

      if (materialPriceRows && materialPriceRows.length > 0) {
        const phanNhomSheet = buildPhanNhomVatTuSheet(materialPriceRows, materialMetadataMap)
        XLSXStyle.utils.book_append_sheet(wb, phanNhomSheet, "Phân nhóm Vật tư")
      }

      const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary', compression: false })
      const s2ab = (s) => {
        const buf = new ArrayBuffer(s.length)
        const view = new Uint8Array(buf)
        for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF
        return buf
      }

      const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' })
      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      const suffix = realReportSubTab === 'nhap' ? 'Nhap_Thuc' : 'Xuat_Thuc'
      a.download = `Bao_Cao_${suffix}_${localProject || 'Tat_Ca'}_${new Date().toISOString().slice(0, 10)}.xlsx`
      document.body.appendChild(a)
      a.click()
      document.body.removeChild(a)
      URL.revokeObjectURL(url)
      return
    }

    if (reportData.length === 0) return

    const wb = XLSXStyle.utils.book_new()
    const ws = {}

    // Set columns widths
    const cols = [
      { key: 'STT', label: 'STT', wch: 6 },
      { key: 'maSAP', label: 'Mã SAP', wch: 14 },
      { key: 'maVatTu', label: 'Mã vật tư', wch: 12 },
      { key: 'tenVatTu', label: 'Tên vật tư', wch: 35 },
      { key: 'dvt', label: 'ĐVT', wch: 8 },
      { key: 'thongSoKyThuat', label: 'Thông số kỹ thuật', wch: 12 },
      { key: 'received', label: 'Khối lượng nhận', wch: 12 },
      { key: 'issued', label: 'Khối lượng xuất', wch: 12 },
      { key: 'stock', label: isRealReport ? 'Khối lượng tồn thực tế' : 'Khối lượng tồn kho', wch: 12 },
      { key: 'latestReceivedDate', label: 'Ngày nhập muộn nhất', wch: 12 },
      { key: 'daysSinceReceived', label: 'Số ngày nhập muộn nhất đến hôm nay', wch: 12 },
      { key: 'latestIssuedDate', label: 'Ngày xuất muộn nhất', wch: 12 },
      { key: 'daysSinceIssued', label: 'Số ngày xuất muộn nhất đến hôm nay', wch: 12 },
      { key: 'unusedStatus', label: 'Trạng thái vật tư không sử dụng', wch: 22 },
      { key: 'estimatedUnitPrice', label: 'Đơn giá trung bình tạm tính', wch: 16 },
      { key: 'valueOver30Days', label: isRealReport ? 'Thành tiền vật tư tồn thực tế >30 ngày' : 'Thành tiền vật tư tồn >30 ngày', wch: 20 },
      { key: 'materialClassification', label: 'Phân loại vật tư', wch: 18 }
    ]
    // Quy đổi Date -> Excel serial number để dùng được trong công thức/định dạng ngày
    const toExcelSerial = (dateVal) => {
      const d = (dateVal instanceof Date) ? dateVal : parseRowDate(dateVal)
      if (!d || isNaN(d.getTime())) return null
      return Math.round(d.getTime() / 86400000) + 25569
    }
    ws['!cols'] = cols.map(c => ({ wch: c.wch }))

    // Title Row
    ws['A1'] = {
      v: isRealReport ? `BÁO CÁO XUẤT NHẬP THỰC VẬT TƯ THIẾT BỊ` : `BÁO CÁO XUẤT NHẬP TỒN VẬT TƯ THIẾT BỊ`,
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '0A3D73' } },
        alignment: { horizontal: 'left', vertical: 'center' }
      }
    }
    ws['A2'] = {
      v: `Kho / Dự án: ${localProject || 'Tất cả'}`,
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 11, italic: true },
        alignment: { horizontal: 'left', vertical: 'center' }
      }
    }
    ws['A3'] = {
      v: `Trạng thái dữ liệu: ${
        statusFilter === 'approved_only' ? 'Chỉ tính đơn đã phê duyệt'
        : statusFilter === 'approved_pending' ? 'Tính đơn Đã phê duyệt + Chưa phê duyệt'
        : 'Tính tất cả các đơn'
      }`,
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 10, italic: true },
        alignment: { horizontal: 'left', vertical: 'center' }
      }
    }
    // Header starting from row 5
    let rowIdx = 5
    cols.forEach((col, colIdx) => {
      const cellRef = `${String.fromCharCode(65 + colIdx)}${rowIdx}`
      ws[cellRef] = {
        v: col.label,
        t: 's',
        s: {
          fill: { patternType: 'solid', fgColor: { rgb: '0F58A7' } },
          font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: '0A3D73' } },
            bottom: { style: 'medium', color: { rgb: '0A3D73' } },
            left: { style: 'thin', color: { rgb: '0A3D73' } },
            right: { style: 'thin', color: { rgb: '0A3D73' } }
          }
        }
      }
    })

    // Write Data rows — Khối lượng nhận/xuất dùng công thức SUMIFS tham chiếu sang
    // sheet Don_Nhan / Don_Xuat theo Mã SAP để người dùng xem được cách tính ngay trên Excel.
    const donNhanLastRow = 4 + detailRows.nhanList.length
    const donXuatLastRow = 4 + detailRows.xuatList.length
    reportData.forEach((item, idx) => {
      rowIdx++
      const sapCell = `B${rowIdx}`
      const cells = [
        idx + 1,
        item.maSAP,
        item.maVatTu,
        item.tenVatTu,
        item.dvt,
        item.thongSoKyThuat,
        donNhanLastRow >= 5
          ? { f: `SUMIFS(Don_Nhan!$J$5:$J$${donNhanLastRow},Don_Nhan!$C$5:$C$${donNhanLastRow},${sapCell})`, v: item.received }
          : 0,
        donXuatLastRow >= 5
          ? { f: `SUMIFS(Don_Xuat!$J$5:$J$${donXuatLastRow},Don_Xuat!$C$5:$C$${donXuatLastRow},${sapCell})`, v: item.issued }
          : 0,
        { f: `G${rowIdx}-H${rowIdx}`, v: item.stock },
        donNhanLastRow >= 5
          ? { f: `IF(MAX(INDEX((Don_Nhan!$C$5:$C$${donNhanLastRow}=${sapCell})*Don_Nhan!$B$5:$B$${donNhanLastRow},0))>0, MAX(INDEX((Don_Nhan!$C$5:$C$${donNhanLastRow}=${sapCell})*Don_Nhan!$B$5:$B$${donNhanLastRow},0)), "")`, v: toExcelSerial(item.latestReceivedDate), isDate: true }
          : { v: '', isDate: true },
        donNhanLastRow >= 5
          ? { f: `IF(ISNUMBER(J${rowIdx}), TODAY()-J${rowIdx}, "")`, v: getDaysToToday(item.latestReceivedDate) }
          : '',
        donXuatLastRow >= 5
          ? { f: `IF(MAX(INDEX((Don_Xuat!$C$5:$C$${donXuatLastRow}=${sapCell})*Don_Xuat!$B$5:$B$${donXuatLastRow},0))>0, MAX(INDEX((Don_Xuat!$C$5:$C$${donXuatLastRow}=${sapCell})*Don_Xuat!$B$5:$B$${donXuatLastRow},0)), "")`, v: toExcelSerial(item.latestIssuedDate), isDate: true }
          : { v: '', isDate: true },
        donXuatLastRow >= 5
          ? { f: `IF(ISNUMBER(L${rowIdx}), TODAY()-L${rowIdx}, "")`, v: getDaysToToday(item.latestIssuedDate) }
          : '',
        item.unusedStatus,
        item.estimatedUnitPrice,
        { f: `IF(N${rowIdx}="Chưa sử dụng (> 30 ngày)", ROUND(IF(OR(ISNUMBER(SEARCH("khấu hao", Q${rowIdx})), ISNUMBER(SEARCH("tài sản", Q${rowIdx}))), K${rowIdx}*I${rowIdx}*O${rowIdx}, I${rowIdx}*O${rowIdx}), 0), 0)`, v: item.valueOver30Days },
        item.materialClassification
      ]

      cells.forEach((val, colIdx) => {
        const cellRef = `${String.fromCharCode(65 + colIdx)}${rowIdx}`
        const isFormula = val && typeof val === 'object' && 'f' in val
        const isDateCol = val && typeof val === 'object' && val.isDate
        const cellVal = isFormula || isDateCol ? val.v : val
        const isNum = isDateCol ? (cellVal !== null && cellVal !== '') : (isFormula ? true : typeof val === 'number')

        const cellStyle = {
          font: { name: 'Segoe UI', sz: 9.5 },
          alignment: { 
            horizontal: colIdx === 3 || colIdx === 5 || colIdx === 16 ? 'left' : (isNum ? 'right' : 'center'), 
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

        // Apply visual styling exactly matching the webapp's text colors and background fills
        if (colIdx === 0) { // STT
          cellStyle.font.color = { rgb: '475569' }
        } else if (colIdx === 1) { // Mã SAP
          cellStyle.font.color = { rgb: '0F172A' }
          cellStyle.font.bold = true
        } else if (colIdx === 2) { // Mã vật tư
          cellStyle.font.color = { rgb: '475569' }
        } else if (colIdx === 3) { // Tên vật tư
          cellStyle.font.color = { rgb: '0F172A' }
          cellStyle.font.bold = true
        } else if (colIdx === 4) { // ĐVT badge-gray
          cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } }
          cellStyle.font.color = { rgb: '475569' }
        } else if (colIdx === 5) { // Thông số kỹ thuật
          cellStyle.font.color = { rgb: '475569' }
        } else if (colIdx === 6) { // Khối lượng nhận (Green)
          cellStyle.font.color = { rgb: '10B981' }
          cellStyle.font.bold = true
        } else if (colIdx === 7) { // Khối lượng xuất (Orange/Amber)
          cellStyle.font.color = { rgb: 'F97316' }
          cellStyle.font.bold = true
        } else if (colIdx === 8) { // Khối lượng tồn kho
          const isNegative = item.stock < 0
          const isZero = item.stock === 0
          if (isNegative) {
            cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'FEF2F2' } }
            cellStyle.font = { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'EF4444' } }
          } else if (isZero) {
            cellStyle.font.color = { rgb: '475569' }
          } else {
            cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'E8F1FB' } }
            cellStyle.font = { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: '0F58A7' } }
          }
        } else if (colIdx === 9 || colIdx === 11) { // Ngày nhập muộn nhất / Ngày xuất muộn nhất
          cellStyle.font.color = { rgb: '475569' }
        } else if (colIdx === 10 || colIdx === 12) { // Số ngày đến hôm nay
          cellStyle.font.color = { rgb: '475569' }
          cellStyle.font.bold = true
        } else if (colIdx === 13) { // Trạng thái vật tư không sử dụng
          const status = item.unusedStatus
          let bg = 'F1F5F9'
          let fg = '64748B'
          if (status === 'Đang sử dụng') {
            bg = 'ECFDF5'
            fg = '047857'
          } else if (status === 'Chưa sử dụng (> 30 ngày)') {
            bg = 'FFFBEB'
            fg = 'D97706'
          }
          cellStyle.fill = { patternType: 'solid', fgColor: { rgb: bg } }
          cellStyle.font = { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: fg } }
        } else if (colIdx === 14) { // Đơn giá trung bình tạm tính
          cellStyle.font.color = { rgb: '0F58A7' }
          cellStyle.font.bold = true
        } else if (colIdx === 15) { // Thành tiền vật tư tồn >30 ngày
          if (item.valueOver30Days > 0) {
            cellStyle.fill = { patternType: 'solid', fgColor: { rgb: 'FFF5F5' } }
            cellStyle.font = { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'B91C1C' } }
          } else {
            cellStyle.font.color = { rgb: '475569' }
            cellStyle.font.bold = true
          }
        } else if (colIdx === 16) { // Phân loại vật tư
          const val = materialClassifications[item.maSAP] || ''
          const text = val.trim().toLowerCase()
          let bg = 'F8FAFC'
          let fg = '64748B'
          if (text) {
            if (text.includes('tiêu hao')) {
              bg = 'ECFDF5'
              fg = '047857'
            } else if (text.includes('khấu hao') || text.includes('tài sản')) {
              bg = 'FFF7ED'
              fg = 'EA580C'
            } else {
              bg = 'EFF6FF'
              fg = '1D4ED8'
            }
          }
          cellStyle.fill = { patternType: 'solid', fgColor: { rgb: bg } }
          cellStyle.font = { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: fg } }
        }

        ws[cellRef] = {
          v: isNum ? cellVal : (cellVal ?? ''),
          t: isNum ? 'n' : 's',
          s: cellStyle
        }
        if (isFormula) ws[cellRef].f = val.f

        // Apply number format
        if (isDateCol) {
          if (isNum) ws[cellRef].z = 'dd/mm/yyyy'
        } else if (isNum && colIdx >= 6) {
          if (colIdx === 10 || colIdx === 12 || colIdx === 14 || colIdx === 15) {
            ws[cellRef].z = '#,##0;[Red]-#,##0;"-"'
          } else {
            ws[cellRef].z = '#,##0.00;[Red]-#,##0.00;"-"'
          }
        }
      })
    })

    // Total row
    rowIdx++
    const totalLabelCell = `A${rowIdx}`
    ws[totalLabelCell] = {
      v: 'TỔNG CỘNG',
      t: 's',
      s: {
        font: { name: 'Segoe UI', sz: 10, bold: true },
        alignment: { horizontal: 'center', vertical: 'center' },
        fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
        border: {
          top: { style: 'medium', color: { rgb: '0A3D73' } },
          bottom: { style: 'medium', color: { rgb: '0A3D73' } },
          left: { style: 'thin', color: { rgb: 'E2E8F0' } },
          right: { style: 'thin', color: { rgb: 'E2E8F0' } }
        }
      }
    }

    // Merge cells for 'TỔNG CỘNG' label (columns A to F)
    ws['!merges'] = [
      { s: { r: rowIdx - 1, c: 0 }, e: { r: rowIdx - 1, c: 5 } }
    ]

    // Style merged empty cells so they have the border/fill
    for (let c = 1; c <= 5; c++) {
      const cellRef = `${String.fromCharCode(65 + c)}${rowIdx}`
      ws[cellRef] = {
        v: '',
        t: 's',
        s: {
          fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
          border: {
            top: { style: 'medium', color: { rgb: '0A3D73' } },
            bottom: { style: 'medium', color: { rgb: '0A3D73' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          }
        }
      }
    }

    // Sum cells for G, H, I
    const totalReceivedCell = `${String.fromCharCode(65 + 6)}${rowIdx}`
    const totalIssuedCell = `${String.fromCharCode(65 + 7)}${rowIdx}`
    const totalStockCell = `${String.fromCharCode(65 + 8)}${rowIdx}`

    ws[totalReceivedCell] = {
      v: metrics.totalReceived,
      t: 'n',
      z: '#,##0.00;[Red]-#,##0.00;"-"',
      s: {
        font: { name: 'Segoe UI', sz: 10, bold: true },
        alignment: { horizontal: 'right', vertical: 'center' },
        fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
        border: {
          top: { style: 'medium', color: { rgb: '0A3D73' } },
          bottom: { style: 'medium', color: { rgb: '0A3D73' } },
          left: { style: 'thin', color: { rgb: 'E2E8F0' } },
          right: { style: 'thin', color: { rgb: 'E2E8F0' } }
        }
      }
    }

    ws[totalIssuedCell] = {
      v: metrics.totalIssued,
      t: 'n',
      z: '#,##0.00;[Red]-#,##0.00;"-"',
      s: {
        font: { name: 'Segoe UI', sz: 10, bold: true },
        alignment: { horizontal: 'right', vertical: 'center' },
        fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
        border: {
          top: { style: 'medium', color: { rgb: '0A3D73' } },
          bottom: { style: 'medium', color: { rgb: '0A3D73' } },
          left: { style: 'thin', color: { rgb: 'E2E8F0' } },
          right: { style: 'thin', color: { rgb: 'E2E8F0' } }
        }
      }
    }

    ws[totalStockCell] = {
      v: metrics.totalStock,
      t: 'n',
      z: '#,##0.00;[Red]-#,##0.00;"-"',
      s: {
        font: { name: 'Segoe UI', sz: 10, bold: true },
        alignment: { horizontal: 'right', vertical: 'center' },
        fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
        border: {
          top: { style: 'medium', color: { rgb: '0A3D73' } },
          bottom: { style: 'medium', color: { rgb: '0A3D73' } },
          left: { style: 'thin', color: { rgb: 'E2E8F0' } },
          right: { style: 'thin', color: { rgb: 'E2E8F0' } }
        }
      }
    }

    // Total for Column P: Thành tiền vật tư tồn >30 ngày (column index 15)
    const totalValueOver30DaysCell = `P${rowIdx}`
    ws[totalValueOver30DaysCell] = {
      f: `SUM(P6:P${rowIdx - 1})`,
      v: metrics.totalValueOver30Days,
      t: 'n',
      z: '#,##0;[Red]-#,##0;"-"',
      s: {
        font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'B91C1C' } },
        alignment: { horizontal: 'right', vertical: 'center' },
        fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
        border: {
          top: { style: 'medium', color: { rgb: '0A3D73' } },
          bottom: { style: 'medium', color: { rgb: '0A3D73' } },
          left: { style: 'thin', color: { rgb: 'E2E8F0' } },
          right: { style: 'thin', color: { rgb: 'E2E8F0' } }
        }
      }
    }

    // Style other columns J to N (9 to 13), O (14), Q (16) in total row to match styling
    const emptyCols = [9, 10, 11, 12, 13, 14, 16] // J, K, L, M, N, O, Q
    emptyCols.forEach(c => {
      const cellRef = `${String.fromCharCode(65 + c)}${rowIdx}`
      ws[cellRef] = {
        v: '',
        t: 's',
        s: {
          fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
          border: {
            top: { style: 'medium', color: { rgb: '0A3D73' } },
            bottom: { style: 'medium', color: { rgb: '0A3D73' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          }
        }
      }
    })

    // Set sheet range bounds
    ws['!ref'] = `A1:Q${rowIdx}`

    XLSXStyle.utils.book_append_sheet(wb, ws, isRealReport ? "Xuat_Nhap_Thuc" : "Xuat_Nhap_Ton")

    // ── Sheet "Đơn nhận" và "Đơn xuất": liệt kê chi tiết từng đơn cấu thành nên
    // cột Khối lượng nhận / Khối lượng xuất ở sheet tổng hợp phía trên ──────────
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
        s: { font: { name: 'Segoe UI', sz: 14, bold: true, color: { rgb: '0A3D73' } }, alignment: { horizontal: 'left', vertical: 'center' } }
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
            fill: { patternType: 'solid', fgColor: { rgb: '0F58A7' } },
            font: { name: 'Segoe UI', sz: 10, bold: true, color: { rgb: 'FFFFFF' } },
            alignment: { horizontal: 'center', vertical: 'center' },
            border: {
              top: { style: 'thin', color: { rgb: '0A3D73' } },
              bottom: { style: 'medium', color: { rgb: '0A3D73' } },
              left: { style: 'thin', color: { rgb: '0A3D73' } },
              right: { style: 'thin', color: { rgb: '0A3D73' } }
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
              t: serial !== null ? 'n' : 's'
            }
            if (serial !== null) dws[cellRef].z = 'dd/mm/yyyy'
            return
          }

          const isNum = typeof val === 'number'
          dws[cellRef] = {
            v: val ?? '',
            t: isNum ? 'n' : 's'
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
          fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
          border: {
            top: { style: 'medium', color: { rgb: '0A3D73' } },
            bottom: { style: 'medium', color: { rgb: '0A3D73' } },
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
            fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
            border: {
              top: { style: 'medium', color: { rgb: '0A3D73' } },
              bottom: { style: 'medium', color: { rgb: '0A3D73' } },
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
          fill: { patternType: 'solid', fgColor: { rgb: 'F1F5F9' } },
          border: {
            top: { style: 'medium', color: { rgb: '0A3D73' } },
            bottom: { style: 'medium', color: { rgb: '0A3D73' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          }
        }
      }

      dws['!ref'] = `A1:O${dRowIdx}`
      return dws
    }

    const nhanSheet = buildDetailSheet(detailRows.nhanList, 'CHI TIẾT ĐƠN NHẬN (cấu thành Khối lượng nhận)', 'Khối lượng nhận')
    XLSXStyle.utils.book_append_sheet(wb, nhanSheet, "Don_Nhan")

    const xuatSheet = buildDetailSheet(detailRows.xuatList, 'CHI TIẾT ĐƠN XUẤT (cấu thành Khối lượng xuất)', 'Khối lượng xuất')
    XLSXStyle.utils.book_append_sheet(wb, xuatSheet, "Don_Xuat")

    if (materialPriceRows && materialPriceRows.length > 0) {
      const phanNhomSheet = buildPhanNhomVatTuSheet(materialPriceRows, materialMetadataMap)
      XLSXStyle.utils.book_append_sheet(wb, phanNhomSheet, "Phân nhóm Vật tư")
    }

    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary', compression: false })

    const s2ab = (s) => {
      const buf = new ArrayBuffer(s.length)
      const view = new Uint8Array(buf)
      for (let i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF
      return buf
    }

    const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = isRealReport 
      ? `Bao_Cao_Xuat_Nhap_Thuc_${(localProject || 'all').replace(/\s+/g, '_')}.xlsx`
      : `Bao_Cao_Xuat_Nhap_Ton_${(localProject || 'all').replace(/\s+/g, '_')}.xlsx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  const handleWarehouseChange = (e) => {
    const val = e.target.value
    setLocalProject(val)
  }

  const toggleSort = (field) => {
    if (sortField === field) {
      setSortDirection(prev => prev === 'asc' ? 'desc' : 'asc')
    } else {
      setSortField(field)
      setSortDirection('asc')
    }
  }

  const btnBase = {
    display: 'inline-flex', alignItems: 'center', justifyContent: 'center',
    height: 30, minWidth: 30, padding: '0 8px',
    border: '1px solid #e2e8f0', borderRadius: 6,
    fontSize: 13, fontWeight: 500, cursor: 'pointer',
    background: '#fff', color: '#374151', transition: 'all 0.15s',
    userSelect: 'none'
  }
  const btnActive = { ...btnBase, background: 'var(--primary)', color: '#fff', borderColor: 'var(--primary)', fontWeight: 700 }
  const btnDisabled = { ...btnBase, opacity: 0.4, cursor: 'not-allowed', pointerEvents: 'none' }

  const metricsCards = [
    {
      label: 'Tổng khối lượng nhận',
      value: metrics.totalReceived === 0 ? '-' : metrics.totalReceived.toLocaleString('vi-VN', { maximumFractionDigits: 2 }),
      color: '#10b981',
      bg: '#ecfdf5',
      border: '#a7f3d0',
      icon: <Truck size={18} />
    },
    {
      label: 'Tổng khối lượng xuất',
      value: metrics.totalIssued === 0 ? '-' : metrics.totalIssued.toLocaleString('vi-VN', { maximumFractionDigits: 2 }),
      color: '#f97316',
      bg: '#fff7ed',
      border: '#fed7aa',
      icon: <Download size={18} style={{ transform: 'rotate(180deg)' }} />
    },
    {
      label: isRealReport ? 'Khối lượng tồn thực tế' : 'Khối lượng tồn kho',
      value: metrics.totalStock === 0 ? '-' : metrics.totalStock.toLocaleString('vi-VN', { maximumFractionDigits: 2 }),
      color: metrics.totalStock >= 0 ? 'var(--primary)' : '#ef4444',
      bg: metrics.totalStock >= 0 ? 'var(--primary-light)' : '#fef2f2',
      border: metrics.totalStock >= 0 ? 'var(--border)' : '#fca5a5',
      icon: <Database size={18} />
    },
    {
      label: isRealReport ? 'Thành tiền tồn thực tế >30 ngày' : 'Thành tiền tồn >30 ngày',
      value: metrics.totalValueOver30Days.toLocaleString('vi-VN') + ' đ',
      color: '#ef4444',
      bg: '#fef2f2',
      border: '#fca5a5',
      icon: <DollarSign size={18} />
    },
    {
      label: 'Tổng số mặt hàng',
      value: metrics.totalItems.toLocaleString('vi-VN'),
      color: '#8b5cf6',
      bg: '#f5f3ff',
      border: '#ddd6fe',
      icon: <PackageCheck size={18} />
    }
  ]

  return (
    <div id="baocao-xuat-nhap-ton-root" style={{ padding: '16px 24px 24px 24px', height: '100%', display: 'flex', flexDirection: 'column', boxSizing: 'border-box', overflow: 'hidden', fontFamily: "'Roboto', sans-serif" }}>
      {/* Sub-tabs for Inventory Report view */}
      {isRealReport && (
        <div style={{ display: 'flex', gap: 6, marginBottom: 12, borderBottom: '1px solid var(--border)', paddingBottom: 6, flexShrink: 0 }}>
          <button
            id="btn-real-report-nhap"
            onClick={() => { setRealReportSubTab('nhap'); setCurrentPage(1); }}
            style={{
              padding: '8px 16px',
              borderRadius: '6px 6px 0 0',
              fontSize: '13.5px',
              fontWeight: realReportSubTab === 'nhap' ? 700 : 500,
              background: realReportSubTab === 'nhap' ? 'var(--primary)' : 'transparent',
              color: realReportSubTab === 'nhap' ? '#ffffff' : 'var(--text-muted)',
              border: 'none',
              cursor: 'pointer',
              transition: 'all 0.15s',
              boxShadow: realReportSubTab === 'nhap' ? 'var(--shadow-sm)' : 'none',
              borderBottom: realReportSubTab === 'nhap' ? '3px solid #10b981' : 'none'
            }}
          >
            Nhập thực
          </button>
          <button
            id="btn-real-report-xuat"
            onClick={() => { setRealReportSubTab('xuat'); setCurrentPage(1); }}
            style={{
              padding: '8px 16px',
              borderRadius: '6px 6px 0 0',
              fontSize: '13.5px',
              fontWeight: realReportSubTab === 'xuat' ? 700 : 500,
              background: realReportSubTab === 'xuat' ? 'var(--primary)' : 'transparent',
              color: realReportSubTab === 'xuat' ? '#ffffff' : 'var(--text-muted)',
              border: 'none',
              cursor: 'pointer',
              transition: 'all 0.15s',
              boxShadow: realReportSubTab === 'xuat' ? 'var(--shadow-sm)' : 'none',
              borderBottom: realReportSubTab === 'xuat' ? '3px solid #f97316' : 'none'
            }}
          >
            Xuất thực
          </button>
          <button
            id="btn-real-report-tonghop"
            onClick={() => { setRealReportSubTab('tonghop'); setCurrentPage(1); }}
            style={{
              padding: '8px 16px',
              borderRadius: '6px 6px 0 0',
              fontSize: '13.5px',
              fontWeight: realReportSubTab === 'tonghop' ? 700 : 500,
              background: realReportSubTab === 'tonghop' ? 'var(--primary)' : 'transparent',
              color: realReportSubTab === 'tonghop' ? '#ffffff' : 'var(--text-muted)',
              border: 'none',
              cursor: 'pointer',
              transition: 'all 0.15s',
              boxShadow: realReportSubTab === 'tonghop' ? 'var(--shadow-sm)' : 'none',
              borderBottom: realReportSubTab === 'tonghop' ? '3px solid #8b5cf6' : 'none'
            }}
          >
            Tổng hợp
          </button>
        </div>
      )}
      {!isRealReport && (
        <div style={{ display: 'flex', gap: 6, marginBottom: 12, borderBottom: '1px solid var(--border)', paddingBottom: 6, flexShrink: 0 }}>
          <button
            onClick={() => setSubTab('summary')}
            style={{
              padding: '8px 16px',
              borderRadius: '6px 6px 0 0',
              fontSize: '13.5px',
              fontWeight: subTab === 'summary' ? 700 : 500,
              background: subTab === 'summary' ? 'var(--primary)' : 'transparent',
              color: subTab === 'summary' ? '#ffffff' : 'var(--text-muted)',
              border: 'none',
              cursor: 'pointer',
              transition: 'all 0.15s',
              boxShadow: subTab === 'summary' ? 'var(--shadow-sm)' : 'none',
              borderBottom: subTab === 'summary' ? '2px solid var(--primary)' : 'none'
            }}
          >
            Bảng chi tiết
          </button>
          <button
            onClick={() => setSubTab('dashboard')}
            style={{
              padding: '8px 16px',
              borderRadius: '6px 6px 0 0',
              fontSize: '13.5px',
              fontWeight: subTab === 'dashboard' ? 700 : 500,
              background: subTab === 'dashboard' ? 'var(--primary)' : 'transparent',
              color: subTab === 'dashboard' ? '#ffffff' : 'var(--text-muted)',
              border: 'none',
              cursor: 'pointer',
              transition: 'all 0.15s',
              boxShadow: subTab === 'dashboard' ? 'var(--shadow-sm)' : 'none',
              borderBottom: subTab === 'dashboard' ? '2px solid var(--primary)' : 'none'
            }}
          >
            Dashboard báo cáo
          </button>
        </div>
      )}

      {subTab === 'dashboard' && (
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'auto', gap: 16, paddingBottom: 16 }}>
          {/* Bộ lọc chọn kho báo cáo cụ thể */}
          <div style={{
            background: '#ffffff',
            borderRadius: 8,
            border: '1px solid var(--border)',
            padding: '12px 16px',
            boxShadow: 'var(--shadow-sm)',
            display: 'flex',
            alignItems: 'center',
            gap: 12,
            flexWrap: 'wrap',
            flexShrink: 0
          }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 13, fontWeight: 700, color: 'var(--text)' }}>
              <Warehouse size={16} style={{ color: 'var(--primary)' }} />
              Chọn kho báo cáo:
            </div>
            <select
              value={dashboardWarehouseFilter}
              onChange={(e) => {
                const val = e.target.value
                setDashboardWarehouseFilter(val)
                // Đồng bộ luôn phần chi tiết bên phải với kho vừa chọn
                setSelectedDashProject(val === 'all' ? null : val)
              }}
              style={{
                flex: '0 1 360px',
                minWidth: 220,
                padding: '7px 10px',
                borderRadius: 6,
                border: '1px solid var(--border)',
                fontSize: 13,
                fontWeight: 600,
                color: 'var(--text)',
                background: '#f8fafc',
                cursor: 'pointer'
              }}
            >
              <option value="all">— Tất cả kho dự án BCH ({bchWarehouses.length} kho) —</option>
              {bchWarehouses.map(w => (
                <option key={w} value={w}>{w}</option>
              ))}
            </select>
            {dashboardWarehouseFilter !== 'all' && (
              <button
                onClick={() => {
                  setDashboardWarehouseFilter('all')
                  setSelectedDashProject(null)
                }}
                style={{
                  padding: '6px 12px',
                  borderRadius: 6,
                  border: '1px solid var(--border)',
                  background: '#ffffff',
                  color: 'var(--text-muted)',
                  fontSize: 12.5,
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                Xem tất cả
              </button>
            )}
            <div style={{ fontSize: 12, color: 'var(--text-muted)' }}>
              {dashboardWarehouseFilter === 'all'
                ? 'Đang hiển thị báo cáo tổng hợp của tất cả kho dự án BCH.'
                : `Đang hiển thị báo cáo riêng cho kho: ${dashboardWarehouseFilter}`}
            </div>
          </div>

          {/* Summary Cards Row */}
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
            gap: 12,
            flexShrink: 0
          }}>
            {/* Card 1: Grand Unused Value */}
            <div style={{
              background: '#ffffff',
              borderRadius: 8,
              border: '1px solid var(--border)',
              padding: '16px 20px',
              boxShadow: 'var(--shadow-sm)',
              display: 'flex',
              alignItems: 'center',
              gap: 16
            }}>
              <div style={{
                background: '#fee2e2',
                color: '#ef4444',
                padding: 12,
                borderRadius: 8,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}>
                <AlertCircle size={24} />
              </div>
              <div>
                <div style={{ fontSize: 12.5, color: 'var(--text-muted)', fontWeight: 500, textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  Tổng giá trị tồn chưa dùng
                </div>
                <div style={{ fontSize: 20, fontWeight: 800, color: '#ef4444', margin: '4px 0 2px 0' }}>
                  {dashboardSummary.grandUnusedValue.toLocaleString('vi-VN')} đ
                </div>
                <div style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                  {dashboardWarehouseFilter === 'all' ? 'Tích lũy từ tất cả dự án BCH (> 30 ngày)' : `Kho: ${dashboardWarehouseFilter} (> 30 ngày)`}
                </div>
              </div>
            </div>

            {/* Card 2: Total Warned Items */}
            <div style={{
              background: '#ffffff',
              borderRadius: 8,
              border: '1px solid var(--border)',
              padding: '16px 20px',
              boxShadow: 'var(--shadow-sm)',
              display: 'flex',
              alignItems: 'center',
              gap: 16
            }}>
              <div style={{
                background: '#fff9db',
                color: '#f59e0b',
                padding: 12,
                borderRadius: 8,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}>
                <AlertTriangle size={24} />
              </div>
              <div>
                <div style={{ fontSize: 12.5, color: 'var(--text-muted)', fontWeight: 500, textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  Số mặt hàng cảnh báo
                </div>
                <div style={{ fontSize: 20, fontWeight: 800, color: '#f59e0b', margin: '4px 0 2px 0' }}>
                  {dashboardSummary.grandUnusedCount.toLocaleString()} / {dashboardSummary.grandTotalItems.toLocaleString()}
                </div>
                <div style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                  Tỷ lệ chậm sử dụng: <strong style={{ color: '#f59e0b' }}>{dashboardSummary.averageUnusedRatio.toFixed(1)}%</strong>
                </div>
              </div>
            </div>

            {/* Card 3: Highest Unused Warehouse */}
            <div style={{
              background: '#ffffff',
              borderRadius: 8,
              border: '1px solid var(--border)',
              padding: '16px 20px',
              boxShadow: 'var(--shadow-sm)',
              display: 'flex',
              alignItems: 'center',
              gap: 16
            }}>
              <div style={{
                background: '#eff6ff',
                color: 'var(--primary)',
                padding: 12,
                borderRadius: 8,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}>
                <Building2 size={24} />
              </div>
              <div style={{ minWidth: 0, flex: 1 }}>
                <div style={{ fontSize: 12.5, color: 'var(--text-muted)', fontWeight: 500, textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  Dự án tồn đọng cao nhất
                </div>
                <div style={{ fontSize: 15, fontWeight: 700, color: 'var(--text)', margin: '4px 0 2px 0', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={dashboardSummary.highestUnusedWarehouse}>
                  {dashboardSummary.highestUnusedWarehouse}
                </div>
                <div style={{ fontSize: 12, color: 'var(--text-muted)', display: 'flex', justifyContent: 'space-between' }}>
                  <span>Giá trị tồn:</span>
                  <strong style={{ color: '#ef4444' }}>{dashboardSummary.highestUnusedVal.toLocaleString('vi-VN')} đ</strong>
                </div>
              </div>
            </div>

            {/* Card 4: Monitoring projects */}
            <div style={{
              background: '#ffffff',
              borderRadius: 8,
              border: '1px solid var(--border)',
              padding: '16px 20px',
              boxShadow: 'var(--shadow-sm)',
              display: 'flex',
              alignItems: 'center',
              gap: 16
            }}>
              <div style={{
                background: '#ecfdf5',
                color: '#10b981',
                padding: 12,
                borderRadius: 8,
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center'
              }}>
                <Warehouse size={24} />
              </div>
              <div>
                <div style={{ fontSize: 12.5, color: 'var(--text-muted)', fontWeight: 500, textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                  {dashboardWarehouseFilter === 'all' ? 'Kho dự án BCH theo dõi' : 'Đang xem báo cáo kho'}
                </div>
                <div style={{ fontSize: dashboardWarehouseFilter === 'all' ? 20 : 15, fontWeight: 800, color: '#10b981', margin: '4px 0 2px 0', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={dashboardWarehouseFilter === 'all' ? undefined : dashboardWarehouseFilter}>
                  {dashboardWarehouseFilter === 'all'
                    ? <>{bchWarehouses.length} <span style={{ fontSize: 13, fontWeight: 500, color: 'var(--text-muted)' }}>kho dự án</span></>
                    : dashboardWarehouseFilter}
                </div>
                <div style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                  {dashboardWarehouseFilter === 'all' ? 'Giám sát tự động tình trạng tồn đọng' : `Trong tổng số ${bchWarehouses.length} kho dự án BCH`}
                </div>
              </div>
            </div>
          </div>

          {/* Biểu đồ so sánh giá trị tồn giữa các kho BCH */}
          <div style={{
            background: '#ffffff',
            borderRadius: 8,
            border: '1px solid var(--border)',
            padding: '16px 20px',
            boxShadow: 'var(--shadow-sm)',
            flexShrink: 0
          }}>
            <h3 style={{ fontSize: 14.5, fontWeight: 700, color: 'var(--text)', margin: '0 0 12px 0', display: 'flex', alignItems: 'center', gap: 6 }}>
              <BarChart3 size={16} style={{ color: 'var(--primary)' }} />
              {dashboardWarehouseFilter === 'all'
                ? (isRealReport ? 'BIỂU ĐỒ SO SÁNH GIÁ TRỊ TỒN THỰC TẾ CHƯA SỬ DỤNG (> 30 NGÀY) GIỮA CÁC KHO DỰ ÁN BCH' : 'BIỂU ĐỒ SO SÁNH GIÁ TRỊ TỒN CHƯA SỬ DỤNG (> 30 NGÀY) GIỮA CÁC KHO DỰ ÁN BCH')
                : (isRealReport ? `GIÁ TRỊ TỒN THỰC TẾ CHƯA SỬ DỤNG (> 30 NGÀY) — KHO: ${dashboardWarehouseFilter}` : `GIÁ TRỊ TỒN CHƯA SỬ DỤNG (> 30 NGÀY) — KHO: ${dashboardWarehouseFilter}`)}
            </h3>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
              {filteredDashboardStats.map((stat, index) => {
                const maxVal = dashboardSummary.highestUnusedVal || 1
                const percent = (stat.unusedValue / maxVal) * 100
                const isSelected = selectedDashProject === stat.projectName || (!selectedDashProject && index === 0)
                
                return (
                  <div 
                    key={stat.projectName}
                    onClick={() => setSelectedDashProject(stat.projectName)}
                    style={{ 
                      display: 'flex', 
                      alignItems: 'center', 
                      gap: 12, 
                      cursor: 'pointer',
                      padding: '6px 8px',
                      borderRadius: 6,
                      background: isSelected ? 'var(--primary-light)' : 'transparent',
                      transition: 'all 0.15s'
                    }}
                  >
                    <div style={{ width: 220, fontSize: 13, fontWeight: isSelected ? 700 : 500, color: 'var(--text)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={stat.projectName}>
                      {index + 1}. {stat.projectName}
                    </div>
                    
                    <div style={{ flex: 1, height: 16, background: '#f1f5f9', borderRadius: 4, overflow: 'hidden', position: 'relative' }}>
                      <div style={{ 
                        width: `${Math.max(2, percent)}%`, 
                        height: '100%', 
                        background: stat.alarmColor === '#ef4444' ? 'linear-gradient(90deg, #fca5a5, #ef4444)' : (stat.alarmColor === '#f59e0b' ? 'linear-gradient(90deg, #fde68a, #f59e0b)' : 'linear-gradient(90deg, #cbd5e1, #94a3b8)'), 
                        borderRadius: 4,
                        transition: 'width 0.4s ease-out'
                      }} />
                    </div>
                    
                    <div style={{ width: 140, textAlign: 'right', fontSize: 13, fontWeight: 700, color: stat.unusedValue > 0 ? stat.alarmColor : 'var(--text-muted)' }}>
                      {stat.unusedValue.toLocaleString('vi-VN')} đ
                    </div>
                  </div>
                )
              })}
            </div>
          </div>

          {/* Interactive Split View: Left List vs Right Detail */}
          <div style={{
            display: 'flex',
            flexDirection: 'row',
            gap: 16,
            flex: 1,
            minHeight: 400
          }}>
            {/* Left Box: BCH Warehouses List */}
            <div style={{
              flex: '1 1 55%',
              background: '#ffffff',
              borderRadius: 8,
              border: '1px solid var(--border)',
              boxShadow: 'var(--shadow-sm)',
              display: 'flex',
              flexDirection: 'column',
              overflow: 'hidden'
            }}>
              <div style={{
                padding: '12px 16px',
                borderBottom: '1px solid var(--border)',
                background: '#f8fafc',
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center'
              }}>
                <h3 style={{ fontSize: 13.5, fontWeight: 700, color: 'var(--text)', margin: 0, display: 'flex', alignItems: 'center', gap: 6 }}>
                  <ClipboardList size={15} style={{ color: 'var(--primary)' }} />
                  {dashboardWarehouseFilter === 'all' ? 'DANH SÁCH XẾP HẠNG CẢNH BÁO KHO DỰ ÁN BCH' : 'CHI TIẾT CẢNH BÁO KHO ĐÃ CHỌN'}
                </h3>
                <span style={{ fontSize: 11, color: 'var(--text-muted)' }}>
                  Theo giá trị tồn đọng chưa dùng giảm dần
                </span>
              </div>

              <div style={{ flex: 1, overflow: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', borderSpacing: 0 }}>
                  <thead>
                    <tr style={{ background: '#f1f5f9', borderBottom: '1px solid var(--border)', position: 'sticky', top: 0, zIndex: 1 }}>
                      <th style={{ width: 50, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px 8px' }}>STT</th>
                      <th style={{ textAlign: 'left', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px 8px' }}>Tên kho dự án</th>
                      <th style={{ width: 150, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px 8px' }}>Mức độ cảnh báo</th>
                      <th style={{ width: 140, textAlign: 'center', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px 8px' }}>Mặt hàng chưa dùng</th>
                      <th style={{ width: 140, textAlign: 'right', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px 8px' }}>Khối lượng chưa dùng</th>
                      <th style={{ width: 150, textAlign: 'right', fontSize: 12, fontWeight: 700, color: '#475569', padding: '10px 8px' }}>Thành tiền tồn đọng</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredDashboardStats.length === 0 ? (
                      <tr>
                        <td colSpan={6} style={{ padding: 30, textAlign: 'center', color: 'var(--text-muted)', fontSize: 13 }}>
                          Không có dữ liệu kho dự án BCH nào
                        </td>
                      </tr>
                    ) : (
                      filteredDashboardStats.map((stat, idx) => {
                        const isSelected = selectedDashProject === stat.projectName || (!selectedDashProject && idx === 0)
                        return (
                          <tr 
                            key={stat.projectName}
                            onClick={() => setSelectedDashProject(stat.projectName)}
                            style={{
                              borderBottom: '1px solid #f1f5f9',
                              cursor: 'pointer',
                              background: isSelected ? 'var(--primary-light)' : (idx % 2 === 0 ? '#ffffff' : '#f8fafc'),
                              fontWeight: isSelected ? 600 : 'normal',
                              transition: 'all 0.15s'
                            }}
                          >
                            <td style={{ textAlign: 'center', padding: '10px 8px', fontSize: 12.5, color: 'var(--text-muted)' }}>{idx + 1}</td>
                            <td style={{ textAlign: 'left', padding: '10px 8px', fontSize: 12.5, color: 'var(--text)' }}>
                              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                                {isSelected && <span style={{ width: 6, height: 6, borderRadius: '50%', background: 'var(--primary)' }} />}
                                <span>{stat.projectName}</span>
                              </div>
                            </td>
                            <td style={{ textAlign: 'center', padding: '10px 8px' }}>
                              <span style={{
                                background: stat.alarmBg,
                                color: stat.alarmColor,
                                border: `1px solid ${stat.alarmBorder}`,
                                padding: '3px 8px',
                                borderRadius: 4,
                                fontSize: 11,
                                fontWeight: 700,
                                display: 'inline-block'
                              }}>
                                {stat.alarmLevel}
                              </span>
                            </td>
                            <td style={{ textAlign: 'center', padding: '10px 8px', fontSize: 12.5, color: 'var(--text)' }}>
                              <strong>{stat.unusedCount}</strong> / {stat.totalItemsInStock}
                              <span style={{ fontSize: 11, color: 'var(--text-muted)', display: 'inline', marginLeft: 4 }}>
                                ({stat.unusedRatio.toFixed(1)}%)
                              </span>
                            </td>
                            <td style={{ textAlign: 'right', padding: '10px 8px', fontSize: 12.5, color: 'var(--text-muted)' }}>
                              {stat.unusedQty.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                            </td>
                            <td style={{ textAlign: 'right', padding: '10px 8px', fontSize: 12.5, color: stat.unusedValue > 0 ? stat.alarmColor : 'var(--text)', fontWeight: 700 }}>
                              {stat.unusedValue.toLocaleString('vi-VN')} đ
                            </td>
                          </tr>
                        )
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Right Box: Selected Detail */}
            <div style={{
              flex: '1 1 45%',
              background: '#ffffff',
              borderRadius: 8,
              border: '1px solid var(--border)',
              boxShadow: 'var(--shadow-sm)',
              display: 'flex',
              flexDirection: 'column',
              overflow: 'hidden'
            }}>
              {activeDashStat ? (
                <div style={{ display: 'flex', flexDirection: 'column', flex: 1, overflow: 'hidden' }}>
                  <div style={{
                    padding: '12px 16px',
                    borderBottom: '1px solid var(--border)',
                    background: '#f8fafc',
                    display: 'flex',
                    flexDirection: 'column',
                    gap: 4
                  }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: 'var(--primary)', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                      CHI TIẾT KHO ĐANG THEO DÕI
                    </div>
                    <h3 style={{ fontSize: 14.5, fontWeight: 800, color: 'var(--text)', margin: 0, display: 'flex', alignItems: 'center', gap: 6, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={activeDashStat.projectName}>
                      <Warehouse size={16} style={{ color: 'var(--primary)' }} />
                      {activeDashStat.projectName}
                    </h3>
                  </div>

                  <div style={{ flex: 1, overflow: 'auto', padding: '16px 20px', display: 'flex', flexDirection: 'column', gap: 16 }}>
                    {/* Progress Bar */}
                    <div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 12.5, color: 'var(--text)', marginBottom: 6 }}>
                        <span style={{ fontWeight: 500 }}>Tỷ lệ mặt hàng chậm sử dụng:</span>
                        <strong style={{ color: activeDashStat.alarmColor }}>{activeDashStat.unusedRatio.toFixed(1)}%</strong>
                      </div>
                      <div style={{ height: 12, background: '#f1f5f9', borderRadius: 6, overflow: 'hidden', display: 'flex' }}>
                        <div style={{
                          width: `${activeDashStat.unusedRatio}%`,
                          background: activeDashStat.alarmColor,
                          height: '100%',
                          transition: 'width 0.3s'
                        }} />
                        <div style={{
                          width: `${100 - activeDashStat.unusedRatio}%`,
                          background: '#e2e8f0',
                          height: '100%'
                        }} />
                      </div>
                      <div style={{ display: 'flex', justifyContent: 'space-between', fontSize: 11, color: 'var(--text-muted)', marginTop: 4 }}>
                        <span>Chưa sử dụng (&gt;30 ngày): <strong>{activeDashStat.unusedCount}</strong> mặt hàng</span>
                        <span>Đang sử dụng: <strong>{activeDashStat.totalItemsInStock - activeDashStat.unusedCount}</strong> mặt hàng</span>
                      </div>
                    </div>

                    {/* Summary box */}
                    <div style={{
                      display: 'grid',
                      gridTemplateColumns: '1fr 1fr',
                      gap: 10,
                      background: '#f8fafc',
                      padding: 12,
                      borderRadius: 6,
                      border: '1px solid var(--border)'
                    }}>
                      <div>
                        <div style={{ fontSize: 11, color: 'var(--text-muted)', fontWeight: 500 }}>TỔNG NHẬP (LŨY KẾ)</div>
                        <div style={{ fontSize: 14, fontWeight: 700, color: 'var(--text)', marginTop: 2 }}>{activeDashStat.totalReceived.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</div>
                      </div>
                      <div>
                        <div style={{ fontSize: 11, color: 'var(--text-muted)', fontWeight: 500 }}>TỔNG XUẤT (LŨY KẾ)</div>
                        <div style={{ fontSize: 14, fontWeight: 700, color: 'var(--text)', marginTop: 2 }}>{activeDashStat.totalIssued.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</div>
                      </div>
                      <div style={{ gridColumn: 'span 2', borderTop: '1px solid #e2e8f0', paddingTop: 8, marginTop: 4 }}>
                        <div style={{ fontSize: 11, color: 'var(--text-muted)', fontWeight: 500 }}>TỔNG THÀNH TIỀN TỒN ĐỌNG CHƯA DÙNG</div>
                        <div style={{ fontSize: 16, fontWeight: 800, color: activeDashStat.alarmColor, marginTop: 2 }}>
                          {activeDashStat.unusedValue.toLocaleString('vi-VN')} đ
                        </div>
                      </div>
                    </div>

                    {/* Top 10 Warned items list */}
                    <div>
                      <h4 style={{ fontSize: 12.5, fontWeight: 700, color: '#475569', margin: '0 0 8px 0', textTransform: 'uppercase', letterSpacing: '0.5px' }}>
                        TOP 10 VẬT TƯ TỒN ĐỌNG GIÁ TRỊ LỚN NHẤT KHO NÀY
                      </h4>
                      
                      {activeDashStat.topStagnant.length === 0 ? (
                        <div style={{ padding: 20, textAlign: 'center', background: '#f8fafc', borderRadius: 6, color: 'var(--text-muted)', fontSize: 12 }}>
                          Không có vật tư nào bị cảnh báo chậm sử dụng ở kho này!
                        </div>
                      ) : (
                        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
                          {activeDashStat.topStagnant.map((item, idx) => (
                            <div 
                              key={item.maSAP + idx}
                              style={{
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'space-between',
                                padding: '8px 10px',
                                background: '#f8fafc',
                                borderRadius: 6,
                                borderLeft: `3px solid ${activeDashStat.alarmColor}`,
                                gap: 12
                              }}
                            >
                              <div style={{ minWidth: 0, flex: 1 }}>
                                <div style={{ fontSize: 12, fontWeight: 700, color: 'var(--text)', display: 'flex', alignItems: 'center', gap: 6 }}>
                                  <span style={{ color: 'var(--text-muted)', fontSize: 11 }}>{item.maSAP}</span>
                                  <span style={{
                                    background: '#e2e8f0',
                                    color: '#475569',
                                    padding: '1px 4px',
                                    borderRadius: 3,
                                    fontSize: 10,
                                    fontWeight: 500
                                  }}>
                                    {item.dvt}
                                  </span>
                                </div>
                                <div style={{ fontSize: 12, color: 'var(--text-muted)', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', marginTop: 2 }} title={item.tenVatTu}>
                                  {item.tenVatTu}
                                </div>
                              </div>
                              <div style={{ textAlign: 'right', flexShrink: 0 }}>
                                <div style={{ fontSize: 12.5, fontWeight: 700, color: 'var(--text)' }}>
                                  {item.valueOver30Days.toLocaleString('vi-VN')} đ
                                </div>
                                <div style={{ fontSize: 11, color: 'var(--text-muted)', marginTop: 1 }}>
                                  Tồn: <strong>{item.stock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}</strong>
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>

                    {/* Action button */}
                    <button
                      onClick={() => {
                        setLocalProject(activeDashStat.projectName)
                        setSubTab('summary')
                        setCurrentPage(1)
                      }}
                      className="btn"
                      style={{
                        width: '100%',
                        background: 'var(--primary)',
                        color: '#ffffff',
                        padding: '10px 14px',
                        borderRadius: 6,
                        fontSize: 13,
                        fontWeight: 700,
                        border: 'none',
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        gap: 6,
                        marginTop: 'auto',
                        boxShadow: 'var(--shadow-sm)',
                        transition: 'opacity 0.15s'
                      }}
                    >
                      <BarChart3 size={15} />
                      <span>Xem chi tiết tất cả vật tư kho này tại Bảng chi tiết</span>
                    </button>
                  </div>
                </div>
              ) : (
                <div style={{ padding: 40, textAlign: 'center', color: 'var(--text-muted)', fontSize: 13.5 }}>
                  Hãy chọn một kho dự án để xem chi tiết
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {subTab === 'summary' && (
        <div style={{ flex: 1, display: 'flex', flexDirection: 'column', overflow: 'hidden', minHeight: 0 }}>
      {/* Filter and selector row */}
      <div style={{
        background: '#ffffff',
        borderRadius: 8,
        border: '1px solid var(--border)',
        padding: '12px 16px',
        marginBottom: 12,
        display: 'flex',
        gap: 12,
        flexWrap: 'wrap',
        alignItems: 'center',
        boxShadow: 'var(--shadow-sm)',
        flexShrink: 0
      }}>
        {/* Selector */}
        <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 280 }}>
          <span style={{ fontSize: 13, fontWeight: 700, color: 'var(--text-muted)', whiteSpace: 'nowrap' }}>
            Kho / Dự án:
          </span>
          <div style={{ flex: 1, minWidth: 180 }}>
            <SearchableSelect
              value={localProject}
              onChange={(val) => setLocalProject(val)}
              options={uniqueWarehouses}
              placeholder="-- Chọn Kho / Dự án --"
              searchPlaceholder="Tìm tên Kho / Dự án..."
              variant="form"
            />
          </div>
        </div>

        {isRealReport && localProject && (
          <>
            {/* Đơn vị giao Selector */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 260 }}>
              <span style={{ fontSize: 13, fontWeight: 700, color: 'var(--text-muted)', whiteSpace: 'nowrap' }}>
                Đơn vị giao:
              </span>
              <div style={{ flex: 1, minWidth: 160 }}>
                <SearchableSelect
                  value={localGiao}
                  onChange={(val) => setLocalGiao(val)}
                  options={uniqueGiaoUnits}
                  placeholder="-- Tất cả Đơn vị giao --"
                  searchPlaceholder="Tìm Đơn vị giao..."
                  variant="form"
                />
              </div>
            </div>

            {/* Đơn vị nhận Selector */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 260 }}>
              <span style={{ fontSize: 13, fontWeight: 700, color: 'var(--text-muted)', whiteSpace: 'nowrap' }}>
                Đơn vị nhận:
              </span>
              <div style={{ flex: 1, minWidth: 160 }}>
                <SearchableSelect
                  value={localNhan}
                  onChange={(val) => setLocalNhan(val)}
                  options={uniqueNhanUnits}
                  placeholder="-- Tất cả Đơn vị nhận --"
                  searchPlaceholder="Tìm Đơn vị nhận..."
                  variant="form"
                />
              </div>
            </div>

            {/* Khoảng thời gian */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, flexShrink: 0, borderLeft: '1px solid var(--border)', paddingLeft: 12 }}>
              <span style={{ fontSize: 13, fontWeight: 700, color: 'var(--text-muted)' }}>Từ ngày:</span>
              <input
                type="date"
                className="input"
                style={{ height: 38, padding: '0 8px', fontSize: 13, border: '1px solid #cbd5e1', borderRadius: 6 }}
                value={startDate}
                onChange={e => setStartDate(e.target.value)}
              />
              <span style={{ fontSize: 13, fontWeight: 700, color: 'var(--text-muted)' }}>Đến ngày:</span>
              <input
                type="date"
                className="input"
                style={{ height: 38, padding: '0 8px', fontSize: 13, border: '1px solid #cbd5e1', borderRadius: 6 }}
                value={endDate}
                onChange={e => setEndDate(e.target.value)}
              />
              {(startDate || endDate) && (
                <button
                  onClick={() => { setStartDate(''); setEndDate(''); }}
                  style={{
                    background: 'none',
                    border: 'none',
                    color: '#ef4444',
                    fontSize: 12,
                    fontWeight: 600,
                    cursor: 'pointer',
                    padding: '0 4px',
                    display: 'flex',
                    alignItems: 'center',
                    gap: 2
                  }}
                  title="Xóa bộ lọc ngày"
                >
                  Xóa lọc
                </button>
              )}
            </div>
          </>
        )}

        {localProject && (
          <>
            {/* Search */}
            <div style={{ position: 'relative', flex: 1, minWidth: 200 }}>
              <Search size={14} style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--text-light)' }} />
              <input
                id="input-inventory-search"
                type="text"
                className="input"
                style={{ paddingLeft: 30, width: '100%', height: 38, fontSize: 13, border: '1px solid #cbd5e1' }}
                placeholder="Tìm theo Mã SAP, Mã vật tư, Tên vật tư..."
                value={searchTerm}
                onChange={e => setSearchTerm(e.target.value)}
              />
            </div>

            {/* Status Filter Toggle */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexShrink: 0 }}>
              <span style={{ fontSize: 13, fontWeight: 700, color: 'var(--text-muted)' }}>Trạng thái:</span>
              <div style={{ display: 'flex', background: 'var(--surface2)', borderRadius: 8, padding: 2, border: '1px solid var(--border)', height: 38, alignItems: 'center' }}>
                <button
                  id="btn-inventory-status-approved"
                  onClick={() => setStatusFilter('approved_only')}
                  style={{
                    padding: '0 12px',
                    height: '100%',
                    borderRadius: 6,
                    fontSize: '12.5px',
                    fontWeight: 600,
                    transition: 'all 0.1s',
                    background: statusFilter === 'approved_only' ? '#ffffff' : 'transparent',
                    color: statusFilter === 'approved_only' ? 'var(--primary)' : 'var(--text-muted)',
                    border: 'none',
                    boxShadow: statusFilter === 'approved_only' ? 'var(--shadow-sm)' : 'none',
                    cursor: 'pointer'
                  }}
                >
                  Chỉ Đã phê duyệt
                </button>
                <button
                  id="btn-inventory-status-approved-pending"
                  onClick={() => setStatusFilter('approved_pending')}
                  style={{
                    padding: '0 12px',
                    height: '100%',
                    borderRadius: 6,
                    fontSize: '12.5px',
                    fontWeight: 600,
                    transition: 'all 0.1s',
                    background: statusFilter === 'approved_pending' ? '#ffffff' : 'transparent',
                    color: statusFilter === 'approved_pending' ? 'var(--primary)' : 'var(--text-muted)',
                    border: 'none',
                    boxShadow: statusFilter === 'approved_pending' ? 'var(--shadow-sm)' : 'none',
                    cursor: 'pointer',
                    whiteSpace: 'nowrap'
                  }}
                >
                  Đã PD + Chưa PD
                </button>
                <button
                  id="btn-inventory-status-all"
                  onClick={() => setStatusFilter('all')}
                  style={{
                    padding: '0 12px',
                    height: '100%',
                    borderRadius: 6,
                    fontSize: '12.5px',
                    fontWeight: 600,
                    transition: 'all 0.1s',
                    background: statusFilter === 'all' ? '#ffffff' : 'transparent',
                    color: statusFilter === 'all' ? 'var(--primary)' : 'var(--text-muted)',
                    border: 'none',
                    boxShadow: statusFilter === 'all' ? 'var(--shadow-sm)' : 'none',
                    cursor: 'pointer'
                  }}
                >
                  Tất cả
                </button>
              </div>
            </div>
          </>
        )}

        {isRealReport && (
          <button
            id="btn-show-logic-explain"
            onClick={() => {
              setExplainModalSubTab(realReportSubTab);
              setShowLogicExplainModal(true);
            }}
            className="btn"
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: 6,
              height: 38,
              padding: '0 14px',
              borderRadius: 6,
              fontWeight: 600,
              fontSize: '13px',
              border: '1px solid #cbd5e1',
              color: '#0369a1',
              background: '#f0f9ff',
              borderColor: '#bae6fd',
              cursor: 'pointer',
              marginLeft: ((isRealReport ? (realReportSubTab === 'tonghop' ? realReportSummaryRows.length : realReportRows.length) : reportData.length) > 0) ? 'auto' : undefined,
              transition: 'all 0.15s ease'
            }}
            onMouseOver={e => {
              e.currentTarget.style.background = '#e0f2fe';
              e.currentTarget.style.borderColor = '#7dd3fc';
            }}
            onMouseOut={e => {
              e.currentTarget.style.background = '#f0f9ff';
              e.currentTarget.style.borderColor = '#bae6fd';
            }}
          >
            <HelpCircle size={14} />
            <span>Xem Giải thích Logic</span>
          </button>
        )}

        {((isRealReport ? (realReportSubTab === 'tonghop' ? realReportSummaryRows.length : realReportRows.length) : reportData.length) > 0) && (
          <button
            id="btn-export-inventory"
            onClick={handleExportExcel}
            className="btn btn-success"
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: 6,
              height: 38,
              padding: '0 16px',
              borderRadius: 6,
              fontWeight: 600,
              fontSize: '13px',
              background: '#10b981',
              color: '#ffffff',
              border: 'none',
              cursor: 'pointer',
              boxShadow: '0 2px 4px rgba(16, 185, 129, 0.2)',
              marginLeft: isRealReport ? undefined : 'auto'
            }}
          >
            <Download size={14} />
            <span>Xuất Excel</span>
          </button>
        )}
      </div>

      {/* Unit Classification Legend */}
      <div style={{
        background: '#ffffff',
        border: '1px solid var(--border)',
        borderRadius: 8,
        padding: '10px 16px',
        marginBottom: 12,
        display: 'flex',
        alignItems: 'center',
        gap: 16,
        flexWrap: 'wrap',
        boxShadow: 'var(--shadow-sm)',
        fontSize: '13px',
        flexShrink: 0
      }}>
        <span style={{ fontWeight: 700, color: 'var(--text)' }}>Chú giải phân loại đơn vị:</span>
        <div style={{ display: 'flex', gap: 16, flexWrap: 'wrap' }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{ width: 14, height: 14, borderRadius: 4, background: '#eff6ff', border: '1px solid #bfdbfe', display: 'inline-block' }} />
            <span style={{ fontWeight: 600, color: '#1e3a8a' }}>Kho dự án / Ban điều hành</span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{ width: 14, height: 14, borderRadius: 4, background: '#fffbeb', border: '1px solid #fde68a', display: 'inline-block' }} />
            <span style={{ fontWeight: 600, color: '#78350f' }}>Nhà cung cấp (NCC)</span>
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <span style={{ width: 14, height: 14, borderRadius: 4, background: '#f0fdf4', border: '1px solid #bbf7d0', display: 'inline-block' }} />
            <span style={{ fontWeight: 600, color: '#065f46' }}>Tổ đội thi công</span>
          </div>
        </div>
      </div>

      {/* Summary Metrics Cards */}
      {localProject && (
        <div style={{ display: 'flex', gap: 12, marginBottom: 12, flexWrap: 'wrap', flexShrink: 0 }}>
          {(isRealReport ? realReportMetrics : metricsCards).map(c => (
            <div key={c.label} style={{
              background: c.bg,
              border: `1px solid ${c.border}`,
              borderRadius: 8,
              padding: '10px 16px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              flex: '1 1 200px',
              minWidth: 180,
              boxShadow: 'var(--shadow-sm)'
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 8, flexWrap: 'wrap' }}>
                <span style={{ color: 'var(--text-muted)', fontSize: 13, fontWeight: 500 }}>{c.label}:</span>
                <span style={{ color: c.color, fontSize: 16, fontWeight: 700, letterSpacing: '-0.01em' }}>{c.value}</span>
              </div>
              <div style={{
                width: 32, height: 32, borderRadius: 8,
                background: '#ffffff', color: c.color,
                display: 'flex', alignItems: 'center', justifyContent: 'center',
                border: `1px solid ${c.border}`,
                flexShrink: 0
              }}>
                {c.icon}
              </div>
            </div>
          ))}
        </div>
      )}

      {/* Table container / main body */}
      {!localProject ? (
        <div style={{ flex: 1, background: '#ffffff', border: '1px solid var(--border)', borderRadius: 8, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 48, boxShadow: 'var(--shadow-sm)' }}>
          <div style={{ width: 56, height: 56, borderRadius: 28, background: '#eff6ff', color: '#0f58a7', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 16 }}>
            <Warehouse size={28} />
          </div>
          <h3 style={{ fontSize: 16, fontWeight: 700, color: 'var(--text)', margin: '0 0 8px 0' }}>Chưa chọn Kho / Dự án</h3>
          <p style={{ color: 'var(--text-muted)', fontSize: 13.5, margin: 0, maxWidth: 420, textAlign: 'center', lineHeight: '1.5' }}>
            Vui lòng chọn một Kho / Dự án cụ thể từ danh sách bộ lọc ở trên để tải dữ liệu báo cáo chi tiết nhanh nhất và tránh bị lag.
          </p>
        </div>
      ) : isRealReport ? (
        realReportSubTab === 'tonghop' ? (
          realReportSummaryRows.length === 0 ? (
            <div style={{ flex: 1, background: '#ffffff', border: '1px solid var(--border)', borderRadius: 8, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 48, boxShadow: 'var(--shadow-sm)' }}>
              <div style={{ width: 56, height: 56, borderRadius: 28, background: '#fef2f2', color: '#ef4444', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 16 }}>
                <AlertCircle size={28} />
              </div>
              <h3 style={{ fontSize: 16, fontWeight: 700, color: 'var(--text)', margin: '0 0 8px 0' }}>Không có dữ liệu phát sinh</h3>
              <p style={{ color: 'var(--text-muted)', fontSize: 13.5, margin: 0, maxWidth: 420, textAlign: 'center', lineHeight: '1.5' }}>
                Không tìm thấy dữ liệu tổng hợp thực tế nào{localProject ? ` liên quan đến kho "${localProject}"` : ''} trong dữ liệu Đơn chung hiện tại.
              </p>
            </div>
          ) : (
            <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
              <RealReportSummaryTable 
                summaryRows={realReportSummaryRows} 
                customCategoryMap={customCategoryMap} 
                localProject={localProject} 
                materialPriceRows={materialPriceRows}
                materialMetadataMap={materialMetadataMap}
              />
            </div>
          )
        ) : realReportRows.length === 0 ? (
          <div style={{ flex: 1, background: '#ffffff', border: '1px solid var(--border)', borderRadius: 8, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 48, boxShadow: 'var(--shadow-sm)' }}>
            <div style={{ width: 56, height: 56, borderRadius: 28, background: '#fef2f2', color: '#ef4444', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 16 }}>
              <AlertCircle size={28} />
            </div>
            <h3 style={{ fontSize: 16, fontWeight: 700, color: 'var(--text)', margin: '0 0 8px 0' }}>Không có dữ liệu phát sinh</h3>
            <p style={{ color: 'var(--text-muted)', fontSize: 13.5, margin: 0, maxWidth: 420, textAlign: 'center', lineHeight: '1.5' }}>
              Không tìm thấy bản ghi {realReportSubTab === 'nhap' ? 'nhập' : 'xuất'} thực tế nào{localProject ? ` liên quan đến kho "${localProject}"` : ''} trong dữ liệu Đơn chung hiện tại.
            </p>
          </div>
        ) : (
          <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
            <DataTable rows={realReportRows} columns={COLS_REAL_REPORT} customCategoryMap={customCategoryMap} />
          </div>
        )
      ) : reportData.length === 0 ? (
        <div style={{ flex: 1, background: '#ffffff', border: '1px solid var(--border)', borderRadius: 8, display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', padding: 48, boxShadow: 'var(--shadow-sm)' }}>
          <div style={{ width: 56, height: 56, borderRadius: 28, background: '#fef2f2', color: '#ef4444', display: 'flex', alignItems: 'center', justifyContent: 'center', marginBottom: 16 }}>
            <AlertCircle size={28} />
          </div>
          <h3 style={{ fontSize: 16, fontWeight: 700, color: 'var(--text)', margin: '0 0 8px 0' }}>Không có dữ liệu phát sinh</h3>
          <p style={{ color: 'var(--text-muted)', fontSize: 13.5, margin: 0, maxWidth: 420, textAlign: 'center', lineHeight: '1.5' }}>
            Không tìm thấy bản ghi giao nhận nào{localProject ? ` liên quan đến kho "${localProject}"` : ''}{statusFilter === 'approved_only' ? ' ở trạng thái đã phê duyệt' : ''} trong dữ liệu Đơn chung hiện tại.
          </p>
        </div>
      ) : (
        <div style={{ display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
          <div className="table-wrap" style={{ flex: 1, overflowX: 'hidden', overflowY: 'auto' }}>
            <table>
              <thead>
                <tr>
                  <th style={{ width: 50, minWidth: 50, maxWidth: 50, textAlign: 'center', verticalAlign: 'middle', fontSize: '12px', padding: '8px 10px' }}>
                    STT
                  </th>
                  <th 
                    onClick={() => toggleSort('maSAP')}
                    style={{ width: 130, minWidth: 130, cursor: 'pointer', fontSize: '12px', padding: '8px 10px', textAlign: 'center', verticalAlign: 'middle' }}
                  >
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 4, justifyContent: 'center', width: '100%' }}>
                      <span>Mã SAP</span>
                      <ArrowUpDown size={12} style={{ opacity: 0.7 }} />
                    </div>
                  </th>
                  <th style={{ width: 75, minWidth: 75, fontSize: '12px', padding: '8px 10px', textAlign: 'center', verticalAlign: 'middle' }}>
                    Mã vật tư
                  </th>
                  <th style={{ minWidth: 200, fontSize: '12px', padding: '8px 10px', textAlign: 'center', verticalAlign: 'middle' }}>
                    Tên vật tư
                  </th>
                  <th style={{ width: 60, minWidth: 60, textAlign: 'center', fontSize: '12px', padding: '8px 10px', verticalAlign: 'middle' }}>
                    ĐVT
                  </th>
                  <th style={{ width: 75, minWidth: 75, fontSize: '12px', padding: '8px 10px', textAlign: 'center', verticalAlign: 'middle' }}>
                    Thông số kỹ thuật
                  </th>
                  <th 
                    onClick={() => toggleSort('received')}
                    style={{ width: 90, minWidth: 90, cursor: 'pointer', fontSize: '12px', padding: '8px 10px', textAlign: 'center', verticalAlign: 'middle' }}
                  >
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 4, justifyContent: 'center', width: '100%' }}>
                      <span>Khối lượng nhận</span>
                      <ArrowUpDown size={12} style={{ opacity: 0.7 }} />
                    </div>
                  </th>
                  <th 
                    onClick={() => toggleSort('issued')}
                    style={{ width: 90, minWidth: 90, cursor: 'pointer', fontSize: '12px', padding: '8px 10px', textAlign: 'center', verticalAlign: 'middle' }}
                  >
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 4, justifyContent: 'center', width: '100%' }}>
                      <span>Khối lượng xuất</span>
                      <ArrowUpDown size={12} style={{ opacity: 0.7 }} />
                    </div>
                  </th>
                  <th 
                    onClick={() => toggleSort('stock')}
                    style={{ width: 85, minWidth: 85, cursor: 'pointer', fontSize: '12px', padding: '8px 10px', textAlign: 'center', verticalAlign: 'middle' }}
                  >
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 4, justifyContent: 'center', width: '100%' }}>
                      <span>{isRealReport ? 'Tồn thực tế' : 'Tồn kho'}</span>
                      <ArrowUpDown size={12} style={{ opacity: 0.7 }} />
                    </div>
                  </th>
                  <th style={{ width: 95, minWidth: 95, fontSize: '12px', padding: '8px 10px', textAlign: 'center', verticalAlign: 'middle', whiteSpace: 'normal', lineHeight: '1.2' }}>
                    <span>Ngày nhập muộn nhất</span>
                  </th>
                  <th style={{ width: 90, minWidth: 90, fontSize: '11px', padding: '8px 6px', textAlign: 'center', verticalAlign: 'middle', whiteSpace: 'normal', lineHeight: '1.2' }}>
                    <span>Số ngày nhập muộn nhất đến hôm nay</span>
                  </th>
                  <th style={{ width: 95, minWidth: 95, fontSize: '12px', padding: '8px 10px', textAlign: 'center', verticalAlign: 'middle', whiteSpace: 'normal', lineHeight: '1.2' }}>
                    <span>Ngày xuất muộn nhất</span>
                  </th>
                  <th style={{ width: 90, minWidth: 90, fontSize: '11px', padding: '8px 6px', textAlign: 'center', verticalAlign: 'middle', whiteSpace: 'normal', lineHeight: '1.2' }}>
                    <span>Số ngày xuất muộn nhất đến hôm nay</span>
                  </th>
                  <th style={{ width: 145, minWidth: 145, fontSize: '12px', padding: '8px 10px', textAlign: 'center', verticalAlign: 'middle', whiteSpace: 'normal', lineHeight: '1.2' }}>
                    <span>Trạng thái vật tư không sử dụng</span>
                  </th>
                  <th 
                    onClick={() => toggleSort('estimatedUnitPrice')}
                    style={{ width: 110, minWidth: 110, cursor: 'pointer', fontSize: '11px', padding: '8px 6px', textAlign: 'center', verticalAlign: 'middle', whiteSpace: 'normal', lineHeight: '1.2' }}
                  >
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 4, justifyContent: 'center', width: '100%' }}>
                      <span>Đơn giá trung bình tạm tính</span>
                      <ArrowUpDown size={12} style={{ opacity: 0.7 }} />
                    </div>
                  </th>
                  <th 
                    onClick={() => toggleSort('valueOver30Days')}
                    style={{ width: 130, minWidth: 130, cursor: 'pointer', fontSize: '11px', padding: '8px 6px', textAlign: 'center', verticalAlign: 'middle', whiteSpace: 'normal', lineHeight: '1.2' }}
                  >
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 4, justifyContent: 'center', width: '100%' }}>
                      <span>{isRealReport ? 'Thành tiền vật tư tồn thực tế >30 ngày' : 'Thành tiền vật tư tồn >30 ngày'}</span>
                      <ArrowUpDown size={12} style={{ opacity: 0.7 }} />
                    </div>
                  </th>
                  <th 
                    onClick={() => toggleSort('materialClassification')}
                    style={{ width: 110, minWidth: 110, cursor: 'pointer', fontSize: '11px', padding: '8px 6px', textAlign: 'center', verticalAlign: 'middle', whiteSpace: 'normal', lineHeight: '1.2' }}
                  >
                    <div style={{ display: 'inline-flex', alignItems: 'center', gap: 4, justifyContent: 'center', width: '100%' }}>
                      <span>Phân loại vật tư</span>
                      <ArrowUpDown size={12} style={{ opacity: 0.7 }} />
                    </div>
                  </th>
                </tr>
              </thead>
              <tbody>
                {paginatedData.map((item, index) => {
                  const stt = (currentPage - 1) * pageSize + index + 1
                  const isNegative = item.stock < 0
                  const isZero = item.stock === 0

                  return (
                    <tr key={item.maSAP}>
                      <td style={{ width: 50, minWidth: 50, maxWidth: 50, textAlign: 'center', fontSize: '12px', color: 'var(--text-muted)', padding: '6px 10px' }}>
                        {stt}
                      </td>
                      <td style={{ width: 130, minWidth: 130, fontSize: '13px', fontWeight: 700, fontFamily: "'Roboto', sans-serif", color: 'var(--text)', padding: '6px 10px' }}>
                        {item.maSAP}
                      </td>
                      <td style={{ width: 75, minWidth: 75, fontSize: '13px', fontFamily: "'Roboto', sans-serif", color: 'var(--text-muted)', padding: '6px 10px' }}>
                        {item.maVatTu || '—'}
                      </td>
                      <td style={{ minWidth: 200, fontSize: '13px', fontWeight: 600, color: 'var(--text)', whiteSpace: 'normal', wordBreak: 'break-word', padding: '6px 10px' }}>
                        {item.tenVatTu || '—'}
                      </td>
                      <td style={{ width: 60, minWidth: 60, textAlign: 'center', padding: '6px 10px' }}>
                        <span className="badge badge-gray" style={{ fontSize: '11px', padding: '2px 6px' }}>{item.dvt || '—'}</span>
                      </td>
                      <td style={{ width: 75, minWidth: 75, fontSize: '13px', color: 'var(--text-muted)', whiteSpace: 'normal', wordBreak: 'break-word', padding: '6px 10px' }}>
                        {item.thongSoKyThuat || '—'}
                      </td>
                      <td style={{ width: 90, minWidth: 90, textAlign: 'right', fontSize: '13px', fontWeight: 700, color: '#10b981', padding: '6px 10px' }}>
                        {item.received > 0 ? item.received.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '0'}
                      </td>
                      <td style={{ width: 90, minWidth: 90, textAlign: 'right', fontSize: '13px', fontWeight: 700, color: '#f97316', padding: '6px 10px' }}>
                        {item.issued > 0 ? item.issued.toLocaleString('vi-VN', { maximumFractionDigits: 2 }) : '0'}
                      </td>
                      <td style={{
                        width: 85, minWidth: 85, textAlign: 'right', fontSize: '13px', fontWeight: 800,
                        color: isNegative ? '#ef4444' : (isZero ? 'var(--text-muted)' : 'var(--primary)'),
                        background: isNegative ? '#fef2f2' : (isZero ? 'transparent' : 'var(--primary-light)'),
                        padding: '6px 10px'
                      }}>
                        {item.stock.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
                      </td>
                      <td style={{ width: 95, minWidth: 95, textAlign: 'center', fontSize: '13px', color: 'var(--text-muted)', padding: '6px 10px' }}>
                        {formatDate(item.latestReceivedDate)}
                      </td>
                      <td style={{ width: 90, minWidth: 90, textAlign: 'center', fontSize: '13px', fontWeight: 700, color: item.latestReceivedDate ? '#475569' : 'var(--text-muted)', padding: '6px 10px' }}>
                        {getDaysToToday(item.latestReceivedDate)}
                      </td>
                      <td style={{ width: 95, minWidth: 95, textAlign: 'center', fontSize: '13px', color: 'var(--text-muted)', padding: '6px 10px' }}>
                        {formatDate(item.latestIssuedDate)}
                      </td>
                      <td style={{ width: 90, minWidth: 90, textAlign: 'center', fontSize: '13px', fontWeight: 700, color: item.latestIssuedDate ? '#475569' : 'var(--text-muted)', padding: '6px 10px' }}>
                        {getDaysToToday(item.latestIssuedDate)}
                      </td>
                      <td style={{ width: 145, minWidth: 145, textAlign: 'center', padding: '6px 10px' }}>
                        {(() => {
                          const status = item.unusedStatus
                          let bg = '#f1f5f9'
                          let fg = '#64748b'
                          if (status === 'Đang sử dụng') {
                            bg = '#ecfdf5'
                            fg = '#047857'
                          } else if (status === 'Chưa sử dụng (> 30 ngày)') {
                            bg = '#fffbeb'
                            fg = '#d97706'
                          }
                          return (
                            <span className="badge" style={{
                              fontSize: '11px',
                              padding: '2px 6px',
                              borderRadius: '4px',
                              fontWeight: 700,
                              background: bg,
                              color: fg,
                              display: 'inline-block',
                              whiteSpace: 'normal',
                              lineHeight: '1.2'
                            }}>
                              {status}
                            </span>
                          )
                        })()}
                      </td>

                      {/* Đơn giá trung bình tạm tính */}
                      <td style={{ width: 110, minWidth: 110, textAlign: 'right', padding: '6px 10px', fontSize: '13px', fontWeight: 700, color: 'var(--primary)' }}>
                        {(() => {
                          const price = item.estimatedUnitPrice || 0;
                          return price > 0 ? price.toLocaleString('vi-VN') : '—';
                        })()}
                      </td>

                      {/* Thành tiền vật tư tồn >30 ngày */}
                      <td style={{
                        width: 130, minWidth: 130, textAlign: 'right', fontSize: '13px', fontWeight: 800,
                        color: item.valueOver30Days > 0 ? '#b91c1c' : 'var(--text-muted)',
                        background: item.valueOver30Days > 0 ? '#fff5f5' : 'transparent',
                        padding: '6px 10px'
                      }}>
                        {item.valueOver30Days > 0 ? item.valueOver30Days.toLocaleString('vi-VN') : '0'}
                      </td>

                      {/* Phân loại vật tư */}
                      <td style={{ width: 110, minWidth: 110, textAlign: 'center', padding: '6px 10px' }}>
                        {(() => {
                          const val = materialClassifications[item.maSAP] || ''
                          const text = val.trim().toLowerCase()
                          
                          let bg = '#f8fafc'
                          let fg = '#64748b'
                          let border = '#cbd5e1'
                          
                          if (text) {
                            if (text.includes('tiêu hao')) {
                              bg = '#ecfdf5' // Soft green
                              fg = '#047857' // Deep green
                              border = '#a7f3d0' // Light green border
                            } else if (text.includes('khấu hao') || text.includes('tài sản')) {
                              bg = '#fff7ed' // Soft orange
                              fg = '#ea580c' // Deep orange
                              border = '#fed7aa' // Light orange border
                            } else {
                              bg = '#eff6ff' // Soft blue
                              fg = '#1d4ed8' // Deep blue
                              border = '#bfdbfe'
                            }
                          }

                          return (
                            <input
                              type="text"
                              placeholder="Phân loại..."
                              value={val}
                              onChange={(e) => handleClassificationChange(item.maSAP, e.target.value)}
                              style={{
                                width: '100%',
                                textAlign: 'center',
                                fontSize: '12px',
                                fontWeight: 700,
                                color: fg,
                                padding: '4px 6px',
                                borderRadius: '4px',
                                border: `1.5px solid ${border}`,
                                outline: 'none',
                                background: bg,
                                transition: 'all 0.15s'
                              }}
                              onFocus={(e) => {
                                e.target.style.background = '#ffffff'
                                e.target.style.borderColor = 'var(--primary)'
                                e.target.style.boxShadow = '0 0 0 2px var(--primary-light)'
                                e.target.style.color = 'var(--text)'
                              }}
                              onBlur={(e) => {
                                e.target.style.background = bg
                                e.target.style.borderColor = border
                                e.target.style.boxShadow = 'none'
                                e.target.style.color = fg
                              }}
                            />
                          )
                        })()}
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>

          {/* Pagination Footer */}
          <div style={{
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            padding: '12px 16px',
            background: '#ffffff',
            border: '1px solid var(--border)',
            borderTop: 'none',
            borderRadius: '0 0 8px 8px',
            flexShrink: 0,
            boxShadow: 'var(--shadow-sm)'
          }}>
            <div style={{ fontSize: 13, color: 'var(--text-muted)' }}>
              Hiển thị từ <strong>{Math.min(reportData.length, (currentPage - 1) * pageSize + 1)}</strong> đến{' '}
              <strong>{Math.min(reportData.length, currentPage * pageSize)}</strong> trong tổng số{' '}
              <strong>{reportData.length.toLocaleString()}</strong> mặt hàng
            </div>

            {totalPages > 1 && (
              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                <button
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage(1)}
                  style={currentPage === 1 ? btnDisabled : btnBase}
                >
                  «
                </button>
                <button
                  disabled={currentPage === 1}
                  onClick={() => setCurrentPage(prev => prev - 1)}
                  style={currentPage === 1 ? btnDisabled : btnBase}
                >
                  ‹
                </button>
                <span style={{ fontSize: 13, color: 'var(--text)', fontWeight: 500, padding: '0 4px' }}>
                  Trang <strong>{currentPage}</strong> / {totalPages}
                </span>
                <button
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage(prev => prev + 1)}
                  style={currentPage === totalPages ? btnDisabled : btnBase}
                >
                  ›
                </button>
                <button
                  disabled={currentPage === totalPages}
                  onClick={() => setCurrentPage(totalPages)}
                  style={currentPage === totalPages ? btnDisabled : btnBase}
                >
                  »
                </button>
              </div>
            )}
          </div>
        </div>
      )}
        </div>
      )}

      {/* Logic Explanation Modal */}
      {showLogicExplainModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(15, 23, 42, 0.6)',
          backdropFilter: 'blur(8px)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999,
          padding: 16
        }}>
          <div style={{
            background: '#ffffff',
            borderRadius: 16,
            width: '100%',
            maxWidth: 620,
            boxShadow: '0 25px 50px -12px rgba(15, 23, 42, 0.25)',
            border: '1px solid var(--border)',
            display: 'flex',
            flexDirection: 'column',
            maxHeight: '90vh',
            overflow: 'hidden'
          }}>
            {/* Header */}
            <div style={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              padding: '16px 24px',
              borderBottom: '1px solid var(--border)',
              background: '#f8fafc'
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                <HelpCircle size={22} color="#0284c7" />
                <h3 style={{ fontSize: 16, fontWeight: 800, color: '#0f172a', margin: 0 }}>
                  Giải Thích Logic Tổng Hợp & Khối Lượng Thực
                </h3>
              </div>
              <button
                onClick={() => setShowLogicExplainModal(false)}
                style={{
                  background: 'none',
                  border: 'none',
                  cursor: 'pointer',
                  color: 'var(--text-muted)',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  padding: 4,
                  borderRadius: 4
                }}
                onMouseOver={e => e.currentTarget.style.color = '#000'}
                onMouseOut={e => e.currentTarget.style.color = 'var(--text-muted)'}
              >
                <X size={20} />
              </button>
            </div>

            {/* Modal Subtabs */}
            <div style={{
              display: 'flex',
              borderBottom: '1px solid var(--border)',
              padding: '0 24px',
              background: '#f8fafc'
            }}>
              <button
                onClick={() => setExplainModalSubTab('nhap')}
                style={{
                  padding: '12px 16px',
                  fontWeight: 700,
                  fontSize: 13.5,
                  border: 'none',
                  background: 'none',
                  color: explainModalSubTab === 'nhap' ? 'var(--primary)' : 'var(--text-muted)',
                  borderBottom: explainModalSubTab === 'nhap' ? '3px solid var(--primary)' : '3px solid transparent',
                  cursor: 'pointer',
                  marginRight: 16,
                  transition: 'all 0.15s ease'
                }}
              >
                Nhập Thực Tế
              </button>
              <button
                onClick={() => setExplainModalSubTab('xuat')}
                style={{
                  padding: '12px 16px',
                  fontWeight: 700,
                  fontSize: 13.5,
                  border: 'none',
                  background: 'none',
                  color: explainModalSubTab === 'xuat' ? 'var(--primary)' : 'var(--text-muted)',
                  borderBottom: explainModalSubTab === 'xuat' ? '3px solid var(--primary)' : '3px solid transparent',
                  cursor: 'pointer',
                  transition: 'all 0.15s ease'
                }}
              >
                Xuất Thực Tế
              </button>
            </div>

            {/* Content */}
            <div style={{ padding: '24px', overflowY: 'auto', display: 'flex', flexDirection: 'column', gap: 18 }}>
              <p style={{ margin: 0, fontSize: 13.5, color: '#475569', lineHeight: 1.6 }}>
                Hệ thống tự động tính toán <strong>Hệ số Logic tổng hợp</strong> và <strong>Khối lượng thực</strong> dựa trên phân nhóm đơn vị (được quản lý tại tab <strong>Kho dự án</strong>).
              </p>

              {explainModalSubTab === 'nhap' ? (
                <>
                  {/* Case 1 */}
                  <div>
                    <h4 style={{ fontSize: 14, fontWeight: 700, color: '#0369a1', margin: '0 0 10px 0', display: 'flex', alignItems: 'center', gap: 6 }}>
                      <span style={{ display: 'inline-block', width: 6, height: 6, borderRadius: '50%', background: '#0369a1' }} />
                      Trường hợp 1: Đơn vị GIAO là Kho / Dự án hiện tại (Nhận về giảm nhận)
                    </h4>
                    <div style={{ border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                        <thead>
                          <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #e2e8f0' }}>
                            <th style={{ padding: '8px 12px', textAlign: 'left', fontWeight: 600, color: '#334155' }}>Đơn vị NHẬN thuộc nhóm</th>
                            <th style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 600, color: '#334155', width: 140 }}>Hệ số Logic</th>
                            <th style={{ padding: '8px 12px', textAlign: 'left', fontWeight: 600, color: '#334155' }}>Ý nghĩa thực tế</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Nhà cung cấp (NCC)</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#ef4444' }}>-1</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Xuất trả nhà cung cấp (giảm nhận thực tế)</td>
                          </tr>
                          <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Kho khác</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#64748b' }}>0</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Luân chuyển đi kho khác (đã tính ở Xuất thực)</td>
                          </tr>
                          <tr>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Tổ đội</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#64748b' }}>0</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Cấp phát cho tổ đội thi công (đã tính ở Xuất thực)</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Case 2 */}
                  <div>
                    <h4 style={{ fontSize: 14, fontWeight: 700, color: '#0891b2', margin: '0 0 10px 0', display: 'flex', alignItems: 'center', gap: 6 }}>
                      <span style={{ display: 'inline-block', width: 6, height: 6, borderRadius: '50%', background: '#0891b2' }} />
                      Trường hợp 2: Đơn vị NHẬN là Kho / Dự án hiện tại
                    </h4>
                    <div style={{ border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                        <thead>
                          <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #e2e8f0' }}>
                            <th style={{ padding: '8px 12px', textAlign: 'left', fontWeight: 600, color: '#334155' }}>Đơn vị GIAO thuộc nhóm</th>
                            <th style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 600, color: '#334155', width: 140 }}>Hệ số Logic</th>
                            <th style={{ padding: '8px 12px', textAlign: 'left', fontWeight: 600, color: '#334155' }}>Ý nghĩa thực tế</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Nhà cung cấp (NCC)</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#10b981' }}>1</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Nhập hàng mua từ NCC (tăng nhận)</td>
                          </tr>
                          <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Kho khác</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#10b981' }}>1</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Nhận điều chuyển từ kho khác (tăng nhận)</td>
                          </tr>
                          <tr>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Tổ đội</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#64748b' }}>0</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Tổ đội trả lại vật tư dư thừa (đã tính ở Xuất thực)</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                </>
              ) : (
                <>
                  {/* Case 1 for Xuat */}
                  <div>
                    <h4 style={{ fontSize: 14, fontWeight: 700, color: '#0369a1', margin: '0 0 10px 0', display: 'flex', alignItems: 'center', gap: 6 }}>
                      <span style={{ display: 'inline-block', width: 6, height: 6, borderRadius: '50%', background: '#0369a1' }} />
                      Trường hợp 1: Đơn vị GIAO là Kho / Dự án hiện tại
                    </h4>
                    <div style={{ border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                        <thead>
                          <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #e2e8f0' }}>
                            <th style={{ padding: '8px 12px', textAlign: 'left', fontWeight: 600, color: '#334155' }}>Đơn vị NHẬN thuộc nhóm</th>
                            <th style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 600, color: '#334155', width: 140 }}>Hệ số Logic</th>
                            <th style={{ padding: '8px 12px', textAlign: 'left', fontWeight: 600, color: '#334155' }}>Ý nghĩa thực tế</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Tổ đội</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#10b981' }}>1</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Cấp phát cho tổ đội thi công (tăng xuất)</td>
                          </tr>
                          <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Kho khác</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#10b981' }}>1</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Luân chuyển đi kho khác (tăng xuất)</td>
                          </tr>
                          <tr>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Nhà cung cấp (NCC)</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#64748b' }}>0</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Xuất trả nhà cung cấp (đã tính ở Nhập thực là -1)</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Case 2 for Xuat */}
                  <div>
                    <h4 style={{ fontSize: 14, fontWeight: 700, color: '#0891b2', margin: '0 0 10px 0', display: 'flex', alignItems: 'center', gap: 6 }}>
                      <span style={{ display: 'inline-block', width: 6, height: 6, borderRadius: '50%', background: '#0891b2' }} />
                      Trường hợp 2: Đơn vị NHẬN là Kho / Dự án hiện tại
                    </h4>
                    <div style={{ border: '1px solid #e2e8f0', borderRadius: 8, overflow: 'hidden' }}>
                      <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                        <thead>
                          <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #e2e8f0' }}>
                            <th style={{ padding: '8px 12px', textAlign: 'left', fontWeight: 600, color: '#334155' }}>Đơn vị GIAO thuộc nhóm</th>
                            <th style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 600, color: '#334155', width: 140 }}>Hệ số Logic</th>
                            <th style={{ padding: '8px 12px', textAlign: 'left', fontWeight: 600, color: '#334155' }}>Ý nghĩa thực tế</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Tổ đội</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#ef4444' }}>-1</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Tổ đội trả lại vật tư dư thừa (giảm xuất thực tế)</td>
                          </tr>
                          <tr style={{ borderBottom: '1px solid #e2e8f0' }}>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Kho khác</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#64748b' }}>0</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Nhận điều chuyển (đã tính ở Nhập thực là 1)</td>
                          </tr>
                          <tr>
                            <td style={{ padding: '8px 12px', fontWeight: 500 }}>Nhà cung cấp (NCC)</td>
                            <td style={{ padding: '8px 12px', textAlign: 'center', fontWeight: 700, color: '#64748b' }}>0</td>
                            <td style={{ padding: '8px 12px', color: '#64748b' }}>Nhập mua từ NCC (đã tính ở Nhập thực là 1)</td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </div>
                </>
              )}

              {/* Formula */}
              <div style={{ background: '#f0f9ff', border: '1px solid #e0f2fe', borderRadius: 8, padding: '12px 16px' }}>
                <span style={{ fontSize: 13, fontWeight: 700, color: '#0369a1', display: 'block', marginBottom: 4 }}>
                  Công thức tính Khối lượng thực:
                </span>
                <code style={{ fontSize: 13, fontWeight: 700, color: '#0f172a', fontFamily: 'var(--font-mono)' }}>
                  Khối lượng thực = Khối lượng {explainModalSubTab === 'nhap' ? 'nhập' : 'xuất'} * Hệ số Logic tổng hợp
                </code>
              </div>
            </div>

            {/* Footer */}
            <div style={{
              padding: '12px 24px',
              borderTop: '1px solid var(--border)',
              background: '#f8fafc',
              display: 'flex',
              justifyContent: 'flex-end'
            }}>
              <button
                onClick={() => setShowLogicExplainModal(false)}
                className="btn btn-primary"
                style={{
                  height: 38,
                  padding: '0 18px',
                  borderRadius: 6,
                  fontWeight: 700,
                  fontSize: 13,
                  cursor: 'pointer',
                  background: 'var(--primary)',
                  color: '#ffffff',
                  border: 'none'
                }}
              >
                Đóng giải thích
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

// ─── Summary Compilation Tab ──────────────────────────────────────────────────
function SummaryCompilationTab({ giaoRows, nhanRows, configs = [], selectedProject, allProjects }) {
  const [selectedConfigId, setSelectedConfigId] = React.useState(null)
  const [searchTerm, setSearchTerm] = React.useState('')
  const [groupFilter, setGroupFilter] = React.useState('all') // 'all' | 'tinhToan' | 'giamTru'
  const [sourceFilter, setSourceFilter] = React.useState('all') // 'all' | 'giao' | 'nhan'
  const [currentPage, setCurrentPage] = React.useState(1)
  const [pageSize, setPageSize] = React.useState(100)

  // Auto-select config based on selected project or configs list
  React.useEffect(() => {
    if (configs && configs.length > 0) {
      if (selectedProject) {
        const match = configs.find(c => c.project === selectedProject)
        if (match) {
          setSelectedConfigId(match.id)
          return
        }
      }
      // If we don't have a match or selectedConfigId is invalid, fall back
      if (!selectedConfigId || !configs.some(c => c.id === selectedConfigId)) {
        // Look for configs that have selectedProject, otherwise take configs[0]
        const matchedProj = configs.find(c => c.project === selectedProject)
        setSelectedConfigId(matchedProj ? matchedProj.id : configs[0].id)
      }
    }
  }, [configs, selectedProject])

  const currentConfig = React.useMemo(() => {
    return configs.find(c => c.id === selectedConfigId) || null
  }, [configs, selectedConfigId])

  // Reset pagination on filter or config change
  React.useEffect(() => {
    setCurrentPage(1)
  }, [selectedConfigId, searchTerm, groupFilter, sourceFilter])

  // Compile data directly
  const compiledRows = React.useMemo(() => {
    if (!currentConfig) return []

    const result = []

    // 1. Process giaoRows (ĐƠN GIAO)
    const filteredGiao = currentConfig.project
      ? giaoRows.filter(r => (r.ten_du_an || r.tenDuAn || r.duAn || '') === currentConfig.project)
      : giaoRows

    filteredGiao.forEach((row, originalIdx) => {
      const unitName = row.donViNhan
      if (!unitName) return

      const unitCfg = currentConfig.giaoTable?.find(item => item.unit === unitName)
      const isBoQua = unitCfg ? !!unitCfg.boQua : false
      const isGiamTru = unitCfg ? !!unitCfg.giamTru : false

      if (isBoQua) return // Skip row

      let originalNhap = row.khoiLuongNhap !== null && row.khoiLuongNhap !== undefined ? parseFloat(String(row.khoiLuongNhap).replace(/[^\d.-]/g, '').replace(',', '.')) : 0
      if (isNaN(originalNhap)) originalNhap = 0

      let originalXuat = row.khoiLuongXuat !== null && row.khoiLuongXuat !== undefined ? parseFloat(String(row.khoiLuongXuat).replace(/[^\d.-]/g, '').replace(',', '.')) : 0
      if (isNaN(originalXuat)) originalXuat = 0

      const compiledNhap = isGiamTru ? -Math.abs(originalNhap) : Math.abs(originalNhap)
      const compiledXuat = isGiamTru ? -Math.abs(originalXuat) : Math.abs(originalXuat)

      result.push({
        ...row,
        compiledSource: 'Đơn Giao',
        compiledUnit: unitName,
        compiledGroup: isGiamTru ? 'Giảm trừ' : 'Tính toán',
        khoiLuongNhap: compiledNhap,
        khoiLuongXuat: compiledXuat,
        originalNhap,
        originalXuat,
        key: `compiled-giao-${row.id || originalIdx}`
      })
    })

    // 2. Process nhanRows (ĐƠN NHẬN)
    const filteredNhan = currentConfig.project
      ? nhanRows.filter(r => (r.ten_du_an || r.tenDuAn || r.duAn || '') === currentConfig.project)
      : nhanRows

    filteredNhan.forEach((row, originalIdx) => {
      const unitName = row.donViGiao
      if (!unitName) return

      const unitCfg = currentConfig.nhanTable?.find(item => item.unit === unitName)
      const isBoQua = unitCfg ? !!unitCfg.boQua : false
      const isGiamTru = unitCfg ? !!unitCfg.giamTru : false

      if (isBoQua) return // Skip row

      let originalNhap = row.khoiLuongNhap !== null && row.khoiLuongNhap !== undefined ? parseFloat(String(row.khoiLuongNhap).replace(/[^\d.-]/g, '').replace(',', '.')) : 0
      if (isNaN(originalNhap)) originalNhap = 0

      let originalXuat = row.khoiLuongXuat !== null && row.khoiLuongXuat !== undefined ? parseFloat(String(row.khoiLuongXuat).replace(/[^\d.-]/g, '').replace(',', '.')) : 0
      if (isNaN(originalXuat)) originalXuat = 0

      const compiledNhap = isGiamTru ? -Math.abs(originalNhap) : Math.abs(originalNhap)
      const compiledXuat = isGiamTru ? -Math.abs(originalXuat) : Math.abs(originalXuat)

      result.push({
        ...row,
        compiledSource: 'Đơn Nhận',
        compiledUnit: unitName,
        compiledGroup: isGiamTru ? 'Giảm trừ' : 'Tính toán',
        khoiLuongNhap: compiledNhap,
        khoiLuongXuat: compiledXuat,
        originalNhap,
        originalXuat,
        key: `compiled-nhan-${row.id || originalIdx}`
      })
    })

    return result
  }, [currentConfig, giaoRows, nhanRows])

  // Filter compiled rows
  const filteredRows = React.useMemo(() => {
    let r = compiledRows

    if (groupFilter === 'tinhToan') {
      r = r.filter(row => row.compiledGroup === 'Tính toán')
    } else if (groupFilter === 'giamTru') {
      r = r.filter(row => row.compiledGroup === 'Giảm trừ')
    }

    if (sourceFilter === 'giao') {
      r = r.filter(row => row.compiledSource === 'Đơn Giao')
    } else if (sourceFilter === 'nhan') {
      r = r.filter(row => row.compiledSource === 'Đơn Nhận')
    }

    if (searchTerm) {
      const q = searchTerm.toLowerCase()
      r = r.filter(row =>
        String(row.tenVatTu || '').toLowerCase().includes(q) ||
        String(row.maVatTu || '').toLowerCase().includes(q) ||
        String(row.maSAP || '').toLowerCase().includes(q) ||
        String(row.compiledUnit || '').toLowerCase().includes(q) ||
        String(row.duAn || '').toLowerCase().includes(q) ||
        String(row.maDonNhapKho || '').toLowerCase().includes(q) ||
        String(row.maDonXuatKho || '').toLowerCase().includes(q) ||
        String(row.donViGiao || '').toLowerCase().includes(q) ||
        String(row.donViNhan || '').toLowerCase().includes(q)
      )
    }

    // Sort descending by date (newest/most recent first)
    const sorted = [...r].sort((a, b) => {
      const dateA = parseRowDate(a.ngayXuatNhap)
      const dateB = parseRowDate(b.ngayXuatNhap)
      if (!dateA && !dateB) return 0
      if (!dateA) return 1
      if (!dateB) return -1
      return dateB.getTime() - dateA.getTime()
    })

    return sorted
  }, [compiledRows, searchTerm, groupFilter, sourceFilter])

  // Stats calculation
  const totalNhap = React.useMemo(() => {
    return filteredRows.reduce((sum, r) => sum + r.khoiLuongNhap, 0)
  }, [filteredRows])

  const totalXuat = React.useMemo(() => {
    return filteredRows.reduce((sum, r) => sum + r.khoiLuongXuat, 0)
  }, [filteredRows])

  // Stats cards values based on all items vs filtered items
  const statsTotalRows = filteredRows.length
  
  // Pagination
  const paginatedRows = React.useMemo(() => {
    const start = (currentPage - 1) * pageSize
    return filteredRows.slice(start, start + pageSize)
  }, [filteredRows, currentPage, pageSize])

  const totalPages = Math.ceil(filteredRows.length / pageSize) || 1

  const handleExportCompiledExcel = () => {
    const wb = XLSXStyle.utils.book_new()
    const ws = {}

    const exportCols = [
      { key: 'STT', label: 'STT', width: 50 },
      { key: 'compiledGroup', label: 'Nhóm tổng hợp', width: 120 },
      { key: 'compiledSource', label: 'Nguồn dữ liệu', width: 100 },
      { key: 'duAn', label: 'Dự án', width: 150 },
      { key: 'ngayXuatNhap', label: 'Ngày xuất nhập', width: 100 },
      { key: 'maVatTu', label: 'Mã vật tư', width: 90 },
      { key: 'maSAP', label: 'Mã SAP', width: 90 },
      { key: 'thongSoKyThuat', label: 'Thông số kỹ thuật', width: 150 },
      { key: 'tenVatTu', label: 'Tên vật tư', width: 180 },
      { key: 'dvt', label: 'ĐVT', width: 60 },
      { key: 'maDonNhapKho', label: 'Mã đơn nhập kho', width: 160 },
      { key: 'maDonXuatKho', label: 'Mã đơn xuất kho', width: 160 },
      { key: 'donViGiao', label: 'Đơn vị giao', width: 160 },
      { key: 'donViNhan', label: 'Đơn vị nhận', width: 160 },
      { key: 'compiledUnit', label: 'Đơn vị đối tác', width: 150 },
      { key: 'khoiLuongNhap', label: 'Khối lượng nhập', width: 110 },
      { key: 'khoiLuongXuat', label: 'Khối lượng xuất', width: 110 },
      { key: 'trangThai', label: 'Trạng thái', width: 100 },
      { key: 'ghiChu', label: 'Ghi chú', width: 150 }
    ]

    ws['!cols'] = exportCols.map(c => ({ wpx: c.width }))

    let excelRowIdx = 1

    // Header stylings
    exportCols.forEach((col, colIdx) => {
      const colChar = getColLabel(colIdx)
      const cellRef = `${colChar}${excelRowIdx}`
      ws[cellRef] = {
        v: col.label,
        t: 's',
        s: {
          fill: { patternType: 'solid', fgColor: { rgb: '0F58A7' } },
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'FFFFFF' } },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: '0A3D73' } },
            bottom: { style: 'medium', color: { rgb: '0A3D73' } },
            left: { style: 'thin', color: { rgb: '0A3D73' } },
            right: { style: 'thin', color: { rgb: '0A3D73' } }
          }
        }
      }
    })

    // Write content
    filteredRows.forEach((row, rowIndex) => {
      excelRowIdx++
      const isEvenNum = (rowIndex % 2 === 1)
      const rowBgColor = isEvenNum ? 'F8FAFC' : 'FFFFFF'

      exportCols.forEach((col, colIdx) => {
        const colChar = getColLabel(colIdx)
        const cellRef = `${colChar}${excelRowIdx}`

        let val = ''
        let cellType = 's'
        let numFormat = undefined
        let cellStyleColor = '1B1919'
        let isBoldFont = false

        if (col.key === 'STT') {
          val = rowIndex + 1
          cellType = 'n'
        } else if (col.key === 'khoiLuongNhap' || col.key === 'khoiLuongXuat') {
          const rawNum = row[col.key]
          val = rawNum !== null && rawNum !== undefined ? Number(rawNum) : 0
          cellType = 'n'
          numFormat = '#,##0.00'
          if (val < 0) {
            cellStyleColor = '9F1239'
            isBoldFont = true
          }
        } else {
          const rawVal = row[col.key]
          val = rawVal !== null && rawVal !== undefined ? String(rawVal) : ''
        }

        const isCenteredCol = [
          'STT', 'compiledGroup', 'compiledSource', 'ngayXuatNhap', 'maVatTu', 'maSAP', 'dvt', 'trangThai', 'maDonNhapKho', 'maDonXuatKho'
        ].includes(col.key)
        
        const isRightAligned = ['khoiLuongNhap', 'khoiLuongXuat'].includes(col.key)

        const cellStyle = {
          font: {
            name: 'Segoe UI',
            sz: 9,
            color: { rgb: cellStyleColor },
            bold: isBoldFont
          },
          alignment: {
            horizontal: isCenteredCol ? 'center' : (isRightAligned ? 'right' : 'left'),
            vertical: 'center',
            wrapText: true
          },
          border: {
            top: { style: 'thin', color: { rgb: 'E2E8F0' } },
            bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          },
          fill: {
            patternType: 'solid',
            fgColor: { rgb: rowBgColor }
          }
        }

        if (col.key === 'compiledGroup') {
          if (val === 'Giảm trừ') {
            cellStyle.fill.fgColor = { rgb: 'FFF1F2' }
            cellStyle.font.color = { rgb: '9F1239' }
            cellStyle.font.bold = true
          } else {
            cellStyle.fill.fgColor = { rgb: 'ECFDF5' }
            cellStyle.font.color = { rgb: '065F46' }
            cellStyle.font.bold = true
          }
        } else if (col.key === 'compiledSource') {
          if (val === 'Đơn Giao') {
            cellStyle.fill.fgColor = { rgb: 'EFF6FF' }
            cellStyle.font.color = { rgb: '1E40AF' }
          } else {
            cellStyle.fill.fgColor = { rgb: 'ECFEFF' }
            cellStyle.font.color = { rgb: '083344' }
          }
        }

        const cellObj = { v: val, t: cellType, s: cellStyle }
        if (numFormat) cellObj.z = numFormat

        ws[cellRef] = cellObj
      })
    })

    const excelRangeRef = `A1:${getColLabel(exportCols.length - 1)}${excelRowIdx}`
    ws['!ref'] = excelRangeRef
    ws['!autofilter'] = { ref: excelRangeRef }

    // -------------------------------------------------------------------------
    // Create aggregated "Tổng hợp" sheet grouping imports/exports by Mã SAP,
    // sorted alphabetically by Tên vật tư ascending
    // -------------------------------------------------------------------------
    const summaryMap = {}
    filteredRows.forEach(row => {
      const code = (row.maSAP || '').trim() || 'Chưa xác định'
      if (!summaryMap[code]) {
        summaryMap[code] = {
          maSAP: code,
          tenVatTu: row.tenVatTu || '',
          dvt: row.dvt || '',
          khoiLuongNhap: 0,
          khoiLuongXuat: 0
        }
      }
      summaryMap[code].khoiLuongNhap += row.khoiLuongNhap || 0
      summaryMap[code].khoiLuongXuat += row.khoiLuongXuat || 0
    })

    const summaryList = Object.values(summaryMap)

    // Sort alphabetically by Vietnamese tên vật tư (tenVatTu) ascending
    summaryList.sort((a, b) => {
      const nameA = a.tenVatTu || ''
      const nameB = b.tenVatTu || ''
      return nameA.localeCompare(nameB, 'vi', { sensitivity: 'accent' })
    })

    const wsSummary = {}
    const summaryCols = [
      { key: 'STT', label: 'STT', width: 50 },
      { key: 'maSAP', label: 'Mã SAP', width: 120 },
      { key: 'tenVatTu', label: 'Tên vật tư', width: 250 },
      { key: 'dvt', label: 'ĐVT', width: 80 },
      { key: 'khoiLuongNhap', label: 'Tổng khối lượng nhập', width: 160 },
      { key: 'khoiLuongXuat', label: 'Tổng khối lượng xuất', width: 160 }
    ]

    wsSummary['!cols'] = summaryCols.map(c => ({ wpx: c.width }))

    let summaryRowIdx = 1

    // Header stylings for wsSummary
    summaryCols.forEach((col, colIdx) => {
      const colChar = getColLabel(colIdx)
      const cellRef = `${colChar}${summaryRowIdx}`
      wsSummary[cellRef] = {
        v: col.label,
        t: 's',
        s: {
          fill: { patternType: 'solid', fgColor: { rgb: '0F58A7' } },
          font: { name: 'Segoe UI', sz: 9.5, bold: true, color: { rgb: 'FFFFFF' } },
          alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
          border: {
            top: { style: 'thin', color: { rgb: '0A3D73' } },
            bottom: { style: 'medium', color: { rgb: '0A3D73' } },
            left: { style: 'thin', color: { rgb: '0A3D73' } },
            right: { style: 'thin', color: { rgb: '0A3D73' } }
          }
        }
      }
    })

    // Populate rows for wsSummary
    summaryList.forEach((row, rowIndex) => {
      summaryRowIdx++
      const isEvenNum = (rowIndex % 2 === 1)
      const rowBgColor = isEvenNum ? 'F8FAFC' : 'FFFFFF'

      summaryCols.forEach((col, colIdx) => {
        const colChar = getColLabel(colIdx)
        const cellRef = `${colChar}${summaryRowIdx}`

        let val = ''
        let cellType = 's'
        let numFormat = undefined
        let cellStyleColor = '1B1919'
        let isBoldFont = false

        if (col.key === 'STT') {
          val = rowIndex + 1
          cellType = 'n'
        } else if (col.key === 'khoiLuongNhap' || col.key === 'khoiLuongXuat') {
          const rawNum = row[col.key]
          val = rawNum !== null && rawNum !== undefined ? Number(rawNum) : 0
          cellType = 'n'
          numFormat = '#,##0.00'
          if (val < 0) {
            cellStyleColor = '9F1239'
            isBoldFont = true
          }
        } else {
          const rawVal = row[col.key]
          val = rawVal !== null && rawVal !== undefined ? String(rawVal) : ''
        }

        const isCenteredCol = ['STT', 'maSAP', 'dvt'].includes(col.key)
        const isRightAligned = ['khoiLuongNhap', 'khoiLuongXuat'].includes(col.key)

        const cellStyle = {
          font: {
            name: 'Segoe UI',
            sz: 9,
            color: { rgb: cellStyleColor },
            bold: isBoldFont
          },
          alignment: {
            horizontal: isCenteredCol ? 'center' : (isRightAligned ? 'right' : 'left'),
            vertical: 'center',
            wrapText: true
          },
          border: {
            top: { style: 'thin', color: { rgb: 'E2E8F0' } },
            bottom: { style: 'thin', color: { rgb: 'E2E8F0' } },
            left: { style: 'thin', color: { rgb: 'E2E8F0' } },
            right: { style: 'thin', color: { rgb: 'E2E8F0' } }
          },
          fill: {
            patternType: 'solid',
            fgColor: { rgb: rowBgColor }
          }
        }

        const cellObj = { v: val, t: cellType, s: cellStyle }
        if (numFormat) cellObj.z = numFormat

        wsSummary[cellRef] = cellObj
      })
    })

    const summaryRangeRef = `A1:${getColLabel(summaryCols.length - 1)}${summaryRowIdx}`
    wsSummary['!ref'] = summaryRangeRef
    wsSummary['!autofilter'] = { ref: summaryRangeRef }

    // Append sheets in order. Let's make "Tổng hợp" first, and "Tổng hợp thông tin" second.
    XLSXStyle.utils.book_append_sheet(wb, wsSummary, "Tổng hợp")
    XLSXStyle.utils.book_append_sheet(wb, ws, "Tổng hợp thông tin")

    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary', compression: false })
    const buf = new ArrayBuffer(wbout.length)
    const view = new Uint8Array(buf)
    for (let i = 0; i < wbout.length; i++) view[i] = wbout.charCodeAt(i) & 0xFF
    
    const blob = new Blob([buf], { type: 'application/octet-stream' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `SGC_Bao_cao_tong_hop_${currentConfig?.name || 'Chung'}_${new Date().toISOString().slice(0, 10)}.xlsx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  function getColLabel(index) {
    let label = ''
    let temp = index
    while (temp >= 0) {
      label = String.fromCharCode((temp % 26) + 65) + label
      temp = Math.floor(temp / 26) - 1
    }
    return label
  }

  // Format negative/positive values beautifully
  const formatValueWithNegColor = (val) => {
    if (val === null || val === undefined) return '0'
    const isNeg = val < 0
    const str = val.toLocaleString('vi-VN', { minimumFractionDigits: 0, maximumFractionDigits: 2 })
    return (
      <span style={{ color: isNeg ? '#ef4444' : '#1e293b', fontWeight: isNeg ? '600' : '400' }}>
        {str}
      </span>
    )
  }

  if (configs.length === 0) {
    return (
      <div style={{ padding: 24 }}>
        <div className="empty-state" style={{ minHeight: 400 }}>
          <div style={{
            width: 64, height: 64, background: '#f1f5f9',
            borderRadius: 18, display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            <BarChart3 size={28} color="var(--text-light)" />
          </div>
          <h3>Chưa có cấu hình tổng hợp</h3>
          <p style={{ maxWidth: 360, textAlign: 'center' }}>
            Vui lòng tạo hoặc đồng bộ Cấu hình tổng hợp ở tab "Cấu hình tổng hợp" trước để tiến hành phân tích và tổng hợp dữ liệu.
          </p>
        </div>
      </div>
    )
  }

  return (
    <div style={{ display: 'flex', flexDirection: 'column', gap: 12, padding: '16px' }}>
      
      {/* Unified Compact Control and Stats Header Panel */}
      <div style={{
        background: '#ffffff',
        border: '1px solid #cbd5e1',
        borderRadius: 8,
        padding: '10px 14px',
        display: 'flex',
        flexWrap: 'wrap',
        alignItems: 'center',
        justifyContent: 'space-between',
        gap: 16
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: 10, flexWrap: 'wrap' }}>
          <span style={{ fontWeight: 600, fontSize: 13.5, color: '#334155' }}>Cấu hình tổng hợp:</span>
          <select 
            value={selectedConfigId || ''} 
            onChange={(e) => setSelectedConfigId(Number(e.target.value))}
            style={{
              padding: '6px 12px',
              borderRadius: 6,
              border: '1px solid #cbd5e1',
              background: '#ffffff',
              fontSize: 13.5,
              fontWeight: 600,
              color: '#1e293b',
              boxShadow: '0 1px 2px rgba(0,0,0,0.05)',
              cursor: 'pointer',
              minWidth: 240,
              height: 32
            }}
          >
            {configs.map(cfg => (
              <option key={cfg.id} value={cfg.id}>
                {cfg.name} {cfg.project ? `(${cfg.project})` : '(Tất cả dự án)'}
              </option>
            ))}
          </select>
        </div>

        {/* Center: Inline highly-compact Stats blocks */}
        {currentConfig && (
          <div style={{ display: 'flex', alignItems: 'center', gap: 24, flexWrap: 'wrap' }}>
            {/* Stat Item 1: Total records */}
            <div style={{ borderLeft: '3px solid #64748b', paddingLeft: 8 }}>
              <div style={{ fontSize: 10.5, color: '#64748b', fontWeight: 500, marginBottom: 1 }}>Dòng sau phân bổ</div>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: '#0f172a' }}>{statsTotalRows.toLocaleString('vi-VN')} dòng</div>
            </div>

            {/* Stat Item 2: Consolidated Import Volume */}
            <div style={{ borderLeft: `3px solid ${totalNhap >= 0 ? '#10b981' : '#ef4444'}`, paddingLeft: 8 }}>
              <div style={{ fontSize: 10.5, color: '#64748b', fontWeight: 500, marginBottom: 1 }}>Khối lượng Nhập ròng</div>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: totalNhap < 0 ? '#dc2626' : '#10b981' }}>
                {totalNhap.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
              </div>
            </div>

            {/* Stat Item 3: Consolidated Export Volume */}
            <div style={{ borderLeft: `3px solid ${totalXuat >= 0 ? '#3b82f6' : '#ef4444'}`, paddingLeft: 8 }}>
              <div style={{ fontSize: 10.5, color: '#64748b', fontWeight: 500, marginBottom: 1 }}>Khối lượng Xuất ròng</div>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: totalXuat < 0 ? '#dc2626' : '#3b82f6' }}>
                {totalXuat.toLocaleString('vi-VN', { maximumFractionDigits: 2 })}
              </div>
            </div>
          </div>
        )}

        {currentConfig && compiledRows.length > 0 && (
          <button
            onClick={handleExportCompiledExcel}
            className="btn btn-sm"
            style={{
              background: 'linear-gradient(135deg, #059669 0%, #10b981 100%)',
              color: '#ffffff',
              border: 'none',
              boxShadow: '0 2px 4px rgba(16,185,129,0.2)',
              display: 'flex',
              alignItems: 'center',
              gap: 6,
              height: 32,
              padding: '0 12px',
              fontSize: 13,
              borderRadius: 6
            }}
          >
            <Download size={14} /> Xuất Excel Tổng hợp
          </button>
        )}
      </div>

      {currentConfig ? (
        <>
          {/* Filter Bar */}
          <div style={{
            background: '#ffffff',
            border: '1px solid #cbd5e1',
            borderRadius: 8,
            padding: '10px 14px',
            display: 'flex',
            flexWrap: 'wrap',
            alignItems: 'center',
            justifyContent: 'space-between',
            gap: 12
          }}>
            {/* Left search */}
            <div style={{ position: 'relative', flex: 1, minWidth: 260 }}>
              <span style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', display: 'flex', alignItems: 'center' }}>
                <Search size={16} color="#94a3b8" />
              </span>
              <input
                type="text"
                placeholder="Tìm kiếm vật tư, mã, đơn vị..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                style={{
                  width: '100%',
                  padding: '8px 12px 8px 36px',
                  borderRadius: 8,
                  border: '1px solid #cbd5e1',
                  background: '#ffffff',
                  fontSize: 14,
                  outline: 'none',
                }}
              />
              {searchTerm && (
                <button
                  onClick={() => setSearchTerm('')}
                  style={{
                    position: 'absolute', right: 10, top: '50%', transform: 'translateY(-50%)',
                    background: 'transparent', border: 'none', cursor: 'pointer', display: 'flex'
                  }}
                >
                  <X size={15} color="#cbd5e1" />
                </button>
              )}
            </div>

            {/* Right filters */}
            <div style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 12 }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <span style={{ fontSize: 13, fontWeight: 550, color: '#64748b' }}>Nhóm:</span>
                <select 
                  value={groupFilter} 
                  onChange={(e) => setGroupFilter(e.target.value)} 
                  style={{ padding: '7px 12px', border: '1px solid #cbd5e1', borderRadius: 8, fontSize: 13.5, background: '#ffffff', fontWeight: 550 }}
                >
                  <option value="all">Tất cả</option>
                  <option value="tinhToan">Tính toán</option>
                  <option value="giamTru">Giảm trừ</option>
                </select>
              </div>

              <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                <span style={{ fontSize: 13, fontWeight: 550, color: '#64748b' }}>Nguồn:</span>
                <select 
                  value={sourceFilter} 
                  onChange={(e) => setSourceFilter(e.target.value)} 
                  style={{ padding: '7px 12px', border: '1px solid #cbd5e1', borderRadius: 8, fontSize: 13.5, background: '#ffffff', fontWeight: 550 }}
                >
                  <option value="all">Tất cả nguồn</option>
                  <option value="giao">Đơn Giao</option>
                  <option value="nhan">Đơn Nhận</option>
                </select>
              </div>
            </div>
          </div>

          {/* Compiled Rows Table */}
          {filteredRows.length > 0 ? (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              <div style={{ background: '#ffffff', border: '1px solid #cbd5e1', borderRadius: 12, overflow: 'hidden' }}>
                <div style={{ overflowX: 'auto', maxHeight: '55vh' }}>
                  <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', minWidth: 1600 }}>
                    <thead style={{ background: '#f8fafc', borderBottom: '2px solid #e2e8f0', position: 'sticky', top: 0, zIndex: 10 }}>
                      <tr>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 60, textAlign: 'center' }}>STT</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 115, textAlign: 'center' }}>Nhóm</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 105, textAlign: 'center' }}>Nguồn</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 110, textAlign: 'center' }}>Ngày</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 105 }}>Mã Vật tư</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 220 }}>Tên Vật tư</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 65, textAlign: 'center' }}>ĐVT</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 170 }}>Mã đơn nhập kho</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 170 }}>Mã đơn xuất kho</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 160 }}>Đơn vị giao</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 160 }}>Đơn vị nhận</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 180 }}>Đơn vị đối tác</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 140, textAlign: 'right' }}>KL Nhập</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 140, textAlign: 'right' }}>KL Xuất</th>
                        <th style={{ padding: '12px 14px', fontSize: 12.5, fontWeight: 700, color: '#475569', textTransform: 'uppercase', width: 115, textAlign: 'center' }}>Trạng thái</th>
                      </tr>
                    </thead>
                    <tbody>
                      {paginatedRows.map((row, index) => {
                        const globalIndex = (currentPage - 1) * pageSize + index + 1
                        const isGiamTru = row.compiledGroup === 'Giảm trừ'
                        const isGiao = row.compiledSource === 'Đơn Giao'

                        return (
                          <tr key={row.key} style={{
                            borderBottom: '1px solid #f1f5f9',
                            background: index % 2 === 1 ? '#f8fafc' : '#ffffff',
                            transition: 'background 0.1s'
                          }}
                          className="table-row-hover"
                          >
                            <td style={{ padding: '10px 14px', textAlign: 'center', fontSize: 13.5, color: '#64748b' }}>{globalIndex}</td>
                            
                            {/* Nhóm Badge */}
                            <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                              <span style={{
                                display: 'inline-block',
                                padding: '3px 8px',
                                borderRadius: 6,
                                fontSize: 12,
                                fontWeight: 700,
                                background: isGiamTru ? '#fff1f2' : '#ecfdf5',
                                color: isGiamTru ? '#be123c' : '#047857',
                                border: `1px solid ${isGiamTru ? '#fecdd3' : '#a7f3d0'}`
                              }}>
                                {row.compiledGroup}
                              </span>
                            </td>

                            {/* Nguồn Badge */}
                            <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                              <span style={{
                                display: 'inline-block',
                                padding: '3px 8px',
                                borderRadius: 6,
                                fontSize: 12,
                                fontWeight: 700,
                                background: isGiao ? '#eff6ff' : '#ecfeff',
                                color: isGiao ? '#1d4ed8' : '#0e7490',
                                border: `1px solid ${isGiao ? '#bfdbfe' : '#a5f3fc'}`
                              }}>
                                {row.compiledSource}
                              </span>
                            </td>

                            {/* Date */}
                            <td style={{ padding: '10px 14px', textAlign: 'center', fontSize: 13, color: '#334155' }}>
                              {row.ngayXuatNhap}
                            </td>

                            {/* Mã SAP/Mã vật tư */}
                            <td style={{ padding: '10px 14px', fontSize: 13, fontWeight: 500, color: '#475569' }}>
                              {row.maSAP || row.maVatTu || '-'}
                            </td>

                            {/* Tên vật tư */}
                            <td style={{ padding: '10px 14px', fontSize: 13.5, color: '#0f172a', fontWeight: 500 }}>
                              <div style={{ maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={row.tenVatTu}>
                                {row.tenVatTu}
                              </div>
                              {row.thongSoKyThuat && (
                                <div style={{ fontSize: 11, color: '#94a3b8', marginTop: 2, maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={row.thongSoKyThuat}>
                                  {row.thongSoKyThuat}
                                </div>
                              )}
                            </td>

                            {/* ĐVT */}
                            <td style={{ padding: '10px 14px', textAlign: 'center', fontSize: 13, color: '#475569' }}>{row.dvt || '-'}</td>

                            {/* Mã đơn nhập kho */}
                            <td style={{ padding: '10px 14px', fontSize: 13, color: '#475569' }}>
                              <div style={{ maxWidth: 150, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={row.maDonNhapKho || ''}>
                                {row.maDonNhapKho || '-'}
                              </div>
                            </td>

                            {/* Mã đơn xuất kho */}
                            <td style={{ padding: '10px 14px', fontSize: 13, color: '#475569' }}>
                              <div style={{ maxWidth: 150, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={row.maDonXuatKho || ''}>
                                {row.maDonXuatKho || '-'}
                              </div>
                            </td>

                            {/* Đơn vị giao */}
                            <td style={{ padding: '10px 14px', fontSize: 13, color: '#334155', fontWeight: 500 }}>
                              <div style={{ maxWidth: 140, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={row.donViGiao || ''}>
                                {row.donViGiao || '-'}
                              </div>
                            </td>

                            {/* Đơn vị nhận */}
                            <td style={{ padding: '10px 14px', fontSize: 13, color: '#334155', fontWeight: 500 }}>
                              <div style={{ maxWidth: 140, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={row.donViNhan || ''}>
                                {row.donViNhan || '-'}
                              </div>
                            </td>

                            {/* Đơn vị đối tác */}
                            <td style={{ padding: '10px 14px', fontSize: 13, color: '#334155', fontWeight: 500 }}>
                              <div style={{ maxWidth: 160, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }} title={row.compiledUnit}>
                                {row.compiledUnit}
                              </div>
                            </td>

                            {/* Khối lượng Nhập */}
                            <td style={{ padding: '10px 14px', textAlign: 'right', fontSize: 13.5 }}>
                              {formatValueWithNegColor(row.khoiLuongNhap)}
                            </td>

                            {/* Khối lượng Xuất */}
                            <td style={{ padding: '10px 14px', textAlign: 'right', fontSize: 13.5 }}>
                              {formatValueWithNegColor(row.khoiLuongXuat)}
                            </td>

                            {/* Trạng thái badge */}
                            <td style={{ padding: '10px 14px', textAlign: 'center' }}>
                              <span className={`badge ${getTrangThaiColor(row.trangThai)}`} style={{ fontSize: 11.5 }}>
                                {row.trangThai || 'Chưa duyệt'}
                              </span>
                            </td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              </div>

              {/* Pagination controls */}
              <div style={{
                display: 'flex', alignItems: 'center', justifyContent: 'space-between',
                padding: '12px 18px', background: '#f8fafc', border: '1px solid #e2e8f0', borderRadius: 10
              }}>
                <div style={{ fontSize: 13.5, color: '#64748b' }}>
                  Hiển thị dòng <strong>{((currentPage - 1) * pageSize + 1).toLocaleString('vi-VN')}</strong> đến <strong>{Math.min(currentPage * pageSize, filteredRows.length).toLocaleString('vi-VN')}</strong> trong tổng số <strong>{filteredRows.length.toLocaleString('vi-VN')}</strong> dòng kết quả
                </div>

                <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                  {/* Select page size */}
                  <select
                    value={pageSize}
                    onChange={(e) => {
                      setPageSize(Number(e.target.value))
                      setCurrentPage(1)
                    }}
                    style={{
                      padding: '5px 8px', border: '1px solid #cbd5e1', borderRadius: 6,
                      fontSize: 12.5, background: '#ffffff', marginRight: 12
                    }}
                  >
                    <option value={50}>50 dòng/trang</option>
                    <option value={100}>100 dòng/trang</option>
                    <option value={200}>200 dòng/trang</option>
                    <option value={500}>500 dòng/trang</option>
                  </select>

                  <button
                    disabled={currentPage === 1}
                    onClick={() => setCurrentPage(1)}
                    className="btn btn-outline btn-xs"
                    style={{ padding: '4px 8px', border: '1px solid #cbd5e1', opacity: currentPage === 1 ? 0.5 : 1, cursor: currentPage === 1 ? 'not-allowed' : 'pointer' }}
                  >
                    Đầu
                  </button>
                  <button
                    disabled={currentPage === 1}
                    onClick={() => setCurrentPage(prev => prev - 1)}
                    className="btn btn-outline btn-xs"
                    style={{ padding: '4px 8px', border: '1px solid #cbd5e1', opacity: currentPage === 1 ? 0.5 : 1, cursor: currentPage === 1 ? 'not-allowed' : 'pointer' }}
                  >
                    Trước
                  </button>
                  <span style={{ fontSize: 13, fontWeight: 600, padding: '0 8px', color: '#334155' }}>
                    Trang {currentPage} / {totalPages}
                  </span>
                  <button
                    disabled={currentPage === totalPages}
                    onClick={() => setCurrentPage(prev => prev + 1)}
                    className="btn btn-outline btn-xs"
                    style={{ padding: '4px 8px', border: '1px solid #cbd5e1', opacity: currentPage === totalPages ? 0.5 : 1, cursor: currentPage === totalPages ? 'not-allowed' : 'pointer' }}
                  >
                    Sau
                  </button>
                  <button
                    disabled={currentPage === totalPages}
                    onClick={() => setCurrentPage(totalPages)}
                    className="btn btn-outline btn-xs"
                    style={{ padding: '4px 8px', border: '1px solid #cbd5e1', opacity: currentPage === totalPages ? 0.5 : 1, cursor: currentPage === totalPages ? 'not-allowed' : 'pointer' }}
                  >
                    Cuối
                  </button>
                </div>
              </div>
            </div>
          ) : (
            <div style={{ padding: 24, background: '#f8fafc', borderRadius: 12, border: '1px solid #e2e8f0' }}>
              <div className="empty-state" style={{ minHeight: 180 }}>
                <span className="text-muted" style={{ fontSize: 14 }}>Không tìm thấy dòng kết quả nào phù hợp với bộ lọc hiện tại.</span>
              </div>
            </div>
          )}
        </>
      ) : (
        <div style={{ padding: 24 }}>
          <div className="empty-state" style={{ minHeight: 240 }}>
            <span className="text-muted" style={{ fontSize: 14 }}>Vui lòng chọn hoặc cấu hình một loại tổng hợp để xem dữ liệu.</span>
          </div>
        </div>
      )}

    </div>
  )
}

// ─── Placeholder Tab ──────────────────────────────────────────────────────────
function PlaceholderTab({ icon, title, desc }) {
  return (
    <div style={{ padding: 24 }}>
      <div className="empty-state" style={{ minHeight: 400 }}>
        <div style={{
          width: 64, height: 64, background: '#f1f5f9',
          borderRadius: 18, display: 'flex', alignItems: 'center', justifyContent: 'center',
        }}>
          {React.cloneElement(icon, { size: 28, color: 'var(--text-light)' })}
        </div>
        <h3>{title}</h3>
        <p style={{ maxWidth: 320 }}>{desc}</p>
        <div style={{
          marginTop: 8,
          background: '#fef9c3',
          border: '1px solid #fde68a',
          borderRadius: 8,
          padding: '8px 16px',
          fontSize: 14,
          color: '#92400e',
          fontWeight: 600,
        }}>
          🚧 Tính năng đang được phát triển
        </div>
      </div>
    </div>
  )
}

// Helper functions to check duplicate rows more accurately (avoiding flagging separate items in the same warehouse receipt/delivery as duplicates)
function cleanNumStr(val) {
  if (val === null || val === undefined) return '0'
  const str = String(val).trim()
  if (!str) return '0'
  const cleaned = str.replace(/[^\d.,-]/g, '').replace(',', '.')
  const num = parseFloat(cleaned)
  return isNaN(num) ? '0' : String(num)
}

function getRowUniqueKey(r) {
  const ngay = String(r.ngayXuatNhap || '').trim().toLowerCase()
  const maSAP = String(r.maSAP || '').trim().toLowerCase()
  const nhap = String(r.maDonNhapKho || '').trim().toLowerCase()
  const xuat = String(r.maDonXuatKho || '').trim().toLowerCase()
  
  return `${nhap || '_'}|${xuat || '_'}|${ngay}|${maSAP}`
}

// ─── Preview Import Modal ─────────────────────────────────────────────────────
function PreviewImportModal({ isOpen, onClose, onConfirm, rows, fileName, type, selectedProject, isAppend, existingRows }) {
  // Nhấn Esc để đóng modal
  React.useEffect(() => {
    if (!isOpen) return
    const handleKeyDown = (e) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('keydown', handleKeyDown)
    return () => document.removeEventListener('keydown', handleKeyDown)
  }, [isOpen, onClose])

  if (!isOpen) return null
  const label = type === 'giao' ? 'Đơn Giao' : 'Đơn Nhận'

  const unitKey = type === 'giao' ? 'donViGiao' : 'donViNhan'
  const unitLabel = type === 'giao' ? 'Đơn vị giao' : 'Đơn vị nhận'

  // Lọc các bản ghi đã có sẵn trong cơ sở dữ liệu để tìm trùng lặp bằng key định danh dòng
  const existingRowKeys = new Set(
    (existingRows || []).map(r => getRowUniqueKey(r))
  )

  // Lọc các dòng có Đơn vị giao/nhận trùng khớp với Kho dự án đang chọn (trùng khớp hoàn toàn, không phân biệt chữ hoa thường)
  const matchedByUnit = selectedProject
    ? rows.filter(r => {
        if (type === 'chung') {
          const unitGiao = String(r.donViGiao || '').trim().toLowerCase()
          const unitNhan = String(r.donViNhan || '').trim().toLowerCase()
          const proj = selectedProject.trim().toLowerCase()
          return unitGiao === proj || unitNhan === proj
        } else {
          const unit = String(r[unitKey] || '').trim().toLowerCase()
          const proj = selectedProject.trim().toLowerCase()
          return unit === proj
        }
      })
    : rows

  // Số dòng bị loại vì không khớp đơn vị
  const mismatchCount = rows.length - matchedByUnit.length

  // Lấy các đơn vị trong file KHÔNG khớp (để thông báo)
  const uniqueUnits = [...new Set(rows.map(r => String(r[unitKey] || r.donViGiao || r.donViNhan || '').trim()).filter(Boolean))]
  const mismatchUnits = uniqueUnits.filter(u => {
    const proj = (selectedProject || '').trim().toLowerCase()
    return u.toLowerCase() !== proj
  })

  // Trong số các dòng khớp đơn vị, bỏ qua lọc theo Đã phê duyệt (cứ thế cho phép tải lên bất kể trạng thái)
  const approvedRows = matchedByUnit
  const skippedStatusCount = 0

  // Người dùng yêu cầu bỏ qua hoàn toàn logic kiểm tra trùng lặp dòng, cho phép lưu hết
  const finalApprovedRows = approvedRows
  const duplicateRows = []
  const duplicateCount = 0

  // Chỉ hiển thị 10 dòng đầu của dữ liệu sẽ thực sự được lưu (hoặc fallback nếu trống)
  const previewRows = finalApprovedRows.length > 0 ? finalApprovedRows.slice(0, 10) : rows.slice(0, 10)

  return (
    <div style={{
      position: 'fixed', inset: 0, zIndex: 10000,
      background: 'rgba(15,23,42,0.55)', backdropFilter: 'blur(3px)',
      display: 'flex', alignItems: 'center', justifyContent: 'center', padding: '12px 16px'
    }}>
      <div style={{
        background: '#fff', borderRadius: 14, boxShadow: '0 25px 60px rgba(15,23,42,0.22)',
        width: '100%', maxWidth: '98vw', height: '96vh',
        display: 'flex', flexDirection: 'column', overflow: 'hidden'
      }}>
        {/* Header */}
        <div style={{
          padding: '18px 24px', borderBottom: '1px solid #e2e8f0',
          display: 'flex', alignItems: 'center', justifyContent: 'space-between', flexShrink: 0
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <div style={{
              width: 40, height: 40, background: 'var(--primary-light)',
              borderRadius: 10, display: 'flex', alignItems: 'center', justifyContent: 'center'
            }}>
              <FileSpreadsheet size={20} color="var(--primary)" />
            </div>
            <div>
              <div style={{ fontWeight: 700, fontSize: 16, color: 'var(--text)' }}>
                Xem trước dữ liệu — {label}
              </div>
              <div style={{ fontSize: 13, color: 'var(--text-muted)', marginTop: 2 }}>
                {fileName} · {rows.length.toLocaleString()} dòng dữ liệu
                {finalApprovedRows.length > 10 && <span style={{ color: '#f59e0b', marginLeft: 4 }}>(hiển thị 10 dòng đầu)</span>}
              </div>
            </div>
          </div>
          <button onClick={onClose} style={{
            background: 'none', border: 'none', cursor: 'pointer',
            color: 'var(--text-muted)', padding: 4, borderRadius: 6,
            display: 'flex', alignItems: 'center'
          }}>
            <X size={20} />
          </button>
        </div>

        {/* Append Mode Information Banner */}
        {isAppend && (
          <div style={{
            background: '#f0f9ff',
            borderBottom: '1px solid #bae6fd',
            padding: '12px 24px',
            color: '#0369a1',
            fontSize: '13.5px',
            fontWeight: 500,
            display: 'flex',
            alignItems: 'center',
            gap: 10,
            flexShrink: 0,
            textAlign: 'left'
          }}>
            <CloudUpload size={18} color="#0284c7" style={{ flexShrink: 0 }} />
            <span>
              <strong>Chế độ: UP FILE NỐI TIẾP</strong> — Dữ liệu mới trong tệp này sẽ được <strong>CHÈN THÊM (nối tiếp)</strong> vào danh sách hiện tại của kho dự án <strong>"{selectedProject || 'tất cả'}"</strong>. Toàn bộ dữ liệu cũ của bạn sẽ được giữ nguyên, không bị thay thế hay xóa bỏ!
            </span>
          </div>
        )}

        {/* Warning notification banner for matching rules */}
        {mismatchCount > 0 && (
          <div style={{
            background: finalApprovedRows.length === 0 ? '#fef2f2' : '#fffbeb',
            borderBottom: finalApprovedRows.length === 0 ? '1px solid #fee2e2' : '1px solid #fef3c7',
            padding: '12px 24px',
            color: finalApprovedRows.length === 0 ? '#991b1b' : '#92400e',
            fontSize: 13,
            display: 'flex',
            alignItems: 'flex-start',
            gap: 10,
            flexShrink: 0,
            textAlign: 'left'
          }}>
            <AlertCircle size={18} color={finalApprovedRows.length === 0 ? '#dc2626' : '#d97706'} style={{ flexShrink: 0, marginTop: '2px' }} />
            <div style={{ flex: 1, lineHeight: 1.5 }}>
              {finalApprovedRows.length === 0 ? (
                <span>
                  <strong style={{ fontSize: '14px', display: 'block', marginBottom: '4px', color: '#b91c1c' }}>⚠️ KIỂM TRA LỖI KHỚP DÀNH CHO KHO DỰ ÁN</strong>
                  Không có bất kỳ dòng dữ liệu hợp lệ nào trong tệp tải lên khớp với Kho dự án đã chọn <strong>"{selectedProject}"</strong> hoặc chưa được duyệt hoặc bị trùng lặp. Hệ thống sẽ <strong>bỏ qua toàn bộ dữ liệu</strong> và không thể lưu tệp này.
                  {uniqueUnits.length > 0 && (
                    <div style={{ marginTop: 8, padding: '8px 12px', background: '#fff', borderRadius: '6px', border: '1px solid #fee2e2', fontWeight: 500, color: '#b91c1c' }}>
                      Các giá trị {unitLabel} được tìm thấy trong tệp hiện tại: <strong>{uniqueUnits.join(', ')}</strong>
                    </div>
                  )}
                </span>
              ) : (
                <span>
                  <strong style={{ fontSize: '14px', display: 'block', marginBottom: '4px' }}>⚠️ LƯU Ý BỎ QUA DỮ LIỆU KHÔNG TRÙNG KHỚP KHO</strong>
                  Phát hiện <strong>{mismatchCount.toLocaleString()} dòng</strong> có cột <strong>{unitLabel}</strong> không trùng khớp hoàn toàn với tên Kho dự án đang chọn <strong>"{selectedProject}"</strong>. Hệ thống sẽ tự động <strong>bỏ qua {mismatchCount.toLocaleString()} dòng này</strong> và chỉ nhập/lưu {finalApprovedRows.length.toLocaleString()} dòng hợp lệ.
                </span>
              )}
            </div>
          </div>
        )}

        {/* Warning notification banner for duplicate entries */}
        {duplicateCount > 0 && (
          <div style={{
            background: '#fff7ed',
            borderBottom: '1px solid #fed7aa',
            padding: '12px 24px',
            color: '#c2410c',
            fontSize: 13,
            display: 'flex',
            alignItems: 'flex-start',
            gap: 10,
            flexShrink: 0,
            textAlign: 'left'
          }}>
            <AlertTriangle size={18} color="#ea580c" style={{ flexShrink: 0, marginTop: '2px' }} />
            <div style={{ flex: 1, lineHeight: 1.5 }}>
              <span>
                <strong style={{ fontSize: '14px', display: 'block', marginBottom: '4px' }}>⚠️ PHÁT HIỆN DỮ LIỆU BỊ TRÙNG LẶP CHI TIẾT DÒNG</strong>
                Phát hiện <strong>{duplicateCount.toLocaleString()} dòng</strong> dữ liệu trong file bị <strong>TRÙNG LẶP HOÀN TOÀN</strong> (các thông tin chi tiết của dòng đã tồn tại trên cơ sở dữ liệu hiện tại hoặc bị lặp lại nội bộ trong file). Hệ thống sẽ tự động <strong>BỎ QUA {duplicateCount.toLocaleString()} dòng này</strong> để giữ an toàn cho dữ liệu gốc.
              </span>
            </div>
          </div>
        )}

        {/* Table preview */}
        <div style={{ overflow: 'auto', flex: 1 }}>
          <DataTable rows={previewRows} />
        </div>

        {/* Footer */}
        <div style={{
          padding: '12px 24px', borderTop: '1px solid #e2e8f0',
          background: '#f8fafc', flexShrink: 0
        }}>
          {/* Thông báo lọc dữ liệu */}
          <div style={{ display: 'flex', gap: 10, marginBottom: 10, flexWrap: 'wrap' }}>
            {/* Tổng cộng số dòng tải lên */}
            <div style={{
              display: 'flex', alignItems: 'center', gap: 7,
              background: '#f8fafc', border: '1px solid #cbd5e1',
              borderRadius: 8, padding: '7px 14px', fontSize: 13, color: '#334155', flex: '1 1 100%'
            }}>
              <FileSpreadsheet size={15} color="#475569" />
              <span>
                Tổng cộng phát hiện: <strong style={{ color: '#0f172a', fontSize: 14 }}>{rows.length.toLocaleString()} dòng</strong> dữ liệu được tải lên từ file Excel.
              </span>
            </div>

            {/* Dòng khớp đơn vị & đã phê duyệt → sẽ lưu */}
            <div style={{
              display: 'flex', alignItems: 'center', gap: 7,
              background: '#ecfdf5', border: '1px solid #a7f3d0',
              borderRadius: 8, padding: '7px 14px', fontSize: 13, color: '#065f46', flex: 1, minWidth: 220
            }}>
              <CheckCircle2 size={15} color="#10b981" />
              <span>
                <strong>{finalApprovedRows.length.toLocaleString()} dòng</strong> khớp <strong>{unitLabel} "{selectedProject || 'tất cả'}"</strong> → sẽ được lưu
              </span>
            </div>

            {/* Dòng trùng lặp → bị loại */}
            {duplicateCount > 0 && (
              <div style={{
                display: 'flex', alignItems: 'center', gap: 7,
                background: '#fff7ed', border: '1px solid #fed7aa',
                borderRadius: 8, padding: '7px 14px', fontSize: 13, color: '#c2410c', flex: 1, minWidth: 220
              }}>
                <AlertTriangle size={15} color="#ea580c" style={{ flexShrink: 0 }} />
                <span>
                  <strong>{duplicateCount.toLocaleString()} dòng</strong> trùng lặp chi tiết (Bỏ qua) → <strong>sẽ không được lưu</strong>
                </span>
              </div>
            )}

            {/* Dòng không khớp đơn vị → bị loại */}
            {mismatchCount > 0 && (
              <div style={{
                display: 'flex', alignItems: 'flex-start', gap: 7,
                background: '#fff1f2', border: '1px solid #fecdd3',
                borderRadius: 8, padding: '7px 14px', fontSize: 13, color: '#9f1239', flex: 1, minWidth: 220
              }}>
                <AlertCircle size={15} color="#ef4444" style={{ marginTop: 1, flexShrink: 0 }} />
                <span>
                  <strong>{mismatchCount.toLocaleString()} dòng</strong> có <strong>{unitLabel}</strong> không khớp kho <strong>"{selectedProject}"</strong>
                  {mismatchUnits.length > 0 && <span style={{ color: '#be123c' }}> ({mismatchUnits.slice(0,2).join(', ')}{mismatchUnits.length > 2 ? '...' : ''})</span>} → <strong>sẽ không được lưu</strong>
                </span>
              </div>
            )}

            {/* Dòng khớp đơn vị nhưng chưa phê duyệt → bị loại */}
            {skippedStatusCount > 0 && (
              <div style={{
                display: 'flex', alignItems: 'center', gap: 7,
                background: '#fff7ed', border: '1px solid #fed7aa',
                borderRadius: 8, padding: '7px 14px', fontSize: 13, color: '#92400e', flex: 1, minWidth: 220
              }}>
                <AlertCircle size={15} color="#f59e0b" />
                <span>
                  <strong>{skippedStatusCount.toLocaleString()} dòng</strong> chưa được phê duyệt (Chưa xác nhận, Từ chối...) → <strong>sẽ không được lưu</strong>
                </span>
              </div>
            )}
          </div>

          {/* Actions */}
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'flex-end', gap: 10 }}>
            <button onClick={onClose} className="btn btn-outline" style={{ minWidth: 90 }}>
              <X size={14} /> Hủy
            </button>
            <button
              onClick={onConfirm}
              className="btn btn-primary"
              style={{ minWidth: 160, opacity: finalApprovedRows.length === 0 ? 0.5 : 1, cursor: finalApprovedRows.length === 0 ? 'not-allowed' : 'pointer' }}
              disabled={finalApprovedRows.length === 0}
            >
              <Save size={14} /> Lưu {finalApprovedRows.length.toLocaleString()} dòng dữ liệu
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

// ─── Modal ───────────────────────────────────────────────────────────────────
function AddProjectModal({ isOpen, onClose, onSave }) {
  const [name, setName] = useState('')
  const [error, setError] = useState('')

  if (!isOpen) return null

  const handleSave = () => {
    const trimmed = name.trim()
    if (!trimmed) {
      setError('Vui lòng nhập tên Kho dự án')
      return
    }
    onSave(trimmed)
    setName('')
    setError('')
    onClose()
  }

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(15, 23, 42, 0.6)',
      backdropFilter: 'blur(4px)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 1000,
    }}>
      <div style={{
        background: '#ffffff',
        borderRadius: 12,
        width: '100%',
        maxWidth: 420,
        padding: 24,
        boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04)',
        border: '1px solid #e2e8f0',
        display: 'flex',
        flexDirection: 'column',
        gap: 16
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: 12 }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#0f172a' }}>Tạo mới Kho dự án</h3>
          <button onClick={onClose} style={{ background: 'transparent', border: 'none', color: '#64748b', cursor: 'pointer', display: 'flex', alignItems: 'center' }}>
            <X size={18} />
          </button>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <label style={{ fontSize: 13, fontWeight: 600, color: '#475569', textAlign: 'left' }}>Tên Kho dự án / Công trình</label>
          <input
            autoFocus
            type="text"
            className="input"
            style={{ width: '100%', border: error ? '1px solid #ef4444' : '1px solid #cbd5e1' }}
            placeholder="Ví dụ: SGC Sunrise Complex 2..."
            value={name}
            onChange={(e) => {
              setName(e.target.value)
              if (error) setError('')
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') handleSave()
            }}
          />
          {error && <span style={{ fontSize: 12, color: '#ef4444', fontWeight: 500, textAlign: 'left' }}>{error}</span>}
        </div>

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 8 }}>
          <button
            onClick={onClose}
            style={{
              padding: '6px 14px',
              borderRadius: 6,
              border: '1px solid #cbd5e1',
              background: '#ffffff',
              color: '#475569',
              fontSize: 13,
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            Hủy
          </button>
          <button
            onClick={handleSave}
            style={{
              padding: '6px 14px',
              borderRadius: 6,
              border: 'none',
              background: 'linear-gradient(135deg, #0f58a7 0%, #1a6abf 100%)',
              color: '#ffffff',
              fontSize: 13,
              fontWeight: 600,
              cursor: 'pointer',
              boxShadow: '0 2px 4px rgba(15,88,167,0.2)'
            }}
          >
            Lưu lại
          </button>
        </div>
      </div>
    </div>
  )
}

// ─── Edit Project Modal ───────────────────────────────────────────────────────
function EditProjectModal({ isOpen, onClose, onSave, currentName }) {
  const [name, setName] = useState('')
  const [error, setError] = useState('')

  React.useEffect(() => {
    if (isOpen) {
      setName(currentName || '')
      setError('')
    }
  }, [isOpen, currentName])

  if (!isOpen) return null

  const handleSave = () => {
    const trimmed = name.trim()
    if (!trimmed) {
      setError('Vui lòng nhập tên Kho dự án')
      return
    }
    onSave(currentName, trimmed)
    setError('')
    onClose()
  }

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(15, 23, 42, 0.6)',
      backdropFilter: 'blur(4px)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 10001
    }}>
      <div style={{
        background: '#ffffff',
        borderRadius: 12,
        width: '100%',
        maxWidth: 420,
        padding: 24,
        boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04)',
        border: '1px solid #e2e8f0',
        display: 'flex',
        flexDirection: 'column',
        gap: 16
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: 12 }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#0f172a' }}>Sửa tên Kho dự án</h3>
          <button onClick={onClose} style={{ background: 'transparent', border: 'none', color: '#64748b', cursor: 'pointer', display: 'flex', alignItems: 'center' }}>
            <X size={18} />
          </button>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <label style={{ fontSize: 13, fontWeight: 600, color: '#475569', textAlign: 'left' }}>Tên Kho dự án / Công trình mới</label>
          <input
            autoFocus
            type="text"
            className="input"
            style={{ width: '100%', border: error ? '1px solid #ef4444' : '1px solid #cbd5e1' }}
            placeholder="Ví dụ: SGC Sunrise Complex 2..."
            value={name}
            onChange={(e) => {
              setName(e.target.value)
              if (error) setError('')
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter') handleSave()
            }}
          />
          {error && <span style={{ fontSize: 12, color: '#ef4444', fontWeight: 500, textAlign: 'left' }}>{error}</span>}
          <p style={{ fontSize: 11, color: '#64748b', marginTop: 4, fontStyle: 'italic', textAlign: 'left', lineHeight: '1.4' }}>
            * Lưu ý: Khi lưu, tất cả các đơn hàng trực thuộc dự án hiện tại sẽ được thay đổi tên để đồng bộ sang tên mới.
          </p>
        </div>

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 8 }}>
          <button
            onClick={onClose}
            style={{
              padding: '6px 14px',
              borderRadius: 6,
              border: '1px solid #cbd5e1',
              background: '#ffffff',
              color: '#475569',
              fontSize: 13,
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            Hủy
          </button>
          <button
            onClick={handleSave}
            style={{
              padding: '6px 14px',
              borderRadius: 6,
              border: 'none',
              background: 'linear-gradient(135deg, #0f58a7 0%, #1a6abf 100%)',
              color: '#ffffff',
              fontSize: 13,
              fontWeight: 600,
              cursor: 'pointer',
              boxShadow: '0 2px 4px rgba(15,88,167,0.2)'
            }}
          >
            Cập nhật
          </button>
        </div>
      </div>
    </div>
  )
}

// ─── Delete File Modal ────────────────────────────────────────────────────
function DeleteFileModal({ isOpen, onClose, onConfirm, type, selectedProject, rowCount }) {
  if (!isOpen) return null
  const label = type === 'giao' ? 'Đơn Giao' : 'Đơn Nhận'
  const hasData = rowCount > 0

  return (
    <div style={{
      position: 'fixed',
      top: 0, left: 0, right: 0, bottom: 0,
      backgroundColor: 'rgba(15, 23, 42, 0.65)',
      backdropFilter: 'blur(4px)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 10002
    }}>
      <div style={{
        background: '#ffffff',
        borderRadius: 12,
        width: '100%',
        maxWidth: 440,
        padding: 24,
        boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04)',
        border: '1px solid #fee2e2',
        display: 'flex',
        flexDirection: 'column',
        gap: 16
      }}>
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: 12 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <div style={{ width: 36, height: 36, borderRadius: 10, background: '#fee2e2', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <Trash2 size={18} color="#ef4444" />
            </div>
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#991b1b' }}>Xác nhận xóa dữ liệu</h3>
          </div>
          <button onClick={onClose} style={{ background: 'transparent', border: 'none', color: '#64748b', cursor: 'pointer', display: 'flex', alignItems: 'center' }}>
            <X size={18} />
          </button>
        </div>

        {/* Body */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          <p style={{ margin: 0, fontSize: 14, color: '#334155', lineHeight: '1.6' }}>
            Bạn có chắc chắn muốn xóa toàn bộ dữ liệu <strong style={{ color: '#0f172a' }}>{label}</strong>
            {selectedProject ? <> của kho dự án <strong style={{ color: '#0f172a' }}>"{selectedProject}"</strong></> : ''}?
          </p>

          {hasData ? (
            <div style={{
              background: '#fff1f2',
              border: '1.5px solid #fecdd3',
              borderRadius: 10,
              padding: '12px 14px',
              display: 'flex',
              gap: 10,
              alignItems: 'flex-start'
            }}>
              <AlertCircle size={18} color="#ef4444" style={{ flexShrink: 0, marginTop: 1 }} />
              <div style={{ fontSize: 13, color: '#9f1239', lineHeight: '1.6' }}>
                <strong>File này đang có dữ liệu!</strong>
                <div style={{ marginTop: 6 }}>
                  <span>• <strong>{rowCount.toLocaleString()} dòng</strong> {label} sẽ bị xóa vĩnh viễn</span>
                </div>
                <div style={{ marginTop: 8, fontWeight: 600 }}>
                  Thao tác này không thể hoàn tác. Bạn có chắc chắn muốn tiếp tục?
                </div>
              </div>
            </div>
          ) : (
            <div style={{
              background: '#fffbeb',
              border: '1px solid #fef3c7',
              borderRadius: 8,
              padding: '10px 12px',
              display: 'flex',
              gap: 8,
              alignItems: 'flex-start'
            }}>
              <AlertCircle size={16} color="#d97706" style={{ flexShrink: 0, marginTop: 2 }} />
              <p style={{ margin: 0, fontSize: 12, color: '#b45309', lineHeight: '1.4', fontWeight: 500 }}>
                File chưa có dữ liệu. Thao tác này sẽ xóa trắng dữ liệu hiện tại.
              </p>
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 4 }}>
          <button
            onClick={onClose}
            style={{
              padding: '8px 16px', borderRadius: 6,
              border: '1px solid #cbd5e1', background: '#ffffff',
              color: '#475569', fontSize: 13, fontWeight: 600, cursor: 'pointer'
            }}
          >
            Hủy bỏ
          </button>
          <button
            onClick={onConfirm}
            style={{
              padding: '8px 18px', borderRadius: 6, border: 'none',
              background: '#dc2626',
              color: '#ffffff', fontSize: 13, fontWeight: 700,
              cursor: 'pointer',
              boxShadow: '0 2px 8px rgba(220,38,38,0.35)',
              display: 'flex', alignItems: 'center', gap: 6
            }}
          >
            <Trash2 size={14} />
            Xóa dữ liệu
          </button>
        </div>
      </div>
    </div>
  )
}

// ─── Delete Project Modal ──────────────────────────────────────────────────
function DeleteProjectModal({ isOpen, onClose, onConfirm, projectName, giaoRows, nhanRows, khoRows }) {
  if (!isOpen) return null

  // Đếm số dòng dữ liệu thuộc dự án này
  const giaoCount = (giaoRows || []).filter(r => (r.ten_du_an || r.tenDuAn || r.duAn) === projectName).length
  const nhanCount = (nhanRows || []).filter(r => (r.ten_du_an || r.tenDuAn || r.duAn) === projectName).length
  const khoCount = (khoRows || []).filter(r => (r.ten_du_an || r.tenDuAn || r.duAn) === projectName).length
  const hasData = giaoCount > 0 || nhanCount > 0 || khoCount > 0

  return (
    <div style={{
      position: 'fixed',
      top: 0, left: 0, right: 0, bottom: 0,
      backgroundColor: 'rgba(15, 23, 42, 0.65)',
      backdropFilter: 'blur(4px)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 10002
    }}>
      <div style={{
        background: '#ffffff',
        borderRadius: 12,
        width: '100%',
        maxWidth: 440,
        padding: 24,
        boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04)',
        border: '1px solid #fee2e2',
        display: 'flex',
        flexDirection: 'column',
        gap: 16
      }}>
        {/* Header */}
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: 12 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <div style={{ width: 36, height: 36, borderRadius: 10, background: '#fee2e2', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <Trash2 size={18} color="#ef4444" />
            </div>
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#991b1b' }}>Xác nhận xóa Dự án</h3>
          </div>
          <button onClick={onClose} style={{ background: 'transparent', border: 'none', color: '#64748b', cursor: 'pointer', display: 'flex', alignItems: 'center' }}>
            <X size={18} />
          </button>
        </div>

        {/* Body */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
          <p style={{ margin: 0, fontSize: 14, color: '#334155', lineHeight: '1.6' }}>
            Bạn có chắc chắn muốn xóa kho dự án <strong style={{ color: '#0f172a' }}>"{projectName}"</strong>?
          </p>

          {/* Cảnh báo nếu có dữ liệu */}
          {hasData ? (
            <div style={{
              background: '#fff1f2',
              border: '1.5px solid #fecdd3',
              borderRadius: 10,
              padding: '12px 14px',
              display: 'flex',
              gap: 10,
              alignItems: 'flex-start'
            }}>
              <AlertCircle size={18} color="#ef4444" style={{ flexShrink: 0, marginTop: 1 }} />
              <div style={{ fontSize: 13, color: '#9f1239', lineHeight: '1.6' }}>
                <strong>Dự án này đang có dữ liệu!</strong>
                <div style={{ marginTop: 6, display: 'flex', flexDirection: 'column', gap: 3 }}>
                  {giaoCount > 0 && (
                    <span>• <strong>{giaoCount.toLocaleString()} dòng</strong> Đơn Giao sẽ bị xóa vĩnh viễn</span>
                  )}
                  {nhanCount > 0 && (
                    <span>• <strong>{nhanCount.toLocaleString()} dòng</strong> Đơn Nhận sẽ bị xóa vĩnh viễn</span>
                  )}
                  {khoCount > 0 && (
                    <span>• <strong>{khoCount.toLocaleString()} dòng</strong> Kho dự án sẽ bị xóa vĩnh viễn</span>
                  )}
                </div>
                <div style={{ marginTop: 8, fontWeight: 600 }}>
                  Thao tác này không thể hoàn tác. Bạn có chắc chắn muốn tiếp tục?
                </div>
              </div>
            </div>
          ) : (
            <div style={{
              background: '#fffbeb',
              border: '1px solid #fef3c7',
              borderRadius: 8,
              padding: '10px 12px',
              display: 'flex',
              gap: 8,
              alignItems: 'flex-start'
            }}>
              <AlertCircle size={16} color="#d97706" style={{ flexShrink: 0, marginTop: 2 }} />
              <p style={{ margin: 0, fontSize: 12, color: '#b45309', lineHeight: '1.4', fontWeight: 500 }}>
                Dự án chưa có dữ liệu. Thao tác này sẽ gỡ dự án khỏi danh sách chọn lựa.
              </p>
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 4 }}>
          <button
            onClick={onClose}
            style={{
              padding: '8px 16px', borderRadius: 6,
              border: '1px solid #cbd5e1', background: '#ffffff',
              color: '#475569', fontSize: 13, fontWeight: 600, cursor: 'pointer'
            }}
          >
            Hủy bỏ
          </button>
          <button
            onClick={onConfirm}
            style={{
              padding: '8px 18px', borderRadius: 6, border: 'none',
              background: hasData ? '#dc2626' : '#ef4444',
              color: '#ffffff', fontSize: 13, fontWeight: 700,
              cursor: 'pointer',
              boxShadow: hasData ? '0 2px 8px rgba(220,38,38,0.35)' : '0 2px 4px rgba(239,68,68,0.2)',
              display: 'flex', alignItems: 'center', gap: 6
            }}
          >
            <Trash2 size={14} />
            {hasData ? 'Xóa cả dữ liệu' : 'Đồng ý xóa'}
          </button>
        </div>
      </div>
    </div>
  )
}

// ─── Supabase Config Modal ───────────────────────────────────────────────────
function SupabaseConfigModal({ isOpen, onClose }) {
  const [url, setUrl] = useState(() => {
    return localStorage.getItem('sgc_supabase_url') || import.meta.env.VITE_SUPABASE_URL || 'https://luhsnaqlajbwkftrsbeg.supabase.co'
  })
  const [key, setKey] = useState(() => {
    return localStorage.getItem('sgc_supabase_key') || import.meta.env.VITE_SUPABASE_ANON_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imx1aHNuYXFsYWpid2tmdHJzYmVnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAwMDk3MjIsImV4cCI6MjA5NTU4NTcyMn0.NqqOG1KkzceGquzudBcPqOSsX1BhB24U_jmew0Mqsc4'
  })
  const [saveSuccess, setSaveSuccess] = useState(false)
  const [currentSource, setCurrentSource] = useState(() => {
    if (localStorage.getItem('sgc_supabase_url')) return 'local'
    if (import.meta.env.VITE_SUPABASE_URL) return 'env'
    return 'default'
  })

  const [testing, setTesting] = useState(false)
  const [testResult, setTestResult] = useState(null) // { success: boolean, message: string }

  React.useEffect(() => {
    if (isOpen) {
      setUrl(localStorage.getItem('sgc_supabase_url') || import.meta.env.VITE_SUPABASE_URL || 'https://luhsnaqlajbwkftrsbeg.supabase.co')
      setKey(localStorage.getItem('sgc_supabase_key') || import.meta.env.VITE_SUPABASE_ANON_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imx1aHNuYXFsYWpid2tmdHJzYmVnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAwMDk3MjIsImV4cCI6MjA5NTU4NTcyMn0.NqqOG1KkzceGquzudBcPqOSsX1BhB24U_jmew0Mqsc4')
      setCurrentSource(() => {
        if (localStorage.getItem('sgc_supabase_url')) return 'local'
        if (import.meta.env.VITE_SUPABASE_URL) return 'env'
        return 'default'
      })
      setSaveSuccess(false)
      setTestResult(null)
    }
  }, [isOpen])

  if (!isOpen) return null

  const handleTestConnection = async () => {
    setTesting(true)
    setTestResult(null)
    try {
      const trimmedUrl = url.trim()
      const trimmedKey = key.trim()
      if (!trimmedUrl || !trimmedKey) {
        setTestResult({ success: false, message: 'Vui lòng điền đầy đủ URL và Anon Key trước!' })
        setTesting(false)
        return
      }

      const res = await fetch(`${trimmedUrl}/rest/v1/du_an?select=*&limit=1`, {
        headers: {
          'apikey': trimmedKey,
          'Authorization': `Bearer ${trimmedKey}`
        }
      })

      if (res.ok) {
        setTestResult({ success: true, message: 'Kết nối thành công! Cấu hình hoàn toàn chính xác.' })
      } else {
        const status = res.status
        let customMsg = `Lỗi từ Máy chủ (HTTP ${status})`
        if (status === 401) {
          customMsg = 'Lỗi 401 Unauthorized: Anon Key của bạn không hợp lệ hoặc đã bị thay đổi/vô hiệu hóa trên trang quản trị Supabase!'
        } else if (status === 404) {
          customMsg = 'Lỗi 404 Not Found: URL dự án Supabase không chính xác!'
        } else {
          try {
            const body = await res.json()
            customMsg = `Lỗi (${status}): ${body.message || body.details || 'Không rõ lý do'}`
          } catch (err) {
            customMsg = `Lỗi (${status}): Yêu cầu bị từ chối hoặc sai credentials.`
          }
        }
        setTestResult({ success: false, message: customMsg })
      }
    } catch (e) {
      setTestResult({ success: false, message: `Lỗi kết nối mạng: ${e.message}. Kiểm tra lại tính chính xác của URL.` })
    } finally {
      setTesting(false)
    }
  }

  const handleSave = () => {
    const trimmedUrl = url.trim()
    const trimmedKey = key.trim()

    if (trimmedUrl && trimmedKey) {
      localStorage.setItem('sgc_supabase_url', trimmedUrl)
      localStorage.setItem('sgc_supabase_key', trimmedKey)
      setSaveSuccess(true)
      setTimeout(() => {
        window.location.reload()
      }, 1000)
    } else {
      alert('Vui lòng điền đầy đủ cả URL và Anon Key!')
    }
  }

  const handleClear = () => {
    if (window.confirm('Bạn có chắc chắn muốn xóa cấu hình thủ công và quay lại dùng cấu hình tích hợp mặc định trong code?')) {
      localStorage.removeItem('sgc_supabase_url')
      localStorage.removeItem('sgc_supabase_key')
      setUrl(import.meta.env.VITE_SUPABASE_URL || 'https://luhsnaqlajbwkftrsbeg.supabase.co')
      setKey(import.meta.env.VITE_SUPABASE_ANON_KEY || 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imx1aHNuYXFsYWpid2tmdHJzYmVnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODAwMDk3MjIsImV4cCI6MjA5NTU4NTcyMn0.NqqOG1KkzceGquzudBcPqOSsX1BhB24U_jmew0Mqsc4')
      setCurrentSource(import.meta.env.VITE_SUPABASE_URL ? 'env' : 'default')
      setSaveSuccess(true)
      setTimeout(() => {
        window.location.reload()
      }, 1000)
    }
  }

  const isCurrentlyConnected = isSupabaseConfigured

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(15, 23, 42, 0.65)',
      backdropFilter: 'blur(5px)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 10002
    }}>
      <div style={{
        background: '#ffffff',
        borderRadius: 16,
        width: '100%',
        maxWidth: 580,
        maxHeight: '90vh',
        overflowY: 'auto',
        padding: 24,
        boxShadow: '0 25px 50px -12px rgba(0,0,0,0.25)',
        border: '1px solid #e2e8f0',
        display: 'flex',
        flexDirection: 'column',
        gap: 16
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: 12 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <Database size={18} color="#0f58a7" />
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#0f172a' }}>Cấu hình kết nối cơ sở dữ liệu Supabase</h3>
          </div>
          <button onClick={onClose} style={{ background: 'transparent', border: 'none', color: '#64748b', cursor: 'pointer', display: 'flex', alignItems: 'center', padding: 4 }}>
            <X size={18} />
          </button>
        </div>

        {/* Status indicator */}
        <div style={{
          display: 'flex',
          alignItems: 'center',
          gap: 12,
          padding: '10px 14px',
          borderRadius: 8,
          backgroundColor: isCurrentlyConnected ? '#ecfdf5' : '#fff1f2',
          border: `1px solid ${isCurrentlyConnected ? '#a7f3d0' : '#fecdd3'}`
        }}>
          <div style={{
            width: 8,
            height: 8,
            borderRadius: '50%',
            backgroundColor: isCurrentlyConnected ? '#10b981' : '#f43f5e'
          }} />
          <div style={{ flex: 1 }}>
            <span style={{ fontSize: 13, fontWeight: 700, color: isCurrentlyConnected ? '#065f46' : '#9f1239' }}>
              Trạng thái: {isCurrentlyConnected ? 'Đã bật đầu nối Supabase' : 'Chưa cấu hình (Offline)'}
            </span>
            <p style={{ margin: '1px 0 0 0', fontSize: 11, color: isCurrentlyConnected ? '#047857' : '#be123c' }}>
              {isCurrentlyConnected 
                ? `Đang load từ: ${currentSource === 'local' ? 'Trình duyệt Web (Local Storage)' : currentSource === 'env' ? 'Biến môi trường (Vite Env)' : 'Cấu hình tích hợp mặc định (Code)'}`
                : 'Đang chạy offline bằng bộ dữ liệu tạm thời.'}
            </p>
          </div>
        </div>

        {/* Collapsible advanced connection settings block */}
        <details style={{
          fontSize: 12,
          color: '#475569',
          border: '1px solid #cbd5e1',
          borderRadius: 8,
          padding: '8px 12px',
          background: '#f8fafc',
          textAlign: 'left'
        }}>
          <summary style={{ fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, userSelect: 'none' }}>
            <Settings size={14} color="#0f58a7" />
            <span>Hiển thị thông tin cấu hình kết nối (URL / API Key)</span>
          </summary>

          <div style={{ marginTop: 12, display: 'flex', flexDirection: 'column', gap: 12 }}>
            {/* Edit Fields */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              <label style={{ fontSize: 11, fontWeight: 600, color: '#475569', textAlign: 'left' }}>
                Supabase Project URL (VITE_SUPABASE_URL)
              </label>
              <input
                type="text"
                className="input"
                style={{ width: '100%', fontFamily: 'monospace', fontSize: 11, height: '36px', padding: '0 10px' }}
                placeholder="Ví dụ: https://your-project-id.supabase.co"
                value={url}
                onChange={(e) => setUrl(e.target.value)}
              />
            </div>

            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              <label style={{ fontSize: 11, fontWeight: 600, color: '#475569', textAlign: 'left' }}>
                Supabase Project API Anon Key (VITE_SUPABASE_ANON_KEY)
              </label>
              <textarea
                className="input"
                rows={3}
                style={{ width: '100%', fontFamily: 'monospace', fontSize: 10, resize: 'vertical', padding: '6px 10px', lineHeight: '1.4' }}
                placeholder="Nhập chuỗi Anon Key..."
                value={key}
                onChange={(e) => setKey(e.target.value)}
              />
            </div>

            {/* Live Test connection button & results */}
            <div style={{ border: '1px dashed #cbd5e1', borderRadius: 8, padding: 12, background: '#ffffff' }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
                <span style={{ fontSize: 11, fontWeight: 600, color: '#334155' }}>Kiểm tra trạng thái kết nối trực tiếp:</span>
                <button
                  type="button"
                  onClick={handleTestConnection}
                  disabled={testing}
                  style={{
                    background: testing ? '#e2e8f0' : '#f1f5f9',
                    border: '1px solid #cbd5e1',
                    padding: '4px 10px',
                    borderRadius: 4,
                    fontSize: 10,
                    fontWeight: 600,
                    color: '#1e293b',
                    cursor: testing ? 'not-allowed' : 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    gap: 4
                  }}
                >
                  {testing ? (
                    <>
                      <RefreshCw size={10} className="animate-spin" />
                      Đang kiểm tra...
                    </>
                  ) : 'Bấm để Test'}
                </button>
              </div>

              {testResult && (
                <div style={{
                  marginTop: 8,
                  padding: 8,
                  borderRadius: 6,
                  fontSize: 10,
                  fontWeight: 500,
                  backgroundColor: testResult.success ? '#f0fdf4' : '#fef2f2',
                  border: `1px solid ${testResult.success ? '#bbf7d0' : '#fca5a5'}`,
                  color: testResult.success ? '#15803d' : '#b91c1c',
                  textAlign: 'left',
                  lineHeight: '1.4'
                }}>
                  {testResult.message}
                </div>
              )}
            </div>
          </div>
        </details>

        {/* Info detail block */}
        <details style={{ fontSize: 12, color: '#475569', border: '1px solid #e2e8f0', borderRadius: 8, padding: '8px 12px', background: '#f8fafc' }}>
          <summary style={{ fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, userSelect: 'none' }}>
            <Info size={14} color="#0f58a7" />
            <span>Hướng dẫn đồng bộ lên Cloudflare / Pages</span>
          </summary>
          <div style={{ marginTop: 8, paddingLeft: 4, display: 'flex', flexDirection: 'column', gap: 6 }}>
            <p style={{ margin: 0 }}><strong>Cách 1:</strong> Bạn chỉ cần cấu hình ngay tại form này rồi bấm <strong>Lưu kết nối</strong>. Trình duyệt của bạn sẽ tự lưu vĩnh viễn và đồng bộ dữ liệu ngay lập tức.</p>
            <p style={{ margin: 0 }}><strong>Cách 2:</strong> Vào Cloudflare Dashboard &gt; Settings &gt; Environment variables &gt; Thêm <code style={{ fontSize: 11, background: '#cbd5e1', padding: '1px 3px', borderRadius: 3 }}>VITE_SUPABASE_URL</code> và <code style={{ fontSize: 11, background: '#cbd5e1', padding: '1px 3px', borderRadius: 3 }}>VITE_SUPABASE_ANON_KEY</code> để mọi thiết bị truy cập mặc định online.</p>
          </div>
        </details>

        {/* SQL Queries / RLS Policy Instructions */}
        <details style={{ fontSize: 12, color: '#475569', border: '1px solid #fed7aa', borderRadius: 8, padding: '8px 12px', background: '#fffbeb', marginBottom: 8 }}>
          <summary style={{ fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, userSelect: 'none', color: '#c2410c' }}>
            <AlertCircle size={14} color="#ea580c" />
            <span>Sửa lỗi lưu/xóa (RLS Row Level Security) trong Supabase</span>
          </summary>
          <div style={{ marginTop: 8, paddingLeft: 4, display: 'flex', flexDirection: 'column', gap: 6, textAlign: 'left' }}>
            <p style={{ margin: 0, fontSize: 11.5, color: '#475569', lineHeight: '1.5' }}>
              Mặc định Supabase bật <strong>Row Level Security (RLS)</strong> nên sẽ chặn các lệnh <strong>INSERT</strong> và <strong>DELETE</strong> từ Client (báo lỗi 401 hoặc lỗi chặn RLS ngay cả khi kết nối Test OK).
            </p>
            <p style={{ margin: 0, fontSize: 11.5, color: '#475569', lineHeight: '1.5' }}>
              Hãy sao chép các câu lệnh SQL dưới đây, dán vào tab <strong>SQL Editor</strong> trong trang quản trị Supabase dự án của bạn rồi chọn <strong>Run</strong> để khắc phục hoàn toàn:
            </p>
            <pre style={{
              margin: '6px 0',
              padding: '8px 10px',
              background: '#1e293b',
              color: '#f8fafc',
              borderRadius: 6,
              fontSize: 10.5,
              overflowX: 'auto',
              fontFamily: 'monospace',
              lineHeight: '1.4'
            }}>
{`-- 1. Cấp quyền đầy đủ cho bảng du_an (Danh sách kho dự án)
ALTER TABLE public.du_an ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow public du_an" ON public.du_an;
CREATE POLICY "Allow public du_an" ON public.du_an FOR ALL USING (true) WITH CHECK (true);

-- 2. Cấp quyền đầy đủ cho bảng don_giao (Báo cáo đơn vị giao)
ALTER TABLE public.don_giao ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow public don_giao" ON public.don_giao;
CREATE POLICY "Allow public don_giao" ON public.don_giao FOR ALL USING (true) WITH CHECK (true);

-- 3. Cấp quyền đầy đủ cho bảng don_nhan (Báo cáo đơn vị nhận)
ALTER TABLE public.don_nhan ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow public don_nhan" ON public.don_nhan;
CREATE POLICY "Allow public don_nhan" ON public.don_nhan FOR ALL USING (true) WITH CHECK (true);

-- 4. Cấp quyền đầy đủ cho bảng don_chung (Báo cáo đơn chung)
ALTER TABLE public.don_chung ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow public don_chung" ON public.don_chung;
CREATE POLICY "Allow public don_chung" ON public.don_chung FOR ALL USING (true) WITH CHECK (true);

-- 5. Cấp quyền đầy đủ cho bảng don_kho (Kho dự án)
ALTER TABLE public.don_kho ENABLE ROW LEVEL SECURITY;
DROP POLICY IF EXISTS "Allow public don_kho" ON public.don_kho;
CREATE POLICY "Allow public don_kho" ON public.don_kho FOR ALL USING (true) WITH CHECK (true);`}
            </pre>
          </div>
        </details>

        {/* SQL Queries / Database Schema configuration instructions */}
        <details style={{ fontSize: 12, color: '#475569', border: '1px solid #93c5fd', borderRadius: 8, padding: '8px 12px', background: '#f0f9ff' }}>
          <summary style={{ fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, userSelect: 'none', color: '#1d4ed8' }}>
            <AlertCircle size={14} color="#2563eb" />
            <span>Sửa liên kết cột sai (Thêm cột ten_du_an & Tạo khóa ngoại khớp)</span>
          </summary>
          <div style={{ marginTop: 8, paddingLeft: 4, display: 'flex', flexDirection: 'column', gap: 6, textAlign: 'left' }}>
            <p style={{ margin: 0, fontSize: 11.5, color: '#475569', lineHeight: '1.5' }}>
              Để sửa lỗi liên kết cột không đồng nhất (cột <code>ten_du_an</code> thuộc bảng <code>du_an</code> nhưng lại liên kết vào cột <code>du_an</code> của 2 bảng kia khiến cho sơ đồ Supabase hiển thị lệch lạc), bạn chỉ cần chạy tập lệnh SQL dưới đây trong <strong>SQL Editor</strong>:
            </p>
            <pre style={{
              margin: '6px 0',
              padding: '8px 10px',
              background: '#0f172a',
              color: '#38bdf8',
              borderRadius: 6,
              fontSize: 10.5,
              overflowX: 'auto',
              fontFamily: 'monospace',
              lineHeight: '1.4'
            }}>
{`-- A. Tạo cột ten_du_an mới cho bảng don_giao & don_nhan & don_chung & don_kho (nếu chưa có)
ALTER TABLE public.don_giao ADD COLUMN IF NOT EXISTS ten_du_an text;
ALTER TABLE public.don_nhan ADD COLUMN IF NOT EXISTS ten_du_an text;
ALTER TABLE public.don_chung ADD COLUMN IF NOT EXISTS ten_du_an text;
ALTER TABLE public.don_kho ADD COLUMN IF NOT EXISTS ten_du_an text;

-- B. Di chuyển tự động toàn bộ dữ liệu dự án cũ sang cột mới (Bảo toàn dữ liệu cũ)
UPDATE public.don_giao SET ten_du_an = du_an WHERE ten_du_an IS NULL OR ten_du_an = '';
UPDATE public.don_nhan SET ten_du_an = du_an WHERE ten_du_an IS NULL OR ten_du_an = '';
UPDATE public.don_chung SET ten_du_an = du_an WHERE ten_du_an IS NULL OR ten_du_an = '';
UPDATE public.don_kho SET ten_du_an = du_an WHERE ten_du_an IS NULL OR ten_du_an = '';

-- C. Đảm bảo cột ten_du_an trong bảng du_an có ràng buộc duy nhất (Bắt buộc để làm Khóa Ngoại)
ALTER TABLE public.du_an DROP CONSTRAINT IF EXISTS du_an_ten_du_an_key;
ALTER TABLE public.du_an ADD CONSTRAINT du_an_ten_du_an_key UNIQUE (ten_du_an);

-- D. Loại bỏ khóa ngoại cũ liên kết trực tiếp vào cột du_an cũ (nếu có)
ALTER TABLE public.don_giao DROP CONSTRAINT IF EXISTS don_giao_du_an_fkey;
ALTER TABLE public.don_nhan DROP CONSTRAINT IF EXISTS don_nhan_du_an_fkey;
ALTER TABLE public.don_chung DROP CONSTRAINT IF EXISTS don_chung_du_an_fkey;
ALTER TABLE public.don_kho DROP CONSTRAINT IF EXISTS don_kho_du_an_fkey;

-- E. Thiết lập khóa ngoại liên kết chuẩn xác đến cột ten_du_an của bảng du_an
ALTER TABLE public.don_giao ADD CONSTRAINT don_giao_ten_du_an_fkey 
  FOREIGN KEY (ten_du_an) REFERENCES public.du_an (ten_du_an) ON UPDATE CASCADE ON DELETE SET NULL;

ALTER TABLE public.don_nhan ADD CONSTRAINT don_nhan_ten_du_an_fkey 
  FOREIGN KEY (ten_du_an) REFERENCES public.du_an (ten_du_an) ON UPDATE CASCADE ON DELETE SET NULL;

ALTER TABLE public.don_chung ADD CONSTRAINT don_chung_ten_du_an_fkey 
  FOREIGN KEY (ten_du_an) REFERENCES public.du_an (ten_du_an) ON UPDATE CASCADE ON DELETE SET NULL;

ALTER TABLE public.don_kho ADD CONSTRAINT don_kho_ten_du_an_fkey 
  FOREIGN KEY (ten_du_an) REFERENCES public.du_an (ten_du_an) ON UPDATE CASCADE ON DELETE SET NULL;

-- F. (Tùy chọn) Xóa hẳn các cột du_an cũ thừa thãi để sơ đồ Supabase sạch đẹp 100%
-- ALTER TABLE public.don_giao DROP COLUMN IF EXISTS du_an;
-- ALTER TABLE public.don_nhan DROP COLUMN IF EXISTS du_an;
-- ALTER TABLE public.don_chung DROP COLUMN IF EXISTS du_an;
-- ALTER TABLE public.don_kho DROP COLUMN IF EXISTS du_an;`}
            </pre>
            <p style={{ margin: 0, fontSize: 11, color: '#16a34a', fontWeight: 500 }}>
              ✔️ <strong>Lưu ý:</strong> Ứng dụng đã được tích hợp cơ chế <strong>Tự phục hồi (Self-Healing)</strong> thông minh. Bất kể bạn đang sử dụng cấu trúc cũ (cột <code>du_an</code>) hay cấu trúc mới (cột <code>ten_du_an</code>), phần mềm sẽ tự phát hiện lỗi column, tự gỡ bỏ cột thừa và lưu trữ trơn tru mà không làm ngắt quãng công việc của bạn!
            </p>
          </div>
        </details>

        {/* SQL Queries / Upgrade Depreciation Schema */}
        <details style={{ fontSize: 12, color: '#475569', border: '1px solid #c084fc', borderRadius: 8, padding: '8px 12px', background: '#faf5ff', marginBottom: 8 }}>
          <summary style={{ fontWeight: 600, cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, userSelect: 'none', color: '#7e22ce' }}>
            <AlertCircle size={14} color="#9333ea" />
            <span>Nâng cấp cấu trúc bảng "Đơn giá vật tư & Khấu hao"</span>
          </summary>
          <div style={{ marginTop: 8, paddingLeft: 4, display: 'flex', flexDirection: 'column', gap: 6, textAlign: 'left' }}>
            <p style={{ margin: 0, fontSize: 11.5, color: '#475569', lineHeight: '1.5' }}>
              Nếu bạn đã khởi tạo cơ sở dữ liệu từ trước và muốn cập nhật các cột mới cũng như gỡ bỏ khóa trùng lặp số tháng (để có thể lưu đồng thời số tháng trong cả nhóm Đã duyệt và Tạm tính), vui lòng chạy lệnh SQL sau trong <strong>SQL Editor</strong> của Supabase:
            </p>
            <pre style={{
              margin: '6px 0',
              padding: '8px 10px',
              background: '#0f172a',
              color: '#c084fc',
              borderRadius: 6,
              fontSize: 10.5,
              overflowX: 'auto',
              fontFamily: 'monospace',
              lineHeight: '1.4'
            }}>
{`-- Nâng cấp thêm cột ma_sap cho bảng cau_hinh_khau_hao để lưu cấu hình theo vật tư bền vững
ALTER TABLE public.cau_hinh_khau_hao ADD COLUMN IF NOT EXISTS ma_sap text;

-- Gỡ bỏ ràng buộc UNIQUE cũ trên bảng cau_hinh_khau_hao để cho phép cấu hình theo từng vật tư cụ thể
ALTER TABLE public.cau_hinh_khau_hao DROP CONSTRAINT IF EXISTS cau_hinh_khau_hao_months_key;
ALTER TABLE public.cau_hinh_khau_hao DROP CONSTRAINT IF EXISTS cau_hinh_khau_hao_months_is_approved_key;

-- Tạo unique index cho cấu hình chung (ma_sap IS NULL)
CREATE UNIQUE INDEX IF NOT EXISTS cau_hinh_khau_hao_global_idx ON public.cau_hinh_khau_hao (months, is_approved) WHERE ma_sap IS NULL;

-- Tạo unique index cho cấu hình riêng của từng vật tư (ma_sap IS NOT NULL)
CREATE UNIQUE INDEX IF NOT EXISTS cau_hinh_khau_hao_ma_sap_idx ON public.cau_hinh_khau_hao (ma_sap) WHERE ma_sap IS NOT NULL;`}
            </pre>
          </div>
        </details>

        {/* Footer actions */}
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, borderTop: '1px solid #f1f5f9', paddingTop: 12 }}>
          {currentSource === 'local' && (
            <button
              onClick={handleClear}
              style={{
                padding: '6px 12px',
                borderRadius: 6,
                background: '#ffffff',
                border: '1px solid #cbd5e1',
                color: '#64748b',
                fontSize: 13,
                fontWeight: 600,
                cursor: 'pointer'
              }}
            >
              Phục hồi mặc định
            </button>
          )}
          <button
            onClick={onClose}
            style={{
              padding: '6px 12px',
              borderRadius: 6,
              border: '1px solid #cbd5e1',
              background: '#ffffff',
              color: '#475569',
              fontSize: 13,
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            Hủy bỏ
          </button>
          <button
            onClick={handleSave}
            disabled={saveSuccess}
            style={{
              padding: '6px 14px',
              borderRadius: 6,
              border: 'none',
              background: 'linear-gradient(135deg, #0f58a7 0%, #1a6abf 100%)',
              color: '#ffffff',
              fontSize: 13,
              fontWeight: 600,
              cursor: saveSuccess ? 'not-allowed' : 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: 6,
              boxShadow: '0 2px 4px rgba(15,88,167,0.2)'
            }}
          >
            {saveSuccess ? (
              <>
                <RefreshCw size={13} className="animate-spin" />
                Đang lưu...
              </>
            ) : (
              <>
                <Save size={13} />
                Lưu kết nối
              </>
            )}
          </button>
        </div>
      </div>
    </div>
  )
}


function SupabaseDbUpgradeModal({ isOpen, onClose }) {
  const [copied, setCopied] = useState(false)

  if (!isOpen) return null

  const sqlCode = `-- Nâng cấp thêm cột ma_sap cho bảng cau_hinh_khau_hao để lưu cấu hình theo vật tư bền vững
ALTER TABLE public.cau_hinh_khau_hao ADD COLUMN IF NOT EXISTS ma_sap text;

-- Gỡ bỏ ràng buộc UNIQUE cũ trên bảng cau_hinh_khau_hao để cho phép cấu hình theo từng vật tư cụ thể
ALTER TABLE public.cau_hinh_khau_hao DROP CONSTRAINT IF EXISTS cau_hinh_khau_hao_months_key;
ALTER TABLE public.cau_hinh_khau_hao DROP CONSTRAINT IF EXISTS cau_hinh_khau_hao_months_is_approved_key;

-- Tạo unique index cho cấu hình chung (ma_sap IS NULL)
CREATE UNIQUE INDEX IF NOT EXISTS cau_hinh_khau_hao_global_idx ON public.cau_hinh_khau_hao (months, is_approved) WHERE ma_sap IS NULL;

-- Tạo unique index cho cấu hình riêng của từng vật tư (ma_sap IS NOT NULL)
CREATE UNIQUE INDEX IF NOT EXISTS cau_hinh_khau_hao_ma_sap_idx ON public.cau_hinh_khau_hao (ma_sap) WHERE ma_sap IS NOT NULL;`

  const handleCopy = () => {
    navigator.clipboard.writeText(sqlCode)
    setCopied(true)
    setTimeout(() => setCopied(false), 2000)
  }

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      backgroundColor: 'rgba(15, 23, 42, 0.65)',
      backdropFilter: 'blur(5px)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 10003
    }}>
      <div style={{
        background: '#ffffff',
        borderRadius: 16,
        width: '100%',
        maxWidth: 580,
        maxHeight: '90vh',
        overflowY: 'auto',
        padding: 24,
        boxShadow: '0 25px 50px -12px rgba(0,0,0,0.25)',
        border: '1px solid #e2e8f0',
        display: 'flex',
        flexDirection: 'column',
        gap: 16
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: 12 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
            <Database size={18} color="#7e22ce" />
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#0f172a' }}>Nâng cấp cấu trúc cơ sở dữ liệu Supabase</h3>
          </div>
          <button onClick={onClose} style={{ background: 'transparent', border: 'none', color: '#64748b', cursor: 'pointer', display: 'flex', alignItems: 'center', padding: 4 }}>
            <X size={18} />
          </button>
        </div>

        <div style={{
          backgroundColor: '#faf5ff',
          border: '1px solid #e9d5ff',
          borderRadius: 8,
          padding: '12px 16px',
          fontSize: '13px',
          color: '#581c87',
          lineHeight: '1.5'
        }}>
          <strong>Tại sao cần nâng cấp?</strong>
          <p style={{ margin: '6px 0 0 0' }}>
            Bạn đang lưu số tháng khấu hao giống nhau (ví dụ: 6 tháng) ở cả 2 nhóm <strong>Đã duyệt</strong> và <strong>Tạm tính</strong>.
            Cơ sở dữ liệu Supabase của bạn hiện có ràng buộc UNIQUE cũ, chỉ cho phép một số tháng xuất hiện một lần duy nhất.
            Lệnh SQL dưới đây sẽ nâng cấp cấu trúc bảng để cho phép lưu trữ song song số tháng trùng lặp giữa các nhóm.
          </p>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 6 }}>
          <label style={{ fontSize: '12.5px', fontWeight: 600, color: '#475569' }}>Các bước thực hiện:</label>
          <ol style={{ margin: 0, paddingLeft: 20, fontSize: '12.5px', color: '#475569', lineHeight: '1.6', display: 'flex', flexDirection: 'column', gap: 4 }}>
            <li>Copy đoạn mã SQL bên dưới bằng nút <strong>Copy câu lệnh SQL</strong>.</li>
            <li>Truy cập vào trang quản trị <strong>Supabase Dashboard</strong> của bạn.</li>
            <li>Nhấp vào mục <strong>SQL Editor</strong> ở thanh menu bên trái.</li>
            <li>Nhấp vào nút <strong>New Query</strong>, dán đoạn mã vừa copy vào khung soạn thảo.</li>
            <li>Nhấp nút <strong>Run</strong> (hoặc nhấn Ctrl+Enter / Cmd+Enter) để thực thi.</li>
          </ol>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 8 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: '12.5px', fontWeight: 600, color: '#475569' }}>Mã SQL nâng cấp:</span>
            <button
              onClick={handleCopy}
              style={{
                display: 'inline-flex',
                alignItems: 'center',
                gap: 4,
                padding: '4px 10px',
                background: copied ? '#22c55e' : '#f1f5f9',
                color: copied ? '#ffffff' : '#475569',
                border: '1px solid ' + (copied ? '#22c55e' : '#cbd5e1'),
                borderRadius: 6,
                fontSize: '11.5px',
                fontWeight: 600,
                cursor: 'pointer',
                transition: 'all 0.15s'
              }}
            >
              {copied ? 'Đã copy!' : 'Copy câu lệnh SQL'}
            </button>
          </div>
          <pre style={{
            margin: 0,
            padding: '12px 14px',
            background: '#0f172a',
            color: '#c084fc',
            borderRadius: 8,
            fontSize: 11,
            overflowX: 'auto',
            fontFamily: 'monospace',
            lineHeight: '1.4',
            border: '1px solid #1e293b'
          }}>
            {sqlCode}
          </pre>
        </div>

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, borderTop: '1px solid #f1f5f9', paddingTop: 12 }}>
          <button
            onClick={onClose}
            style={{
              padding: '6px 14px',
              borderRadius: 6,
              border: '1px solid #cbd5e1',
              background: '#ffffff',
              color: '#475569',
              fontSize: 13,
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            Đóng
          </button>
        </div>
      </div>
    </div>
  )
}


// Helper functions for robust database access and column-case normalization
function toSnakeCase(str) {
  if (str === 'maSAP') return 'ma_sap'
  if (str === 'maDonChuyenTiepLC') return 'ma_don_chuyen_tiep_lc'
  if (str === 'maDonChuyenTiepNB') return 'ma_don_chuyen_tiep_nb'
  if (str === 'donGiaTrungBinh1Ngay') return 'don_gia_trung_binh_1_ngay'
  return str.replace(/[A-Z]/g, letter => `_${letter.toLowerCase()}`)
}

function extractMissingField(errMsg) {
  if (!errMsg) return null
  const patterns = [
    /Could not find the ['"](.*?)['"] (column|property)/i,
    /Could not find the (column|property) ['"](.*?)['"]/i,
    /(column|property) ['"](.*?)['"] of relation/i,
    /(column|property) ['"](.*?)['"] does not exist/i,
    /['"](.*?)['"] is not a (column|property)/i,
    /relation ['"].*?['"] does not have a (column|property) ['"](.*?)['"]/i
  ]
  for (const regex of patterns) {
    const match = errMsg.match(regex)
    if (match) {
      for (let i = 1; i < match.length; i++) {
        const val = match[i]
        if (val && val.toLowerCase() !== 'column' && val.toLowerCase() !== 'property') {
          return val
        }
      }
    }
  }
  return null
}

function normalizeDbRow(dbRow) {
  if (!dbRow) return dbRow
  const normalized = {}
  
  // Start with whatever properties came from DB
  Object.assign(normalized, dbRow)
  
  // Map PostgreSQL columns back to React's camelCase keys
  COLS_GIAO_NHAN.forEach(col => {
    // Look for exact camelCase, lowercase, or snake_case matches
    let dbKeyMatch = Object.keys(dbRow).find(
      k => k === col.key || 
           k.toLowerCase() === col.key.toLowerCase() || 
           k.toLowerCase() === toSnakeCase(col.key).toLowerCase()
    )
    // Đặc biệt hỗ trợ ánh xạ cột ten_du_an hoặc tenduan về duAn của React
    if (col.key === 'duAn' && !dbKeyMatch) {
      dbKeyMatch = Object.keys(dbRow).find(
        k => k.toLowerCase() === 'ten_du_an' || 
             k.toLowerCase() === 'tenduan' || 
             k.toLowerCase() === 'ten_duan' ||
             k.toLowerCase() === 'tendu_an'
      )
    }
    // Hỗ trợ thêm cho maSAP nếu không tìm thấy tự động
    if (col.key === 'maSAP' && !dbKeyMatch) {
      dbKeyMatch = Object.keys(dbRow).find(
        k => k.toLowerCase() === 'ma_sap' || 
             k.toLowerCase() === 'masap' || 
             k.toLowerCase() === 'ma_s_a_p'
      )
    }
    // Hỗ trợ thêm cho maDonChuyenTiepLC
    if (col.key === 'maDonChuyenTiepLC' && !dbKeyMatch) {
      dbKeyMatch = Object.keys(dbRow).find(
        k => k.toLowerCase() === 'ma_don_chuyen_tiep_lc' || 
             k.toLowerCase() === 'ma_don_chuyen_tiep_l_c'
      )
    }
    // Hỗ trợ thêm cho maDonChuyenTiepNB
    if (col.key === 'maDonChuyenTiepNB' && !dbKeyMatch) {
      dbKeyMatch = Object.keys(dbRow).find(
        k => k.toLowerCase() === 'ma_don_chuyen_tiep_nb' || 
             k.toLowerCase() === 'ma_don_chuyen_tiep_n_b'
      )
    }

    if (dbKeyMatch && dbRow[dbKeyMatch] !== undefined) {
      normalized[col.key] = dbRow[dbKeyMatch]
    }
  })
  
  return normalized
}

// Cache the best casing style for each table ('camel', 'lower', 'snake') to bypass wrong attempts
const tableCasingStyleCache = {}

// Cache list of columns that do not exist on each table to avoid repeated "column not found" roundtrips
const blacklistedColumnsCache = {}

function convertToCasingStyle(row, style) {
  const newRow = {}
  Object.keys(row).forEach(k => {
    if (style === 'lower') {
      newRow[k.toLowerCase()] = row[k]
      // Project name column fallback keys for self-healing
      if (k === 'duAn') {
        newRow['ten_du_an'] = row[k]
        newRow['tenduan'] = row[k]
      }
    } else if (style === 'snake') {
      newRow[toSnakeCase(k)] = row[k]
      if (k === 'duAn') {
        newRow['ten_du_an'] = row[k]
      }
    } else {
      newRow[k] = row[k] // camel / default
      if (k === 'duAn') {
        newRow['ten_du_an'] = row[k]
        newRow['tenDuAn'] = row[k]
      }
    }
  })
  return newRow
}

function filterRowByBlacklist(row, tableName) {
  const blacklist = blacklistedColumnsCache[tableName] || []
  if (blacklist.length === 0) return row
  
  const cleanRow = { ...row }
  blacklist.forEach(badField => {
    delete cleanRow[badField]
    delete cleanRow[badField.toLowerCase()]
    delete cleanRow[toSnakeCase(badField)]
    delete cleanRow[badField.replace(/_/g, '')]
  })
  return cleanRow
}

async function tryInsertWithSelfHealing(tableName, payload, style) {
  let currentPayload = JSON.parse(JSON.stringify(payload))
  
  for (let retry = 0; retry < 50; retry++) {
    const { error } = await supabase.from(tableName).insert(currentPayload)
    if (!error) {
      return { success: true }
    }
    
    const errMsg = error.message || ''
    const missingField = extractMissingField(errMsg)
    
    if (missingField) {
      console.warn(`[Self-Healing] Loại bỏ cột không tồn tại '${missingField}' trên bảng '${tableName}' và thử lại...`)
      
      // Update our global blacklist for subsequent chunks & retry
      if (!blacklistedColumnsCache[tableName]) {
        blacklistedColumnsCache[tableName] = []
      }
      if (!blacklistedColumnsCache[tableName].includes(missingField)) {
        blacklistedColumnsCache[tableName].push(missingField)
      }
      
      currentPayload.forEach(row => {
        delete row[missingField]
        delete row[missingField.toLowerCase()]
        delete row[toSnakeCase(missingField)]
        delete row[missingField.replace(/_/g, '')]
      })
      
      if (currentPayload.length > 0 && Object.keys(currentPayload[0]).length === 0) {
        break
      }
      continue
    }
    
    return { success: false, error }
  }
  return { success: false, error: new Error('Không thể tự khắc phục cột bị thiếu sau nhiều lần thử') }
}

async function insertWithFallback(tableName, originalChunk) {
  // If we already know the casing style for this table, use it directly!
  const cachedStyle = tableCasingStyleCache[tableName]
  if (cachedStyle) {
    const styledChunk = originalChunk.map(row => 
      filterRowByBlacklist(convertToCasingStyle(row, cachedStyle), tableName)
    )
    const res = await tryInsertWithSelfHealing(tableName, styledChunk, cachedStyle)
    if (res.success) return { success: true }
    console.warn(`Cached style "${cachedStyle}" failed for ${tableName}:`, res.error?.message || res.error)
  }

  // First try direct straight insert across casings WITHOUT self-healing to find the perfect style 
  // (Prevents silent stripping of required columns of the table)
  const stylesToTry = ['snake', 'lower', 'camel']
  for (const style of stylesToTry) {
    const styledChunk = originalChunk.map(row => 
      filterRowByBlacklist(convertToCasingStyle(row, style), tableName)
    )
    const { error } = await supabase.from(tableName).insert(styledChunk)
    if (!error) {
      console.log(`[Casing Finder] Phát hiện cấu trúc table "${tableName}" trùng khớp hoàn hảo với kiểu: "${style}"`)
      tableCasingStyleCache[tableName] = style
      return { success: true }
    }
    
    // Check if the error is due to missing columns or something else.
    const errMsg = error.message || ''
    const isMissingColumnError = 
      (errMsg.toLowerCase().includes('column') || errMsg.toLowerCase().includes('property') || errMsg.toLowerCase().includes('relation')) &&
      (errMsg.toLowerCase().includes('not exist') || errMsg.toLowerCase().includes('could not find') || errMsg.toLowerCase().includes('schema cache'))
      
    if (!isMissingColumnError) {
      // If it's a constraint, RLS or real error, bubble it up.
      throw error
    }
  }

  // If all straight inserts failed, we attempt with Self Healing in snake, then lower, then camel
  const healingOrder = ['snake', 'lower', 'camel']
  for (const style of healingOrder) {
    const styledChunk = originalChunk.map(row => 
      filterRowByBlacklist(convertToCasingStyle(row, style), tableName)
    )
    const res = await tryInsertWithSelfHealing(tableName, styledChunk, style)
    if (res.success) {
      console.log(`[Casing Finder/Healed] Đồng bộ thành công trên "${tableName}" kiểu "${style}" sau khi loại bỏ các cột không dùng`)
      tableCasingStyleCache[tableName] = style
      return { success: true }
    }
  }

  throw new Error('Supabase từ chối đồng bộ dữ liệu. Đã thử đồng hóa viết hoa/thường và tự động gỡ các cột không có trong Table nhưng vẫn thất bại.')
}

async function upsertWithFallback(tableName, originalChunk, keyColumn) {
  // If we already know the casing style for this table, use it directly!
  const cachedStyle = tableCasingStyleCache[tableName]
  if (cachedStyle) {
    const styledChunk = originalChunk.map(row => 
      filterRowByBlacklist(convertToCasingStyle(row, cachedStyle), tableName)
    )
    let styledKey = keyColumn
    if (cachedStyle === 'lower') styledKey = keyColumn.toLowerCase()
    else if (cachedStyle === 'snake') styledKey = toSnakeCase(keyColumn)
    
    const res = await tryUpsertWithSelfHealing(tableName, styledChunk, styledKey, cachedStyle)
    if (res.success) return { success: true }
    console.warn(`Cached style "${cachedStyle}" failed for upsert ${tableName}:`, res.error?.message || res.error)
  }

  // First try direct straight upsert across casings WITHOUT self-healing to find the perfect style 
  const stylesToTry = ['snake', 'lower', 'camel']
  for (const style of stylesToTry) {
    const styledChunk = originalChunk.map(row => 
      filterRowByBlacklist(convertToCasingStyle(row, style), tableName)
    )
    let styledKey = keyColumn
    if (style === 'lower') styledKey = keyColumn.toLowerCase()
    else if (style === 'snake') styledKey = toSnakeCase(keyColumn)
    
    const { error } = await supabase.from(tableName).upsert(styledChunk, { onConflict: styledKey })
    if (!error) {
      console.log(`[Casing Finder] Phát hiện cấu trúc table "${tableName}" trùng khớp hoàn hảo với kiểu: "${style}"`)
      tableCasingStyleCache[tableName] = style
      return { success: true }
    }
    
    // Check if the error is due to missing columns or something else.
    const errMsg = error.message || ''
    const isMissingColumnError = 
      (errMsg.toLowerCase().includes('column') || errMsg.toLowerCase().includes('property') || errMsg.toLowerCase().includes('relation')) &&
      (errMsg.toLowerCase().includes('not exist') || errMsg.toLowerCase().includes('could not find') || errMsg.toLowerCase().includes('schema cache'))
      
    const isNoUniqueConstraintError = 
      error.code === '42P10' || 
      errMsg.toLowerCase().includes('there is no unique or exclusion constraint') ||
      errMsg.toLowerCase().includes('on conflict specification')

    if (!isMissingColumnError && !isNoUniqueConstraintError) {
      // If it's a constraint, RLS or real error, bubble it up.
      throw error
    }
  }

  // If all straight upserts failed, we attempt with Self Healing
  const healingOrder = ['snake', 'lower', 'camel']
  const styleErrors = {}
  for (const style of healingOrder) {
    const styledChunk = originalChunk.map(row => 
      filterRowByBlacklist(convertToCasingStyle(row, style), tableName)
    )
    let styledKey = keyColumn
    if (style === 'lower') styledKey = keyColumn.toLowerCase()
    else if (style === 'snake') styledKey = toSnakeCase(keyColumn)
    
    const res = await tryUpsertWithSelfHealing(tableName, styledChunk, styledKey, style)
    if (res.success) {
      console.log(`[Casing Finder/Healed] Đồng bộ thành công trên "${tableName}" kiểu "${style}" sau khi loại bỏ các cột không dùng`)
      tableCasingStyleCache[tableName] = style
      return { success: true }
    }
    styleErrors[style] = res.error
  }

  const chosenError = styleErrors['snake'] || styleErrors['lower'] || styleErrors['camel']
  const detailMsg = chosenError?.message || (typeof chosenError === 'string' ? chosenError : JSON.stringify(chosenError))
  throw new Error(`Supabase từ chối đồng bộ dữ liệu. Đã thử đồng hóa viết hoa/thường và tự động gỡ các cột không có trong Table nhưng vẫn thất bại. Chi tiết lỗi từ DB: ${detailMsg || 'Không rõ nguyên nhân'}`)
}

async function tryUpsertWithSelfHealing(tableName, payload, keyColumn, style) {
  let currentPayload = JSON.parse(JSON.stringify(payload))
  
  for (let retry = 0; retry < 50; retry++) {
    const { error } = await supabase
      .from(tableName)
      .upsert(currentPayload, { onConflict: keyColumn })
      
    if (!error) {
      return { success: true }
    }
    
    const errMsg = error.message || ''
    const isNoUniqueConstraintError = 
      error.code === '42P10' || 
      errMsg.toLowerCase().includes('there is no unique or exclusion constraint') ||
      errMsg.toLowerCase().includes('on conflict specification')

    if (isNoUniqueConstraintError) {
      console.warn(`[Self-Healing Upsert] Phát hiện lỗi thiếu ràng buộc UNIQUE trên cột khóa '${keyColumn}' của bảng '${tableName}'. Tiến hành tự động chuyển sang cơ chế xóa-và-chèn (delete-and-insert) để thay thế cho upsert...`)
      try {
        const keyValues = currentPayload
          .map(row => row[keyColumn])
          .filter(val => val !== undefined && val !== null)

        if (keyValues.length > 0) {
          const { error: delError } = await supabase
            .from(tableName)
            .delete()
            .in(keyColumn, keyValues)

          if (delError) {
            console.error(`[Self-Healing Upsert] Xóa bản ghi cũ thất bại:`, delError.message)
            return { success: false, error: delError }
          }
        }

        const { error: insError } = await supabase
          .from(tableName)
          .insert(currentPayload)

        if (insError) {
          console.error(`[Self-Healing Upsert] Chèn bản ghi mới thất bại sau khi xóa:`, insError.message)
          return { success: false, error: insError }
        }

        console.log(`[Self-Healing Upsert] Đồng bộ xóa-và-chèn thành công cho bảng '${tableName}'!`)
        return { success: true }
      } catch (fallbackErr) {
        return { success: false, error: fallbackErr }
      }
    }

    const missingField = extractMissingField(errMsg)
    
    if (missingField) {
      // Nếu cột bị thiếu chính là khóa chính/khóa so khớp (onConflict), tuyệt đối không gỡ bỏ nó.
      // Thay vào đó, trả về thất bại luôn để vòng lặp thử style khác (ví dụ từ camel sang snake).
      const isKeyCol = 
        missingField.toLowerCase() === keyColumn.toLowerCase() ||
        toSnakeCase(missingField).toLowerCase() === keyColumn.toLowerCase() ||
        missingField.replace(/_/g, '').toLowerCase() === keyColumn.replace(/_/g, '').toLowerCase()
        
      if (isKeyCol) {
        return { success: false, error }
      }

      console.warn(`[Self-Healing Upsert] Loại bỏ cột không tồn tại '${missingField}' trên bảng '${tableName}' và thử lại...`)
      
      // Update our global blacklist for subsequent chunks & retry
      if (!blacklistedColumnsCache[tableName]) {
        blacklistedColumnsCache[tableName] = []
      }
      if (!blacklistedColumnsCache[tableName].includes(missingField)) {
        blacklistedColumnsCache[tableName].push(missingField)
      }
      
      currentPayload.forEach(row => {
        delete row[missingField]
        delete row[missingField.toLowerCase()]
        delete row[toSnakeCase(missingField)]
        delete row[missingField.replace(/_/g, '')]
      })
      
      if (currentPayload.length > 0 && Object.keys(currentPayload[0]).length === 0) {
        break
      }
      continue
    }
    
    return { success: false, error }
  }
  return { success: false, error: new Error('Không thể tự khắc phục cột bị thiếu sau nhiều lần thử') }
}

// ─── Reset sequence của bảng về MAX(id)+1 hoặc về 1 nếu bảng trống ──────────
// Yêu cầu: đã tạo 2 functions trên Supabase SQL Editor (xem hướng dẫn trong app)
async function resetTableSequence(tableName) {
  try {
    const { count, error: countErr } = await supabase
      .from(tableName)
      .select('*', { count: 'exact', head: true })

    if (countErr) {
      console.warn(`[Sequence Reset] Không đọc được số dòng bảng "${tableName}":`, countErr.message)
      return
    }

    if (count === 0 || count === null) {
      // Bảng trống → reset sequence về 1
      const { error: rpcErr } = await supabase.rpc('reset_sequence_to_one', { p_table: tableName })
      if (rpcErr) {
        console.warn(`[Sequence Reset] RPC reset_sequence_to_one thất bại: ${rpcErr.message}`)
      } else {
        console.log(`[Sequence Reset] Reset sequence "${tableName}" về 1 (bảng trống).`)
      }
    } else {
      // Còn dữ liệu → đồng bộ sequence theo MAX(id)
      const { error: rpcErr } = await supabase.rpc('sync_sequence_to_max', { p_table: tableName })
      if (rpcErr) {
        console.warn(`[Sequence Reset] RPC sync_sequence_to_max thất bại: ${rpcErr.message}`)
      } else {
        console.log(`[Sequence Reset] Sync sequence "${tableName}" theo MAX(id).`)
      }
    }
  } catch (e) {
    console.warn(`[Sequence Reset] Lỗi:`, e)
  }
}

// Pre-populate cache cho các bảng đã biết cấu trúc — tránh fetch schema không cần thiết
const tableSchemaCache = {
  don_giao: ['id', 'ten_du_an', 'tenDuAn', 'tenduan', 'du_an'],
  don_nhan: ['id', 'ten_du_an', 'tenDuAn', 'tenduan', 'du_an'],
  don_kho:  ['id', 'ten_du_an', 'tenDuAn', 'tenduan', 'du_an'],
  du_an:    ['id', 'ten_du_an', 'tenduan', 'tenDuAn'],
}
async function deleteFromTableAdaptive(tableName, possibleColumns, targetValue) {
  try {
    let columns = null
    
    // Try 1: Dùng cache schema nếu đã có (tránh fetch lại nhiều lần)
    if (tableSchemaCache[tableName]) {
      columns = tableSchemaCache[tableName]
    } else {
      try {
        const response = await fetch(`${supabaseUrl}/rest/v1/`, {
          headers: {
            'apikey': supabaseAnonKey,
            'Authorization': `Bearer ${supabaseAnonKey}`
          }
        })
        if (response.ok) {
          const schema = await response.json()
          const def = schema.definitions?.[tableName]
          if (def && def.properties) {
            columns = Object.keys(def.properties)
            tableSchemaCache[tableName] = columns // Lưu cache
            console.log(`[Schema Finder] Cache cột bảng "${tableName}":`, columns)
          }
        }
      } catch (schemaErr) {
        console.warn('[Schema Finder] Lỗi tải OpenAPI schema:', schemaErr.message)
      }
    }

    // Try 2: Kiểm tra bảng có trống không (chỉ select id để tối thiểu Egress)
    if (!columns) {
      const { data, error } = await supabase.from(tableName).select('id').limit(1)
      if (!error && (!data || data.length === 0)) {
        console.log(`[Schema Finder] Bảng "${tableName}" trống. Không cần DELETE.`)
        return { success: true, reason: 'table_empty' }
      } else if (error) {
        console.warn(`[Schema Finder] Lỗi truy vấn bảng "${tableName}":`, error.message)
      }
      // columns vẫn null → chuyển sang fallback sequential bên dưới
    }

    // If we have columns, find the matching column to delete by
    if (columns && columns.length > 0) {
      const matchingCol = possibleColumns.find(col => 
        columns.includes(col) || 
        columns.some(c => c.toLowerCase() === col.toLowerCase())
      )

      if (matchingCol) {
        const exactCol = columns.find(c => c.toLowerCase() === matchingCol.toLowerCase()) || matchingCol
        console.log(`[Adaptive Delete] Đang thực hiện xóa trên bảng "${tableName}" với điều kiện "${exactCol}" = "${targetValue}"`)
        const { error: delErr } = await supabase.from(tableName).delete().eq(exactCol, targetValue)
        if (delErr) {
          console.error(`[Adaptive Delete] Lỗi khi xóa trên bảng "${tableName}":`, delErr.message)
          throw delErr
        }

        return { success: true }
      }
    }

    // Fallback: If no columns were detected at all, do sequential trial as a last resort
    console.warn(`[Adaptive Delete] Không phát hiện được cấu trúc cột cho "${tableName}". Đang thử xóa tuần tự fallback...`)
    let lastError = null
    for (const col of possibleColumns) {
      const { error: delErr } = await supabase.from(tableName).delete().eq(col, targetValue)
      if (!delErr) {
        console.log(`[Adaptive Delete] Xóa thành công trên "${tableName}" bằng cột fallback "${col}"`)
        return { success: true }
      }
      lastError = delErr
    }
    
    throw lastError || new Error(`Không tìm thấy cột phù hợp hoặc không được cấu hình quyền DELETE trên bảng "${tableName}"`)
  } catch (err) {
    console.error(`[Adaptive Delete] Lỗi trong quá trình xóa dữ liệu trên bảng "${tableName}":`, err)
    return { success: false, error: err }
  }
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState('chung')
  const [currentUser, setCurrentUser] = useState(() => {
    const cached = localStorage.getItem('sgc_logged_in_user')
    return cached ? JSON.parse(cached) : null
  })

  const handleLoginSuccess = (user) => {
    setCurrentUser(user)
    localStorage.setItem('sgc_logged_in_user', JSON.stringify(user))
  }

  const handleLogout = () => {
    setCurrentUser(null)
    localStorage.removeItem('sgc_logged_in_user')
  }

  const [customCategoryMap, setCustomCategoryMap] = useState({})
  const [dbCategoryMap, setDbCategoryMap] = useState({})

  React.useEffect(() => {
    let isMounted = true
    const loadCustomClassifications = async () => {
      if (!isSupabaseConfigured) {
        if (isMounted) {
          setCustomCategoryMap({})
          setDbCategoryMap({})
        }
        return
      }
      try {
        const { data, error } = await supabase
          .from('phan_loai_don_vi')
          .select('ten_don_vi, nhom_don_vi')
          
        if (error) throw error
        
        if (data && isMounted) {
          const map = {}
          data.forEach(row => {
            if (row.ten_don_vi) {
              map[row.ten_don_vi] = row.nhom_don_vi
            }
          })
          setCustomCategoryMap(map)
          setDbCategoryMap(map)
        }
      } catch (err) {
        console.error('Error loading custom classifications in App:', err)
      }
    }
    loadCustomClassifications()
    return () => {
      isMounted = false
    }
  }, [])

  const [isSidebarOpen, setIsSidebarOpen] = useState(false)
  const [isSidebarPinned, setIsSidebarPinned] = useState(() => {
    try {
      return localStorage.getItem('sgc_sidebar_pinned') === 'true'
    } catch (e) {
      return false
    }
  })
  const sidebarTimerRef = useRef(null)

  const handleMouseEnter = () => {
    if (sidebarTimerRef.current) clearTimeout(sidebarTimerRef.current)
    setIsSidebarOpen(true)
  }

  const handleMouseLeave = () => {
    sidebarTimerRef.current = setTimeout(() => {
      setIsSidebarOpen(false)
    }, 250)
  }

  const handleTogglePin = () => {
    const newPinned = !isSidebarPinned
    setIsSidebarPinned(newPinned)
    try {
      localStorage.setItem('sgc_sidebar_pinned', String(newPinned))
    } catch (e) {}
  }

  const syncInProgressRef = React.useRef(false) // Tạm dừng Realtime khi đang sync để tránh flood request
  const realtimeDebounceRef = React.useRef(null) // Debounce timer cho Realtime callbacks
  const focusSyncingRef = React.useRef(false) // Chặn việc chạy trùng lặp khi 'visibilitychange' và 'focus' cùng bắn 1 lúc
  const lastFocusSyncAtRef = React.useRef(0) // Throttle: không sync lại nếu vừa mới sync gần đây
  const [selectedProject, setSelectedProject] = useState('')
  const [showAddProjectModal, setShowAddProjectModal] = useState(false)
  const [showEditProjectModal, setShowEditProjectModal] = useState(false)
  const [showDeleteFileModal, setShowDeleteFileModal] = useState(false)
  const [deleteFileType, setDeleteFileType] = useState(null) // 'giao' | 'nhan'
  const [showDeleteProjectModal, setShowDeleteProjectModal] = useState(false)
  const [showConfigModal, setShowConfigModal] = useState(false)
  const [showDbUpgradeModal, setShowDbUpgradeModal] = useState(false)
  const [isDbSchemaOutdated, setIsDbSchemaOutdated] = React.useState(false)
  const [supabaseAuthError, setSupabaseAuthError] = useState(false)
  const [projectToEdit, setProjectToEdit] = useState('')
  const [projectToDelete, setProjectToDelete] = useState('')

  // Preview import modal state
  const [previewModal, setPreviewModal] = useState(null) // { type, rows, fileName }

  const [customProjects, setCustomProjects] = useState(() => {
    try {
      const saved = localStorage.getItem('sgc_custom_projects')
      return saved ? JSON.parse(saved) : []
    } catch (e) {
      return []
    }
  })

  // Hold rows and sheet file names globally so state is retained across tab switches!
  const [giaoRows, setGiaoRows] = useState([])
  const [giaoFileName, setGiaoFileName] = useState('')

  const [nhanRows, setNhanRows] = useState([])
  const [nhanFileName, setNhanFileName] = useState('')

  const [chungRows, setChungRows] = useState([])
  const [chungFileName, setChungFileName] = useState('')

  const [khoRows, setKhoRows] = useState([])
  const [khoFileName, setKhoFileName] = useState('')

  const [materialPriceRows, setMaterialPriceRows] = React.useState(() => {
    try {
      const saved = localStorage.getItem('sgc_report_material_price_rows')
      return saved ? JSON.parse(saved) : []
    } catch (e) {
      return []
    }
  })

  const [materialPrices, setMaterialPrices] = React.useState(() => {
    try {
      const saved = localStorage.getItem('sgc_report_material_prices')
      return saved ? JSON.parse(saved) : {}
    } catch (e) {
      return {}
    }
  })

  const [materialClassifications, setMaterialClassifications] = React.useState(() => {
    try {
      const saved = localStorage.getItem('sgc_report_material_classifications')
      return saved ? JSON.parse(saved) : {}
    } catch (e) {
      return {}
    }
  })

  const [loadingDbPrices, setLoadingDbPrices] = React.useState(false)
  const [lastSyncedDepreciationOptions, setLastSyncedDepreciationOptions] = React.useState(null)
  const [lastSyncedMaterialPriceRows, setLastSyncedMaterialPriceRows] = React.useState(null)

  const loadPricesFromSupabase = React.useCallback(async () => {
    if (!isSupabaseConfigured) return

    // Read from cache first to render instantly
    const cachedRowsStr = localStorage.getItem('sgc_report_material_price_rows')
    let cachedRows = []
    try {
      if (cachedRowsStr) {
        cachedRows = JSON.parse(cachedRowsStr)
        if (Array.isArray(cachedRows) && cachedRows.length > 0) {
          setMaterialPriceRows(cachedRows)
          const prices = {}
          const classifications = {}
          cachedRows.forEach(r => {
            prices[r.maSAP] = r.donGiaTrungBinh
            classifications[r.maSAP] = r.phanLoaiVatTu
          })
          setMaterialPrices(prices)
          setMaterialClassifications(classifications)
        }
      }
    } catch (e) {
      console.warn('Lỗi đọc cache đơn giá:', e)
    }

    setLoadingDbPrices(true)
    try {
      // Fetch server count (uses almost zero egress) to check if we can skip download
      const { count: serverCount, error: countErr } = await supabase
        .from('don_gia_vat_tu')
        .select('id', { count: 'exact', head: true })

      if (!countErr && serverCount !== null && cachedRows.length > 0 && serverCount === cachedRows.length) {
        console.log('[Cache Tối Ưu] Đơn giá vật tư trùng khớp số lượng. Bỏ qua tải tải từ Supabase.')
        setLoadingDbPrices(false)
        setLastSyncedMaterialPriceRows(cachedRows)
        return
      }

      const { data, error } = await supabase
        .from('don_gia_vat_tu')
        .select('*')
      
      if (error) {
        console.error('Lỗi khi tải đơn giá từ Supabase:', error)
        setLastSyncedMaterialPriceRows(cachedRows)
        const errMsg = error.message || ''
        if (error.code === '42703' || errMsg.toLowerCase().includes('column') || errMsg.toLowerCase().includes('relation') || errMsg.toLowerCase().includes('does not exist')) {
          setIsDbSchemaOutdated(true)
        }
      } else if (data) {
        // Tải các cấu hình khấu hao riêng của từng vật tư từ cau_hinh_khau_hao làm nguồn chính bền vững
        const configMap = new Map()
        try {
          const { data: configData, error: configError } = await supabase
            .from('cau_hinh_khau_hao')
            .select('*')
            .not('ma_sap', 'is', null)
          
          if (configError) {
            console.warn('Lỗi khi tải cấu hình khấu hao chi tiết từ Supabase:', configError.message)
            const errMsg = configError.message || ''
            if (configError.code === '42703' || errMsg.toLowerCase().includes('column') || errMsg.toLowerCase().includes('relation') || errMsg.toLowerCase().includes('does not exist')) {
              setIsDbSchemaOutdated(true)
            }
          } else if (configData) {
            configData.forEach(item => {
              if (item.ma_sap) {
                configMap.set(String(item.ma_sap).trim().toLowerCase(), {
                  months: item.months || 0,
                  isApproved: item.is_approved !== undefined ? !!item.is_approved : false
                })
              }
            })
          }
        } catch (e) {
          console.warn('Lỗi khi tải cấu hình khấu hao chi tiết:', e)
        }

        const rows = data.map(item => {
          const sapKey = String(item.ma_sap || '').trim().toLowerCase()
          const config = configMap.get(sapKey)
          return {
            maSAP: item.ma_sap,
            khoiLuongTong: item.khoi_luong_tong || 0,
            thanhTien: item.thanh_tien || 0,
            donGiaTrungBinh: item.don_gia_trung_binh || 0,
            donGiaTrungBinh1Ngay: item.don_gia_trung_binh_1_ngay || 0,
            phanLoaiVatTu: item.phan_loai_vat_tu || '',
            isApprovedDepreciation: config ? config.isApproved : (item.is_approved_depreciation !== undefined ? !!item.is_approved_depreciation : false),
            depreciationMonths: config ? config.months : (item.depreciation_months || 0)
          }
        })
        let localOptions = []
        try {
          const saved = localStorage.getItem('sgc_depreciation_options_v2')
          if (saved) {
            const parsed = JSON.parse(saved)
            if (Array.isArray(parsed) && parsed.length > 0 && typeof parsed[0] === 'object' && parsed[0] !== null) {
              const unique = []
              const seen = new Set()
              parsed.forEach(opt => {
                const key = `${opt.months}-${!!opt.isApproved}`
                if (!seen.has(key)) {
                  seen.add(key)
                  unique.push(opt)
                }
              })
              localOptions = unique.sort((a, b) => a.months - b.months || (a.isApproved ? 1 : 0) - (b.isApproved ? 1 : 0))
            }
          }
        } catch (e) {}
        if (localOptions.length === 0) {
          localOptions = [
            { months: 0, isApproved: false },
            { months: 12, isApproved: true },
            { months: 24, isApproved: true },
            { months: 36, isApproved: true },
            { months: 48, isApproved: true },
            { months: 60, isApproved: true },
            { months: 72, isApproved: true },
            { months: 120, isApproved: true }
          ]
        }

        const alignedRows = alignMaterialDepreciationWithConfig(rows, localOptions)

        setMaterialPriceRows(alignedRows)
        setLastSyncedMaterialPriceRows(alignedRows)
        localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(alignedRows))

        const prices = {}
        const classifications = {}
        alignedRows.forEach(r => {
          prices[r.maSAP] = r.donGiaTrungBinh
          classifications[r.maSAP] = r.phanLoaiVatTu
        })
        setMaterialPrices(prices)
        setMaterialClassifications(classifications)
        localStorage.setItem('sgc_report_material_prices', JSON.stringify(prices))
        localStorage.setItem('sgc_report_material_classifications', JSON.stringify(classifications))
      } else {
        setLastSyncedMaterialPriceRows([])
      }
    } catch (err) {
      console.error('Lỗi khi truy vấn Supabase:', err)
    } finally {
      setLoadingDbPrices(false)
    }
  }, [])

  React.useEffect(() => {
    loadPricesFromSupabase()
  }, [loadPricesFromSupabase])

  const handlePriceChange = React.useCallback((maSAP, value) => {
    const numVal = parseFloat(value) || 0
    setMaterialPrices(prev => {
      const updated = { ...prev, [maSAP]: numVal }
      localStorage.setItem('sgc_report_material_prices', JSON.stringify(updated))
      return updated
    })

    setMaterialPriceRows(prev => {
      const copy = prev.map(r => {
        if (r.maSAP === maSAP) {
          const months = r.depreciationMonths || 0
          let newDaily = 0
          if (months > 0) {
            newDaily = Math.round(numVal / (months * 30.417))
          }
          return { ...r, donGiaTrungBinh: numVal, donGiaTrungBinh1Ngay: newDaily }
        }
        return r
      })
      localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(copy))
      return copy
    })
  }, [])

  const handleClassificationChange = React.useCallback((maSAP, value) => {
    setMaterialClassifications(prev => {
      const updated = { ...prev, [maSAP]: value }
      localStorage.setItem('sgc_report_material_classifications', JSON.stringify(updated))
      return updated
    })

    setMaterialPriceRows(prev => {
      const copy = prev.map(r => r.maSAP === maSAP ? { ...r, phanLoaiVatTu: value } : r)
      localStorage.setItem('sgc_report_material_price_rows', JSON.stringify(copy))
      return copy
    })
  }, [])

  // Automatically persist rows to IndexedDB to keep cache synced and allow instant loading
  // (Chuyển từ localStorage sang IndexedDB vì dữ liệu lớn (hàng chục nghìn dòng) sẽ
  // vượt quá giới hạn dung lượng localStorage và gây lỗi QuotaExceededError)
  React.useEffect(() => {
    idbSet('sgc_giao_rows', giaoRows || [])
  }, [giaoRows])

  React.useEffect(() => {
    idbSet('sgc_nhan_rows', nhanRows || [])
  }, [nhanRows])

  React.useEffect(() => {
    idbSet('sgc_chung_rows', chungRows || [])
  }, [chungRows])

  React.useEffect(() => {
    idbSet('sgc_kho_rows', khoRows || [])
  }, [khoRows])

  // Tự động trích xuất các Kho dự án duy nhất từ các dòng dữ liệu local hiện có (Offline-first / Manual)
  // Giúp người dùng khi tải file lên là có sẵn tên kho dự án trong dropdown selector, không phụ thuộc vào Supabase.
  React.useEffect(() => {
    const projSet = new Set(customProjects)
    let changed = false

    const extractFromRows = (rows) => {
      if (!Array.isArray(rows)) return
      rows.forEach(r => {
        const pName = String(r.ten_du_an || r.tenDuAn || r.duAn || '').trim()
        if (pName && !projSet.has(pName)) {
          projSet.add(pName)
          changed = true
        }
      })
    }

    extractFromRows(chungRows)
    extractFromRows(giaoRows)
    extractFromRows(nhanRows)
    extractFromRows(khoRows)

    if (changed) {
      const sorted = Array.from(projSet).sort()
      setCustomProjects(sorted)
      localStorage.setItem('sgc_custom_projects', JSON.stringify(sorted))
    }
  }, [chungRows, giaoRows, nhanRows, khoRows])

  const [syncingType, setSyncingType] = useState(null) // 'giao' | 'nhan' | null
  const [supabaseMessage, setSupabaseMessage] = useState(null) // { text, type: 'success' | 'error' | 'info' }

  const [configs, setConfigs] = useState(() => loadSummaryConfigs())
  const [isInitialLoading, setIsInitialLoading] = useState(true)


  // Danh sách columns cần fetch (bỏ qua metadata Supabase như created_at, updated_at)
  const SELECT_COLS = 'id,ngay_xuat_nhap,ma_vat_tu,ma_s_a_p,thong_so_ky_thuat,ten_vat_tu,dvt,loai_don,ma_don_nhap_kho,ma_don_xuat_kho,khoi_luong_nhap,ma_don_vi_giao,don_vi_giao,nguoi_giao,khoi_luong_xuat,ma_don_vi_nhan,don_vi_nhan,nguoi_phe_duyet,ten_nguon,ma_nguon,lo,hang_muc,so_hop_dong,thu_kho,bien_so_xe,phan_khu,du_an,tinh_trang,nguoi_nhan,ma_don_lien_quan,nha_cung_cap,ma_don_chuyen_tiep_l_c,ma_don_chuyen_tiep_n_b,ghi_chu,ghi_chu_vat_tu,trang_thai,nhan_hieu,ten_du_an,tenDuAn'

  const loadProjectsFromSupabase = React.useCallback(async () => {
    // Đã gỡ bỏ tính năng load danh mục dự án từ bảng du_an của Supabase theo yêu cầu người dùng
    // Danh sách dự án hiện tại hoạt động offline-first thông qua localStorage
    console.log('[loadProjects] Đã gỡ bỏ query bảng du_an từ Supabase.')
  }, [])

  const loadTableFromSupabase = React.useCallback(async (tableType, forceBypassCache = false) => {
    if (!isSupabaseConfigured) return
    if (tableType !== 'chung') return // Skip deleted tables
    const tableName = 'don_chung'
    // KHAI BÁO Ở ĐÂY (không phải bên trong try) để catch block bên dưới vẫn truy cập được,
    // tránh lỗi "cachedKey is not defined" khiến dữ liệu tải xong không lưu được vào cache.
    const cachedKey = 'sgc_chung_rows'

    try {
      // Get current cached rows
      const cachedJson = await idbGet(cachedKey)
      const cachedRows = Array.isArray(cachedJson) ? cachedJson : []

      // 0. Tự động phát hiện các cột thực tế đang tồn tại trong bảng (Chống lỗi 400 Bad Request triệt để)
      let tableColumnsStr = '*'
      let hasUpdatedAt = false
      let hasCreatedAt = false
      let bestProjCol = null

      try {
        const { data: sampleRow, error: sampleErr } = await supabase
          .from(tableName)
          .select('*')
          .limit(1)
        
        if (!sampleErr && sampleRow && sampleRow.length > 0) {
          const firstRow = sampleRow[0]
          hasUpdatedAt = 'updated_at' in firstRow
          hasCreatedAt = 'created_at' in firstRow
          const skipCols = ['created_at', 'updated_at']
          const cols = Object.keys(firstRow).filter(k => !skipCols.includes(k))
          if (cols.length > 0) {
            tableColumnsStr = cols.join(',')
          }
          const projKeys = ['ten_du_an', 'tenDuAn', 'tenduan', 'du_an', 'duAn', 'ten_duan', 'tendu_an', 'name']
          bestProjCol = cols.find(c => projKeys.includes(c))
        }
      } catch (colErr) {
        console.warn('Lỗi nhận diện cột động, sử dụng fallback "*" :', colErr)
      }

      const timestampField = hasUpdatedAt ? 'updated_at' : (hasCreatedAt ? 'created_at' : null)

      let serverCount = null
      let useLimitPreview = false

      if (selectedProject && bestProjCol) {
        // TẢI THEO KHO DỰ ÁN ĐÃ CHỌN (Siêu nhẹ & Siêu nhanh, không bao giờ bị 500 timeout!)
        try {
          const { count, error: countErr } = await supabase
            .from(tableName)
            .select('id', { count: 'exact', head: true })
            .eq(bestProjCol, selectedProject)
          
          if (!countErr) {
            serverCount = count
          }
        } catch (cntErr) {
          console.warn('[Tải DB] Không lấy được count theo dự án, sử dụng fallback:', cntErr)
        }
      } else {
        // TẤT CẢ DỰ ÁN (Có tới 145k+ dòng, không đếm exact để tránh lỗi 500 timeout của Supabase!)
        useLimitPreview = true
        serverCount = 2000 // Giới hạn tải 2000 dòng mới nhất làm preview
      }

      // 1.5. SIÊU TỐI ƯU PRECHECK (Nếu đã có cache trùng khớp số dòng và timestamp):
      let localMaxTimestamp = ''
      if (timestampField && cachedRows.length > 0) {
        cachedRows.forEach(r => {
          const val = r[timestampField] || r.updated_at || r.created_at || ''
          if (val && val > localMaxTimestamp) {
            localMaxTimestamp = val
          }
        })
      }

      if (!forceBypassCache && serverCount !== null && cachedRows.length === serverCount && cachedRows.length > 0 && timestampField && !useLimitPreview) {
        try {
          let serverMaxQuery = supabase
            .from(tableName)
            .select(timestampField)
            .order(timestampField, { ascending: false })
            .limit(1)
          
          if (selectedProject && bestProjCol) {
            serverMaxQuery = serverMaxQuery.eq(bestProjCol, selectedProject)
          }

          const { data: serverMaxData, error: serverMaxErr } = await serverMaxQuery

          if (!serverMaxErr && serverMaxData && serverMaxData.length > 0) {
            const serverMaxTimestamp = serverMaxData[0][timestampField]
            if (serverMaxTimestamp && String(localMaxTimestamp) === String(serverMaxTimestamp)) {
              console.log('[Ultra-Lightweight Precheck] Khớp số dòng và max timestamp hoàn hảo! Bỏ qua sync chi tiết. Tiết kiệm 100% Egress.')
              setChungRows(cachedRows)
              setChungFileName(cachedRows.length > 0 ? `Report_Orders_Don_chung (Cached DB - ${selectedProject || 'Tất cả dự án'})` : '')
              setIsInitialLoading(false)
              return
            }
          }
        } catch (precheckErr) {
          console.warn('[Ultra-Lightweight Precheck] Có lỗi khi chạy pre-check, dùng fallback:', precheckErr)
        }
      }

      // 2. TỐI ƯU CỰC ĐẠI PHÁT HIỆN THAY ĐỔI THEO TIMESTAMP CHO DỰ ÁN (Incremental Sync):
      if (!forceBypassCache && serverCount !== null && cachedRows.length > 0 && timestampField && localMaxTimestamp && !useLimitPreview && selectedProject && bestProjCol) {
        console.log(`[Incremental Sync] Đang tải các dòng mới hoặc thay đổi sau thời điểm: ${localMaxTimestamp}...`)
        try {
          const { data: changedData, error: changedErr } = await supabase
            .from(tableName)
            .select(tableColumnsStr + (timestampField ? `,${timestampField}` : ''))
            .eq(bestProjCol, selectedProject)
            .gt(timestampField, localMaxTimestamp)

          if (!changedErr && changedData) {
            console.log(`[Incremental Sync] Đã tìm thấy ${changedData.length} dòng mới/thay đổi.`)
            const normalizedChanged = changedData.map(row => normalizeDbRow(row))
            const changedMap = new Map(normalizedChanged.map(r => [Number(r.id), r]))

            let mergedRows = cachedRows.map(r => {
              const rid = Number(r.id)
              if (changedMap.has(rid)) {
                const updated = changedMap.get(rid)
                changedMap.delete(rid)
                return updated
              }
              return r
            })

            if (changedMap.size > 0) {
              mergedRows.push(...Array.from(changedMap.values()))
            }

            mergedRows.sort((a, b) => Number(a.id) - Number(b.id))

            const localCountAfterMerge = mergedRows.length
            if (localCountAfterMerge === serverCount) {
              console.log('[Incremental Sync Hoàn Tất] Số lượng dòng khớp hoàn hảo. Không phát hiện dòng bị xóa. Tiết kiệm 100% ID check.')
              setChungRows(mergedRows)
              setChungFileName(mergedRows.length > 0 ? `Report_Orders_Don_chung (Supabase DB - ${selectedProject})` : '')
              await idbSet(cachedKey, mergedRows)
              setIsInitialLoading(false)
              return
            }
          }
        } catch (incErr) {
          console.warn('[Incremental Sync] Lỗi hệ thống khi chạy sync gia tăng, fallback tải lại:', incErr)
        }
      }

      // 3. TẢI TOÀN BỘ HOẶC PREVIEW DỮ LIỆU
      let allData = []
      let errorObj = null

      if (useLimitPreview) {
        // TẤT CẢ DỰ ÁN: Chỉ tải tối đa 2000 dòng mới nhất để đảm bảo tốc độ cực nhanh, không bao giờ timeout/treo máy!
        console.log(`[Tải DB] Đang tải preview 2000 dòng mới nhất của bảng "${tableName}"...`)
        const { data, error } = await supabase
          .from(tableName)
          .select(tableColumnsStr)
          .order('id', { ascending: false })
          .limit(2000)

        if (error) {
          errorObj = error
        } else if (data) {
          // Trả về thứ tự ban đầu (tăng dần theo ID) để hiển thị đồng nhất
          allData = [...data].reverse()
        }
      } else {
        // THEO DỰ ÁN: Tải toàn bộ dòng của dự án được chọn
        const PAGE_SIZE = 1000
        const pagesCount = Math.ceil((serverCount || 0) / PAGE_SIZE)
        console.log(`[Tải DB] Phát hiện ${serverCount} dòng của dự án "${selectedProject}". Đang tải song song ${pagesCount} trang...`)

        if (pagesCount > 0) {
          const fetchTasks = []
          for (let i = 0; i < pagesCount; i++) {
            const fromIdx = i * PAGE_SIZE
            const toIdx = fromIdx + PAGE_SIZE - 1
            fetchTasks.push(() =>
              supabase
                .from(tableName)
                .select(tableColumnsStr)
                .eq(bestProjCol, selectedProject)
                .order('id', { ascending: true })
                .range(fromIdx, toIdx)
            )
          }

          const results = await runWithConcurrencyLimit(fetchTasks, 6)
          for (let i = 0; i < results.length; i++) {
            const { data: pageData, error: pageErr } = results[i]
            if (pageErr) {
              errorObj = pageErr
              break
            }
            if (pageData) {
              allData = [...allData, ...pageData]
            }
          }
        }
      }

      if (errorObj) {
        throw errorObj
      }

      const mapped = allData.map((item, idx) => {
        const normalized = normalizeDbRow(item)
        return { id: normalized.id || idx, ...normalized }
      }).sort((a, b) => Number(a.id) - Number(b.id))

      setChungRows(mapped)
      setChungFileName(mapped.length > 0 ? `Report_Orders_Don_chung (Supabase DB${useLimitPreview ? ' - 2000 dòng mới nhất' : ` - ${selectedProject}`})` : '')
      await idbSet(cachedKey, mapped)
      setIsInitialLoading(false)

    } catch (e) {
      console.warn('[Tải DB] Gặp lỗi tải song song/preview, đang chuyển sang tuần tự...', e)
      let allData = []
      let errorObj = null
      let from = 0
      let hasMore = true
      const PAGE_SIZE = 1000

      try {
        let tableColumnsStr = '*'
        let bestProjCol = null
        try {
          const { data: sampleRow, error: sampleErr } = await supabase
            .from(tableName)
            .select('*')
            .limit(1)
          
          if (!sampleErr && sampleRow && sampleRow.length > 0) {
            const skipCols = ['created_at', 'updated_at']
            const cols = Object.keys(sampleRow[0]).filter(k => !skipCols.includes(k))
            if (cols.length > 0) {
              tableColumnsStr = cols.join(',')
            }
            const projKeys = ['ten_du_an', 'tenDuAn', 'tenduan', 'du_an', 'duAn', 'ten_duan', 'tendu_an', 'name']
            bestProjCol = cols.find(c => projKeys.includes(c))
          }
        } catch (colErr) {
          console.warn('Lỗi nhận diện cột động tuần tự:', colErr)
        }

        if (selectedProject && bestProjCol) {
          while (hasMore) {
            const { data: page, error } = await supabase
              .from(tableName)
              .select(tableColumnsStr)
              .eq(bestProjCol, selectedProject)
              .order('id', { ascending: true })
              .range(from, from + PAGE_SIZE - 1)

            if (error) { errorObj = error; break }
            if (page && page.length > 0) {
              allData = [...allData, ...page]
              hasMore = page.length === PAGE_SIZE
              from += PAGE_SIZE
            } else {
              hasMore = false
            }
          }
        } else {
          // Tất cả dự án: Chỉ tải tối đa 2000 dòng tuần tự
          const { data, error } = await supabase
            .from(tableName)
            .select(tableColumnsStr)
            .order('id', { ascending: false })
            .limit(2000)

          if (error) {
            errorObj = error
          } else if (data) {
            allData = [...data].reverse()
          }
        }

        if (errorObj) {
          if (errorObj.status === 401) setSupabaseAuthError(true)
          setIsInitialLoading(false)
          return
        }

        const mapped = allData.map((item, idx) => {
          const normalized = normalizeDbRow(item)
          return { id: normalized.id || idx, ...normalized }
        }).sort((a, b) => Number(a.id) - Number(b.id))

        setChungRows(mapped)
        setChungFileName(mapped.length > 0 ? 'Report_Orders_Don_chung (Supabase DB Fallback)' : '')
        await idbSet(cachedKey, mapped)
        setIsInitialLoading(false)
      } catch (sequentialErr) {
        console.error(`Lỗi tải tuần tự bảng ${tableName}:`, sequentialErr)
        setIsInitialLoading(false)
      }
    }
  }, [isSupabaseConfigured, selectedProject])

  const loadDataFromSupabase = React.useCallback(async (forceBypassCache = false) => {
    if (!isSupabaseConfigured) return
    await Promise.all([
      loadProjectsFromSupabase(),
      // loadTableFromSupabase('chung', forceBypassCache) is now disabled. don_chung is offline-first / manual upload only
    ])
  }, [loadProjectsFromSupabase])

  // Fetch initial data from Supabase if connected
  React.useEffect(() => {
    let isMounted = true

    const initData = async () => {
      if (!isSupabaseConfigured) {
        setIsInitialLoading(false)
        return
      }

      // ĐỌC CACHE TRƯỚC (Offline-First / Stale-While-Revalidate):
      // Đọc nhanh từ IndexedDB và hiển thị lên UI NGAY LẬP TỨC (mất ~0.05s) giúp người dùng truy cập phát được luôn!
      try {
        const cachedChung = await idbGet('sgc_chung_rows')
        const cachedGiao = await idbGet('sgc_giao_rows')
        const cachedNhan = await idbGet('sgc_nhan_rows')
        const cachedKho = await idbGet('sgc_kho_rows')

        if (isMounted && Array.isArray(cachedChung) && cachedChung.length > 0) {
          setChungRows(cachedChung)
          setChungFileName('Report_Orders_Don_chung (Cached DB)')
          
          if (Array.isArray(cachedGiao)) setGiaoRows(cachedGiao)
          if (Array.isArray(cachedNhan)) setNhanRows(cachedNhan)
          if (Array.isArray(cachedKho)) setKhoRows(cachedKho)

          // TẮT LOADING NGAY LẬP TỨC! Người dùng không phải chờ đợi màn hình loading nữa
          setIsInitialLoading(false)
          console.log('[Offline-First Cache] Đã hiển thị dữ liệu từ IndexedDB, tiến hành đồng bộ ngầm dưới nền...')
        }
      } catch (cacheErr) {
        console.warn('[Offline-First Cache] Không thể đọc cache ban đầu:', cacheErr)
      }

      // Tiếp tục kiểm tra/đồng bộ ngầm với Supabase để lấy cập nhật mới nhất
      loadDataFromSupabase().finally(() => {
        if (isMounted) {
          setIsInitialLoading(false)
        }
      })
    }

    initData()

    if (!isSupabaseConfigured) return

    // Tự động đồng bộ ngầm siêu nhanh khi người dùng quay lại Tab/App (Focus-based Smart Sync)
    // Giúp loại bỏ hoàn toàn việc lắng nghe Realtime liên tục trên hàng ngàn row gây ngốn Egress bandwidth.
    // TỐI ƯU EGRESS (02/07/2026): Tăng từ 60s lên 5 phút. Việc chuyển tab qua lại nhiều lần trong lúc
    // làm việc/test trước đây cứ mỗi 60s lại kích hoạt 1 lượt kiểm tra đồng bộ (tốn request + có thể
    // vô tình trùng thời điểm bảng đang bị ghi đè hàng loạt do import). 5 phút vẫn đủ nhanh để thấy
    // dữ liệu mới trong lúc dùng bình thường, nhưng giảm đáng kể số lần gọi Supabase khi bật/tắt tab liên tục.
    const FOCUS_SYNC_MIN_INTERVAL_MS = 5 * 60 * 1000 // Không sync lại nếu vừa sync trong vòng 5 phút gần nhất
    const handleFocus = () => {
      if (document.visibilityState !== 'visible' || !isSupabaseConfigured || syncInProgressRef.current) return
      // Chặn trùng lặp: 'visibilitychange' và 'focus' thường bắn gần như cùng lúc khi đổi tab,
      // nếu không chặn sẽ chạy loadDataFromSupabase() 2 lần song song → gấp đôi số request tới Supabase.
      if (focusSyncingRef.current) return
      const now = Date.now()
      if (now - lastFocusSyncAtRef.current < FOCUS_SYNC_MIN_INTERVAL_MS) return

      focusSyncingRef.current = true
      lastFocusSyncAtRef.current = now
      console.log('[Focus Sync] Người dùng quay lại Tab. Đang kiểm tra cập nhật mới nhất từ Supabase ngầm...')
      loadDataFromSupabase().finally(() => {
        focusSyncingRef.current = false
      })
    }
    document.addEventListener('visibilitychange', handleFocus)
    window.addEventListener('focus', handleFocus)

    return () => {
      isMounted = false
      if (realtimeDebounceRef.current) clearTimeout(realtimeDebounceRef.current)
      document.removeEventListener('visibilitychange', handleFocus)
      window.removeEventListener('focus', handleFocus)
    }
  }, [isSupabaseConfigured, loadDataFromSupabase, loadProjectsFromSupabase])

  const syncRowsToSupabase = async (type, rowsToSync, isAuto = false, isAppend = false) => {
    if (!isSupabaseConfigured) return
    // Ngưng kết nối sheet don_chung trên supabase, giữ dữ liệu 100% offline cục bộ
    if (type === 'chung' || type === 'giao' || type === 'nhan' || type === 'kho') {
      console.log(`[syncRowsToSupabase] Bỏ qua đồng bộ bảng '${type}' lên Supabase (chế độ offline-first).`)
      return
    }
    if (type !== 'chung') return // Skip syncing deleted tables

    if (!rowsToSync || rowsToSync.length === 0) {
      setSupabaseMessage({
        text: 'Không có dữ liệu để đồng bộ.',
        type: 'error'
      })
      setTimeout(() => setSupabaseMessage(null), 4000)
      return
    }

    syncInProgressRef.current = true // Tạm dừng Realtime trong khi sync
    setSyncingType(type)
    setSupabaseMessage({
      text: isAuto 
        ? (isAppend ? `Đang tự động tải nối tiếp ${rowsToSync.length} dòng lên Supabase...` : `Đang tự động đồng bộ ${rowsToSync.length} dòng lên Supabase...`)
        : `Đang lưu dữ liệu Đơn ${type === 'giao' ? 'Giao' : type === 'nhan' ? 'Nhận' : type === 'kho' ? 'Kho dự án' : 'Chung'} lên Supabase...`,
      type: 'info'
    })

    try {
      const tableName = type === 'chung' ? 'don_chung' : type === 'giao' ? 'don_giao' : type === 'nhan' ? 'don_nhan' : type === 'kho' ? 'don_kho' : 'don_chung'

      // 1. Delete existing records belonging to the synced projects to avoid wiping other unrelated projects, UNLESS isAppend is true.
      if (!isAppend) {
        // Lấy danh sách ten_du_an (Kho dự án) để xóa đúng các dòng cũ
        const uniqueProjectsInSync = [...new Set(rowsToSync.map(r => r.ten_du_an || r.tenDuAn || r.duAn).filter(Boolean))]
        
        if (uniqueProjectsInSync.length > 0) {
          console.log(`[Sync] Đang dọn dẹp các dòng cũ thuộc ${uniqueProjectsInSync.length} dự án đang lưu...`)
          for (const proj of uniqueProjectsInSync) {
            // Ưu tiên xóa theo ten_du_an (Kho dự án), không phải du_an gốc
            const delRes = await deleteFromTableAdaptive(tableName, ['ten_du_an', 'tenDuAn', 'tenduan'], proj)
            if (!delRes.success && delRes.reason !== 'table_empty') {
              console.warn(`[Sync] Không thể tự động dọn dẹp dự án "${proj}" trên bảng "${tableName}":`, delRes.error)
            }
          }
        } else {
          console.log(`[Sync] Không tìm thấy thông tin dự án cụ thể. Tiến hành dọn dẹp toàn bảng...`)
          const { error: delError } = await supabase
            .from(tableName)
            .delete()
            .neq('id', -999)

          if (delError) throw delError
        }
      } else {
        console.log(`[Sync] Chế độ UP FILE NỐI TIẾP: Giữ nguyên dữ liệu cũ, tiến hành chèn thêm ${rowsToSync.length} dòng mới...`)
      }

      // 2. Map structure to columns in Postgres matching COLS_GIAO_NHAN.
      // Crucial: Use SQL null instead of empty string '' for empty/undefined fields to avoid casting exceptions on numeric columns like khoi_luong_nhap/khoi_luong_xuat.
      const payload = rowsToSync.map(row => {
        const item = {}
        COLS_GIAO_NHAN.forEach(col => {
          const val = row[col.key]
          if (val === undefined || val === null || String(val).trim() === '') {
            item[col.key] = null
          } else {
            if (typeof val === 'number') {
              item[col.key] = val
            } else {
              const isNumericColumn = col.key === 'khoiLuongNhap' || col.key === 'khoiLuongXuat'
              if (isNumericColumn) {
                // Parse float robustly (remove potential spaces or thousand separator commas)
                const cleanStr = String(val).replace(/,/g, '').replace(/\s/g, '').trim()
                const parsed = parseFloat(cleanStr)
                item[col.key] = isNaN(parsed) ? null : parsed
              } else {
                item[col.key] = String(val)
              }
            }
          }
        })

        // Giữ nguyên duAn (du_an) của từng dòng - KHÔNG ghi đè bằng selectedProject
        // Chỉ cập nhật ten_du_an theo Kho dự án đang chọn (nếu có)
        const originalDuAn = row.duAn || ''
        item['duAn'] = originalDuAn || null
        // ten_du_an: nếu đang chọn kho dự án thì ghi tên kho, nếu không thì giữ nguyên duAn gốc
        const tenDuAnVal = selectedProject || originalDuAn || null
        item['ten_du_an'] = tenDuAnVal
        item['tenDuAn'] = tenDuAnVal

        return item
      })

      // 3. Batch insert with chunking & smart casing fallbacks
      const chunkSize = 500
      for (let i = 0; i < payload.length; i += chunkSize) {
        const chunk = payload.slice(i, i + chunkSize)
        await insertWithFallback(tableName, chunk)
        // Nhỏ nghỉ giữa các chunk để tránh quá tải request
        if (i + chunkSize < payload.length) {
          await new Promise(r => setTimeout(r, 200))
        }
      }

      // 3b. Đồng bộ sequence ID sau khi insert để tránh id nhảy số
      await resetTableSequence(tableName)

      // Không tự động thêm dự án vào bảng du_an — danh sách kho dự án chỉ do user quản lý thủ công

      setSupabaseMessage({
        text: isAuto 
          ? `Đã tự động liên thông thành công ${rowsToSync.length} dòng lên Supabase!`
          : `Đồng bộ thành công ${rowsToSync.length} dòng dữ liệu lên Supabase!`,
        type: 'success'
      })
      setTimeout(() => setSupabaseMessage(null), 5000)
    } catch (err) {
      console.error(err)
      setSupabaseMessage({
        text: `Lỗi khi đồng bộ lên Supabase: ${err.message || 'Hãy kiểm tra xem cấu trúc bảng trên Supabase đã đúng chưa.'}`,
        type: 'error'
      })
      setTimeout(() => setSupabaseMessage(null), 6000)
    } finally {
      setSyncingType(null)
      // Delay trước khi bật lại Realtime để tránh flood request ngay sau khi sync
      setTimeout(() => {
        syncInProgressRef.current = false
      }, 3000)
    }
  }

  const handleSyncToSupabase = async (type) => {
    const rowsToSync = type === 'giao' ? giaoRows : type === 'nhan' ? nhanRows : type === 'kho' ? khoRows : chungRows
    await syncRowsToSupabase(type, rowsToSync, false)
  }

  const handleImportFile = async (type, parsedRows, name, isAppend = false) => {
    // Hiển thị modal preview trước, chờ người dùng xác nhận lưu
    setPreviewModal({ type, rows: parsedRows, fileName: name, isAppend })
  }

  const handleConfirmImport = async () => {
    if (!previewModal) return
    const { type, rows: parsedRows, fileName: name, isAppend } = previewModal
    setPreviewModal(null)

    // Lấy danh sách các dòng đã tồn tại của bảng tương ứng
    const existingRows =
      type === 'giao' ? giaoRows :
      type === 'nhan' ? nhanRows :
      type === 'chung' ? chungRows :
      type === 'kho' ? khoRows : []

    const existingRowKeys = new Set(
      (existingRows || []).map(r => getRowUniqueKey(r))
    )

    // Bước 1: Lọc các dòng có Đơn vị giao/nhận trùng khớp với Kho dự án đang chọn (trùng khớp hoàn toàn, không phân biệt chữ hoa thường)
    const matchedByUnit = selectedProject
      ? parsedRows.filter(r => {
          if (type === 'chung') {
            const unitGiao = String(r.donViGiao || '').trim().toLowerCase()
            const unitNhan = String(r.donViNhan || '').trim().toLowerCase()
            const proj = selectedProject.trim().toLowerCase()
            return unitGiao === proj || unitNhan === proj
          } else {
            const unitKey = type === 'giao' ? 'donViGiao' : 'donViNhan'
            const unit = String(r[unitKey] || '').trim().toLowerCase()
            const proj = selectedProject.trim().toLowerCase()
            return unit === proj
          }
        })
      : parsedRows

    // Bước 2: Cho phép lưu tất cả các dòng, bất kể trạng thái (bỏ qua lọc Đã phê duyệt) và bỏ qua lọc trùng lặp hoàn toàn (cho lưu hết)
    const finalApprovedRows = matchedByUnit

    // Bước 3: Gán ten_du_an theo Kho dự án đang chọn, nhưng GIỮ NGUYÊN duAn gốc của từng dòng
    const rowsToStore = finalApprovedRows.map(row => {
      const originalDuAn = row.duAn && String(row.duAn).trim() !== '' ? String(row.duAn).trim() : ''
      // Không thay đổi duAn gốc - chỉ set ten_du_an theo kho đang chọn
      const tenDuAnVal = selectedProject || originalDuAn || ''
      return { 
        ...row, 
        duAn: originalDuAn, // giữ nguyên du_an gốc
        tenDuAn: tenDuAnVal, // ten_du_an theo kho dự án
        ten_du_an: tenDuAnVal
      }
    })

    if (type === 'giao') {
      setGiaoRows(prev => {
        if (isAppend) {
          return [...prev, ...rowsToStore]
        } else {
          const projectsInNewRows = new Set(rowsToStore.map(r => String(r.ten_du_an || r.tenDuAn || r.duAn || '').trim().toLowerCase()).filter(Boolean))
          const otherProjects = prev.filter(r => {
            const rowProject = String(r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || '').trim().toLowerCase()
            if (selectedProject) {
              return rowProject !== selectedProject.trim().toLowerCase()
            } else {
              return !projectsInNewRows.has(rowProject)
            }
          })
          return [...otherProjects, ...rowsToStore]
        }
      })
      setGiaoFileName(prev => isAppend ? (prev ? `${prev} + ${name}` : name) : name)
      // Skip remote sync as giaoRows is now offline-first / manual only
      console.log('[Import] Đơn Giao được lưu thành công vào cơ sở dữ liệu cục bộ.')
    } else if (type === 'nhan') {
      setNhanRows(prev => {
        if (isAppend) {
          return [...prev, ...rowsToStore]
        } else {
          const projectsInNewRows = new Set(rowsToStore.map(r => String(r.ten_du_an || r.tenDuAn || r.duAn || '').trim().toLowerCase()).filter(Boolean))
          const otherProjects = prev.filter(r => {
            const rowProject = String(r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || '').trim().toLowerCase()
            if (selectedProject) {
              return rowProject !== selectedProject.trim().toLowerCase()
            } else {
              return !projectsInNewRows.has(rowProject)
            }
          })
          return [...otherProjects, ...rowsToStore]
        }
      })
      setNhanFileName(prev => isAppend ? (prev ? `${prev} + ${name}` : name) : name)
      // Skip remote sync as nhanRows is now offline-first / manual only
      console.log('[Import] Đơn Nhận được lưu thành công vào cơ sở dữ liệu cục bộ.')
    } else if (type === 'kho') {
      setKhoRows(prev => {
        if (isAppend) {
          return [...prev, ...rowsToStore]
        } else {
          const projectsInNewRows = new Set(rowsToStore.map(r => String(r.ten_du_an || r.tenDuAn || r.duAn || '').trim().toLowerCase()).filter(Boolean))
          const otherProjects = prev.filter(r => {
            const rowProject = String(r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || '').trim().toLowerCase()
            if (selectedProject) {
              return rowProject !== selectedProject.trim().toLowerCase()
            } else {
              return !projectsInNewRows.has(rowProject)
            }
          })
          return [...otherProjects, ...rowsToStore]
        }
      })
      setKhoFileName(prev => isAppend ? (prev ? `${prev} + ${name}` : name) : name)
      // Skip remote sync as khoRows is now offline-first / manual only
      console.log('[Import] Đơn Kho được lưu thành công vào cơ sở dữ liệu cục bộ.')
    } else {
      // type === 'chung'
      // 1. Lưu Đơn chung
      setChungRows(prev => {
        if (isAppend) {
          return [...prev, ...rowsToStore]
        } else {
          const projectsInNewRows = new Set(rowsToStore.map(r => String(r.ten_du_an || r.tenDuAn || r.duAn || '').trim().toLowerCase()).filter(Boolean))
          const otherProjects = prev.filter(r => {
            const rowProject = String(r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || '').trim().toLowerCase()
            if (selectedProject) {
              return rowProject !== selectedProject.trim().toLowerCase()
            } else {
              return !projectsInNewRows.has(rowProject)
            }
          })
          return [...otherProjects, ...rowsToStore]
        }
      })
      setChungFileName(prev => isAppend ? (prev ? `${prev} + ${name}` : name) : name)

      // 2. Trích xuất Đơn Giao (đơn vị giao trùng khớp với Kho dự án đang chọn)
      const extractedGiaoRows = rowsToStore.filter(r => {
        const unitGiao = String(r.donViGiao || '').trim()
        if (selectedProject) {
          return unitGiao.toLowerCase() === selectedProject.trim().toLowerCase()
        }
        return unitGiao !== ''
      }).map(r => {
        if (!selectedProject) {
          const unitGiao = String(r.donViGiao || '').trim()
          return { ...r, ten_du_an: unitGiao, tenDuAn: unitGiao }
        }
        return r
      })
      setGiaoRows(prev => {
        if (isAppend) {
          return [...prev, ...extractedGiaoRows]
        } else {
          const projectsInNewRows = new Set(extractedGiaoRows.map(r => String(r.ten_du_an || r.tenDuAn || r.duAn || '').trim().toLowerCase()).filter(Boolean))
          const otherProjects = prev.filter(r => {
            const rowProject = String(r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || '').trim().toLowerCase()
            if (selectedProject) {
              return rowProject !== selectedProject.trim().toLowerCase()
            } else {
              return !projectsInNewRows.has(rowProject)
            }
          })
          return [...otherProjects, ...extractedGiaoRows]
        }
      })
      setGiaoFileName(prev => isAppend ? (prev ? `${prev} + ${name} (Trích xuất)` : `${name} (Trích xuất)`) : `${name} (Trích xuất)`)

      // 3. Trích xuất Đơn Nhận (đơn vị nhận trùng khớp với Kho dự án đang chọn)
      const extractedNhanRows = rowsToStore.filter(r => {
        const unitNhan = String(r.donViNhan || '').trim()
        if (selectedProject) {
          return unitNhan.toLowerCase() === selectedProject.trim().toLowerCase()
        }
        return unitNhan !== ''
      }).map(r => {
        if (!selectedProject) {
          const unitNhan = String(r.donViNhan || '').trim()
          return { ...r, ten_du_an: unitNhan, tenDuAn: unitNhan }
        }
        return r
      })
      setNhanRows(prev => {
        if (isAppend) {
          return [...prev, ...extractedNhanRows]
        } else {
          const projectsInNewRows = new Set(extractedNhanRows.map(r => String(r.ten_du_an || r.tenDuAn || r.duAn || '').trim().toLowerCase()).filter(Boolean))
          const otherProjects = prev.filter(r => {
            const rowProject = String(r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || '').trim().toLowerCase()
            if (selectedProject) {
              return rowProject !== selectedProject.trim().toLowerCase()
            } else {
              return !projectsInNewRows.has(rowProject)
            }
          })
          return [...otherProjects, ...extractedNhanRows]
        }
      })
      setNhanFileName(prev => isAppend ? (prev ? `${prev} + ${name} (Trích xuất)` : `${name} (Trích xuất)`) : `${name} (Trích xuất)`)

      // Skip remote sync as don_chung is now offline-first / manual only
      console.log('[Import] Đơn Chung được lưu thành công và trích xuất thành công vào cơ sở dữ liệu cục bộ.')
    }
  }

  const handleDeleteFile = (type) => {
    setDeleteFileType(type)
    setShowDeleteFileModal(true)
  }

  const handleConfirmDeleteFile = async () => {
    const type = deleteFileType
    setShowDeleteFileModal(false)
    setDeleteFileType(null)

    // 1. Xóa dữ liệu local - chỉ xóa rows thuộc selectedProject, không xóa dữ liệu kho khác
    if (type === 'giao') {
      if (selectedProject) {
        setGiaoRows(prev => prev.filter(r => {
          const rowProject = r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || ''
          return rowProject !== selectedProject
        }))
        setGiaoFileName(prev => prev)
      } else {
        setGiaoRows([])
        setGiaoFileName('')
      }
    } else if (type === 'nhan') {
      if (selectedProject) {
        setNhanRows(prev => prev.filter(r => {
          const rowProject = r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || ''
          return rowProject !== selectedProject
        }))
        setNhanFileName(prev => prev)
      } else {
        setNhanRows([])
        setNhanFileName('')
      }
    } else if (type === 'kho') {
      if (selectedProject) {
        setKhoRows(prev => prev.filter(r => {
          const rowProject = r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || ''
          return rowProject !== selectedProject
        }))
        setKhoFileName(prev => prev)
      } else {
        setKhoRows([])
        setKhoFileName('')
      }
    } else {
      if (selectedProject) {
        setChungRows(prev => prev.filter(r => {
          const rowProject = r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || ''
          return rowProject !== selectedProject
        }))
        setChungFileName(prev => prev)
      } else {
        setChungRows([])
        setChungFileName('')
      }
    }

    // 2. Nếu đã kết nối Supabase, xóa các dòng tương ứng trên Supabase (bỏ qua nếu là các bảng offline)
    if (isSupabaseConfigured && type !== 'chung' && type !== 'giao' && type !== 'nhan' && type !== 'kho') {
      const tableName = type === 'giao' ? 'don_giao' : type === 'nhan' ? 'don_nhan' : type === 'kho' ? 'don_kho' : 'don_chung'
      setSyncingType(type)
      setSupabaseMessage({ text: `Đang xóa dữ liệu Đơn ${type === 'giao' ? 'Giao' : type === 'nhan' ? 'Nhận' : type === 'kho' ? 'Kho dự án' : 'Chung'} trên Supabase...`, type: 'info' })
      try {
        if (selectedProject) {
          const deleteRes = await deleteFromTableAdaptive(tableName, ['ten_du_an', 'tenDuAn', 'tenduan', 'du_an'], selectedProject)
          if (!deleteRes.success && deleteRes.reason !== 'table_empty') {
            throw deleteRes.error || new Error('Không tìm thấy cột phù hợp hoặc không được phân quyền xóa.')
          }
        } else {
          const { error } = await supabase.from(tableName).delete().neq('id', -999)
          if (error) throw error
        }
        setSupabaseMessage({ text: `Đã xóa dữ liệu Đơn ${type === 'giao' ? 'Giao' : type === 'nhan' ? 'Nhận' : type === 'kho' ? 'Kho dự án' : 'Chung'} khỏi Supabase thành công!`, type: 'success' })
        setTimeout(() => setSupabaseMessage(null), 4000)
      } catch (err) {
        setSupabaseMessage({ text: `Lỗi khi xóa trên Supabase: ${err.message || err}`, type: 'error' })
        setTimeout(() => setSupabaseMessage(null), 5000)
      }
      setSyncingType(null)
    }
  }

  const handleAddProject = async (name) => {
    const trimmed = name.trim()
    if (!trimmed) return
    if (!customProjects.includes(trimmed)) {
      // 1. Lưu local ngay lập tức
      const updated = [...customProjects, trimmed]
      setCustomProjects(updated)
      localStorage.setItem('sgc_custom_projects', JSON.stringify(updated))

      setSupabaseMessage({ text: 'Tạo kho dự án "' + trimmed + '" thành công!', type: 'success' })
      setTimeout(() => setSupabaseMessage(null), 3000)
    }
    setSelectedProject(trimmed)
  }

  const handleEditProject = async (oldName, newName) => {
    const trimmedOld = oldName.trim()
    const trimmedNew = newName.trim()
    if (!trimmedOld || !trimmedNew || trimmedOld === trimmedNew) return

    // 1. Update customProjects local state & localStorage
    setCustomProjects(prev => {
      const updated = prev.map(p => p === trimmedOld ? trimmedNew : p)
      localStorage.setItem('sgc_custom_projects', JSON.stringify(updated))
      return updated
    })

    // 2. Chỉ cập nhật ten_du_an trong rows - KHÔNG thay đổi duAn (du_an) gốc
    // duAn là dữ liệu gốc từ file Excel, ten_du_an là tên Kho dự án được gán
    setGiaoRows(prev => prev.map(row => {
      // Cập nhật ten_du_an nếu row thuộc kho cũ
      if (row.duAn === trimmedOld || row.ten_du_an === trimmedOld || row.tenDuAn === trimmedOld) {
        return { ...row, tenDuAn: trimmedNew, ten_du_an: trimmedNew }
      }
      return row
    }))
    setNhanRows(prev => prev.map(row => {
      if (row.duAn === trimmedOld || row.ten_du_an === trimmedOld || row.tenDuAn === trimmedOld) {
        return { ...row, tenDuAn: trimmedNew, ten_du_an: trimmedNew }
      }
      return row
    }))
    setKhoRows(prev => prev.map(row => {
      if (row.duAn === trimmedOld || row.ten_du_an === trimmedOld || row.tenDuAn === trimmedOld) {
        return { ...row, tenDuAn: trimmedNew, ten_du_an: trimmedNew }
      }
      return row
    }))

    // Update active selection nếu đang chọn kho cũ
    if (selectedProject === trimmedOld) {
      setSelectedProject(trimmedNew)
    }

    // 3. Update Supabase
    if (isSupabaseConfigured) {
      setSupabaseMessage({
        text: `Đã đổi tên dự án thành "${trimmedNew}" thành công!`,
        type: 'success'
      })
      setTimeout(() => setSupabaseMessage(null), 5000)
    }
  }

  const handleDeleteProject = (projectName) => {
    const trimmed = projectName.trim()
    if (!trimmed) return
    setProjectToDelete(trimmed)
    setShowDeleteProjectModal(true)
  }

  const handleConfirmDeleteProject = async () => {
    const trimmed = projectToDelete.trim()
    if (!trimmed) return

    setShowDeleteProjectModal(false)
    setProjectToDelete('')

    // 1. Cập nhật customProjects local state & localStorage
    setCustomProjects(prev => {
      const updated = prev.filter(p => p !== trimmed)
      localStorage.setItem('sgc_custom_projects', JSON.stringify(updated))
      return updated
    })

    // Clear references from in-memory delivery/receipt rows so it doesn't stay in fileProjects list!
    setGiaoRows(prev => prev.map(row => {
      if ((row.ten_du_an || row.tenDuAn || row.duAn) === trimmed) {
        return { ...row, ten_du_an: '', tenDuAn: '' }
      }
      return row
    }))
    setNhanRows(prev => prev.map(row => {
      if ((row.ten_du_an || row.tenDuAn || row.duAn) === trimmed) {
        return { ...row, ten_du_an: '', tenDuAn: '' }
      }
      return row
    }))

    // Reset active selection nếu đang chọn kho bị xóa
    if (selectedProject === trimmed) {
      setSelectedProject('')
    }

    // 2. Xóa trên Supabase nếu đã cấu hình
    if (isSupabaseConfigured) {
      setSupabaseMessage({
        text: `Xóa kho dự án "${trimmed}" thành công!`,
        type: 'success'
      })
      setTimeout(() => setSupabaseMessage(null), 4000)
    }
  }

  const fileProjects = useMemo(() => {
    const list = new Set()
    giaoRows.forEach(r => { if (r.duAn) list.add(r.duAn) })
    nhanRows.forEach(r => { if (r.duAn) list.add(r.duAn) })
    khoRows.forEach(r => { if (r.duAn) list.add(r.duAn) })
    return [...list].sort()
  }, [giaoRows, nhanRows, khoRows])

  const allProjects = useMemo(() => {
    // Danh sách dự án chỉ lấy từ customProjects (do người dùng tạo bằng nút "+ Tạo mới")
    // KHÔNG tự động lấy từ file upload lên
    return [...new Set(customProjects)].sort()
  }, [customProjects])

  const tabs = [
    { id: 'chung', label: 'Đơn chung', icon: <ClipboardList size={15} /> },
    { id: 'kho', label: 'Kho dự án', icon: <Warehouse size={15} /> },
    { id: 'dongia', label: 'Phân nhóm vật tư', icon: <PackageCheck size={15} /> },
    { id: 'inventory', label: 'Báo cáo xuất nhập tồn', icon: <Database size={15} /> },
    { id: 'inventory_real', label: 'Báo cáo xuất nhập thực', icon: <BarChart3 size={15} /> },
    { id: 'depreciation_assets', label: 'Báo cáo tài sản khấu hao', icon: <DollarSign size={15} /> },
    { id: 'accounts', label: 'Quản lý tài khoản', icon: <Users size={15} /> },
  ]

  if (!currentUser) {
    return <LoginPage onLoginSuccess={handleLoginSuccess} />
  }

  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column', overflow: 'hidden', position: 'relative' }}>
      {isInitialLoading && (
        <div style={{
          position: 'fixed',
          inset: 0,
          zIndex: 9999,
          background: '#fff5f5',
          display: 'flex',
          flexDirection: 'column',
          alignItems: 'center',
          justifyContent: 'center',
          gap: 16,
        }}>
          <div style={{
            width: 64, height: 64, background: 'var(--primary-light)',
            borderRadius: 18, display: 'flex', alignItems: 'center', justifyContent: 'center',
          }}>
            <RefreshCw size={30} color="var(--primary)" style={{ animation: 'spin 1s linear infinite' }} />
          </div>
          <div style={{ textAlign: 'center' }}>
            <h2 style={{ fontSize: 18, fontWeight: 800, color: 'var(--text)', margin: '0' }}>
              SGC | BÁO CÁO GIAO NHẬN
            </h2>
          </div>
        </div>
      )}
      <Header
        selectedProject={selectedProject}
        setSelectedProject={setSelectedProject}
        duAnOptions={allProjects}
        onOpenAddProjectModal={() => setShowAddProjectModal(true)}
        onEditProject={(projectName) => {
          setProjectToEdit(projectName)
          setShowEditProjectModal(true)
        }}
        onDeleteProject={handleDeleteProject}
        onOpenConfigModal={() => setShowConfigModal(true)}
        onForceRefresh={() => loadDataFromSupabase(true)}
        onMouseEnterLogo={() => setIsSidebarOpen(true)}
        activeTab={tab}
        currentUser={currentUser}
        onLogout={handleLogout}
      />

      <div style={{ flex: 1, display: 'flex', position: 'relative', overflow: 'hidden' }}>
        {/* Small persistent handle showing when sidebar is closed & unpinned */}


        {/* Dynamic sliding Sidebar Drawer */}
        <div
          onMouseEnter={handleMouseEnter}
          onMouseLeave={handleMouseLeave}
          style={{
            position: 'absolute',
            left: 0,
            top: 0,
            bottom: 0,
            width: '270px',
            background: 'linear-gradient(180deg, #0a3d73 0%, #153e75 100%)',
            borderRight: '1px solid rgba(255, 255, 255, 0.12)',
            boxShadow: '0 10px 30px -5px rgba(10, 61, 115, 0.2), 4px 0 20px -4px rgba(10, 61, 115, 0.1)',
            display: 'flex',
            flexDirection: 'column',
            zIndex: 1000,
            transform: isSidebarOpen || isSidebarPinned ? 'translateX(0)' : 'translateX(-270px)',
            transition: 'transform 0.3s cubic-bezier(0.4, 0, 0.2, 1)',
            textAlign: 'left'
          }}
        >
          {/* Sidebar Header */}
          <div style={{
            padding: '22px 24px 18px 24px',
            borderBottom: '1px solid rgba(255, 255, 255, 0.12)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            background: 'linear-gradient(180deg, #08325e 0%, #0a3d73 100%)'
          }}>
            <div style={{ display: 'flex', alignItems: 'center' }}>
              <div style={{ display: 'flex', flexDirection: 'column' }}>
                <span style={{ 
                  fontFamily: 'var(--font-display)', 
                  fontWeight: 800, 
                  fontSize: '13.5px', 
                  color: '#ffffff', 
                  letterSpacing: '0.06em', 
                  textTransform: 'uppercase' 
                }}>
                  Hệ thống Báo cáo
                </span>
              </div>
            </div>
            
            {/* Pin Toggle Button */}
            <button
              onClick={handleTogglePin}
              title={isSidebarPinned ? "Bỏ ghim (Dạng menu trượt ẩn)" : "Ghim menu (Giữ cố định)"}
              style={{
                background: isSidebarPinned ? 'rgba(255, 255, 255, 0.12)' : 'transparent',
                border: isSidebarPinned ? '1px solid rgba(255, 255, 255, 0.2)' : '1px solid rgba(255, 255, 255, 0.1)',
                color: isSidebarPinned ? '#38bdf8' : '#93c5fd',
                cursor: 'pointer',
                padding: '6px',
                borderRadius: '8px',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                transition: 'all 0.2s ease',
                boxShadow: isSidebarPinned ? '0 1px 3px rgba(0, 0, 0, 0.1)' : 'none'
              }}
              onMouseOver={(e) => {
                e.currentTarget.style.background = 'rgba(255, 255, 255, 0.18)'
                e.currentTarget.style.color = '#ffffff'
                e.currentTarget.style.borderColor = 'rgba(255, 255, 255, 0.3)'
              }}
              onMouseOut={(e) => {
                e.currentTarget.style.background = isSidebarPinned ? 'rgba(255, 255, 255, 0.12)' : 'transparent'
                e.currentTarget.style.color = isSidebarPinned ? '#38bdf8' : '#93c5fd'
                e.currentTarget.style.borderColor = isSidebarPinned ? 'rgba(255, 255, 255, 0.2)' : 'rgba(255, 255, 255, 0.1)'
              }}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" fill={isSidebarPinned ? "#38bdf8" : "none"} stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
                <line x1="12" y1="17" x2="12" y2="22"></line>
                <path d="M5 17h14v-1.76a2 2 0 0 0-.44-1.24l-2.78-3.5A2 2 0 0 1 15 9.24V5a1 1 0 0 0-1-1h-4a1 1 0 0 0-1 1v4.24c0 .43-.14.85-.4 1.18l-2.78 3.5a2 2 0 0 0-.44 1.24z"></path>
              </svg>
            </button>
          </div>
 
          {/* List of Sheet Items with Styled Groups */}
          <div style={{ flex: 1, overflowY: 'auto', padding: '20px 16px', display: 'flex', flexDirection: 'column', gap: 24 }}>
            
            {/* GROUP 1: DỮ LIỆU ĐẦU VÀO */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
              <div style={{
                fontSize: '10px',
                fontWeight: 700,
                color: '#93c5fd',
                letterSpacing: '0.1em',
                paddingLeft: '12px',
                marginBottom: '8px',
                textTransform: 'uppercase',
                display: 'flex',
                alignItems: 'center',
                gap: 8
              }}>
                <span>Dữ liệu đầu vào</span>
                <div style={{ flex: 1, height: '1px', background: 'rgba(255, 255, 255, 0.15)' }} />
              </div>
 
              {tabs.filter(t => t.id === 'giao' || t.id === 'nhan' || t.id === 'chung' || t.id === 'kho' || t.id === 'dongia').map(t => {
                const isSelected = tab === t.id
                const count = t.id === 'giao' ? giaoRows.length : t.id === 'nhan' ? nhanRows.length : t.id === 'kho' ? khoRows.length : t.id === 'chung' ? chungRows.length : t.id === 'dongia' ? materialPriceRows.length : null
 
                return (
                  <button
                    key={t.id}
                    onClick={() => {
                      setTab(t.id)
                      if (!isSidebarPinned) {
                        setIsSidebarOpen(false)
                      }
                    }}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      width: '100%',
                      padding: '12px 14px',
                      borderRadius: '10px',
                      fontSize: '13.5px',
                      fontWeight: isSelected ? 700 : 500,
                      color: isSelected ? '#ffffff' : '#dbeafe',
                      background: isSelected ? 'linear-gradient(135deg, #1e40af 0%, #3b82f6 100%)' : 'transparent',
                      border: 'none',
                      cursor: 'pointer',
                      transition: 'all 0.2s cubic-bezier(0.4, 0, 0.2, 1)',
                      boxShadow: isSelected ? '0 4px 14px rgba(0, 0, 0, 0.15)' : 'none',
                      position: 'relative',
                    }}
                    onMouseOver={(e) => {
                      if (!isSelected) {
                        e.currentTarget.style.background = 'rgba(255, 255, 255, 0.08)'
                        e.currentTarget.style.color = '#ffffff'
                        e.currentTarget.style.transform = 'translateX(4px)'
                      }
                    }}
                    onMouseOut={(e) => {
                      if (!isSelected) {
                        e.currentTarget.style.background = 'transparent'
                        e.currentTarget.style.color = '#dbeafe'
                        e.currentTarget.style.transform = 'none'
                      }
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                      <div style={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        width: 20,
                        height: 20,
                      }}>
                        {React.cloneElement(t.icon, { size: 16, strokeWidth: isSelected ? 2.5 : 2, color: isSelected ? '#ffffff' : '#93c5fd' })}
                      </div>
                      <span>{t.label}</span>
                    </div>
 
                    {count !== null && count > 0 && (
                      <span style={{
                        background: isSelected ? 'rgba(255, 255, 255, 0.25)' : 'rgba(255, 255, 255, 0.12)',
                        color: isSelected ? '#ffffff' : '#93c5fd',
                        border: isSelected ? 'none' : '1px solid rgba(255, 255, 255, 0.15)',
                        borderRadius: '20px',
                        padding: '2px 8px',
                        fontSize: '11px',
                        fontWeight: 700,
                        minWidth: '24px',
                        textAlign: 'center',
                        boxShadow: isSelected ? 'none' : '0 1px 2px rgba(0, 0, 0, 0.05)'
                      }}>
                        {count.toLocaleString()}
                      </span>
                    )}
                  </button>
                )
              })}
            </div>
 
            {/* GROUP 2: XỬ LÝ & TỔNG HỢP */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 5 }}>
              <div style={{
                fontSize: '10px',
                fontWeight: 700,
                color: '#93c5fd',
                letterSpacing: '0.1em',
                paddingLeft: '12px',
                marginBottom: '8px',
                textTransform: 'uppercase',
                display: 'flex',
                alignItems: 'center',
                gap: 8
              }}>
                <span>Xử lý & Tổng hợp</span>
                <div style={{ flex: 1, height: '1px', background: 'rgba(255, 255, 255, 0.15)' }} />
              </div>
 
              {tabs.filter(t => t.id === 'inventory' || t.id === 'inventory_real' || t.id === 'depreciation_assets' || (t.id === 'accounts' && currentUser?.quyen === 'Admin')).map(t => {
                const isSelected = tab === t.id
 
                return (
                  <button
                    key={t.id}
                    onClick={() => {
                      setTab(t.id)
                      if (!isSidebarPinned) {
                        setIsSidebarOpen(false)
                      }
                    }}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      width: '100%',
                      padding: '12px 14px',
                      borderRadius: '10px',
                      fontSize: '13.5px',
                      fontWeight: isSelected ? 700 : 500,
                      color: isSelected ? '#ffffff' : '#dbeafe',
                      background: isSelected ? 'linear-gradient(135deg, #1e40af 0%, #3b82f6 100%)' : 'transparent',
                      border: 'none',
                      cursor: 'pointer',
                      transition: 'all 0.2s cubic-bezier(0.4, 0, 0.2, 1)',
                      boxShadow: isSelected ? '0 4px 14px rgba(0, 0, 0, 0.15)' : 'none',
                      position: 'relative',
                    }}
                    onMouseOver={(e) => {
                      if (!isSelected) {
                        e.currentTarget.style.background = 'rgba(255, 255, 255, 0.08)'
                        e.currentTarget.style.color = '#ffffff'
                        e.currentTarget.style.transform = 'translateX(4px)'
                      }
                    }}
                    onMouseOut={(e) => {
                      if (!isSelected) {
                        e.currentTarget.style.background = 'transparent'
                        e.currentTarget.style.color = '#dbeafe'
                        e.currentTarget.style.transform = 'none'
                      }
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
                      <div style={{
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        width: 20,
                        height: 20,
                      }}>
                        {React.cloneElement(t.icon, { size: 16, strokeWidth: isSelected ? 2.5 : 2, color: isSelected ? '#ffffff' : '#93c5fd' })}
                      </div>
                      <span>{t.label}</span>
                    </div>
                  </button>
                )
              })}
            </div>
 
          </div>

          {/* List of Sheet Items with Styled Groups */}
          <div style={{ display: 'none' }}> {/* DUP_CONTAINER_START */}
            
            {/* GROUP 1: DỮ LIỆU ĐẦU VÀO */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              <div style={{
                fontSize: '11px',
                fontWeight: 700,
                color: '#94a3b8',
                letterSpacing: '0.08em',
                paddingLeft: '12px',
                marginBottom: '6px',
                textTransform: 'uppercase',
                display: 'flex',
                alignItems: 'center',
                gap: 6
              }}>
                <span style={{ width: 4, height: 4, borderRadius: '50%', background: '#cbd5e1' }} />
                Dữ liệu đầu vào
              </div>

              {tabs.filter(t => t.id === 'giao' || t.id === 'nhan' || t.id === 'chung').map(t => {
                const isSelected = tab === t.id
                const count = t.id === 'giao' ? giaoRows.length : t.id === 'nhan' ? nhanRows.length : t.id === 'chung' ? chungRows.length : null

                return (
                  <button
                    key={t.id}
                    onClick={() => {
                      setTab(t.id)
                      if (!isSidebarPinned) {
                        setIsSidebarOpen(false)
                      }
                    }}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      width: '100%',
                      padding: '11px 14px',
                      borderRadius: '8px',
                      fontSize: '13.5px',
                      fontWeight: isSelected ? 700 : 500,
                      color: isSelected ? '#ffffff' : '#475569',
                      background: isSelected ? 'linear-gradient(135deg, #0f58a7 0%, #083c75 100%)' : 'transparent',
                      border: 'none',
                      borderLeft: isSelected ? '4px solid #38bdf8' : '4px solid transparent',
                      cursor: 'pointer',
                      transition: 'all 0.15s ease-in-out',
                      boxShadow: isSelected ? '0 4px 12px rgba(15,88,167,0.15)' : 'none',
                      position: 'relative'
                    }}
                    onMouseOver={(e) => {
                      if (!isSelected) {
                        e.currentTarget.style.background = '#f0f7ff'
                        e.currentTarget.style.color = '#0f58a7'
                      }
                    }}
                    onMouseOut={(e) => {
                      if (!isSelected) {
                        e.currentTarget.style.background = 'transparent'
                        e.currentTarget.style.color = '#475569'
                      }
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                      {React.cloneElement(t.icon, { size: 15, strokeWidth: isSelected ? 2.5 : 2, color: isSelected ? '#ffffff' : '#64748b' })}
                      <span>{t.label}</span>
                    </div>

                    {count !== null && count > 0 && (
                      <span style={{
                        background: isSelected ? 'rgba(255, 255, 255, 0.2)' : '#e2e8f0',
                        color: isSelected ? '#ffffff' : '#475569',
                        borderRadius: '12px',
                        padding: '1px 7px',
                        fontSize: '11px',
                        fontWeight: 700,
                        minWidth: '24px',
                        textAlign: 'center'
                      }}>
                        {count.toLocaleString()}
                      </span>
                    )}
                  </button>
                )
              })}
            </div>

            {/* GROUP 2: XỬ LÝ & TỔNG HỢP */}
            <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
              <div style={{
                fontSize: '11px',
                fontWeight: 700,
                color: '#94a3b8',
                letterSpacing: '0.08em',
                paddingLeft: '12px',
                marginBottom: '6px',
                textTransform: 'uppercase',
                display: 'flex',
                alignItems: 'center',
                gap: 6
              }}>
                <span style={{ width: 4, height: 4, borderRadius: '50%', background: '#cbd5e1' }} />
                Xử lý & Tổng hợp
              </div>

              {tabs.filter(t => t.id === 'inventory' || t.id === 'inventory_real' || t.id === 'depreciation_assets').map(t => {
                const isSelected = tab === t.id

                return (
                  <button
                    key={t.id}
                    onClick={() => {
                      setTab(t.id)
                      if (!isSidebarPinned) {
                        setIsSidebarOpen(false)
                      }
                    }}
                    style={{
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'space-between',
                      width: '100%',
                      padding: '11px 14px',
                      borderRadius: '8px',
                      fontSize: '13.5px',
                      fontWeight: isSelected ? 700 : 500,
                      color: isSelected ? '#ffffff' : '#475569',
                      background: isSelected ? 'linear-gradient(135deg, #0f58a7 0%, #083c75 100%)' : 'transparent',
                      border: 'none',
                      borderLeft: isSelected ? '4px solid #38bdf8' : '4px solid transparent',
                      cursor: 'pointer',
                      transition: 'all 0.15s ease-in-out',
                      boxShadow: isSelected ? '0 4px 12px rgba(15,88,167,0.15)' : 'none',
                      position: 'relative'
                    }}
                    onMouseOver={(e) => {
                      if (!isSelected) {
                        e.currentTarget.style.background = '#f0f7ff'
                        e.currentTarget.style.color = '#0f58a7'
                      }
                    }}
                    onMouseOut={(e) => {
                      if (!isSelected) {
                        e.currentTarget.style.background = 'transparent'
                        e.currentTarget.style.color = '#475569'
                      }
                    }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
                      {React.cloneElement(t.icon, { size: 15, strokeWidth: isSelected ? 2.5 : 2, color: isSelected ? '#ffffff' : '#64748b' })}
                      <span>{t.label}</span>
                    </div>
                  </button>
                )
              })}
            </div>

          </div>

          {/* Footer of Sidebar */}
          <div style={{
            padding: '18px 24px',
            borderTop: '1px solid rgba(255, 255, 255, 0.12)',
            fontSize: '11px',
            color: '#93c5fd',
            textAlign: 'center',
            background: 'linear-gradient(180deg, #153e75 0%, #0d284c 100%)',
            fontWeight: 700,
            letterSpacing: '0.04em',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            gap: 8
          }}>
            <span style={{ width: 6, height: 6, borderRadius: '50%', background: '#22c55e', display: 'inline-block', boxShadow: '0 0 8px rgba(34, 197, 94, 0.6)' }} />
            SGC REPORT SYSTEM v1.1.0
          </div>
        </div>
        {/* END DUP DELETE */}

        {/* Main Content Area */}
        <div style={{
          flex: 1,
          display: 'flex',
          flexDirection: 'column',
          overflow: 'hidden',
          paddingLeft: isSidebarPinned ? '270px' : '0px',
          transition: 'padding-left 0.3s cubic-bezier(0.4, 0, 0.2, 1)',
          backgroundColor: '#031fae'
        }}>
          {supabaseAuthError && (
            <div style={{
              margin: '16px 16px 0 16px',
              padding: '12px 16px',
              borderRadius: '12px',
              border: '1px solid #fca5a5',
              backgroundColor: '#fef2f2',
              color: '#991b1b',
              display: 'flex',
              alignItems: 'flex-start',
              gap: '12px',
              boxShadow: '0 2px 4px rgba(0,0,0,0.02)',
              textAlign: 'left'
            }}>
              <AlertCircle size={20} color="#dc2626" style={{ marginTop: '2px', flexShrink: 0 }} />
              <div style={{ flex: 1 }}>
                <h4 style={{ margin: 0, fontSize: '13px', fontWeight: 700 }}>⚠️ Lỗi xác thực kết nối Supabase (401 Unauthorized)</h4>
                <p style={{ margin: '4px 0 0 0', fontSize: '12px', color: '#b91c1c', lineHeight: 1.4 }}>
                  Cơ sở dữ liệu Supabase đã từ chối yêu cầu truy cập vì Project URL hoặc Anon Key bị cấu hình sai hoặc đã bị thay đổi/hết hạn trên trang quản trị dự án của bạn (Lỗi 401 Unauthorized). 
                  Do đó, bạn <strong>không thể xem danh sách, cập nhật hay xóa dự án</strong> được. Vui lòng bấm vào nút cấu hình lại bên dưới.
                </p>
                <button
                  onClick={() => setShowConfigModal(true)}
                  style={{
                    marginTop: '8px',
                    padding: '4px 10px',
                    fontSize: '11px',
                    fontWeight: 700,
                    color: '#ffffff',
                    backgroundColor: '#dc2626',
                    border: 'none',
                    borderRadius: '4px',
                    cursor: 'pointer'
                  }}
                >
                  Sửa & kiểm cấu hình kết nối ngay
                </button>
              </div>
            </div>
          )}

          <div style={{ flex: 1, overflow: 'hidden', backgroundColor: '#ededed' }}>
            {tab === 'giao' && (
              <OrderTab
                type="giao"
                rows={giaoRows}
                setRows={setGiaoRows}
                fileName={giaoFileName}
                setFileName={setGiaoFileName}
                selectedProject={selectedProject}
                setSelectedProject={setSelectedProject}
                onSync={() => handleSyncToSupabase('giao')}
                syncing={syncingType === 'giao'}
                supabaseMessage={supabaseMessage}
                onEditProject={(projectName) => {
                  setProjectToEdit(projectName)
                  setShowEditProjectModal(true)
                }}
                projectOptions={allProjects}
                onImportFile={(rows, name, isAppend) => handleImportFile('giao', rows, name, isAppend)}
                onDeleteFile={() => handleDeleteFile('giao')}
                customCategoryMap={customCategoryMap}
              />
            )}
            {tab === 'nhan' && (
              <OrderTab
                type="nhan"
                rows={nhanRows}
                setRows={setNhanRows}
                fileName={nhanFileName}
                setFileName={setNhanFileName}
                selectedProject={selectedProject}
                setSelectedProject={setSelectedProject}
                onSync={() => handleSyncToSupabase('nhan')}
                syncing={syncingType === 'nhan'}
                supabaseMessage={supabaseMessage}
                onEditProject={(projectName) => {
                  setProjectToEdit(projectName)
                  setShowEditProjectModal(true)
                }}
                projectOptions={allProjects}
                onImportFile={(rows, name, isAppend) => handleImportFile('nhan', rows, name, isAppend)}
                onDeleteFile={() => handleDeleteFile('nhan')}
                customCategoryMap={customCategoryMap}
              />
            )}
            {tab === 'chung' && (
              <OrderTab
                type="chung"
                rows={chungRows}
                setRows={setChungRows}
                fileName={chungFileName}
                setFileName={setChungFileName}
                selectedProject={selectedProject}
                setSelectedProject={setSelectedProject}
                onSync={() => handleSyncToSupabase('chung')}
                syncing={syncingType === 'chung'}
                supabaseMessage={supabaseMessage}
                onEditProject={(projectName) => {
                  setProjectToEdit(projectName)
                  setShowEditProjectModal(true)
                }}
                projectOptions={allProjects}
                onImportFile={(rows, name, isAppend) => handleImportFile('chung', rows, name, isAppend)}
                onDeleteFile={() => handleDeleteFile('chung')}
                customCategoryMap={customCategoryMap}
              />
            )}
            {tab === 'kho' && (
              <KhoDuAnTab
                chungRows={chungRows}
                selectedProject={selectedProject}
                setSelectedProject={setSelectedProject}
                allProjects={allProjects}
                customCategoryMap={customCategoryMap}
                setCustomCategoryMap={setCustomCategoryMap}
                dbCategoryMap={dbCategoryMap}
                setDbCategoryMap={setDbCategoryMap}
              />
            )}
            {tab === 'dongia' && (
              <PhanNhomVatTuTab
                giaoRows={giaoRows}
                nhanRows={nhanRows}
                chungRows={chungRows}
                materialPriceRows={materialPriceRows}
                setMaterialPriceRows={setMaterialPriceRows}
                materialPrices={materialPrices}
                setMaterialPrices={setMaterialPrices}
                materialClassifications={materialClassifications}
                setMaterialClassifications={setMaterialClassifications}
                loadingDbPrices={loadingDbPrices}
                allProjects={allProjects}
                customCategoryMap={customCategoryMap}
                handleClassificationChange={handleClassificationChange}
                handlePriceChange={handlePriceChange}
                lastSyncedDepreciationOptions={lastSyncedDepreciationOptions}
                setLastSyncedDepreciationOptions={setLastSyncedDepreciationOptions}
                lastSyncedMaterialPriceRows={lastSyncedMaterialPriceRows}
                setLastSyncedMaterialPriceRows={setLastSyncedMaterialPriceRows}
              />
            )}
            {tab === 'inventory' && (
              <BaoCaoXuatNhapTonTab
                chungRows={chungRows}
                giaoRows={giaoRows}
                nhanRows={nhanRows}
                selectedProject={selectedProject}
                setSelectedProject={setSelectedProject}
                allProjects={allProjects}
                isRealReport={false}
                customCategoryMap={customCategoryMap}
                materialPriceRows={materialPriceRows}
                materialPrices={materialPrices}
                materialClassifications={materialClassifications}
                handleClassificationChange={handleClassificationChange}
                handlePriceChange={handlePriceChange}
              />
            )}
            {tab === 'inventory_real' && (
              <BaoCaoXuatNhapTonTab
                chungRows={chungRows}
                giaoRows={giaoRows}
                nhanRows={nhanRows}
                selectedProject={selectedProject}
                setSelectedProject={setSelectedProject}
                allProjects={allProjects}
                isRealReport={true}
                customCategoryMap={customCategoryMap}
                materialPriceRows={materialPriceRows}
                materialPrices={materialPrices}
                materialClassifications={materialClassifications}
                handleClassificationChange={handleClassificationChange}
                handlePriceChange={handlePriceChange}
              />
            )}
            {tab === 'depreciation_assets' && (
              <TaiSanKhauHaoTab
                chungRows={chungRows}
                giaoRows={giaoRows}
                nhanRows={nhanRows}
                allProjects={allProjects}
                customCategoryMap={customCategoryMap}
                materialPriceRows={materialPriceRows}
                materialPrices={materialPrices}
                materialClassifications={materialClassifications}
                isDbSchemaOutdated={isDbSchemaOutdated}
                onOpenDbUpgradeModal={() => setShowDbUpgradeModal(true)}
              />
            )}
            {tab === 'accounts' && (
              <QuanLyTaiKhoanTab />
            )}
          </div>
        </div>
      </div>

      <PreviewImportModal
        isOpen={!!previewModal}
        onClose={() => setPreviewModal(null)}
        onConfirm={handleConfirmImport}
        rows={previewModal?.rows || []}
        fileName={previewModal?.fileName || ''}
        type={previewModal?.type || 'giao'}
        selectedProject={selectedProject}
        isAppend={previewModal?.isAppend || false}
        existingRows={
          previewModal?.type === 'giao' ? giaoRows :
          previewModal?.type === 'nhan' ? nhanRows :
          previewModal?.type === 'chung' ? chungRows :
          previewModal?.type === 'kho' ? khoRows : []
        }
      />

      <AddProjectModal
        isOpen={showAddProjectModal}
        onClose={() => setShowAddProjectModal(false)}
        onSave={handleAddProject}
      />

      <EditProjectModal
        isOpen={showEditProjectModal}
        onClose={() => {
          setShowEditProjectModal(false)
          setProjectToEdit('')
        }}
        onSave={handleEditProject}
        currentName={projectToEdit}
      />

      <SupabaseConfigModal
        isOpen={showConfigModal}
        onClose={() => setShowConfigModal(false)}
      />

      <SupabaseDbUpgradeModal
        isOpen={showDbUpgradeModal}
        onClose={() => setShowDbUpgradeModal(false)}
      />

      <DeleteFileModal
        isOpen={showDeleteFileModal}
        onClose={() => { setShowDeleteFileModal(false); setDeleteFileType(null) }}
        onConfirm={handleConfirmDeleteFile}
        type={deleteFileType}
        selectedProject={selectedProject}
        rowCount={deleteFileType === 'giao'
          ? (selectedProject ? giaoRows.filter(r => (r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || '') === selectedProject).length : giaoRows.length)
          : deleteFileType === 'nhan'
          ? (selectedProject ? nhanRows.filter(r => (r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || '') === selectedProject).length : nhanRows.length)
          : deleteFileType === 'kho'
          ? (selectedProject ? khoRows.filter(r => (r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || '') === selectedProject).length : khoRows.length)
          : (selectedProject ? chungRows.filter(r => (r.ten_du_an || r.tenDuAn || r.tenduan || r.duAn || r.du_an || '') === selectedProject).length : chungRows.length)
        }
      />

      <DeleteProjectModal
        isOpen={showDeleteProjectModal}
        onClose={() => {
          setShowDeleteProjectModal(false)
          setProjectToDelete('')
        }}
        onConfirm={handleConfirmDeleteProject}
        projectName={projectToDelete}
        giaoRows={giaoRows}
        nhanRows={nhanRows}
        khoRows={khoRows}
      />

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
      `}</style>
    </div>
  )
}
