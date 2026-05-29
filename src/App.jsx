import React, { useState, useRef, useCallback, useMemo } from 'react'
import * as XLSX from 'xlsx'
import * as XLSXStyleRaw from 'xlsx-js-style'
const XLSXStyle = XLSXStyleRaw.default || XLSXStyleRaw
import {
  Upload, FileSpreadsheet, Search, X, RefreshCw, Info,
  ChevronDown, Download, Truck, PackageCheck, Settings, BarChart3,
  AlertCircle, CheckCircle2, Filter, ArrowUpDown, Clock, CloudUpload, Database, Save,
  Pencil, Trash2, Lock
} from 'lucide-react'
import { COLS_GIAO_NHAN, parseXlsxToRows, formatVal, getTrangThaiColor, isApprovedStatus, isPendingStatus, isRejectedStatus } from './constants.js'
import { supabase, isSupabaseConfigured, supabaseUrl, supabaseAnonKey } from './supabaseClient.js'

// ─── Searchable Select ────────────────────────────────────────────────────────
function SearchableSelect({ value, onChange, options, placeholder = 'Tất cả dự án', variant = 'header', onEditProject, onDeleteProject }) {
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
    maxWidth: 460,
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
    left: 0,
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
    <div ref={containerRef} style={{ position: 'relative', display: 'inline-block' }}>
      {/* Trigger */}
      <div onClick={() => setIsOpen(!isOpen)} style={triggerStyle}>
        <span style={{
          textOverflow: 'ellipsis',
          whiteSpace: 'nowrap',
          overflow: 'hidden',
          marginRight: 4,
          display: 'flex',
          alignItems: 'center',
          height: '100%',
          lineHeight: 'normal'
        }}>
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
              placeholder="Tìm kiếm dự án..."
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
                  {onEditProject && (
                    <button
                      onClick={(e) => {
                        e.stopPropagation()
                        onEditProject(opt)
                      }}
                      title="Sửa tên dự án"
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
                        transition: 'all 0.1s'
                      }}
                      onMouseOver={e => {
                        e.stopPropagation()
                        e.currentTarget.style.color = '#0f58a7'
                        e.currentTarget.style.background = '#e2e8f0'
                      }}
                      onMouseOut={e => {
                        e.stopPropagation()
                        e.currentTarget.style.color = '#64748b'
                        e.currentTarget.style.background = 'transparent'
                      }}
                    >
                      <Pencil size={11} />
                    </button>
                  )}
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
function Header({ selectedProject, setSelectedProject, duAnOptions, onOpenAddProjectModal, onEditProject, onDeleteProject, onOpenConfigModal }) {
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
      <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
        <div style={{
          width: 36, height: 36, background: '#ffffff',
          borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center',
          boxShadow: '0 3px 12px rgba(10,61,115,0.2)',
        }}>
          <span style={{ color: '#0f58a7', fontWeight: 900, fontSize: 15, letterSpacing: '0.02em' }}>SGC</span>
        </div>
        <div>
          <div style={{ color: '#ffffff', fontWeight: 800, fontSize: 18, letterSpacing: '0.03em', display: 'flex', alignItems: 'center', gap: 8 }}>
            SGC | BÁO CÁO GIAO NHẬN
          </div>
        </div>
      </div>

      {/* Kho Dự Án Widget */}
      <div style={{
        display: 'flex',
        alignItems: 'center',
        gap: 12,
        background: 'rgba(255, 255, 255, 0.12)',
        padding: '6px 14px',
        borderRadius: 8,
        border: '1px solid rgba(255, 255, 255, 0.20)'
      }}>
        <span style={{ color: '#ffffff', fontWeight: 700, fontSize: 14, textTransform: 'uppercase', letterSpacing: '0.04em', whiteSpace: 'nowrap' }}>
          Kho dự án:
        </span>
        <div style={{ display: 'flex', alignItems: 'center' }}>
          <SearchableSelect
            value={selectedProject}
            onChange={setSelectedProject}
            options={duAnOptions}
            placeholder="Tất cả dự án"
            variant="header"
            onEditProject={onEditProject}
            onDeleteProject={onDeleteProject}
          />
        </div>
        <button
          onClick={onOpenAddProjectModal}
          style={{
            background: 'linear-gradient(180deg, #10b981 0%, #059669 100%)',
            border: 'none',
            color: '#ffffff',
            fontSize: 14,
            fontWeight: 700,
            padding: '4px 12px',
            borderRadius: 6,
            cursor: 'pointer',
            display: 'flex',
            alignItems: 'center',
            gap: 4,
            height: 32,
            boxShadow: '0 2px 4px rgba(5,150,105,0.2)',
            transition: 'opacity 0.15s',
            whiteSpace: 'nowrap'
          }}
          onMouseOver={(e) => e.currentTarget.style.opacity = '0.9'}
          onMouseOut={(e) => e.currentTarget.style.opacity = '1'}
        >
          <span>+ Tạo mới</span>
        </button>
      </div>

      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <button
          onClick={onOpenConfigModal}
          title="Cấu hình kết nối cơ sở dữ liệu Supabase"
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
            cursor: 'pointer',
            transition: 'all 0.15s ease',
            height: 28,
            boxSizing: 'border-box'
          }}
          onMouseOver={(e) => {
            e.currentTarget.style.background = isSupabaseConfigured ? 'rgba(16, 185, 129, 0.3)' : 'rgba(245, 158, 11, 0.3)'
          }}
          onMouseOut={(e) => {
            e.currentTarget.style.background = isSupabaseConfigured ? 'rgba(16, 185, 129, 0.15)' : 'rgba(245, 158, 11, 0.15)'
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
        </button>

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

  const handle = useCallback((file) => {
    if (disabled || !file) return
    const reader = new FileReader()
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
      onFile(data, file.name)
    }
    reader.readAsArrayBuffer(file)
  }, [onFile, disabled])

  return (
    <div
      className={`upload-zone${drag && !disabled ? ' drag-over' : ''}${disabled ? ' disabled' : ''}`}
      onClick={() => { if (!disabled) ref.current.click() }}
      onDragOver={e => { e.preventDefault(); if (!disabled) setDrag(true) }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); if (!disabled) handle(e.dataTransfer.files[0]) }}
      style={{
        cursor: disabled ? 'not-allowed' : 'pointer',
        opacity: disabled ? 0.75 : 1,
        backgroundColor: disabled ? '#fafafa' : undefined,
        borderColor: disabled ? '#fca5a5' : undefined,
        borderStyle: disabled ? 'dashed' : 'dashed',
        position: 'relative'
      }}
    >
      <input ref={ref} type="file" accept={accept} disabled={disabled} onChange={e => { if (!disabled) handle(e.target.files[0]) }} style={{ display: 'none' }} />
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
              <>Kéo thả file vào đây hoặc <span style={{ color: 'var(--primary)', fontWeight: 600 }}>chọn file từ thiết bị</span></>
            )}
          </div>
          <div style={{ color: disabled ? '#b91c1c' : 'var(--text-light)', fontSize: 12, marginTop: 6, fontStyle: 'italic', fontWeight: disabled ? 500 : 'normal' }}>
            {disabled 
              ? 'Hãy chọn một dự án cụ thể bên danh sách "KHO DỰ ÁN" ở thanh Menu trên trước để tiếp tục tải lên.' 
              : 'Hỗ trợ định dạng chuẩn: .xlsx, .xls'}
          </div>
        </div>
      </div>
    </div>
  )
}

// ─── Stats Bar ────────────────────────────────────────────────────────────────
function StatsBar({ rows }) {
  const stats = useMemo(() => {
    const total = rows.length
    const daPheDuyet = rows.filter(r => isApprovedStatus(r.trangThai)).length;
    const choPheDuyet = rows.filter(r => isPendingStatus(r.trangThai)).length;
    const tuChoi = rows.filter(r => isRejectedStatus(r.trangThai)).length;

    return { total, daPheDuyet, choPheDuyet, tuChoi }
  }, [rows])

  const cards = [
    { label: 'Tổng số lượng đơn', value: stats.total.toLocaleString() + ' đơn', color: 'var(--primary)', bg: 'var(--primary-light)', border: 'var(--border)', icon: <FileSpreadsheet size={18} /> },
    { label: 'Đã phê duyệt', value: stats.daPheDuyet.toLocaleString() + ' đơn', color: '#10b981', bg: '#ecfdf5', border: '#a7f3d0', icon: <CheckCircle2 size={18} /> },
    { label: 'Chờ phê duyệt', value: stats.choPheDuyet.toLocaleString() + ' đơn', color: '#b45309', bg: '#fffbeb', border: '#fde68a', icon: <Clock size={18} /> },
    { label: 'Từ chối', value: stats.tuChoi.toLocaleString() + ' đơn', color: '#ef4444', bg: '#fef2f2', border: '#fca5a5', icon: <AlertCircle size={18} /> },
  ]

  return (
    <div style={{ display: 'flex', gap: 12, marginBottom: 16, flexWrap: 'wrap' }}>
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
    </div>
  )
}

// ─── Filter Bar ───────────────────────────────────────────────────────────────
function FilterBar({ search, setSearch, trangThai, setTrangThai, trangThaiOptions, duAn, setDuAn, duAnOptions, onClear, onEditProject }) {
  return (
    <div style={{ display: 'flex', gap: 8, marginBottom: 12, flexWrap: 'wrap', alignItems: 'center' }}>
      <div style={{ position: 'relative', flex: '1 1 200px', minWidth: 160 }}>
        <Search size={14} style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)', color: 'var(--text-light)' }} />
        <input
          className="input"
          style={{ paddingLeft: 32, width: '100%' }}
          placeholder="Tìm kiếm tên vật tư, mã đơn, đơn vị..."
          value={search}
          onChange={e => setSearch(e.target.value)}
        />
      </div>
      <select className="input" style={{ flex: '0 0 160px' }} value={trangThai} onChange={e => setTrangThai(e.target.value)}>
        <option value="">Tất cả trạng thái</option>
        {trangThaiOptions.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
      <div style={{ flex: '0 0 180px', minWidth: 180 }}>
        <SearchableSelect
          value={duAn}
          onChange={setDuAn}
          options={duAnOptions}
          placeholder="Tất cả dự án"
          variant="filter"
          onEditProject={onEditProject}
        />
      </div>
      {(search || trangThai || duAn) && (
        <button className="btn btn-outline btn-sm" onClick={onClear}>
          <X size={12} /> Xóa lọc
        </button>
      )}
    </div>
  )
}

// ─── Data Table ───────────────────────────────────────────────────────────────
const PAGE_SIZE_OPTIONS = [50, 100, 200, 500]

function DataTable({ rows }) {
  const [pageSize, setPageSize] = React.useState(100)
  const [currentPage, setCurrentPage] = React.useState(1)

  // Reset to page 1 when rows change (filter applied)
  React.useEffect(() => { setCurrentPage(1) }, [rows])

  if (rows.length === 0) return (
    <div className="empty-state">
      <Search size={48} />
      <h3>Không có dữ liệu</h3>
      <p>Không tìm thấy kết quả phù hợp với bộ lọc hiện tại.</p>
    </div>
  )

  const totalPages = Math.ceil(rows.length / pageSize)
  const startIdx = (currentPage - 1) * pageSize
  const endIdx = Math.min(startIdx + pageSize, rows.length)
  const pageRows = rows.slice(startIdx, endIdx)

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
    <div>
      {/* Pagination top bar */}
      <div style={{
        display: 'flex', alignItems: 'center', justifyContent: 'space-between',
        marginBottom: 10, gap: 12, flexWrap: 'wrap'
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

          {/* Prev */}
          <button style={currentPage === 1 ? btnDisabled : btnBase} onClick={() => setCurrentPage(p => p - 1)}>‹</button>

          {/* Page numbers */}
          {getPageNums().map((p, i) =>
            p === '...'
              ? <span key={`ellipsis-${i}`} style={{ fontSize: 13, color: '#94a3b8', padding: '0 2px' }}>…</span>
              : <button key={p} style={currentPage === p ? btnActive : btnBase} onClick={() => setCurrentPage(p)}>{p}</button>
          )}

          {/* Next */}
          <button style={currentPage === totalPages ? btnDisabled : btnBase} onClick={() => setCurrentPage(p => p + 1)}>›</button>
        </div>
      </div>

      {/* Table */}
      <div className="table-wrap">
        <table>
          <thead>
            <tr>
              <th style={{ width: 50, minWidth: 50, maxWidth: 50, textAlign: 'center', verticalAlign: 'middle', fontSize: '12px', padding: '8px 10px' }}>
                STT
              </th>
              {COLS_GIAO_NHAN.map(c => (
                <th
                  key={c.key}
                  style={{
                    width: c.width, minWidth: c.width, maxWidth: c.width,
                    textAlign: 'center', verticalAlign: 'middle',
                    fontSize: '12px', padding: '8px 10px',
                    whiteSpace: 'normal', wordBreak: 'break-word', lineHeight: '1.2'
                  }}
                >
                  {c.label}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {pageRows.map((row, i) => (
              <tr key={row.id}>
                <td style={{ width: 50, minWidth: 50, maxWidth: 50, textAlign: 'center', fontSize: '12px', color: '#1b1919', padding: '6px 10px' }}>
                  {startIdx + i + 1}
                </td>
                {COLS_GIAO_NHAN.map(col => {
                  const isCenteredCol = [
                    'ngayXuatNhap', 'maVatTu', 'maSAP', 'dvt', 'loaiDon',
                    'maDonViGiao', 'donViGiao', 'nguoiGiao',
                    'maDonViNhan', 'donViNhan', 'nguoiPheDuyet', 'nguoiNhan',
                    'soHopDong', 'thuKho', 'tinhTrang'
                  ].includes(col.key)
                  const isRightAligned = ['khoiLuongNhap', 'khoiLuongXuat'].includes(col.key) || col.key.toLowerCase().includes('khoiluong')
                  return (
                    <td
                      key={col.key}
                      style={{
                        width: col.width, minWidth: col.width, maxWidth: col.width,
                        color: '#1b1919',
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

      {/* Pagination bottom bar */}
      {totalPages > 1 && (
        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: 6, marginTop: 12 }}>
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
      )}
    </div>
  )
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
  onDeleteFile
}) {
  const isGiao = type === 'giao'
  const [loading, setLoading] = useState(false)
  const [search, setSearch] = useState('')
  const [trangThai, setTrangThai] = useState('')

  const label = isGiao ? 'Đơn Giao' : 'Đơn Nhận'
  const uploadLabel = isGiao
    ? 'Tải lên file Report_Orders_Đơn giao'
    : 'Tải lên file Report_Orders_Đơn nhận'

  const handleFile = useCallback((data, name) => {
    setLoading(true)
    setTimeout(() => {
      const parsed = parseXlsxToRows(data)
      if (onImportFile) {
        onImportFile(parsed, name)
      } else {
        setRows(parsed)
        setFileName(name)
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
    if (selectedProject) r = r.filter(row => row.duAn === selectedProject)
    return r
  }, [rows, search, trangThai, selectedProject])

  const projectFilteredRowsForStats = useMemo(() => {
    if (!selectedProject) return rows
    return rows.filter(r => r.duAn === selectedProject)
  }, [rows, selectedProject])

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
      ws[cellRef] = {
        v: col.label,
        t: 's',
        s: {
          fill: {
            patternType: 'solid',
            fgColor: { rgb: '0F58A7' }
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
            top: { style: 'thin', color: { rgb: '0A3D73' } },
            bottom: { style: 'medium', color: { rgb: '0A3D73' } },
            left: { style: 'thin', color: { rgb: '0A3D73' } },
            right: { style: 'thin', color: { rgb: '0A3D73' } }
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
          'STT', 'ngayXuatNhap', 'maVatTu', 'maSAP', 'dvt', 'loaiDon',
          'maDonViGiao', 'donViGiao', 'nguoiGiao',
          'maDonViNhan', 'donViNhan', 'nguoiPheDuyet', 'nguoiNhan',
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
    ws['!ref'] = `A1:${getColLabel(columns.length - 1)}${excelRowIdx}`

    XLSXStyle.utils.book_append_sheet(wb, ws, isGiao ? "Đơn Giao" : "Đơn Nhận")
    
    // Save
    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary' })
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
    a.download = `Export_${isGiao ? 'Don_giao' : 'Don_nhan'}_${new Date().toISOString().slice(0, 10)}.xlsx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }, [filtered, isGiao])

  return (
    <div style={{ padding: '16px 24px 24px 24px', height: '100%', display: 'flex', flexDirection: 'column', boxSizing: 'border-box', overflow: 'hidden' }}>
      {!rows.length && !loading ? (
        <div style={{ maxWidth: 560, margin: '40px auto' }}>
          <div style={{ marginBottom: 20, textAlign: 'center' }}>
            <div style={{
              width: 56, height: 56, background: 'var(--primary-light)',
              borderRadius: 16, display: 'flex', alignItems: 'center', justifyContent: 'center',
              margin: '0 auto 12px'
            }}>
              {isGiao ? <Truck size={26} color="var(--primary)" /> : <PackageCheck size={26} color="var(--primary)" />}
            </div>
            <h2 style={{ fontSize: 19, fontWeight: 800, color: 'var(--text)', marginBottom: 6 }}>
              Tab {label}
            </h2>
            <p style={{ color: 'var(--text-muted)', fontSize: 15 }}>
              Import file Excel để xem dữ liệu báo cáo theo dõi giao nhận.
            </p>
          </div>

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
              background: '#fff8e1',
              border: '1px solid #f59e0b',
              borderRadius: 10,
              padding: '10px 14px',
              marginBottom: 16,
              display: 'flex',
              gap: 8,
              alignItems: 'flex-start',
            }}>
              <Info size={14} color="#d97706" style={{ marginTop: 1, flexShrink: 0 }} />
              <div style={{ fontSize: 14, color: '#92400e' }}>
                <strong>Lưu ý:</strong> Vui lòng chọn <strong>Kho dự án</strong> ở góc trên trước khi upload file. Dữ liệu sẽ được gán vào dự án đang chọn. Chức năng tải file đang tạm thời bị khóa cho đến khi chọn dự án.
              </div>
            </div>
          )}
          <UploadZone onFile={handleFile} label={uploadLabel} disabled={!selectedProject} />
        </div>
      ) : loading ? (
        <div className="empty-state">
          <RefreshCw size={40} style={{ animation: 'spin 1s linear infinite', opacity: 0.5 }} />
          <h3>Đang xử lý dữ liệu...</h3>
        </div>
      ) : (
        <>
          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: 16, gap: 12 }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
              <div style={{
                background: 'var(--primary-light)', borderRadius: 8, padding: '6px 12px',
                display: 'flex', alignItems: 'center', gap: 6
              }}>
                <FileSpreadsheet size={14} color="var(--primary)" />
                <span style={{ fontSize: 14, fontWeight: 600, color: 'var(--primary)', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {fileName}
                </span>
              </div>
              <span style={{ fontSize: 14, color: 'var(--text-muted)' }}>
                {rows.length.toLocaleString()} dòng dữ liệu · {filtered.length.toLocaleString()} hiển thị
              </span>
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              {rows.length > 0 && (
                <button
                  className="btn btn-sm"
                  onClick={handleExportExcel}
                  style={{
                    background: 'linear-gradient(135deg, #059669 0%, #10b981 100%)',
                    color: '#ffffff',
                    border: 'none',
                    boxShadow: '0 2px 4px rgba(16,185,129,0.2)',
                  }}
                >
                  <Download size={12} /> Xuất Excel
                </button>
              )}
              <button className="btn btn-outline btn-sm" onClick={() => { if (onDeleteFile) { onDeleteFile() } else { setRows([]); setFileName(''); setSearch(''); setTrangThai('') } }} style={{ display: selectedProject ? 'flex' : 'none' }}>
                <X size={12} /> Xóa file
              </button>
            </div>
          </div>

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

          <StatsBar rows={projectFilteredRowsForStats} />

          <FilterBar
            search={search} setSearch={setSearch}
            trangThai={trangThai} setTrangThai={setTrangThai}
            trangThaiOptions={trangThaiOptions}
            duAn={selectedProject} setDuAn={setSelectedProject}
            duAnOptions={duAnOptions}
            onClear={() => { setSearch(''); setTrangThai(''); setSelectedProject('') }}
            onEditProject={onEditProject}
          />

          <DataTable rows={filtered} />
        </>
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

// ─── Preview Import Modal ─────────────────────────────────────────────────────
function PreviewImportModal({ isOpen, onClose, onConfirm, rows, fileName, type, selectedProject }) {
  // Nhấn Esc để đóng modal
  React.useEffect(() => {
    if (!isOpen) return
    const handleKeyDown = (e) => { if (e.key === 'Escape') onClose() }
    document.addEventListener('keydown', handleKeyDown)
    return () => document.removeEventListener('keydown', handleKeyDown)
  }, [isOpen, onClose])

  if (!isOpen) return null
  const label = type === 'giao' ? 'Đơn Giao' : 'Đơn Nhận'
  const previewRows = rows.slice(0, 10)

  const unitKey = type === 'giao' ? 'donViGiao' : 'donViNhan'
  const unitLabel = type === 'giao' ? 'Đơn vị giao' : 'Đơn vị nhận'

  // Lọc các dòng có Đơn vị giao/nhận trùng khớp với Kho dự án đang chọn (trùng khớp hoàn toàn, không phân biệt chữ hoa thường)
  const matchedByUnit = selectedProject
    ? rows.filter(r => {
        const unit = String(r[unitKey] || '').trim().toLowerCase()
        const proj = selectedProject.trim().toLowerCase()
        return unit === proj
      })
    : rows

  // Số dòng bị loại vì không khớp đơn vị
  const mismatchCount = rows.length - matchedByUnit.length

  // Lấy các đơn vị trong file KHÔNG khớp (để thông báo)
  const uniqueUnits = [...new Set(rows.map(r => String(r[unitKey] || '').trim()).filter(Boolean))]
  const mismatchUnits = uniqueUnits.filter(u => {
    const proj = (selectedProject || '').trim().toLowerCase()
    return u.toLowerCase() !== proj
  })

  // Trong số các dòng khớp đơn vị, chỉ lấy dòng Đã phê duyệt
  const approvedRows = matchedByUnit.filter(r => isApprovedStatus(r.trangThai))
  const skippedStatusCount = matchedByUnit.length - approvedRows.length

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
                {rows.length > 10 && <span style={{ color: '#f59e0b', marginLeft: 4 }}>(hiển thị 10 dòng đầu)</span>}
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

        {/* Warning notification banner for matching rules */}
        {mismatchCount > 0 && (
          <div style={{
            background: approvedRows.length === 0 ? '#fef2f2' : '#fffbeb',
            borderBottom: approvedRows.length === 0 ? '1px solid #fee2e2' : '1px solid #fef3c7',
            padding: '12px 24px',
            color: approvedRows.length === 0 ? '#991b1b' : '#92400e',
            fontSize: 13,
            display: 'flex',
            alignItems: 'flex-start',
            gap: 10,
            flexShrink: 0,
            textAlign: 'left'
          }}>
            <AlertCircle size={18} color={approvedRows.length === 0 ? '#dc2626' : '#d97706'} style={{ flexShrink: 0, marginTop: '2px' }} />
            <div style={{ flex: 1, lineHeight: 1.5 }}>
              {approvedRows.length === 0 ? (
                <span>
                  <strong style={{ fontSize: '14px', display: 'block', marginBottom: '4px', color: '#b91c1c' }}>⚠️ KIỂM TRA LỖI KHỚP DÀNH CHO KHO DỰ ÁN</strong>
                  Không có bất kỳ dòng dữ liệu nào trong tệp tải lên có cột <strong>{unitLabel}</strong> khớp hoàn toàn với Kho dự án đã chọn <strong>"{selectedProject}"</strong>. Hệ thống sẽ <strong>bỏ qua toàn bộ dữ liệu</strong> và không thể lưu tệp này.
                  {uniqueUnits.length > 0 && (
                    <div style={{ marginTop: 8, padding: '8px 12px', background: '#fff', borderRadius: '6px', border: '1px solid #fee2e2', fontWeight: 500, color: '#b91c1c' }}>
                      Các giá trị {unitLabel} được tìm thấy trong tệp hiện tại: <strong>{uniqueUnits.join(', ')}</strong>
                    </div>
                  )}
                </span>
              ) : (
                <span>
                  <strong style={{ fontSize: '14px', display: 'block', marginBottom: '4px' }}>⚠️ LƯU Ý BỎ QUA DỮ LIỆU KHÔNG TRÙNG KHỚP</strong>
                  Phát hiện <strong>{mismatchCount.toLocaleString()} dòng</strong> có cột <strong>{unitLabel}</strong> không trùng khớp hoàn toàn với tên Kho dự án đang chọn <strong>"{selectedProject}"</strong>. Hệ thống sẽ tự động <strong>bỏ qua {mismatchCount.toLocaleString()} dòng này</strong> và chỉ nhập/lưu {approvedRows.length.toLocaleString()} dòng trùng khớp có trạng thái đã phê duyệt.
                </span>
              )}
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
                <strong>{approvedRows.length.toLocaleString()} dòng</strong> khớp <strong>{unitLabel} "{selectedProject || 'tất cả'}"</strong> &amp; trạng thái <strong>Đã phê duyệt</strong> → sẽ được lưu
              </span>
            </div>

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
              style={{ minWidth: 160, opacity: approvedRows.length === 0 ? 0.5 : 1, cursor: approvedRows.length === 0 ? 'not-allowed' : 'pointer' }}
              disabled={approvedRows.length === 0}
            >
              <Save size={14} /> Lưu {approvedRows.length.toLocaleString()} dòng đã duyệt
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

// ─── Delete Project Modal ──────────────────────────────────────────────────
function DeleteProjectModal({ isOpen, onClose, onConfirm, projectName, giaoRows, nhanRows }) {
  if (!isOpen) return null

  // Đếm số dòng dữ liệu thuộc dự án này
  const giaoCount = (giaoRows || []).filter(r => r.duAn === projectName).length
  const nhanCount = (nhanRows || []).filter(r => r.duAn === projectName).length
  const hasData = giaoCount > 0 || nhanCount > 0

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

        {/* Edit Fields */}
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: '#475569', textAlign: 'left' }}>
              Supabase Project URL (VITE_SUPABASE_URL)
            </label>
            <input
              type="text"
              className="input"
              style={{ width: '100%', fontFamily: 'monospace', fontSize: 12 }}
              placeholder="Ví dụ: https://your-project-id.supabase.co"
              value={url}
              onChange={(e) => setUrl(e.target.value)}
            />
          </div>

          <div style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
            <label style={{ fontSize: 12, fontWeight: 600, color: '#475569', textAlign: 'left' }}>
              Supabase Project API Anon Key (VITE_SUPABASE_ANON_KEY)
            </label>
            <textarea
              className="input"
              rows={3}
              style={{ width: '100%', fontFamily: 'monospace', fontSize: 11, resize: 'vertical', padding: '6px 10px', lineHeight: '1.4' }}
              placeholder="Nhập chuỗi Anon Key..."
              value={key}
              onChange={(e) => setKey(e.target.value)}
            />
          </div>
        </div>

        {/* Live Test connection button & results */}
        <div style={{ border: '1px dashed #cbd5e1', borderRadius: 8, padding: 12, background: '#f8fafc' }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span style={{ fontSize: 12, fontWeight: 600, color: '#334155' }}>Kiểm tra trạng thái kết nối trực tiếp:</span>
            <button
              type="button"
              onClick={handleTestConnection}
              disabled={testing}
              style={{
                background: testing ? '#e2e8f0' : '#f1f5f9',
                border: '1px solid #cbd5e1',
                padding: '4px 10px',
                borderRadius: 4,
                fontSize: 11,
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
                  <RefreshCw size={11} className="animate-spin" />
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
              fontSize: 11,
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
CREATE POLICY "Allow public don_nhan" ON public.don_nhan FOR ALL USING (true) WITH CHECK (true);`}
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
{`-- A. Tạo cột ten_du_an mới cho bảng don_giao & don_nhan (nếu chưa có)
ALTER TABLE public.don_giao ADD COLUMN IF NOT EXISTS ten_du_an text;
ALTER TABLE public.don_nhan ADD COLUMN IF NOT EXISTS ten_du_an text;

-- B. Di chuyển tự động toàn bộ dữ liệu dự án cũ sang cột mới (Bảo toàn dữ liệu cũ)
UPDATE public.don_giao SET ten_du_an = du_an WHERE ten_du_an IS NULL OR ten_du_an = '';
UPDATE public.don_nhan SET ten_du_an = du_an WHERE ten_du_an IS NULL OR ten_du_an = '';

-- C. Đảm bảo cột ten_du_an trong bảng du_an có ràng buộc duy nhất (Bắt buộc để làm Khóa Ngoại)
ALTER TABLE public.du_an DROP CONSTRAINT IF EXISTS du_an_ten_du_an_key;
ALTER TABLE public.du_an ADD CONSTRAINT du_an_ten_du_an_key UNIQUE (ten_du_an);

-- D. Loại bỏ khóa ngoại cũ liên kết trực tiếp vào cột du_an cũ (nếu có)
ALTER TABLE public.don_giao DROP CONSTRAINT IF EXISTS don_giao_du_an_fkey;
ALTER TABLE public.don_nhan DROP CONSTRAINT IF EXISTS don_nhan_du_an_fkey;

-- E. Thiết lập khóa ngoại liên kết chuẩn xác đến cột ten_du_an của bảng du_an
ALTER TABLE public.don_giao ADD CONSTRAINT don_giao_ten_du_an_fkey 
  FOREIGN KEY (ten_du_an) REFERENCES public.du_an (ten_du_an) ON UPDATE CASCADE ON DELETE SET NULL;

ALTER TABLE public.don_nhan ADD CONSTRAINT don_nhan_ten_du_an_fkey 
  FOREIGN KEY (ten_du_an) REFERENCES public.du_an (ten_du_an) ON UPDATE CASCADE ON DELETE SET NULL;

-- F. (Tùy chọn) Xóa hẳn hai cột du_an cũ thừa thãi để sơ đồ Supabase sạch đẹp 100%
-- ALTER TABLE public.don_giao DROP COLUMN IF EXISTS du_an;
-- ALTER TABLE public.don_nhan DROP COLUMN IF EXISTS du_an;`}
            </pre>
            <p style={{ margin: 0, fontSize: 11, color: '#16a34a', fontWeight: 500 }}>
              ✔️ <strong>Lưu ý:</strong> Ứng dụng đã được tích hợp cơ chế <strong>Tự phục hồi (Self-Healing)</strong> thông minh. Bất kể bạn đang sử dụng cấu trúc cũ (cột <code>du_an</code>) hay cấu trúc mới (cột <code>ten_du_an</code>), phần mềm sẽ tự phát hiện lỗi column, tự gỡ bỏ cột thừa và lưu trữ trơn tru mà không làm ngắt quãng công việc của bạn!
            </p>
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


// Helper functions for robust database access and column-case normalization
function toSnakeCase(str) {
  if (str === 'maSAP') return 'ma_sap'
  if (str === 'maDonChuyenTiepLC') return 'ma_don_chuyen_tiep_lc'
  if (str === 'maDonChuyenTiepNB') return 'ma_don_chuyen_tiep_nb'
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

// Adaptive database row deletion helper which detects table columns from OpenAPI schema or fallback methods
async function deleteFromTableAdaptive(tableName, possibleColumns, targetValue) {
  try {
    let columns = null
    
    // Try 1: OpenAPI Schema detection (highly robust, zero 400 bad requests)
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
          console.log(`[Schema Finder] Phát hiện các cột của bảng "${tableName}" từ OpenAPI schema:`, columns)
        }
      }
    } catch (schemaErr) {
      console.warn('[Schema Finder] Lỗi khi tải OpenAPI schema từ Supabase:', schemaErr.message)
    }

    // Try 2: Row inspection fallback (if OpenAPI schema was unreachable or failed)
    if (!columns) {
      const { data, error } = await supabase.from(tableName).select('*').limit(1)
      if (!error && data && data.length > 0) {
        columns = Object.keys(data[0])
        console.log(`[Schema Finder] Phát hiện các cột của bảng "${tableName}" từ truy vấn dòng mẫu:`, columns)
      } else if (!error && (!data || data.length === 0)) {
        // Table is empty, no-op is safe
        console.log(`[Schema Finder] Bảng "${tableName}" đang trống hoàn toàn. Không cần gửi lệnh DELETE.`)
        return { success: true, reason: 'table_empty' }
      } else if (error) {
        console.warn(`[Schema Finder] Lỗi truy vấn dòng mẫu từ bảng "${tableName}":`, error.message)
      }
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

        // Verify if the item actually deleted from Supabase to catch silent RLS denials (PostgreSQL RLS silent bypass returns 200 OK)
        const { data: checkData, error: checkErr } = await supabase.from(tableName).select(exactCol).eq(exactCol, targetValue).limit(1)
        if (!checkErr && checkData && checkData.length > 0) {
          console.warn(`[Adaptive Delete] Phát hiện dòng có "${exactCol}" = "${targetValue}" vẫn hiện diện trên "${tableName}" mặc dù lệnh DELETE trả về thành công! Đây chắc chắn là do chính sách bảo mật RLS (Row Level Security) đang chặn quyền DELETE.`)
          throw new Error(`Supabase từ chối xóa dòng này (bị chặn bởi cấu hình bảo mật RLS - Row Level Security). Vui lòng tạo chính sách "DELETE" cho vai trò "anon" (hoặc "public") trên bảng "${tableName}" trong trang quản trị.`)
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
        // Verify delete on fallback
        const { data: checkData, error: checkErr } = await supabase.from(tableName).select(col).eq(col, targetValue).limit(1)
        if (!checkErr && checkData && checkData.length > 0) {
          lastError = new Error(`Cập nhật bị chặn bởi RLS: Bảng "${tableName}" không cho phép xóa cột "${col}" bằng vai trò hiện tại.`)
          continue
        }
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
  const [tab, setTab] = useState('giao')
  const [selectedProject, setSelectedProject] = useState('')
  const [showAddProjectModal, setShowAddProjectModal] = useState(false)
  const [showEditProjectModal, setShowEditProjectModal] = useState(false)
  const [showDeleteProjectModal, setShowDeleteProjectModal] = useState(false)
  const [showConfigModal, setShowConfigModal] = useState(false)
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

  const [syncingType, setSyncingType] = useState(null) // 'giao' | 'nhan' | null
  const [supabaseMessage, setSupabaseMessage] = useState(null) // { text, type: 'success' | 'error' | 'info' }

  // Fetch data from Supabase - hoisted so it can be called anywhere (e.g. after delete)
  const loadDataFromSupabase = React.useCallback(async () => {
    if (!isSupabaseConfigured) return
    try {
      const { data: projData, error: projError } = await supabase
        .from('du_an')
        .select('*')
      
      if (projError) {
        console.error('Lỗi tải danh mục dự án từ Supabase:', projError)
        if (projError.status === 401 || projError.message?.includes('Unauthorized') || projError.message?.includes('JWT') || projError.message?.includes('key') || projError.message?.includes('Invalid API key')) {
          setSupabaseAuthError(true)
        }
      } else {
        setSupabaseAuthError(false)
      }

      if (!projError && projData) {
        const list = projData.map(d => {
          const key = Object.keys(d).find(k =>
            k.toLowerCase() === 'ten_du_an' ||
            k.toLowerCase() === 'tenduan' ||
            k.toLowerCase() === 'ten_duan' ||
            k.toLowerCase() === 'tendu_an' ||
            k.toLowerCase() === 'tenduan'
          ) || Object.keys(d)[0]
          return d[key]
        }).filter(Boolean)
        const sortedList = [...list].sort()
        setCustomProjects(sortedList)
        localStorage.setItem('sgc_custom_projects', JSON.stringify(sortedList))
      }

      // Paginate and load ALL rows for don_giao without limit constraints
      let gData = []
      let fromG = 0
      const PAGE_SIZE = 1000
      let hasMoreG = true
      let gErrorObj = null

      while (hasMoreG) {
        const { data: gPage, error: gError } = await supabase
          .from('don_giao')
          .select('*')
          .order('id', { ascending: true })
          .range(fromG, fromG + PAGE_SIZE - 1)
        
        if (gError) {
          gErrorObj = gError
          break
        }
        if (gPage && gPage.length > 0) {
          gData = [...gData, ...gPage]
          if (gPage.length < PAGE_SIZE) {
            hasMoreG = false
          } else {
            fromG += PAGE_SIZE
          }
        } else {
          hasMoreG = false
        }
      }

      if (gErrorObj) {
        if (gErrorObj.status === 401) {
          setSupabaseAuthError(true)
        }
      } else {
        if (gData && gData.length > 0) {
          const mappedG = gData.map((item, idx) => {
            const normalized = normalizeDbRow(item)
            return { id: normalized.id || idx, ...normalized }
          })
          // Sắp xếp tăng dần theo ID để đảm bảo thứ tự dòng từ trên xuống như file gốc
          mappedG.sort((a, b) => {
            const idA = Number(a.id)
            const idB = Number(b.id)
            if (!isNaN(idA) && !isNaN(idB)) {
              return idA - idB
            }
            return 0
          })
          setGiaoRows(mappedG)
          setGiaoFileName('Report_Orders_Don_giao (Supabase DB)')
        } else {
          setGiaoRows([])
          setGiaoFileName('')
        }
      }

      // Paginate and load ALL rows for don_nhan without limit constraints
      let nData = []
      let fromN = 0
      let hasMoreN = true
      let nErrorObj = null

      while (hasMoreN) {
        const { data: nPage, error: nError } = await supabase
          .from('don_nhan')
          .select('*')
          .order('id', { ascending: true })
          .range(fromN, fromN + PAGE_SIZE - 1)
        
        if (nError) {
          nErrorObj = nError
          break
        }
        if (nPage && nPage.length > 0) {
          nData = [...nData, ...nPage]
          if (nPage.length < PAGE_SIZE) {
            hasMoreN = false
          } else {
            fromN += PAGE_SIZE
          }
        } else {
          hasMoreN = false
        }
      }

      if (nErrorObj) {
        if (nErrorObj.status === 401) {
          setSupabaseAuthError(true)
        }
      } else {
        if (nData && nData.length > 0) {
          const mappedN = nData.map((item, idx) => {
            const normalized = normalizeDbRow(item)
            return { id: normalized.id || idx, ...normalized }
          })
          // Sắp xếp tăng dần theo ID để đảm bảo thứ tự dòng từ trên xuống như file gốc
          mappedN.sort((a, b) => {
            const idA = Number(a.id)
            const idB = Number(b.id)
            if (!isNaN(idA) && !isNaN(idB)) {
              return idA - idB
            }
            return 0
          })
          setNhanRows(mappedN)
          setNhanFileName('Report_Orders_Don_nhan (Supabase DB)')
        } else {
          setNhanRows([])
          setNhanFileName('')
        }
      }
    } catch (e) {
      console.error('Lỗi khi tải dữ liệu từ Supabase:', e)
    }
  }, [])

  // Fetch initial data from Supabase if connected
  React.useEffect(() => {
    loadDataFromSupabase()

    if (!isSupabaseConfigured) return

    // Đăng ký nhận thông tin Realtime tức thời từ Supabase
    const channel = supabase
      .channel('schema-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'don_giao' }, () => {
        console.log('[Supabase Realtime] Tự động cập nhật bảng Đơn Giao...')
        loadDataFromSupabase()
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'don_nhan' }, () => {
        console.log('[Supabase Realtime] Tự động cập nhật bảng Đơn Nhận...')
        loadDataFromSupabase()
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'du_an' }, () => {
        console.log('[Supabase Realtime] Tự động cập nhật danh sách Dự Án...')
        loadDataFromSupabase()
      })
      .subscribe()

    return () => {
      supabase.removeChannel(channel)
    }
  }, [loadDataFromSupabase])

  const syncRowsToSupabase = async (type, rowsToSync, isAuto = false) => {
    if (!isSupabaseConfigured) return

    if (!rowsToSync || rowsToSync.length === 0) {
      setSupabaseMessage({
        text: 'Không có dữ liệu để đồng bộ.',
        type: 'error'
      })
      setTimeout(() => setSupabaseMessage(null), 4000)
      return
    }

    setSyncingType(type)
    setSupabaseMessage({
      text: isAuto 
        ? `Đang tự động đồng bộ ${rowsToSync.length} dòng lên Supabase...`
        : `Đang lưu dữ liệu Đơn ${type === 'giao' ? 'Giao' : 'Nhận'} lên Supabase...`,
      type: 'info'
    })

    try {
      const tableName = type === 'giao' ? 'don_giao' : 'don_nhan'

      // 1. Delete existing records belonging to the synced projects to avoid wiping other unrelated projects.
      // If no projects are found in the sync list, delete all table rows as fallback.
      const uniqueProjectsInSync = [...new Set(rowsToSync.map(r => r.duAn).filter(Boolean))]
      
      if (uniqueProjectsInSync.length > 0) {
        console.log(`[Sync] Đang dọn dẹp các dòng cũ thuộc ${uniqueProjectsInSync.length} dự án đang lưu...`)
        for (const proj of uniqueProjectsInSync) {
          const delRes = await deleteFromTableAdaptive(tableName, ['duAn', 'du_an', 'duan', 'ten_du_an', 'tenduan', 'tenDuAn'], proj)
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

        // Tự động điền Kho dự án tương ứng vào cột ten_du_an và du_an (duAn)
        const projectVal = row.duAn || selectedProject || ''
        item['duAn'] = projectVal || null
        item['ten_du_an'] = selectedProject || projectVal || null
        item['tenDuAn'] = selectedProject || projectVal || null

        return item
      })

      // 3. Batch insert with chunking & smart casing fallbacks
      const chunkSize = 1000
      for (let i = 0; i < payload.length; i += chunkSize) {
        const chunk = payload.slice(i, i + chunkSize)
        await insertWithFallback(tableName, chunk)
      }

      // 4. Ensure project names exist in 'du_an' database safely.
      // Instead of an unsafe upsert that may trigger 409 Conflict errors (due to lack of unique constraint/index),
      // we query existing project names first and insert only the ones that do not exist.
      const uniqueProjects = [...new Set(rowsToSync.map(r => r.duAn).filter(Boolean))]
      if (uniqueProjects.length > 0) {
        try {
          const { data: existingProj, error: fetchErr } = await supabase
            .from('du_an')
            .select('ten_du_an, tenduan, tenDuAn')
          
          if (!fetchErr) {
            const existingNames = new Set(
              (existingProj || [])
                .map(p => p.ten_du_an || p.tenduan || p.tenDuAn)
                .filter(Boolean)
                .map(n => n.trim().toLowerCase())
            )

            const newProjects = uniqueProjects.filter(p => !existingNames.has(p.trim().toLowerCase()))

            if (newProjects.length > 0) {
              let colName = 'ten_du_an'
              if (existingProj && existingProj.length > 0) {
                const firstRowKeys = Object.keys(existingProj[0])
                if (firstRowKeys.includes('tenduan')) colName = 'tenduan'
                else if (firstRowKeys.includes('tenDuAn')) colName = 'tenDuAn'
              }
              
              const insertPayload = newProjects.map(p => ({ [colName]: p }))
              const { error: insErr } = await supabase.from('du_an').insert(insertPayload)
              if (insErr) {
                console.warn('[Sync Projects] Lỗi khi chèn dự án mới:', insErr.message)
              }
            }
          }
        } catch (e) {
          console.warn('Lỗi đồng bộ danh sách dự án:', e)
        }
      }

      // Refresh project list locally
      if (uniqueProjects.length > 0) {
        setCustomProjects(prev => {
          const merged = new Set([...prev, ...uniqueProjects])
          return [...merged].sort()
        })
      }

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
    }
  }

  const handleSyncToSupabase = async (type) => {
    const rowsToSync = type === 'giao' ? giaoRows : nhanRows
    await syncRowsToSupabase(type, rowsToSync, false)
  }

  const handleImportFile = async (type, parsedRows, name) => {
    // Hiển thị modal preview trước, chờ người dùng xác nhận lưu
    setPreviewModal({ type, rows: parsedRows, fileName: name })
  }

  const handleConfirmImport = async () => {
    if (!previewModal) return
    const { type, rows: parsedRows, fileName: name } = previewModal
    setPreviewModal(null)

    const unitKey = type === 'giao' ? 'donViGiao' : 'donViNhan'

    // Bước 1: Lọc các dòng có Đơn vị giao/nhận trùng khớp với Kho dự án đang chọn (trùng khớp hoàn toàn, không phân biệt chữ hoa thường)
    const matchedByUnit = selectedProject
      ? parsedRows.filter(r => {
          const unit = String(r[unitKey] || '').trim().toLowerCase()
          const proj = selectedProject.trim().toLowerCase()
          return unit === proj
        })
      : parsedRows

    // Bước 2: Chỉ lấy các dòng có trạng thái "Đã phê duyệt"
    const approvedRows = matchedByUnit.filter(r => isApprovedStatus(r.trangThai))

    // Bước 3: Gán duAn và ten_du_an cho tất cả rows đã lọc - tự động điền thông tin Kho dự án đã chọn làm mặc định hoặc giữ nguyên nếu khớp
    const rowsToStore = approvedRows.map(row => {
      const originalDuAn = row.duAn && String(row.duAn).trim() !== '' ? String(row.duAn).trim() : ''
      const finalDuAn = selectedProject || originalDuAn || ''
      return { 
        ...row, 
        duAn: finalDuAn,
        tenDuAn: finalDuAn,
        ten_du_an: finalDuAn
      }
    })

    if (type === 'giao') {
      setGiaoRows(rowsToStore)
      setGiaoFileName(name)
    } else {
      setNhanRows(rowsToStore)
      setNhanFileName(name)
    }

    if (isSupabaseConfigured) {
      await syncRowsToSupabase(type, rowsToStore, true)
    }
  }

  const handleDeleteFile = async (type) => {
    // 1. Xóa dữ liệu local - chỉ xóa rows thuộc selectedProject, không xóa dữ liệu kho khác
    if (type === 'giao') {
      if (selectedProject) {
        setGiaoRows(prev => prev.filter(r => r.duAn !== selectedProject))
      } else {
        setGiaoRows([])
        setGiaoFileName('')
      }
    } else {
      if (selectedProject) {
        setNhanRows(prev => prev.filter(r => r.duAn !== selectedProject))
      } else {
        setNhanRows([])
        setNhanFileName('')
      }
    }

    // 2. Nếu đã kết nối Supabase, xóa các dòng tương ứng trên Supabase
    if (isSupabaseConfigured) {
      const tableName = type === 'giao' ? 'don_giao' : 'don_nhan'
      setSyncingType(type)
      setSupabaseMessage({ text: `Đang xóa dữ liệu Đơn ${type === 'giao' ? 'Giao' : 'Nhận'} trên Supabase...`, type: 'info' })
      try {
        let query = supabase.from(tableName).delete()
        if (selectedProject) {
          // Xóa theo dự án đang chọn sử dụng cơ chế Adaptive cực kỳ tối ưu
          const deleteRes = await deleteFromTableAdaptive(tableName, ['duAn', 'du_an', 'duan', 'ten_du_an', 'tenduan', 'tenDuAn'], selectedProject)
          if (!deleteRes.success && deleteRes.reason !== 'table_empty') {
            throw deleteRes.error || new Error('Không tìm thấy cột phù hợp hoặc không được phân quyền xóa.')
          }
        } else {
          // Không có dự án cụ thể → xóa hết bảng
          await query.neq('id', -999)
        }
        setSupabaseMessage({ text: `Đã xóa dữ liệu Đơn ${type === 'giao' ? 'Giao' : 'Nhận'} khỏi Supabase thành công!`, type: 'success' })
        setTimeout(() => setSupabaseMessage(null), 4000)
      } catch (err) {
        setSupabaseMessage({ text: `Lỗi khi xóa trên Supabase: ${err.message || err}`, type: 'error' })
        setTimeout(() => setSupabaseMessage(null), 5000)
      }
      setSyncingType(null)
    }
  }

  const handleAddProject = (name) => {
    const trimmed = name.trim()
    if (!trimmed) return
    if (!customProjects.includes(trimmed)) {
      const updated = [...customProjects, trimmed]
      setCustomProjects(updated)
      localStorage.setItem('sgc_custom_projects', JSON.stringify(updated))

      if (isSupabaseConfigured) {
        // Try inserting as ten_du_an (snake_case)
        supabase.from('du_an')
          .insert([{ ten_du_an: trimmed }])
          .then(({ error }) => {
            if (error) {
              // Try tenduan (lowercase)
              supabase.from('du_an')
                .insert([{ tenduan: trimmed }])
                .then(({ error: error2 }) => {
                  if (error2) {
                    // Try tenDuAn (camelCase)
                    supabase.from('du_an')
                      .insert([{ tenDuAn: trimmed }])
                      .catch(e => console.error('Lỗi lưu dự án:', e))
                  }
                })
            }
          })
      }
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

    // 2. Update existing rows locally so user sees immediate results on screen
    setGiaoRows(prev => prev.map(row => {
      if (row.duAn === trimmedOld) {
        return { ...row, duAn: trimmedNew }
      }
      return row
    }))
    setNhanRows(prev => prev.map(row => {
      if (row.duAn === trimmedOld) {
        return { ...row, duAn: trimmedNew }
      }
      return row
    }))

    // Update active selection if user had the old name selected
    if (selectedProject === trimmedOld) {
      setSelectedProject(trimmedNew)
    }

    // 3. Update Supabase
    if (isSupabaseConfigured) {
      try {
        setSupabaseMessage({
          text: `Đang kết nối để sửa tên dự án thành "${trimmedNew}" trên Supabase...`,
          type: 'info'
        })

        // -- Update 'du_an' table
        // Try all casing structures for public schema matching:
        const { error: pErr1 } = await supabase
          .from('du_an')
          .update({ ten_du_an: trimmedNew })
          .eq('ten_du_an', trimmedOld)
        
        if (pErr1) {
          const { error: pErr2 } = await supabase
            .from('du_an')
            .update({ tenduan: trimmedNew })
            .eq('tenduan', trimmedOld)
          
          if (pErr2) {
            await supabase
              .from('du_an')
              .update({ tenDuAn: trimmedNew })
              .eq('tenDuAn', trimmedOld)
              .catch(e => console.warn('Ignore: failed camelCase du_an rename', e))
          }
        }

        // -- Update 'don_giao' table
        const { error: gErr0 } = await supabase
          .from('don_giao')
          .update({ ten_du_an: trimmedNew })
          .eq('ten_du_an', trimmedOld)

        if (gErr0) {
          const { error: gErr1 } = await supabase
            .from('don_giao')
            .update({ du_an: trimmedNew })
            .eq('du_an', trimmedOld)
          
          if (gErr1) {
            const { error: gErr2 } = await supabase
              .from('don_giao')
              .update({ duan: trimmedNew })
              .eq('duan', trimmedOld)
            
            if (gErr2) {
              await supabase
                .from('don_giao')
                .update({ duAn: trimmedNew })
                .eq('duAn', trimmedOld)
                .catch(e => console.warn('Ignore: failed camelCase don_giao rename', e))
            }
          }
        }

        // -- Update 'don_nhan' table
        const { error: nErr0 } = await supabase
          .from('don_nhan')
          .update({ ten_du_an: trimmedNew })
          .eq('ten_du_an', trimmedOld)

        if (nErr0) {
          const { error: nErr1 } = await supabase
            .from('don_nhan')
            .update({ du_an: trimmedNew })
            .eq('du_an', trimmedOld)
          
          if (nErr1) {
            const { error: nErr2 } = await supabase
              .from('don_nhan')
              .update({ duan: trimmedNew })
              .eq('duan', trimmedOld)
            
            if (nErr2) {
              await supabase
                .from('don_nhan')
                .update({ duAn: trimmedNew })
                .eq('duAn', trimmedOld)
                .catch(e => console.warn('Ignore: failed camelCase don_nhan rename', e))
            }
          }
        }

        setSupabaseMessage({
          text: `Cập nhật dự án thành công & đã đồng bộ dữ liệu trên Supabase!`,
          type: 'success'
        })
        setTimeout(() => setSupabaseMessage(null), 5000)
      } catch (err) {
        console.error('Lỗi khi cập nhật Supabase:', err)
        setSupabaseMessage({
          text: `Sửa cục bộ thành công! Nhưng gặp lỗi đồng bộ Supabase: ${err.message || err}`,
          type: 'error'
        })
        setTimeout(() => setSupabaseMessage(null), 6000)
      }
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
      if (row.duAn === trimmed) {
        return { ...row, duAn: '' }
      }
      return row
    }))
    setNhanRows(prev => prev.map(row => {
      if (row.duAn === trimmed) {
        return { ...row, duAn: '' }
      }
      return row
    }))

    // Reset active selection if the deleted project is currently selected
    if (selectedProject === trimmed) {
      setSelectedProject('')
    }

    // 2. Xóa trên Supabase nếu đã cấu hình
    if (isSupabaseConfigured) {
      try {
        setSupabaseMessage({
          text: `Đang kết nối để xóa kho dự án "${trimmed}" trên Supabase...`,
          type: 'info'
        })

        // Xóa tuần tự: Xóa các tham chiếu ở con trước để tránh Foreign Key Violation, sau đó mới xóa ở cha
        console.log('[Supabase Delete] Thực hiện xóa liên kết don_giao...')
        const resGiao = await deleteFromTableAdaptive('don_giao', ['duAn', 'du_an', 'duan', 'ten_du_an', 'tenduan', 'tenDuAn'], trimmed)
        if (!resGiao.success && resGiao.reason !== 'table_empty') {
          console.warn('[Supabase Delete] Cảnh báo lỗi xóa don_giao (có thể bỏ qua nếu bảng không dùng/không có dữ liệu):', resGiao.error)
        }

        console.log('[Supabase Delete] Thực hiện xóa liên kết don_nhan...')
        const resNhan = await deleteFromTableAdaptive('don_nhan', ['duAn', 'du_an', 'duan', 'ten_du_an', 'tenduan', 'tenDuAn'], trimmed)
        if (!resNhan.success && resNhan.reason !== 'table_empty') {
          console.warn('[Supabase Delete] Cảnh báo lỗi xóa don_nhan (col có thể bỏ qua nếu bảng không dùng/không có dữ liệu):', resNhan.error)
        }

        console.log('[Supabase Delete] Thực hiện xóa dự án trong bảng du_an...')
        const resDuAn = await deleteFromTableAdaptive('du_an', ['ten_du_an', 'tenduan', 'tenDuAn', 'ten_duan', 'tendu_an', 'name'], trimmed)
        if (!resDuAn.success) {
          const errObj = resDuAn.error || {}
          const errorMsg = errObj.message || errObj.details || 'Không rõ nguyên nhân. Vui lòng kiểm tra quyền RLS (Row Level Security) cho chính sách DELETE hoặc ràng buộc khóa ngoại (Foreign Key Constraints).'
          throw new Error(errorMsg)
        }

        // Reload toàn bộ dữ liệu từ Supabase để đồng bộ sau khi xóa
        await loadDataFromSupabase()

        setSupabaseMessage({
          text: `Xóa kho dự án "${trimmed}" thành công và đã đồng bộ!`,
          type: 'success'
        })
        setTimeout(() => setSupabaseMessage(null), 4000)
      } catch (err) {
        console.error('Lỗi khi xóa dự án trên Supabase:', err)
        // Dù lỗi vẫn thử reload để đồng bộ
        await loadDataFromSupabase().catch(() => {})
        setSupabaseMessage({
          text: `Đã xóa cục bộ thành công! Nhưng gặp lỗi đồng bộ Supabase: ${err.message || err}`,
          type: 'error'
        })
        setTimeout(() => setSupabaseMessage(null), 5000)
      }
    }
  }

  const fileProjects = useMemo(() => {
    const list = new Set()
    giaoRows.forEach(r => { if (r.duAn) list.add(r.duAn) })
    nhanRows.forEach(r => { if (r.duAn) list.add(r.duAn) })
    return [...list].sort()
  }, [giaoRows, nhanRows])

  const allProjects = useMemo(() => {
    // Danh sách dự án chỉ lấy từ customProjects (do người dùng tạo bằng nút "+ Tạo mới")
    // KHÔNG tự động lấy từ file upload lên
    return [...new Set(customProjects)].sort()
  }, [customProjects])

  const tabs = [
    { id: 'giao', label: 'Đơn Giao', icon: <Truck size={15} /> },
    { id: 'nhan', label: 'Đơn Nhận', icon: <PackageCheck size={15} /> },
    { id: 'summary', label: 'Tổng hợp', icon: <BarChart3 size={15} /> },
  ]

  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column', overflow: 'hidden' }}>
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
      />
      <TabBar tabs={tabs} active={tab} onChange={setTab} />

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

      <div style={{ flex: 1, overflow: 'hidden' }}>
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
            onImportFile={(rows, name) => handleImportFile('giao', rows, name)}
            onDeleteFile={() => handleDeleteFile('giao')}
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
            onImportFile={(rows, name) => handleImportFile('nhan', rows, name)}
            onDeleteFile={() => handleDeleteFile('nhan')}
          />
        )}
        {tab === 'summary' && (
          <PlaceholderTab
            icon={<BarChart3 />}
            title="Tổng hợp"
            desc="Tính năng tổng hợp và phân tích dữ liệu sẽ được triển khai trong phiên bản tiếp theo."
          />
        )}
      </div>

      <PreviewImportModal
        isOpen={!!previewModal}
        onClose={() => setPreviewModal(null)}
        onConfirm={handleConfirmImport}
        rows={previewModal?.rows || []}
        fileName={previewModal?.fileName || ''}
        type={previewModal?.type || 'giao'}
        selectedProject={selectedProject}
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
      />

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
      `}</style>
    </div>
  )
}
