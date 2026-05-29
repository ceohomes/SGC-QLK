import React, { useState, useRef, useCallback, useMemo } from 'react'
import * as XLSX from 'xlsx'
import {
  Upload, FileSpreadsheet, Search, X, RefreshCw, Info,
  ChevronDown, Download, Truck, PackageCheck, Settings, BarChart3,
  AlertCircle, CheckCircle2, Filter, ArrowUpDown, Clock, CloudUpload, Database, Save,
  Pencil, Trash2
} from 'lucide-react'
import { COLS_GIAO_NHAN, parseXlsxToRows, formatVal, getTrangThaiColor } from './constants.js'
import { MOCK_GIAO_ROWS, MOCK_NHAN_ROWS } from './mockData.js'
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
function UploadZone({ onFile, label, accept = '.xlsx,.xls' }) {
  const [drag, setDrag] = useState(false)
  const ref = useRef()

  const handle = useCallback((file) => {
    if (!file) return
    const reader = new FileReader()
    reader.onload = (e) => {
      const wb = XLSX.read(e.target.result, { type: 'array' })
      const ws = wb.Sheets[wb.SheetNames[0]]
      const data = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null })
      onFile(data, file.name)
    }
    reader.readAsArrayBuffer(file)
  }, [onFile])

  return (
    <div
      className={`upload-zone${drag ? ' drag-over' : ''}`}
      onClick={() => ref.current.click()}
      onDragOver={e => { e.preventDefault(); setDrag(true) }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); handle(e.dataTransfer.files[0]) }}
    >
      <input ref={ref} type="file" accept={accept} onChange={e => handle(e.target.files[0])} />
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 12 }}>
        <div style={{
          width: 56, height: 56, borderRadius: 12,
          background: 'var(--primary-light)',
          display: 'flex', alignItems: 'center', justifyContent: 'center',
          boxShadow: '0 2px 8px rgba(37,99,235,0.08)'
        }}>
          <Upload size={24} color="var(--primary)" />
        </div>
        <div>
          <div style={{ fontWeight: 700, fontSize: 16, color: 'var(--text)', marginBottom: 6 }}>
            {label}
          </div>
          <div style={{ color: 'var(--text-muted)', fontSize: 14 }}>
            Kéo thả file vào đây hoặc <span style={{ color: 'var(--primary)', fontWeight: 600 }}>chọn file từ thiết bị</span>
          </div>
          <div style={{ color: 'var(--text-light)', fontSize: 12, marginTop: 6, fontStyle: 'italic' }}>
            Hỗ trợ định dạng chuẩn: .xlsx, .xls
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
    const daPheDuyet = rows.filter(r => {
      const ts = String(r.trangThai || '').toLowerCase();
      return ts.includes('đã phê duyệt') || ts.includes('approved') || ts.includes('hoàn thành');
    }).length;
    
    const choPheDuyet = rows.filter(r => {
      const ts = String(r.trangThai || '').toLowerCase();
      return ts.includes('chờ') || ts.includes('pending') || ts.includes('chưa');
    }).length;

    const tuChoi = rows.filter(r => {
      const ts = String(r.trangThai || '').toLowerCase();
      return ts.includes('từ') || ts.includes('rejected') || ts.includes('hủy');
    }).length;

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
function DataTable({ rows }) {
  if (rows.length === 0) return (
    <div className="empty-state">
      <Search size={48} />
      <h3>Không có dữ liệu</h3>
      <p>Không tìm thấy kết quả phù hợp với bộ lọc hiện tại.</p>
    </div>
  )

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th style={{ width: 50, minWidth: 50, maxWidth: 50, textAlign: 'center', verticalAlign: 'middle', fontSize: '12px', padding: '8px 10px' }}>
              STT
            </th>
            {COLS_GIAO_NHAN.map(c => {
              return (
                <th
                  key={c.key}
                  style={{
                    width: c.width,
                    minWidth: c.width,
                    maxWidth: c.width,
                    textAlign: 'center',
                    verticalAlign: 'middle',
                    fontSize: '12px',
                    padding: '8px 10px',
                    whiteSpace: 'normal',
                    wordBreak: 'break-word',
                    lineHeight: '1.2'
                  }}
                >
                  {c.label}
                </th>
              );
            })}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={row.id}>
              <td style={{ width: 50, minWidth: 50, maxWidth: 50, textAlign: 'center', fontSize: '12px', color: '#1b1919', padding: '6px 10px' }}>{i + 1}</td>
              {COLS_GIAO_NHAN.map(col => {
                const isCenteredCol = [
                  'ngayXuatNhap', 'maVatTu', 'maSAP', 'dvt', 'loaiDon',
                  'maDonViGiao', 'donViGiao', 'nguoiGiao',
                  'maDonViNhan', 'donViNhan', 'nguoiPheDuyet', 'nguoiNhan',
                  'soHopDong', 'thuKho', 'tinhTrang'
                ].includes(col.key);
                const isRightAligned = ['khoiLuongNhap', 'khoiLuongXuat'].includes(col.key) || col.key.toLowerCase().includes('khoiluong');
                return (
                  <td
                    key={col.key}
                    style={{
                      width: col.width,
                      minWidth: col.width,
                      maxWidth: col.width,
                      color: '#1b1919',
                      textAlign: isCenteredCol ? 'center' : (isRightAligned ? 'right' : 'left'),
                      fontSize: '12px',
                      padding: '6px 10px',
                      wordBreak: 'break-word',
                      whiteSpace: 'normal'
                    }}
                    title={String(formatVal(row[col.key]) || '')}
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
                      formatVal(row[col.key]) !== null && formatVal(row[col.key]) !== undefined ? formatVal(row[col.key]) : ''
                    )}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
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
  onImportFile
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

  return (
    <div style={{ padding: 24 }}>
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
                <strong>Lưu ý:</strong> Vui lòng chọn <strong>Kho dự án</strong> ở góc trên trước khi upload file. Dữ liệu sẽ được gán vào dự án đang chọn.
              </div>
            </div>
          )}
          <UploadZone onFile={handleFile} label={uploadLabel} />
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
              {isSupabaseConfigured && rows.length > 0 && (
                <button
                  className="btn btn-sm"
                  onClick={onSync}
                  disabled={syncing}
                  style={{
                    background: 'linear-gradient(135deg, #059669 0%, #10b981 100%)',
                    color: '#ffffff',
                    border: 'none',
                    boxShadow: '0 2px 4px rgba(16,185,129,0.2)',
                  }}
                >
                  {syncing ? (
                    <RefreshCw size={12} style={{ animation: 'spin 1s linear infinite' }} />
                  ) : (
                    <CloudUpload size={12} />
                  )}
                  {syncing ? 'Đang lưu...' : 'Lưu vào Supabase'}
                </button>
              )}
              <button className="btn btn-outline btn-sm" onClick={() => { setRows([]); setFileName(''); setSearch(''); setTrangThai(''); setSelectedProject('') }}>
                <X size={12} /> Xóa file
              </button>
              <label className="btn btn-primary btn-sm" style={{ cursor: 'pointer' }}>
                <Upload size={12} /> Đổi file
                <input type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
                  onChange={e => { if (e.target.files[0]) handleFile(e.target.files[0], e.target.files[0].name); e.target.value = '' }} />
              </label>
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
function DeleteProjectModal({ isOpen, onClose, onConfirm, projectName }) {
  if (!isOpen) return null

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
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
        maxWidth: 420,
        padding: 24,
        boxShadow: '0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04)',
        border: '1px solid #fee2e2',
        display: 'flex',
        flexDirection: 'column',
        gap: 16
      }}>
        <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid #f1f5f9', paddingBottom: 12 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <AlertCircle size={20} color="#ef4444" />
            <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, color: '#991b1b' }}>Xác nhận xóa Dự án</h3>
          </div>
          <button onClick={onClose} style={{ background: 'transparent', border: 'none', color: '#64748b', cursor: 'pointer', display: 'flex', alignItems: 'center' }}>
            <X size={18} />
          </button>
        </div>

        <div style={{ display: 'flex', flexDirection: 'column', gap: 10, textAlign: 'left' }}>
          <p style={{ margin: 0, fontSize: 14, color: '#334155', lineHeight: '1.5' }}>
            Bạn có chắc chắn muốn xóa kho dự án <strong style={{ color: '#0f172a' }}>"{projectName}"</strong>?
          </p>
          <div style={{
            backgroundColor: '#fffbeb',
            border: '1px solid #fef3c7',
            borderRadius: 8,
            padding: '10px 12px',
            display: 'flex',
            gap: 8,
            alignItems: 'flex-start'
          }}>
            <AlertCircle size={16} color="#d97706" style={{ flexShrink: 0, marginTop: 2 }} />
            <p style={{ margin: 0, fontSize: 12, color: '#b45309', lineHeight: '1.4', fontWeight: 500 }}>
              Lưu ý: Thao tác này sẽ gỡ ký danh dự án khỏi danh sách chọn lựa của bạn.
            </p>
          </div>
        </div>

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 8 }}>
          <button
            onClick={onClose}
            style={{
              padding: '8px 16px',
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
            onClick={onConfirm}
            style={{
              padding: '8px 16px',
              borderRadius: 6,
              border: 'none',
              background: '#ef4444',
              color: '#ffffff',
              fontSize: 13,
              fontWeight: 600,
              cursor: 'pointer',
              boxShadow: '0 2px 4px rgba(239, 68, 68, 0.2)',
              display: 'flex',
              alignItems: 'center',
              gap: 6
            }}
          >
            <Trash2 size={14} />
            Đồng ý xóa
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
    }
  }, [isOpen])

  if (!isOpen) return null

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
              Trạng thái: {isCurrentlyConnected ? 'Đã kết nối với Database' : 'Chưa kết nối (Offline)'}
            </span>
            <p style={{ margin: '1px 0 0 0', fontSize: 11, color: isCurrentlyConnected ? '#047857' : '#be123c' }}>
              {isCurrentlyConnected 
                ? `Đầu nối từ: ${currentSource === 'local' ? 'Trình duyệt Web (Local Storage)' : currentSource === 'env' ? 'Biến môi trường (Vite Env)' : 'Cấu hình tích hợp mặc định (Code)'}`
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
  return str.replace(/[A-Z]/g, letter => `_${letter.toLowerCase()}`)
}

function normalizeDbRow(dbRow) {
  if (!dbRow) return dbRow
  const normalized = {}
  
  // Start with whatever properties came from DB
  Object.assign(normalized, dbRow)
  
  // Map PostgreSQL columns back to React's camelCase keys
  COLS_GIAO_NHAN.forEach(col => {
    // Look for exact camelCase, lowercase, or snake_case matches
    const dbKeyMatch = Object.keys(dbRow).find(
      k => k === col.key || 
           k.toLowerCase() === col.key.toLowerCase() || 
           k.toLowerCase() === toSnakeCase(col.key).toLowerCase()
    )
    if (dbKeyMatch && dbRow[dbKeyMatch] !== undefined) {
      normalized[col.key] = dbRow[dbKeyMatch]
    }
  })
  
  return normalized
}

async function insertWithFallback(tableName, originalChunk) {
  // Helper to run a self-healing insert-and-retry loop for a specific payload structure
  const tryInsertWithHealing = async (payload) => {
    // Deep clone the payload so modifications don't leak or cross-contaminate different formats
    let currentPayload = JSON.parse(JSON.stringify(payload))
    
    for (let retry = 0; retry < 50; retry++) {
      const { error } = await supabase.from(tableName).insert(currentPayload)
      if (!error) {
        return { success: true }
      }
      
      const errMsg = error.message || ''
      let missingField = null
      
      // Match various PostgreSQL / PostgREST missing column error formats
      const match1 = errMsg.match(/Could not find the '(.*?)' column of/i)
      const match2 = errMsg.match(/column "(.*?)" of relation/i)
      const match3 = errMsg.match(/column "(.*?)" does not exist/i)
      
      if (match1) {
        missingField = match1[1]
      } else if (match2) {
        missingField = match2[1]
      } else if (match3) {
        missingField = match3[1]
      }
      
      if (missingField) {
        console.warn(`[Self-Healing] Loại bỏ cột không tồn tại '${missingField}' trên bảng '${tableName}' và thử lại...`)
        currentPayload.forEach(row => {
          // Remove case-insensitively & snake-case-insensitively
          const keys = Object.keys(row)
          keys.forEach(k => {
            if (
              k.toLowerCase() === missingField.toLowerCase() ||
              toSnakeCase(k).toLowerCase() === toSnakeCase(missingField).toLowerCase() ||
              k.replace(/_/g, '').toLowerCase() === missingField.replace(/_/g, '').toLowerCase()
            ) {
              delete row[k]
            }
          })
          // Fallback direct deletes
          delete row[missingField]
          delete row[missingField.toLowerCase()]
          delete row[toSnakeCase(missingField)]
        })
        
        // If there are no keys left to attempt, stop to prevent infinite looping
        if (currentPayload.length > 0 && Object.keys(currentPayload[0]).length === 0) {
          break
        }
        continue
      }
      
      return { success: false, error }
    }
    return { success: false, error: new Error('Không thể lưu do loại bỏ cột không thành công hoặc lỗi DB nghiêm trọng') }
  }

  // Try 1: Exact CamelCase payload
  {
    const res = await tryInsertWithHealing(originalChunk)
    if (res.success) return { success: true }
    console.warn(`CamelCase insert failed for ${tableName}:`, res.error?.message || res.error)
  }

  // Try 2: Lowercase payload (unquoted PostgreSQL creates lowercase columns by default)
  {
    const lowercaseChunk = originalChunk.map(row => {
      const newRow = {}
      Object.keys(row).forEach(k => {
        newRow[k.toLowerCase()] = row[k]
      })
      return newRow
    })
    const res = await tryInsertWithHealing(lowercaseChunk)
    if (res.success) return { success: true }
    console.warn(`Lowercase insert failed for ${tableName}:`, res.error?.message || res.error)
  }

  // Try 3: SnakeCase payload (common PostgreSQL naming style)
  {
    const snakeChunk = originalChunk.map(row => {
      const newRow = {}
      Object.keys(row).forEach(k => {
        newRow[toSnakeCase(k)] = row[k]
      })
      return newRow
    })
    const res = await tryInsertWithHealing(snakeChunk)
    if (res.success) return { success: true }
    console.warn(`SnakeCase insert failed for ${tableName}:`, res.error?.message || res.error)
  }

  throw new Error('Supabase từ chối đồng bộ dữ liệu. Đã thử đồng hóa viết hoa/thường và tự động gỡ các cột không có trong Table nhưng vẫn thất bại.')
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState('giao')
  const [selectedProject, setSelectedProject] = useState('')
  const [showAddProjectModal, setShowAddProjectModal] = useState(false)
  const [showEditProjectModal, setShowEditProjectModal] = useState(false)
  const [showDeleteProjectModal, setShowDeleteProjectModal] = useState(false)
  const [showConfigModal, setShowConfigModal] = useState(false)
  const [projectToEdit, setProjectToEdit] = useState('')
  const [projectToDelete, setProjectToDelete] = useState('')

  const [customProjects, setCustomProjects] = useState(() => {
    try {
      const saved = localStorage.getItem('sgc_custom_projects')
      return saved ? JSON.parse(saved) : []
    } catch (e) {
      return []
    }
  })

  // Hold rows and sheet file names globally so state is retained across tab switches!
  const [giaoRows, setGiaoRows] = useState(MOCK_GIAO_ROWS)
  const [giaoFileName, setGiaoFileName] = useState('Report_Orders_Don_giao.xlsx')

  const [nhanRows, setNhanRows] = useState(MOCK_NHAN_ROWS)
  const [nhanFileName, setNhanFileName] = useState('Report_Orders_Don_nhan.xlsx')

  const [syncingType, setSyncingType] = useState(null) // 'giao' | 'nhan' | null
  const [supabaseMessage, setSupabaseMessage] = useState(null) // { text, type: 'success' | 'error' | 'info' }

  // Fetch initial data from Supabase if connected
  React.useEffect(() => {
    async function loadData() {
      if (!isSupabaseConfigured) return
      try {
        // Fetch projects supporting both snake_case, lowercase, and camelCase column naming
        const { data: projData, error: projError } = await supabase
          .from('du_an')
          .select('*')
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

        // Fetch don_giao rows
        const { data: gData, error: gError } = await supabase
          .from('don_giao')
          .select('*')
        
        if (!gError && gData && gData.length > 0) {
          const mappedG = gData.map((item, idx) => {
            const normalized = normalizeDbRow(item)
            return {
              id: normalized.id || idx,
              ...normalized
            }
          })
          setGiaoRows(mappedG)
          setGiaoFileName('Report_Orders_Don_giao (Supabase DB)')
        }

        // Fetch don_nhan rows
        const { data: nData, error: nError } = await supabase
          .from('don_nhan')
          .select('*')
        
        if (!nError && nData && nData.length > 0) {
          const mappedN = nData.map((item, idx) => {
            const normalized = normalizeDbRow(item)
            return {
              id: normalized.id || idx,
              ...normalized
            }
          })
          setNhanRows(mappedN)
          setNhanFileName('Report_Orders_Don_nhan (Supabase DB)')
        }
      } catch (e) {
        console.error('Lỗi khi tải dữ liệu từ Supabase:', e)
      }
    }
    
    loadData()

    if (!isSupabaseConfigured) return

    // Đăng ký nhận thông tin Realtime tức thời từ Supabase
    const channel = supabase
      .channel('schema-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'don_giao' }, () => {
        console.log('[Supabase Realtime] Tự động cập nhật bảng Đơn Giao...')
        loadData()
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'don_nhan' }, () => {
        console.log('[Supabase Realtime] Tự động cập nhật bảng Đơn Nhận...')
        loadData()
      })
      .on('postgres_changes', { event: '*', schema: 'public', table: 'du_an' }, () => {
        console.log('[Supabase Realtime] Tự động cập nhật danh sách Dự Án...')
        loadData()
      })
      .subscribe()

    return () => {
      supabase.removeChannel(channel)
    }
  }, [])

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

      // 1. Delete all records from this table (filter by serial id ensures lowercase success on any postgres setup)
      const { error: delError } = await supabase
        .from(tableName)
        .delete()
        .neq('id', -999)

      if (delError) throw delError

      // 2. Map structure to columns in Postgres matching COLS_GIAO_NHAN
      const payload = rowsToSync.map(row => {
        const item = {}
        COLS_GIAO_NHAN.forEach(col => {
          item[col.key] = row[col.key] !== undefined ? String(row[col.key]) : ''
        })
        return item
      })

      // 3. Batch insert with chunking & smart casing fallbacks
      const chunkSize = 200
      for (let i = 0; i < payload.length; i += chunkSize) {
        const chunk = payload.slice(i, i + chunkSize)
        await insertWithFallback(tableName, chunk)
      }

      // Also ensure project names exist/upserted in 'du_an' database
      const uniqueProjects = [...new Set(rowsToSync.map(r => r.duAn).filter(Boolean))]
      if (uniqueProjects.length > 0) {
        try {
          // Attempt upsert on snake_case
          const projPayload = uniqueProjects.map(p => ({ ten_du_an: p }))
          const { error: pErr1 } = await supabase
            .from('du_an')
            .upsert(projPayload, { onConflict: 'ten_du_an' })
          
          if (pErr1) {
            // Attempt upsert on lowercase
            const projPayload2 = uniqueProjects.map(p => ({ tenduan: p }))
            const { error: pErr2 } = await supabase
              .from('du_an')
              .upsert(projPayload2, { onConflict: 'tenduan' })
            
            if (pErr2) {
              // Attempt upsert on camelCase
              const projPayload3 = uniqueProjects.map(p => ({ tenDuAn: p }))
              await supabase
                .from('du_an')
                .upsert(projPayload3, { onConflict: 'tenDuAn' })
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
    // Nếu đang chọn một dự án cụ thể, chỉ lấy rows thuộc dự án đó
    // và gán duAn cho tất cả rows bằng selectedProject (đảm bảo khớp)
    let rowsToStore = parsedRows
    if (selectedProject) {
      // Gán dự án hiện tại cho tất cả rows được upload (rows thuộc dự án đang xem)
      rowsToStore = parsedRows.map(row => ({ ...row, duAn: selectedProject }))
    }

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

        // -- Update 'don_nhan' table
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

        // Xóa theo các trường hợp cột có thể xảy ra trong bảng 'du_an'
        const { error: err1 } = await supabase
          .from('du_an')
          .delete()
          .eq('ten_du_an', trimmed)

        if (err1) {
          const { error: err2 } = await supabase
            .from('du_an')
            .delete()
            .eq('tenduan', trimmed)

          if (err2) {
            const { error: err3 } = await supabase
              .from('du_an')
              .delete()
              .eq('tenDuAn', trimmed)
              
            if (err3) {
              await supabase
                .from('du_an')
                .delete()
                .eq('ten_duan', trimmed)
                .catch(e => console.warn('Ignore: failed lowercase snake cases delete', e))
            }
          }
        }

        // -- Clear referencing columns in 'don_giao' table on Supabase to keep it synchronized!
        const { error: gErr1 } = await supabase
          .from('don_giao')
          .update({ du_an: '' })
          .eq('du_an', trimmed)
        
        if (gErr1) {
          const { error: gErr2 } = await supabase
            .from('don_giao')
            .update({ duan: '' })
            .eq('duan', trimmed)
          
          if (gErr2) {
            await supabase
              .from('don_giao')
              .update({ duAn: '' })
              .eq('duAn', trimmed)
              .catch(e => console.warn('Ignore: failed camelCase don_giao clear references', e))
          }
        }

        // -- Clear referencing columns in 'don_nhan' table on Supabase to keep it synchronized!
        const { error: nErr1 } = await supabase
          .from('don_nhan')
          .update({ du_an: '' })
          .eq('du_an', trimmed)
        
        if (nErr1) {
          const { error: nErr2 } = await supabase
            .from('don_nhan')
            .update({ duan: '' })
            .eq('duan', trimmed)
          
          if (nErr2) {
            await supabase
              .from('don_nhan')
              .update({ duAn: '' })
              .eq('duAn', trimmed)
              .catch(e => console.warn('Ignore: failed camelCase don_nhan clear references', e))
          }
        }

        setSupabaseMessage({
          text: `Xóa kho dự án "${trimmed}" thành công!`,
          type: 'success'
        })
        setTimeout(() => setSupabaseMessage(null), 4000)
      } catch (err) {
        console.error('Lỗi khi xóa dự án trên Supabase:', err)
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
    <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>
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
      />

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
      `}</style>
    </div>
  )
}
