import React, { useState, useRef, useCallback, useMemo } from 'react'
import * as XLSX from 'xlsx'
import * as XLSXStyleRaw from 'xlsx-js-style'
const XLSXStyle = XLSXStyleRaw.default || XLSXStyleRaw
import {
  Upload, FileSpreadsheet, Search, X, RefreshCw, Info,
  ChevronDown, ChevronRight, Download, Truck, PackageCheck, Settings, BarChart3,
  AlertCircle, CheckCircle2, Filter, ArrowUpDown, Clock, CloudUpload, Database, Save,
  Pencil, Trash2, Lock, ClipboardList, Warehouse, Building2, Users, HelpCircle,
  Calendar, AlertTriangle
} from 'lucide-react'
import { COLS_GIAO_NHAN, parseXlsxToRows, formatVal, getTrangThaiColor, isApprovedStatus, isPendingStatus, isRejectedStatus } from './constants.js'
import { supabase, isSupabaseConfigured, supabaseUrl, supabaseAnonKey } from './supabaseClient.js'

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
function Header({ selectedProject, setSelectedProject, duAnOptions, onOpenAddProjectModal, onEditProject, onDeleteProject, onOpenConfigModal, onForceRefresh, onMouseEnterLogo }) {
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
            SGC | BÁO CÁO GIAO NHẬN
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
            style={{ display: 'none' }}
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) {
                const reader = new FileReader();
                reader.onload = (evt) => {
                  if (evt.target?.result) {
                    const wb = XLSX.read(evt.target.result, { type: 'array' });
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
                    onAppendFile(data, file.name);
                  }
                };
                reader.readAsArrayBuffer(file);
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

// ─── Data Table ───────────────────────────────────────────────────────────────
const PAGE_SIZE_OPTIONS = [50, 100, 200, 500]

function DataTable({ rows, setRows, type }) {
  const [pageSize, setPageSize] = React.useState(100)
  const [currentPage, setCurrentPage] = React.useState(1)
  const [selectedIds, setSelectedIds] = React.useState(new Set())
  const [showConfirmDelete, setShowConfirmDelete] = React.useState(false)
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
  }, [rows, pageSize])

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

  const pageRowIds = React.useMemo(() => pageRows.map(r => r.id), [pageRows])

  const isAllSelected = React.useMemo(() => {
    if (rows.length === 0) return false
    return rows.every(r => selectedIds.has(r.id))
  }, [rows, selectedIds])

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
    if (isSupabaseConfigured) {
      const tableName = type === 'chung' ? 'don_chung' : type === 'giao' ? 'don_giao' : type === 'nhan' ? 'don_nhan' : type === 'kho' ? 'don_kho' : 'don_chung'
      try {
        const { error } = await supabase
          .from(tableName)
          .delete()
          .in('id', idsArr)
        
        if (error) throw error
      } catch (err) {
        alert(`Lỗi khi xóa dòng trên Supabase: ${err.message || err}`)
        return
      }
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
              {COLS_GIAO_NHAN.map(c => {
                const isNhapKhoGroup = ['khoiLuongNhap', 'maDonViGiao', 'donViGiao', 'nguoiGiao'].includes(c.key)
                const isXuatKhoGroup = ['khoiLuongXuat', 'maDonViNhan', 'donViNhan', 'nguoiPheDuyet'].includes(c.key)
                const thBg = isNhapKhoGroup ? '#0f766e' : isXuatKhoGroup ? '#c2410c' : undefined
                const thBorderBottom = isNhapKhoGroup ? '2px solid #115e59' : isXuatKhoGroup ? '2px solid #9a3412' : undefined

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
                {COLS_GIAO_NHAN.map(col => {
                  const isCenteredCol = [
                    'ngayXuatNhap', 'maVatTu', 'maSAP', 'dvt', 'loaiDon',
                    'maDonViGiao', 'nguoiGiao',
                    'maDonViNhan', 'nguoiPheDuyet', 'nguoiNhan',
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
  onDeleteFile
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

  const handleFile = useCallback((data, name) => {
    setLoading(true)
    setTimeout(() => {
      const parsed = parseXlsxToRows(data)
      if (onImportFile) {
        onImportFile(parsed, name, false)
      } else {
        setRows(parsed)
        setFileName(name)
      }
      setLoading(false)
    }, 80)
  }, [setRows, setFileName, onImportFile])

  const handleAppendFile = useCallback((data, name) => {
    setLoading(true)
    setTimeout(() => {
      const parsed = parseXlsxToRows(data)
      if (onImportFile) {
        onImportFile(parsed, name, true)
      } else {
        setRows(prev => [...prev, ...parsed])
        setFileName(prev => prev ? `${prev} + ${name}` : name)
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
          'STT', 'ngayXuatNhap', 'maVatTu', 'maSAP', 'dvt', 'loaiDon',
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
            <DataTable rows={filtered} setRows={setRows} type={type} />
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
function KhoDuAnTab({ chungRows, selectedProject, setSelectedProject, allProjects = [] }) {
  const [search, setSearch] = useState('')
  const [saveStatus, setSaveStatus] = useState('idle') // 'idle' | 'saving' | 'success' | 'error'
  const [errorMessage, setErrorMessage] = useState('')

  // Drag and drop / Custom classifications state
  const [customCategoryMap, setCustomCategoryMap] = useState({})
  const [dbCategoryMap, setDbCategoryMap] = useState({}) // Stores the initial/saved state from Supabase
  const [draggedItemName, setDraggedItemName] = useState(null)
  const [dragOverCol, setDragOverCol] = useState(null)

  // Load custom classifications from Supabase table phan_loai_don_vi
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
        console.error('Error loading custom classifications:', err)
      }
    }
    loadCustomClassifications()
    return () => {
      isMounted = false
    }
  }, [])

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
      // Upsert toàn bộ danh sách theo ten_don_vi (yêu cầu cột ten_don_vi có ràng buộc UNIQUE
      // trong Supabase). Cách này để Postgres tự xử lý insert-hoặc-update theo đúng ràng buộc,
      // tránh lỗi 409 do tự fetch rồi so sánh/update thủ công như trước đây.
      const payload = uniqueDonVi.map(item => ({
        ten_don_vi: item.name,
        nhom_don_vi: customCategoryMap[item.name] || getUnitCategory(item.name)
      }))

      if (payload.length > 0) {
        const chunkSize = 500
        for (let i = 0; i < payload.length; i += chunkSize) {
          const chunk = payload.slice(i, i + chunkSize)
          const { error: upsertErr } = await supabase
            .from('phan_loai_don_vi')
            .upsert(chunk, { onConflict: 'ten_don_vi' })
          if (upsertErr) throw upsertErr
        }
      }
      
      // Update our reference map representing the database
      setDbCategoryMap(prev => {
        const next = { ...prev }
        uniqueDonVi.forEach(item => {
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
  }, [uniqueDonVi, customCategoryMap])

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
  setConfigs,
  isSgcSummaryConfigsMissing,
  setIsSgcSummaryConfigsMissing 
}) {
  const [showCreateModal, setShowCreateModal] = React.useState(false)
  const [expandedConfig, setExpandedConfig] = React.useState(null) // index of expanded config
  const [configToDeleteIdx, setConfigToDeleteIdx] = React.useState(null)
  
  const filteredConfigs = selectedProject
    ? configs.filter(cfg => cfg.project === selectedProject)
    : []
  
  // Supabase sync states
  const [dbState, setDbState] = React.useState('idle') // 'idle' | 'loading' | 'success' | 'saving' | 'error'
  const [dbMessage, setDbMessage] = React.useState('')
  const lastSyncedRef = React.useRef(null)

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

  // Load configs from Supabase
  const fetchConfigsFromSupabase = React.useCallback(async () => {
    if (!isSupabaseConfigured) return
    setDbState('loading')
    setDbMessage('Đang tải cấu hình tổng hợp từ Supabase...')
    try {
      const { data, error } = await supabase
        .from('sgc_summary_configs')
        .select('*')
        .order('id', { ascending: true })

      if (error) {
        if (error.message.includes('Could not find the table') || error.message.includes('sgc_summary_configs')) {
          console.warn('Lỗi tải cấu hình tổng hợp từ Supabase: Bảng sgc_summary_configs chưa được tạo trong database Supabase.')
          if (setIsSgcSummaryConfigsMissing) setIsSgcSummaryConfigsMissing(true)
        } else {
          console.error('Lỗi tải cấu hình tổng hợp từ Supabase:', error.message)
        }
        setDbState('error')
        setDbMessage('Lỗi tải từ Supabase: ' + error.message)
        return
      }

      if (setIsSgcSummaryConfigsMissing) setIsSgcSummaryConfigsMissing(false)
      if (data) {
        const mapped = data.map(dbRow => ({
          id: dbRow.id,
          name: dbRow.name,
          project: dbRow.project,
          giaoTable: dbRow.giao_table || [],
          nhanTable: dbRow.nhan_table || [],
          bgColor: dbRow.bg_color || '#eff6ff',
        }))
        setConfigs(mapped)
        saveSummaryConfigs(mapped)
        lastSyncedRef.current = JSON.parse(JSON.stringify(mapped))
        setDbState('success')
        setDbMessage('Đồng bộ thành công từ Supabase')
        setTimeout(() => setDbMessage(''), 3000)
      }
    } catch (err) {
      console.error('Lỗi kết nối Supabase khi tải configs:', err)
      setDbState('error')
      setDbMessage('Lỗi kết nối Supabase')
    }
  }, [])

  // Initial load
  React.useEffect(() => {
    fetchConfigsFromSupabase()
  }, [fetchConfigsFromSupabase])

  // Save config update to Supabase
  const saveConfigToSupabase = async (cfg) => {
    if (!isSupabaseConfigured || !cfg || typeof cfg.id !== 'number') return
    setDbState('saving')
    setDbMessage('Đang tự động lưu cấu hình lên Supabase...')
    try {
      const payload = {
        name: cfg.name,
        project: cfg.project,
        giao_table: cfg.giaoTable,
        nhan_table: cfg.nhanTable,
      }

      let error
      // 1. Try updating with bg_color column first
      const res = await supabase
        .from('sgc_summary_configs')
        .update({ ...payload, bg_color: cfg.bgColor })
        .eq('id', cfg.id)

      if (res.error && (res.error.message.includes('bg_color') || res.error.message.includes('column') || res.error.message.includes('not exist'))) {
        // 2. Retry without bg_color if column does not exist
        const res2 = await supabase
          .from('sgc_summary_configs')
          .update(payload)
          .eq('id', cfg.id)
        error = res2.error
      } else {
        error = res.error
      }

      if (error) {
        console.error(`Lỗi cập nhật cấu hình ID=${cfg.id} trên Supabase:`, error.message)
        setDbState('error')
        setDbMessage('Không thể lưu lên Supabase: ' + error.message)
      } else {
        setDbState('success')
        setDbMessage('Đã đồng bộ tự động lên Supabase')
        setTimeout(() => setDbMessage(''), 2500)
      }
    } catch (err) {
      console.error('Lỗi khi thực hiện lưu cấu hình:', err)
      setDbState('error')
      setDbMessage('Lỗi liên kết cơ sở dữ liệu')
    }
  }

  // Monitor configs state to auto-save
  React.useEffect(() => {
    // Save to local storage as fallback/cache
    saveSummaryConfigs(configs)

    // Bypass check during initial loading to prevent overwrites
    if (dbState === 'loading') {
      return
    }

    if (lastSyncedRef.current === null) {
      lastSyncedRef.current = JSON.parse(JSON.stringify(configs))
      return
    }

    if (isSupabaseConfigured) {
      const lastSynced = lastSyncedRef.current
      configs.forEach(cfg => {
        const prev = lastSynced.find(c => c.id === cfg.id)
        if (!prev) return // Handled in creation or loaded just now
        
        if (JSON.stringify(prev) !== JSON.stringify(cfg)) {
          saveConfigToSupabase(cfg)
        }
      })
    }

    lastSyncedRef.current = JSON.parse(JSON.stringify(configs))
  }, [configs])

  const handleCreateConfig = async (name, proj, bgColor) => {
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

    // Add locally for instant UI response
    setConfigs(prev => [...prev, newConfig])
    setShowCreateModal(false)

    if (isSupabaseConfigured) {
      setDbState('saving')
      setDbMessage('Đang lưu cấu hình mới lên Supabase...')
      try {
        const payload = {
          name,
          project: proj,
          giao_table: newConfig.giaoTable,
          nhan_table: newConfig.nhanTable,
        }

        let data, error
        const res = await supabase
          .from('sgc_summary_configs')
          .insert([{ ...payload, bg_color: bgColor || '#eff6ff' }])
          .select()

        if (res.error && (res.error.message.includes('bg_color') || res.error.message.includes('column') || res.error.message.includes('not exist'))) {
          // Retry without bg_color if column is missing
          const res2 = await supabase
            .from('sgc_summary_configs')
            .insert([payload])
            .select()
          data = res2.data
          error = res2.error
        } else {
          data = res.data
          error = res.error
        }

        if (error) {
          console.error('Lỗi khi lưu cấu hình mới lên Supabase:', error.message)
          setDbState('error')
          setDbMessage('Lỗi tạo cấu hình Supabase: ' + error.message)
        } else if (data && data.length > 0) {
          const dbRow = data[0]
          setConfigs(prev => {
            const updated = prev.map(c => c.id === tempId ? {
              ...c,
              id: dbRow.id,
              bgColor: dbRow.bg_color || bgColor || '#eff6ff',
            } : c)
            
            // Expand newly created config
            const idx = updated.findIndex(c => c.id === dbRow.id)
            if (idx !== -1) setExpandedConfig(idx)
            return updated
          })
          setDbState('success')
          setDbMessage('Tạo cấu hình mới thành công')
          setTimeout(() => setDbMessage(''), 2500)
          return
        }
      } catch (err) {
        console.error('Lỗi insert Supabase:', err)
        setDbState('error')
        setDbMessage('Lỗi đồng bộ cấu hình mới')
      }
    }

    // Offline / fallback expand
    setConfigs(prev => {
      const idx = prev.findIndex(c => c.id === tempId)
      if (idx !== -1) setExpandedConfig(idx)
      return prev
    })
  }

  const handleDeleteConfig = async (idx) => {
    const cfg = configs[idx]
    const updated = configs.filter((_, i) => i !== idx)
    setConfigs(updated)
    if (expandedConfig === idx) setExpandedConfig(null)
    else if (expandedConfig > idx) setExpandedConfig(expandedConfig - 1)
    setConfigToDeleteIdx(null)

    if (isSupabaseConfigured && cfg && typeof cfg.id === 'number') {
      setDbState('saving')
      setDbMessage('Đang xóa cấu hình trên Supabase...')
      try {
        const { error } = await supabase
          .from('sgc_summary_configs')
          .delete()
          .eq('id', cfg.id)

        if (error) {
          console.error('Lỗi xóa cấu hình trên Supabase:', error.message)
          setDbState('error')
          setDbMessage('Lỗi xóa Supabase: ' + error.message)
        } else {
          setDbState('success')
          setDbMessage('Đã xóa thành công cấu hình trên Supabase')
          setTimeout(() => setDbMessage(''), 2500)
        }
      } catch (err) {
        console.error('Lỗi kết nối khi xóa cấu hình:', err)
        setDbState('error')
        setDbMessage('Lỗi liên kết dữ liệu khi xóa')
      }
    }
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
          {isSupabaseConfigured ? (
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginTop: 8, fontSize: 12 }}>
              <span style={{
                display: 'inline-block', width: 8, height: 8, borderRadius: '50%',
                background: dbState === 'error' ? '#ef4444' : dbState === 'saving' || dbState === 'loading' ? '#f59e0b' : '#10b981',
                boxShadow: `0 0 8px ${dbState === 'error' ? '#ef4444' : dbState === 'saving' || dbState === 'loading' ? '#f59e0b' : '#10b981'}`
              }} />
              <span style={{ color: '#475569', fontWeight: 500 }}>
                {dbMessage || 'Đồng bộ tự động với Supabase (sgc_summary_configs)'}
              </span>
              <button
                type="button"
                onClick={fetchConfigsFromSupabase}
                disabled={dbState === 'loading' || dbState === 'saving'}
                title="Tải lại từ Supabase"
                style={{
                  background: 'transparent', border: 'none', padding: '2px 4px', cursor: 'pointer',
                  color: '#0f58a7', fontSize: 11, display: 'flex', alignItems: 'center', gap: 2,
                  outline: 'none', fontWeight: 600
                }}
              >
                <RefreshCw size={11} className={(dbState === 'loading' || dbState === 'saving') ? 'animate-spin' : ''} />
                Làm mới
              </button>
            </div>
          ) : (
            <div style={{ display: 'flex', alignItems: 'center', gap: 6, marginTop: 8, fontSize: 12, color: '#f59e0b', fontWeight: 600 }}>
              <AlertCircle size={12} />
              <span>Chưa kết nối Supabase, dữ liệu lưu tạm ở trình duyệt</span>
            </div>
          )}
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

      {/* Missing table SQL installation card */}
      {isSgcSummaryConfigsMissing && (
        <div style={{
          background: '#fffbeb',
          border: '1px solid #fef3c7',
          borderRadius: 12,
          padding: '20px',
          marginBottom: '24px',
          boxShadow: '0 1px 3px rgba(0,0,0,0.05)',
        }}>
          <div style={{ display: 'flex', gap: '12px', alignItems: 'flex-start' }}>
            <AlertTriangle color="#d97706" size={24} style={{ flexShrink: 0, marginTop: '2px' }} />
            <div style={{ flex: 1 }}>
              <h3 style={{ margin: '0 0 8px', fontSize: 16, fontWeight: 700, color: '#92400e' }}>
                Thiếu bảng sgc_summary_configs trong Supabase
              </h3>
              <p style={{ margin: '0 0 16px', color: '#b45309', fontSize: 14, lineHeight: '1.6' }}>
                Hệ thống phát hiện cơ sở dữ liệu Supabase của bạn chưa được khởi tạo bảng cấu hình tổng hợp <code>sgc_summary_configs</code>. 
                Vui lòng sao chép câu lệnh SQL bên dưới, dán và chạy trong phần <strong>SQL Editor</strong> trên trang quản lý Supabase của bạn để kích hoạt đầy đủ tính năng đồng bộ trực tuyến:
              </p>
              
              <div style={{ position: 'relative', background: '#1e293b', borderRadius: '8px', padding: '16px', marginBottom: '16px' }}>
                <pre style={{
                  margin: 0, 
                  fontFamily: 'JetBrains Mono, monospace', 
                  fontSize: '12px', 
                  color: '#f8fafc', 
                  overflowX: 'auto',
                  whiteSpace: 'pre-wrap',
                  wordBreak: 'break-all'
                }}>
{`CREATE TABLE IF NOT EXISTS public.sgc_summary_configs (
  id bigint generated by default as identity primary key,
  name text not null,
  project text not null,
  giao_table jsonb default '[]'::jsonb,
  nhan_table jsonb default '[]'::jsonb,
  bg_color text default '#eff6ff',
  created_at timestamp with time zone default timezone('utc'::text, now()) not null
);

-- Bật Row Level Security (RLS)
ALTER TABLE public.sgc_summary_configs ENABLE ROW LEVEL SECURITY;

-- Tạo chính sách cho phép đọc ghi công khai
CREATE POLICY "Allow public read and write" ON public.sgc_summary_configs FOR ALL USING (true) WITH CHECK (true);`}
                </pre>
                <button
                  onClick={() => {
                    navigator.clipboard.writeText(`CREATE TABLE IF NOT EXISTS public.sgc_summary_configs (
  id bigint generated by default as identity primary key,
  name text not null,
  project text not null,
  giao_table jsonb default '[]'::jsonb,
  nhan_table jsonb default '[]'::jsonb,
  bg_color text default '#eff6ff',
  created_at timestamp with time zone default timezone('utc'::text, now()) not null
);

ALTER TABLE public.sgc_summary_configs ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Allow public read and write" ON public.sgc_summary_configs FOR ALL USING (true) WITH CHECK (true);`);
                    const btn = document.getElementById('copy-sql-btn');
                    if (btn) {
                      const oldTxt = btn.innerText;
                      btn.innerText = 'Đã sao chép!';
                      setTimeout(() => { btn.innerText = oldTxt; }, 3000);
                    }
                  }}
                  id="copy-sql-btn"
                  style={{
                    position: 'absolute',
                    top: '12px',
                    right: '12px',
                    background: '#334155',
                    color: '#f8fafc',
                    border: 'none',
                    borderRadius: '6px',
                    padding: '6px 12px',
                    fontSize: '12px',
                    fontWeight: 600,
                    cursor: 'pointer',
                    transition: 'background 0.2s'
                  }}
                  onMouseOver={(e) => e.currentTarget.style.background = '#475569'}
                  onMouseOut={(e) => e.currentTarget.style.background = '#334155'}
                >
                  Sao chép SQL
                </button>
              </div>
              
              <p style={{ margin: 0, color: '#b45309', fontSize: 13, fontStyle: 'italic' }}>
                * Lưu ý: Khi chưa tạo bảng này, mọi thay đổi cấu hình vẫn hoạt động bình thường và được tự động lưu tạm tại trình duyệt này (LocalStorage).
              </p>
            </div>
          </div>
        </div>
      )}

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

// ─── Inventory Report (Báo cáo xuất nhập tồn) ──────────────────────────────────
function BaoCaoXuatNhapTonTab({ chungRows = [], giaoRows = [], nhanRows = [], selectedProject, setSelectedProject, allProjects = [] }) {
  const [localProject, setLocalProject] = React.useState(selectedProject || '')
  const [searchTerm, setSearchTerm] = React.useState('')
  const [statusFilter, setStatusFilter] = React.useState('approved_only') // 'approved_only' | 'all'
  const [currentPage, setCurrentPage] = React.useState(1)
  const [pageSize, setPageSize] = React.useState(50)
  const [sortField, setSortField] = React.useState('maSAP') // 'maSAP' | 'received' | 'issued' | 'stock'
  const [sortDirection, setSortDirection] = React.useState('asc') // 'asc' | 'desc'

  // Synchronize localProject with selectedProject prop when it changes
  React.useEffect(() => {
    if (selectedProject !== localProject) {
      setLocalProject(selectedProject || '')
    }
  }, [selectedProject])

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

  // Process rows and group by maSAP
  const reportData = React.useMemo(() => {
    const groups = {}
    const projLower = localProject ? localProject.trim().toLowerCase() : ''

    // Helper to parse double values safely
    const parseVal = (val) => {
      if (val === null || val === undefined) return 0
      if (typeof val === 'number') return val
      const cleaned = String(val).replace(/[^\d.-]/g, '').replace(',', '.')
      const num = parseFloat(cleaned)
      return isNaN(num) ? 0 : num
    }

    const sourceRows = (chungRows && chungRows.length > 0) 
      ? chungRows 
      : [...giaoRows, ...nhanRows]

    sourceRows.forEach(r => {
      // Apply status filter
      if (statusFilter === 'approved_only' && !isApprovedStatus(r.trangThai)) {
        return
      }

      const nhanUnit = String(r.donViNhan || '').trim().toLowerCase()
      const giaoUnit = String(r.donViGiao || '').trim().toLowerCase()

      let isNhan = false
      let isGiao = false

      if (localProject) {
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
      if (!sap) return // Skip rows without SAP code as per requirement: "thống kê báo cáo với Mã SAP"

      // Initialize group if not exists
      if (!groups[sap]) {
        groups[sap] = {
          maSAP: sap,
          maVatTu: String(r.maVatTu || '').trim(),
          tenVatTu: String(r.tenVatTu || '').trim(),
          dvt: String(r.dvt || '').trim(),
          thongSoKyThuat: String(r.thongSoKyThuat || '').trim(),
          received: 0,
          issued: 0
        }
      } else {
        // Fallbacks for empty info
        if (!groups[sap].maVatTu && r.maVatTu) groups[sap].maVatTu = String(r.maVatTu).trim()
        if (!groups[sap].tenVatTu && r.tenVatTu) groups[sap].tenVatTu = String(r.tenVatTu).trim()
        if (!groups[sap].dvt && r.dvt) groups[sap].dvt = String(r.dvt).trim()
        if (!groups[sap].thongSoKyThuat && r.thongSoKyThuat) groups[sap].thongSoKyThuat = String(r.thongSoKyThuat).trim()
      }

      // Add to received/issued
      if (localProject) {
        if (isNhan) {
          // Receiving unit: add to received
          groups[sap].received += parseVal(r.khoiLuongNhap) || parseVal(r.khoiLuongXuat)
        }
        if (isGiao) {
          // Issuing/delivery unit: add to issued
          groups[sap].issued += parseVal(r.khoiLuongXuat) || parseVal(r.khoiLuongNhap)
        }
      } else {
        // Global aggregation
        groups[sap].received += parseVal(r.khoiLuongNhap)
        groups[sap].issued += parseVal(r.khoiLuongXuat)
      }
    })

    // Map groups to list and compute stock/inventory
    const result = Object.values(groups).map(g => {
      const stock = g.received - g.issued
      return {
        ...g,
        stock
      }
    })

    // Filter by search term
    const s = searchTerm.trim().toLowerCase()
    const filtered = s
      ? result.filter(item => 
          item.maSAP.toLowerCase().includes(s) ||
          item.maVatTu.toLowerCase().includes(s) ||
          item.tenVatTu.toLowerCase().includes(s) ||
          item.thongSoKyThuat.toLowerCase().includes(s)
        )
      : result

    // Sort
    filtered.sort((a, b) => {
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
      }

      if (valA < valB) return sortDirection === 'asc' ? -1 : 1
      if (valA > valB) return sortDirection === 'asc' ? 1 : -1
      return 0
    })

    return filtered
  }, [chungRows, giaoRows, nhanRows, localProject, searchTerm, statusFilter, sortField, sortDirection])

  // Metrics
  const metrics = React.useMemo(() => {
    let totalReceived = 0
    let totalIssued = 0
    let totalStock = 0
    reportData.forEach(item => {
      totalReceived += item.received
      totalIssued += item.issued
      totalStock += item.stock
    })
    return {
      totalReceived,
      totalIssued,
      totalStock,
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
    if (reportData.length === 0) return

    const wb = XLSXStyle.utils.book_new()
    const ws = {}

    // Set columns widths
    const cols = [
      { key: 'STT', label: 'STT', width: 60 },
      { key: 'maSAP', label: 'Mã SAP', width: 120 },
      { key: 'maVatTu', label: 'Mã vật tư', width: 120 },
      { key: 'tenVatTu', label: 'Tên vật tư', width: 250 },
      { key: 'dvt', label: 'ĐVT', width: 80 },
      { key: 'thongSoKyThuat', label: 'Thông số kỹ thuật', width: 200 },
      { key: 'received', label: 'Khối lượng nhận', width: 150 },
      { key: 'issued', label: 'Khối lượng xuất', width: 150 },
      { key: 'stock', label: 'Khối lượng tồn kho', width: 150 }
    ]
    ws['!cols'] = cols.map(c => ({ wpx: c.width }))

    // Title Row
    ws['A1'] = {
      v: `BÁO CÁO XUẤT NHẬP TỒN VẬT TƯ THIẾT BỊ`,
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
      v: `Trạng thái dữ liệu: ${statusFilter === 'approved_only' ? 'Chỉ tính đơn đã phê duyệt' : 'Tính tất cả các đơn'}`,
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

    // Write Data rows
    reportData.forEach((item, idx) => {
      rowIdx++
      const cells = [
        idx + 1,
        item.maSAP,
        item.maVatTu,
        item.tenVatTu,
        item.dvt,
        item.thongSoKyThuat,
        item.received,
        item.issued,
        item.stock
      ]

      cells.forEach((val, colIdx) => {
        const cellRef = `${String.fromCharCode(65 + colIdx)}${rowIdx}`
        const isNum = typeof val === 'number'

        ws[cellRef] = {
          v: val,
          t: isNum ? 'n' : 's',
          s: {
            font: { name: 'Segoe UI', sz: 9.5 },
            alignment: { 
              horizontal: colIdx === 3 || colIdx === 5 ? 'left' : (isNum ? 'right' : 'center'), 
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

        // Apply number format
        if (isNum && colIdx >= 6) {
          ws[cellRef].z = '#,##0.000'
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
      z: '#,##0.000',
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
      z: '#,##0.000',
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
      z: '#,##0.000',
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

    // Set sheet range bounds
    ws['!ref'] = `A1:I${rowIdx}`

    XLSXStyle.utils.book_append_sheet(wb, ws, "Xuat_Nhap_Ton")
    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary' })

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
    a.download = `Bao_Cao_Xuat_Nhap_Ton_${(localProject || 'all').replace(/\s+/g, '_')}.xlsx`
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)
    URL.revokeObjectURL(url)
  }

  const handleWarehouseChange = (e) => {
    const val = e.target.value
    setLocalProject(val)
    setSelectedProject(val) // Sync to global selectedProject
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
      value: metrics.totalReceived.toLocaleString('vi-VN', { maximumFractionDigits: 3 }),
      color: '#10b981',
      bg: '#ecfdf5',
      border: '#a7f3d0',
      icon: <Truck size={18} />
    },
    {
      label: 'Tổng khối lượng xuất',
      value: metrics.totalIssued.toLocaleString('vi-VN', { maximumFractionDigits: 3 }),
      color: '#f97316',
      bg: '#fff7ed',
      border: '#fed7aa',
      icon: <Download size={18} style={{ transform: 'rotate(180deg)' }} />
    },
    {
      label: 'Khối lượng tồn kho',
      value: metrics.totalStock.toLocaleString('vi-VN', { maximumFractionDigits: 3 }),
      color: metrics.totalStock >= 0 ? 'var(--primary)' : '#ef4444',
      bg: metrics.totalStock >= 0 ? 'var(--primary-light)' : '#fef2f2',
      border: metrics.totalStock >= 0 ? 'var(--border)' : '#fca5a5',
      icon: <Database size={18} />
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
    <div id="baocao-xuat-nhap-ton-root" style={{ padding: '16px 24px 24px 24px', height: '100%', display: 'flex', flexDirection: 'column', boxSizing: 'border-box', overflow: 'hidden' }}>
      {/* Title Header Card */}
      <div style={{
        background: '#ffffff',
        borderRadius: 8,
        border: '1px solid var(--border)',
        padding: '14px 20px',
        marginBottom: 12,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'space-between',
        boxShadow: 'var(--shadow-sm)',
        flexShrink: 0
      }}>
        <div>
          <h2 style={{ fontSize: 18, fontWeight: 800, color: 'var(--text)', margin: 0 }}>
            Báo cáo xuất nhập tồn vật tư thiết bị
          </h2>
          <p style={{ color: 'var(--text-muted)', fontSize: 13, margin: '4px 0 0 0' }}>
            Báo cáo chi tiết theo Mã SAP về khối lượng nhập, khối lượng xuất và tồn kho tương ứng của mỗi dự án.
          </p>
        </div>
        {reportData.length > 0 && (
          <button
            id="btn-export-inventory"
            onClick={handleExportExcel}
            className="btn btn-success"
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: 6,
              height: 36,
              padding: '0 16px',
              borderRadius: 6,
              fontWeight: 600,
              fontSize: '13px',
              background: '#10b981',
              color: '#ffffff',
              border: 'none',
              cursor: 'pointer',
              boxShadow: '0 2px 4px rgba(16, 185, 129, 0.2)'
            }}
          >
            <Download size={14} />
            <span>Xuất Excel</span>
          </button>
        )}
      </div>

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
          <select
            id="select-inventory-project"
            value={localProject}
            onChange={handleWarehouseChange}
            className="input"
            style={{ height: 38, flex: 1, padding: '0 10px', fontSize: 13, minWidth: 180, border: '1px solid #cbd5e1' }}
          >
            <option value="">-- Tất cả Kho / Dự án --</option>
            {uniqueWarehouses.map(w => (
              <option key={w} value={w}>{w}</option>
            ))}
          </select>
        </div>

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
      </div>

      {/* Summary Metrics Cards */}
      <div style={{ display: 'flex', gap: 12, marginBottom: 12, flexWrap: 'wrap', flexShrink: 0 }}>
        {metricsCards.map(c => (
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

      {/* Table container / main body */}
      {reportData.length === 0 ? (
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
                    style={{ width: 130, minWidth: 130, cursor: 'pointer', fontSize: '12px', padding: '8px 10px', textAlign: 'left', verticalAlign: 'middle' }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 4 }}>
                      <span>Mã SAP</span>
                      <ArrowUpDown size={12} style={{ opacity: 0.7 }} />
                    </div>
                  </th>
                  <th style={{ width: 130, minWidth: 130, fontSize: '12px', padding: '8px 10px', textAlign: 'left', verticalAlign: 'middle' }}>
                    Mã vật tư
                  </th>
                  <th style={{ minWidth: 200, fontSize: '12px', padding: '8px 10px', textAlign: 'left', verticalAlign: 'middle' }}>
                    Tên vật tư
                  </th>
                  <th style={{ width: 80, minWidth: 80, textAlign: 'center', fontSize: '12px', padding: '8px 10px', verticalAlign: 'middle' }}>
                    ĐVT
                  </th>
                  <th style={{ minWidth: 150, fontSize: '12px', padding: '8px 10px', textAlign: 'left', verticalAlign: 'middle' }}>
                    Thông số kỹ thuật
                  </th>
                  <th 
                    onClick={() => toggleSort('received')}
                    style={{ width: 150, minWidth: 150, cursor: 'pointer', fontSize: '12px', padding: '8px 10px', textAlign: 'right', verticalAlign: 'middle' }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 4, justifyContent: 'flex-end' }}>
                      <span>Khối lượng nhận</span>
                      <ArrowUpDown size={12} style={{ opacity: 0.7 }} />
                    </div>
                  </th>
                  <th 
                    onClick={() => toggleSort('issued')}
                    style={{ width: 150, minWidth: 150, cursor: 'pointer', fontSize: '12px', padding: '8px 10px', textAlign: 'right', verticalAlign: 'middle' }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 4, justifyContent: 'flex-end' }}>
                      <span>Khối lượng xuất</span>
                      <ArrowUpDown size={12} style={{ opacity: 0.7 }} />
                    </div>
                  </th>
                  <th 
                    onClick={() => toggleSort('stock')}
                    style={{ width: 160, minWidth: 160, cursor: 'pointer', fontSize: '12px', padding: '8px 10px', textAlign: 'right', verticalAlign: 'middle' }}
                  >
                    <div style={{ display: 'flex', alignItems: 'center', gap: 4, justifyContent: 'flex-end' }}>
                      <span>Tồn kho</span>
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
                      <td style={{ width: 130, minWidth: 130, fontSize: '13px', fontWeight: 700, fontFamily: 'monospace', color: 'var(--text)', padding: '6px 10px' }}>
                        {item.maSAP}
                      </td>
                      <td style={{ width: 130, minWidth: 130, fontSize: '13px', fontFamily: 'monospace', color: 'var(--text-muted)', padding: '6px 10px' }}>
                        {item.maVatTu || '—'}
                      </td>
                      <td style={{ minWidth: 200, fontSize: '13px', fontWeight: 600, color: 'var(--text)', whiteSpace: 'normal', wordBreak: 'break-word', padding: '6px 10px' }}>
                        {item.tenVatTu || '—'}
                      </td>
                      <td style={{ width: 80, minWidth: 80, textAlign: 'center', padding: '6px 10px' }}>
                        <span className="badge badge-gray" style={{ fontSize: '11px', padding: '2px 6px' }}>{item.dvt || '—'}</span>
                      </td>
                      <td style={{ minWidth: 150, fontSize: '13px', color: 'var(--text-muted)', whiteSpace: 'normal', wordBreak: 'break-word', padding: '6px 10px' }}>
                        {item.thongSoKyThuat || '—'}
                      </td>
                      <td style={{ width: 150, minWidth: 150, textAlign: 'right', fontSize: '13px', fontWeight: 700, color: '#10b981', padding: '6px 10px' }}>
                        {item.received > 0 ? item.received.toLocaleString('vi-VN', { maximumFractionDigits: 3 }) : '0'}
                      </td>
                      <td style={{ width: 150, minWidth: 150, textAlign: 'right', fontSize: '13px', fontWeight: 700, color: '#f97316', padding: '6px 10px' }}>
                        {item.issued > 0 ? item.issued.toLocaleString('vi-VN', { maximumFractionDigits: 3 }) : '0'}
                      </td>
                      <td style={{
                        width: 160, minWidth: 160, textAlign: 'right', fontSize: '13px', fontWeight: 800,
                        color: isNegative ? '#ef4444' : (isZero ? 'var(--text-muted)' : 'var(--primary)'),
                        background: isNegative ? '#fef2f2' : (isZero ? 'transparent' : 'var(--primary-light)'),
                        padding: '6px 10px'
                      }}>
                        {item.stock.toLocaleString('vi-VN', { maximumFractionDigits: 3 })}
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
          numFormat = '#,##0.000'
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
          numFormat = '#,##0.000'
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

    const wbout = XLSXStyle.write(wb, { bookType: 'xlsx', type: 'binary' })
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
    const str = val.toLocaleString('vi-VN', { minimumFractionDigits: 0, maximumFractionDigits: 3 })
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
                {totalNhap.toLocaleString('vi-VN', { maximumFractionDigits: 3 })}
              </div>
            </div>

            {/* Stat Item 3: Consolidated Export Volume */}
            <div style={{ borderLeft: `3px solid ${totalXuat >= 0 ? '#3b82f6' : '#ef4444'}`, paddingLeft: 8 }}>
              <div style={{ fontSize: 10.5, color: '#64748b', fontWeight: 500, marginBottom: 1 }}>Khối lượng Xuất ròng</div>
              <div style={{ fontSize: 14.5, fontWeight: 700, color: totalXuat < 0 ? '#dc2626' : '#3b82f6' }}>
                {totalXuat.toLocaleString('vi-VN', { maximumFractionDigits: 3 })}
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
  const [selectedProject, setSelectedProject] = useState('')
  const [showAddProjectModal, setShowAddProjectModal] = useState(false)
  const [showEditProjectModal, setShowEditProjectModal] = useState(false)
  const [showDeleteFileModal, setShowDeleteFileModal] = useState(false)
  const [deleteFileType, setDeleteFileType] = useState(null) // 'giao' | 'nhan'
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

  const [chungRows, setChungRows] = useState([])
  const [chungFileName, setChungFileName] = useState('')

  const [khoRows, setKhoRows] = useState([])
  const [khoFileName, setKhoFileName] = useState('')

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

  const [syncingType, setSyncingType] = useState(null) // 'giao' | 'nhan' | null
  const [supabaseMessage, setSupabaseMessage] = useState(null) // { text, type: 'success' | 'error' | 'info' }

  const [configs, setConfigs] = useState(() => loadSummaryConfigs())
  const [isSgcSummaryConfigsMissing, setIsSgcSummaryConfigsMissing] = useState(false)
  const [isInitialLoading, setIsInitialLoading] = useState(true)

  const fetchConfigsFromSupabaseInApp = React.useCallback(async () => {
    if (!isSupabaseConfigured) return
    try {
      const { data, error } = await supabase
        .from('sgc_summary_configs')
        .select('*')
        .order('id', { ascending: true })

      if (error) {
        if (error.message.includes('Could not find the table') || error.message.includes('sgc_summary_configs')) {
          console.warn('Lỗi tải cấu hình tổng hợp từ Supabase (App): Bảng sgc_summary_configs chưa được tạo trong database Supabase.')
          setIsSgcSummaryConfigsMissing(true)
        } else {
          console.error('Lỗi tải cấu hình tổng hợp từ Supabase (App):', error.message)
        }
        return
      }

      setIsSgcSummaryConfigsMissing(false)
      if (data) {
        const mapped = data.map(dbRow => ({
          id: dbRow.id,
          name: dbRow.name,
          project: dbRow.project,
          giaoTable: dbRow.giao_table || [],
          nhanTable: dbRow.nhan_table || [],
          bgColor: dbRow.bg_color || '#eff6ff',
        }))
        setConfigs(mapped)
        saveSummaryConfigs(mapped)
      }
    } catch (err) {
      console.error('Lỗi kết nối Supabase khi tải configs (App):', err)
    }
  }, [])

  // Danh sách columns cần fetch (bỏ qua metadata Supabase như created_at, updated_at)
  const SELECT_COLS = 'id,ngay_xuat_nhap,ma_vat_tu,ma_s_a_p,thong_so_ky_thuat,ten_vat_tu,dvt,loai_don,ma_don_nhap_kho,ma_don_xuat_kho,khoi_luong_nhap,ma_don_vi_giao,don_vi_giao,nguoi_giao,khoi_luong_xuat,ma_don_vi_nhan,don_vi_nhan,nguoi_phe_duyet,ten_nguon,ma_nguon,lo,hang_muc,so_hop_dong,thu_kho,bien_so_xe,phan_khu,du_an,tinh_trang,nguoi_nhan,ma_don_lien_quan,nha_cung_cap,ma_don_chuyen_tiep_l_c,ma_don_chuyen_tiep_n_b,ghi_chu,ghi_chu_vat_tu,trang_thai,nhan_hieu,ten_du_an,tenDuAn'

  const loadProjectsFromSupabase = React.useCallback(async () => {
    if (!isSupabaseConfigured) return
    try {
      // Select * để tự detect cột tên dự án bất kể snake_case hay camelCase
      const { data: projData, error: projError } = await supabase
        .from('du_an')
        .select('*')

      if (projError) {
        if (projError.message && (projError.message.includes('Could not find the table') || projError.message.includes('du_an'))) {
          // Bỏ qua lỗi thiếu bảng du_an
        } else {
          console.error('Lỗi tải danh mục dự án:', projError.message || projError)
        }
        if (projError.status === 401) setSupabaseAuthError(true)
        return
      }
      setSupabaseAuthError(false)
      if (projData && projData.length > 0) {
        console.log('[loadProjects] Raw Supabase du_an columns:', Object.keys(projData[0]))
        // Tự detect cột tên dự án theo thứ tự ưu tiên
        const NAME_COLS = ['ten_du_an', 'tenduan', 'tenDuAn', 'ten_duan', 'name', 'tendu_an']
        const detectedCol = NAME_COLS.find(col => col in projData[0])
        if (detectedCol) {
          console.log('[loadProjects] Dùng cột:', detectedCol)
          const list = projData.map(d => d[detectedCol]).filter(Boolean).sort()
          setCustomProjects(list)
          localStorage.setItem('sgc_custom_projects', JSON.stringify(list))
        } else {
          // Fallback: thử tất cả giá trị string đầu tiên tìm được
          console.warn('[loadProjects] Không tìm thấy cột tên quen thuộc, thử fallback...')
          const firstRow = projData[0]
          const firstStrKey = Object.keys(firstRow).find(k => k !== 'id' && typeof firstRow[k] === 'string')
          if (firstStrKey) {
            const list = projData.map(d => d[firstStrKey]).filter(Boolean).sort()
            console.log('[loadProjects] Fallback dùng cột:', firstStrKey, '→', list)
            setCustomProjects(list)
            localStorage.setItem('sgc_custom_projects', JSON.stringify(list))
          }
        }
      } else if (projData && projData.length === 0) {
        // Bảng du_an rỗng — giữ nguyên localStorage
      }
    } catch (e) {
      console.error('Lỗi tải dự án:', e)
    }
  }, [])

  const loadTableFromSupabase = React.useCallback(async (tableType, forceBypassCache = false) => {
    if (!isSupabaseConfigured) return
    if (tableType !== 'chung') return // Skip deleted tables
    const tableName = 'don_chung'

    if (!forceBypassCache) {
      try {
        // 1. Fetch exact total row count on Supabase (uses almost zero bandwidth/egress since there is no data body)
        const { count, error: countErr } = await supabase
          .from(tableName)
          .select('*', { count: 'exact', head: true })

        if (!countErr && count !== null) {
          // 2. Fetch highest record ID (uses almost zero bandwidth/egress)
          const { data: maxIdData, error: maxIdErr } = await supabase
            .from(tableName)
            .select('id')
            .order('id', { ascending: false })
            .limit(1)

          const maxRemoteId = maxIdData && maxIdData.length > 0 ? Number(maxIdData[0].id) : 0

          // Get our current cached rows
          const cachedKey = tableType === 'chung' ? 'sgc_chung_rows' : tableType === 'giao' ? 'sgc_giao_rows' : tableType === 'nhan' ? 'sgc_nhan_rows' : 'sgc_kho_rows'
          const cachedJson = await idbGet(cachedKey)
          if (Array.isArray(cachedJson)) {
            const maxLocalId = cachedJson.length > 0 ? Math.max(...cachedJson.map(r => Number(r.id) || 0)) : 0

            // If count and high ID match, we can skip fetching from Supabase!
            if (count === cachedJson.length && maxRemoteId === maxLocalId) {
              console.log(`[Cache Tối Ưu] ${tableName} trùng khớp dữ liệu (Egress = 0). Dùng cache gốc.`);
              if (tableType === 'giao') {
                setGiaoRows(cachedJson)
                setGiaoFileName(cachedJson.length > 0 ? 'Report_Orders_Don_giao (Cached DB)' : '')
              } else if (tableType === 'nhan') {
                setNhanRows(cachedJson)
                setNhanFileName(cachedJson.length > 0 ? 'Report_Orders_Don_nhan (Cached DB)' : '')
              } else if (tableType === 'kho') {
                setKhoRows(cachedJson)
                setKhoFileName(cachedJson.length > 0 ? 'Report_Orders_Kho_du_an (Cached DB)' : '')
              } else {
                setChungRows(cachedJson)
                setChungFileName(cachedJson.length > 0 ? 'Report_Orders_Don_chung (Cached DB)' : '')
              }
              return
            }
          }
        }
      } catch (cacheErr) {
        console.warn(`[Cache Check Bypass] Lỗi kiểm tra trùng khớp cho ${tableName}:`, cacheErr)
      }
    }

    // Cache miss or force reload triggered -> Fetch fresh rows
    console.log(`[Tải DB] Bắt đầu tải mới từ Supabase cho bảng "${tableName}"...`)
    const PAGE_SIZE = 1000
    let allData = []
    let from = 0
    let hasMore = true
    let errorObj = null

    try {
      while (hasMore) {
        const { data: page, error } = await supabase
          .from(tableName)
          .select('*')
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

      if (errorObj) {
        if (errorObj.status === 401) setSupabaseAuthError(true)
        return
      }

      const mapped = allData.map((item, idx) => {
        const normalized = normalizeDbRow(item)
        return { id: normalized.id || idx, ...normalized }
      }).sort((a, b) => Number(a.id) - Number(b.id))

      if (tableType === 'giao') {
        setGiaoRows(mapped)
        setGiaoFileName(mapped.length > 0 ? 'Report_Orders_Don_giao (Supabase DB)' : '')
      } else if (tableType === 'nhan') {
        setNhanRows(mapped)
        setNhanFileName(mapped.length > 0 ? 'Report_Orders_Don_nhan (Supabase DB)' : '')
      } else if (tableType === 'kho') {
        setKhoRows(mapped)
        setKhoFileName(mapped.length > 0 ? 'Report_Orders_Kho_du_an (Supabase DB)' : '')
      } else {
        setChungRows(mapped)
        setChungFileName(mapped.length > 0 ? 'Report_Orders_Don_chung (Supabase DB)' : '')
      }
    } catch (e) {
      console.error(`Lỗi tải bảng ${tableName}:`, e)
    }
  }, [])

  const loadDataFromSupabase = React.useCallback(async (forceBypassCache = false) => {
    if (!isSupabaseConfigured) return
    await Promise.all([
      loadProjectsFromSupabase(),
      loadTableFromSupabase('chung', forceBypassCache),
      fetchConfigsFromSupabaseInApp(),
    ])
  }, [loadProjectsFromSupabase, loadTableFromSupabase, fetchConfigsFromSupabaseInApp])

  // Fetch initial data from Supabase if connected
  React.useEffect(() => {
    if (!isSupabaseConfigured) {
      setIsInitialLoading(false)
    } else {
      loadDataFromSupabase().finally(() => setIsInitialLoading(false))
    }

    if (!isSupabaseConfigured) return

    // Realtime: CHỈ lắng nghe bảng du_an (rất nhẹ, vài row)
    // KHÔNG lắng nghe don_giao/don_nhan — tránh fetch lại hàng nghìn row mỗi lần có thay đổi
    // Dữ liệu don_giao/don_nhan được cập nhật qua local state ngay sau khi sync/xóa
    const debouncedLoadProjects = () => {
      if (syncInProgressRef.current) return
      if (realtimeDebounceRef.current) clearTimeout(realtimeDebounceRef.current)
      realtimeDebounceRef.current = setTimeout(() => {
        loadProjectsFromSupabase()
      }, 3000)
    }

    const channel = supabase
      .channel('du_an-changes')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'du_an' }, () => {
        if (syncInProgressRef.current) return
        console.log('[Realtime] Cập nhật danh sách Dự Án...')
        debouncedLoadProjects()
      })
      .subscribe()

    return () => {
      if (realtimeDebounceRef.current) clearTimeout(realtimeDebounceRef.current)
      supabase.removeChannel(channel)
    }
  }, [loadDataFromSupabase, loadProjectsFromSupabase])

  const syncRowsToSupabase = async (type, rowsToSync, isAuto = false, isAppend = false) => {
    if (!isSupabaseConfigured) return
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
      if (isSupabaseConfigured) {
        await syncRowsToSupabase('giao', rowsToStore, true, isAppend)
      }
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
      if (isSupabaseConfigured) {
        await syncRowsToSupabase('nhan', rowsToStore, true, isAppend)
      }
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
      if (isSupabaseConfigured) {
        await syncRowsToSupabase('kho', rowsToStore, true, isAppend)
      }
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

      // 4. Đồng bộ lên Supabase nếu có kết nối
      if (isSupabaseConfigured) {
        await syncRowsToSupabase('chung', rowsToStore, true, isAppend)
        if (extractedGiaoRows.length > 0) {
          await syncRowsToSupabase('giao', extractedGiaoRows, true, isAppend)
        }
        if (extractedNhanRows.length > 0) {
          await syncRowsToSupabase('nhan', extractedNhanRows, true, isAppend)
        }
      }
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

    // 2. Nếu đã kết nối Supabase, xóa các dòng tương ứng trên Supabase
    if (isSupabaseConfigured) {
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

      // 2. Đồng bộ lên Supabase với fallback cột, rồi reload lại đúng danh sách
      if (isSupabaseConfigured) {
        let insertOk = false
        const colVariants = ['ten_du_an', 'tenduan', 'tenDuAn']
        for (const col of colVariants) {
          const { error } = await supabase.from('du_an').insert([{ [col]: trimmed }])
          if (!error) {
            insertOk = true
            console.log('[AddProject] Lưu dự án "' + trimmed + '" thành công với cột "' + col + '"')
            break
          } else {
            const msg = error.message || ''
            const isColMissing = msg.includes('column') || msg.includes('not exist') || msg.includes('schema cache') || msg.includes('violates')
            if (!isColMissing) {
              console.error('[AddProject] Lỗi khi insert dự án:', error)
              setSupabaseMessage({ text: 'Đã tạo dự án cục bộ nhưng lỗi lưu Supabase: ' + msg, type: 'error' })
              setTimeout(() => setSupabaseMessage(null), 5000)
              break
            }
          }
        }
        // 3. Reload lại danh sách từ Supabase để đảm bảo hiển thị đúng
        if (insertOk) {
          await loadProjectsFromSupabase()
          setSupabaseMessage({ text: 'Tạo kho dự án "' + trimmed + '" và đồng bộ Supabase thành công!', type: 'success' })
          setTimeout(() => setSupabaseMessage(null), 3000)
        }
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
      try {
        setSupabaseMessage({
          text: `Đang cập nhật tên dự án trên Supabase...`,
          type: 'info'
        })

        const errors = []

        // Helper: thử update một bảng với nhiều kiểu cột
        const updateTable = async (table, possibleCols) => {
          for (const col of possibleCols) {
            const { error } = await supabase
              .from(table)
              .update({ [col]: trimmedNew })
              .eq(col, trimmedOld)
            if (!error) return true // thành công
            // Nếu lỗi do cột không tồn tại → thử cột tiếp theo
            const msg = error.message || ''
            const isColMissing = msg.includes('column') || msg.includes('not exist') || msg.includes('schema cache')
            if (!isColMissing) {
              // Lỗi thực sự (quyền, constraint...) → báo lỗi ngay
              errors.push(`[${table}] ${msg}`)
              return false
            }
          }
          // Không có cột nào khớp → coi như không có gì cần update (bảng dùng cột khác)
          return true
        }

        // Cập nhật bảng du_an
        await updateTable('du_an', ['ten_du_an', 'tenduan', 'tenDuAn'])

        // Cập nhật bảng don_chung
        await updateTable('don_chung', ['ten_du_an', 'tenDuAn', 'tenduan'])

        if (errors.length > 0) {
          setSupabaseMessage({
            text: `Cập nhật cục bộ thành công! Một số lỗi Supabase: ${errors.join(' | ')}`,
            type: 'error'
          })
        } else {
          setSupabaseMessage({
            text: `Đã đổi tên dự án thành "${trimmedNew}" và đồng bộ toàn bộ dữ liệu lên Supabase!`,
            type: 'success'
          })
        }
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
      try {
        setSupabaseMessage({
          text: `Đang kết nối để xóa kho dự án "${trimmed}" trên Supabase...`,
          type: 'info'
        })

        // Xóa tuần tự: Xóa các tham chiếu ở con trước để tránh Foreign Key Violation, sau đó mới xóa ở cha
        console.log('[Supabase Delete] Thực hiện xóa liên kết don_chung...')
        const resChung = await deleteFromTableAdaptive('don_chung', ['ten_du_an', 'tenDuAn', 'tenduan'], trimmed)
        if (!resChung.success && resChung.reason !== 'table_empty') {
          console.warn('[Supabase Delete] Cảnh báo lỗi xóa don_chung:', resChung.error)
        }

        console.log('[Supabase Delete] Thực hiện xóa dự án trong bảng du_an...')
        const resDuAn = await deleteFromTableAdaptive('du_an', ['ten_du_an', 'tenduan', 'tenDuAn', 'ten_duan', 'tendu_an', 'name'], trimmed)
        if (!resDuAn.success) {
          const errObj = resDuAn.error || {}
          const errorMsg = errObj.message || errObj.details || 'Không rõ nguyên nhân. Vui lòng kiểm tra quyền RLS (Row Level Security) cho chính sách DELETE hoặc ràng buộc khóa ngoại (Foreign Key Constraints).'
          throw new Error(errorMsg)
        }

        // Chỉ reload danh sách dự án (nhẹ), KHÔNG reload toàn bộ don_giao/don_nhan
        await loadProjectsFromSupabase()

        setSupabaseMessage({
          text: `Xóa kho dự án "${trimmed}" thành công và đã đồng bộ!`,
          type: 'success'
        })
        setTimeout(() => setSupabaseMessage(null), 4000)
      } catch (err) {
        console.error('Lỗi khi xóa dự án trên Supabase:', err)
        await loadProjectsFromSupabase().catch(() => {})
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
    { id: 'inventory', label: 'Báo cáo xuất nhập tồn', icon: <Database size={15} /> },
  ]

  return (
    <div style={{ height: '100vh', display: 'flex', flexDirection: 'column', overflow: 'hidden', position: 'relative' }}>
      {isInitialLoading && (
        <div style={{
          position: 'fixed',
          inset: 0,
          zIndex: 9999,
          background: '#f8fafc',
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
            <h2 style={{ fontSize: 18, fontWeight: 800, color: 'var(--text)', margin: '0 0 6px' }}>
              SGC | BÁO CÁO GIAO NHẬN
            </h2>
            <p style={{ color: 'var(--text-muted)', fontSize: 14, margin: 0 }}>
              Đang tải dữ liệu từ Supabase, vui lòng chờ trong giây lát...
            </p>
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
 
              {tabs.filter(t => t.id === 'giao' || t.id === 'nhan' || t.id === 'chung' || t.id === 'kho').map(t => {
                const isSelected = tab === t.id
                const count = t.id === 'giao' ? giaoRows.length : t.id === 'nhan' ? nhanRows.length : t.id === 'kho' ? khoRows.length : t.id === 'chung' ? chungRows.length : null
 
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
 
              {tabs.filter(t => t.id === 'inventory').map(t => {
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

              {tabs.filter(t => t.id === 'inventory').map(t => {
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
              />
            )}
            {tab === 'kho' && (
              <KhoDuAnTab
                chungRows={chungRows}
                selectedProject={selectedProject}
                setSelectedProject={setSelectedProject}
                allProjects={allProjects}
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
              />
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
