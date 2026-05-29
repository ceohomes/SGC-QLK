import React, { useState, useRef, useCallback, useMemo } from 'react'
import * as XLSX from 'xlsx'
import {
  Upload, FileSpreadsheet, Search, X, RefreshCw, Info,
  ChevronDown, Download, Truck, PackageCheck, Settings, BarChart3,
  AlertCircle, CheckCircle2, Filter, ArrowUpDown
} from 'lucide-react'
import { COLS_GIAO_NHAN, parseXlsxToRows, formatVal, getTrangThaiColor } from './constants.js'

// ─── Header ──────────────────────────────────────────────────────────────────
function Header() {
  return (
    <header style={{
      background: 'linear-gradient(135deg, #0a3d73 0%, #0f58a7 60%, #1a6abf 100%)',
      padding: '0 24px',
      height: 56,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'space-between',
      boxShadow: '0 2px 12px rgba(10,61,115,0.4)',
      position: 'sticky',
      top: 0,
      zIndex: 100,
    }}>
      <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
        <div style={{
          width: 32, height: 32, background: '#fff',
          borderRadius: 8, display: 'flex', alignItems: 'center', justifyContent: 'center',
          boxShadow: '0 2px 8px rgba(0,0,0,0.15)'
        }}>
          <span style={{ color: '#0f58a7', fontWeight: 900, fontSize: 13 }}>SGC</span>
        </div>
        <div>
          <div style={{ color: '#fff', fontWeight: 800, fontSize: 15, letterSpacing: '0.02em' }}>
            SGC | BÁO CÁO GIAO NHẬN
          </div>
          <div style={{ color: 'rgba(255,255,255,0.55)', fontSize: 10, fontWeight: 500, letterSpacing: '0.08em' }}>
            QUẢN LÝ VẬT TƯ & MÁY MÓC THIẾT BỊ
          </div>
        </div>
      </div>
      <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
        <div style={{
          background: 'rgba(255,255,255,0.12)',
          border: '1px solid rgba(255,255,255,0.2)',
          borderRadius: 8,
          padding: '4px 12px',
          color: '#fff',
          fontSize: 11,
          fontWeight: 600
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
      gap: 0,
      background: 'var(--surface)',
      borderBottom: '2px solid var(--border)',
      padding: '0 24px',
    }}>
      {tabs.map(t => (
        <button
          key={t.id}
          onClick={() => onChange(t.id)}
          style={{
            display: 'flex',
            alignItems: 'center',
            gap: 7,
            padding: '12px 20px',
            fontWeight: active === t.id ? 700 : 500,
            fontSize: 13,
            color: active === t.id ? 'var(--primary)' : 'var(--text-muted)',
            background: 'transparent',
            borderBottom: active === t.id ? '2px solid var(--primary)' : '2px solid transparent',
            marginBottom: -2,
            transition: 'all 0.15s',
            whiteSpace: 'nowrap',
          }}
        >
          {t.icon}
          {t.label}
          {t.count != null && t.count > 0 && (
            <span style={{
              background: active === t.id ? 'var(--primary)' : '#e2e8f0',
              color: active === t.id ? '#fff' : 'var(--text-muted)',
              borderRadius: 20,
              padding: '1px 7px',
              fontSize: 10,
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
      <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: 10 }}>
        <div style={{
          width: 52, height: 52, borderRadius: 14,
          background: 'var(--primary-light)',
          display: 'flex', alignItems: 'center', justifyContent: 'center'
        }}>
          <Upload size={22} color="var(--primary)" />
        </div>
        <div>
          <div style={{ fontWeight: 700, fontSize: 14, color: 'var(--text)', marginBottom: 4 }}>
            {label}
          </div>
          <div style={{ color: 'var(--text-muted)', fontSize: 12 }}>
            Kéo thả file vào đây hoặc <span style={{ color: 'var(--primary)', fontWeight: 600 }}>chọn file</span>
          </div>
          <div style={{ color: 'var(--text-light)', fontSize: 11, marginTop: 4 }}>
            Hỗ trợ: .xlsx, .xls
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
    const daPD = rows.filter(r => String(r.trangThai).toLowerCase().includes('đã phê duyệt') || String(r.trangThai).toLowerCase().includes('hoàn thành')).length
    const chuaPD = rows.filter(r => String(r.trangThai).toLowerCase().includes('chưa')).length
    const kls = rows.reduce((s, r) => s + (parseFloat(r.khoiLuongNhap) || 0), 0)
    return { total, daPD, chuaPD, kls }
  }, [rows])

  const cards = [
    { label: 'Tổng đơn', value: stats.total.toLocaleString(), color: '#0f58a7', bg: '#e8f1fb' },
    { label: 'Đã phê duyệt', value: stats.daPD.toLocaleString(), color: '#16a34a', bg: '#dcfce7' },
    { label: 'Chưa xác nhận', value: stats.chuaPD.toLocaleString(), color: '#d97706', bg: '#fef9c3' },
    { label: 'Tổng KL nhập', value: stats.kls.toLocaleString('vi-VN'), color: '#7c3aed', bg: '#ede9fe' },
  ]

  return (
    <div style={{ display: 'flex', gap: 12, marginBottom: 16 }}>
      {cards.map(c => (
        <div key={c.label} style={{
          background: c.bg,
          borderRadius: 10,
          padding: '10px 16px',
          display: 'flex',
          flexDirection: 'column',
          gap: 2,
          flex: 1,
          minWidth: 0,
        }}>
          <div style={{ color: c.color, fontSize: 20, fontWeight: 800 }}>{c.value}</div>
          <div style={{ color: c.color, fontSize: 11, fontWeight: 600, opacity: 0.75 }}>{c.label}</div>
        </div>
      ))}
    </div>
  )
}

// ─── Filter Bar ───────────────────────────────────────────────────────────────
function FilterBar({ search, setSearch, trangThai, setTrangThai, trangThaiOptions, duAn, setDuAn, duAnOptions, onClear }) {
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
      <select className="input" style={{ flex: '0 0 180px' }} value={duAn} onChange={e => setDuAn(e.target.value)}>
        <option value="">Tất cả dự án</option>
        {duAnOptions.map(o => <option key={o} value={o}>{o}</option>)}
      </select>
      {(search || trangThai || duAn) && (
        <button className="btn btn-outline btn-sm" onClick={onClear}>
          <X size={12} /> Xóa lọc
        </button>
      )}
    </div>
  )
}

// ─── Data Table ───────────────────────────────────────────────────────────────
const GROUPS = [
  { key: null, cols: COLS_GIAO_NHAN.filter(c => !c.group), label: null },
  { key: 'Nhập kho', cols: COLS_GIAO_NHAN.filter(c => c.group === 'Nhập kho'), label: 'Nhập kho' },
  { key: 'Xuất kho', cols: COLS_GIAO_NHAN.filter(c => c.group === 'Xuất kho'), label: 'Xuất kho' },
  { key: 'Thông tin chung', cols: COLS_GIAO_NHAN.filter(c => c.group === 'Thông tin chung'), label: 'Thông tin chung' },
]

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
            <th style={{ width: 48, textAlign: 'center' }}>#</th>
            {GROUPS.map(g =>
              g.label
                ? <th key={g.key} colSpan={g.cols.length} style={{ textAlign: 'center' }}>{g.label}</th>
                : g.cols.map(c => <th key={c.key} rowSpan={2} style={{ minWidth: c.width }}>{c.label}</th>)
            )}
          </tr>
          <tr>
            {GROUPS.filter(g => g.label).map(g =>
              g.cols.map(c => <th key={c.key} style={{ minWidth: c.width }}>{c.label}</th>)
            )}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={row.id}>
              <td style={{ textAlign: 'center', color: 'var(--text-light)', fontSize: 11 }}>{i + 1}</td>
              {COLS_GIAO_NHAN.map(col => (
                <td key={col.key} style={{ maxWidth: col.width + 40 }} title={String(formatVal(row[col.key]) || '')}>
                  {col.key === 'trangThai' ? (
                    <span className={`badge ${getTrangThaiColor(row[col.key])}`}>
                      {row[col.key] || '—'}
                    </span>
                  ) : col.key === 'tinhTrang' ? (
                    row[col.key]
                      ? <span className={`badge ${row[col.key] === 'NEW' ? 'badge-green' : row[col.key] === 'USED' ? 'badge-yellow' : 'badge-gray'}`}>{row[col.key]}</span>
                      : '—'
                  ) : (
                    formatVal(row[col.key]) || <span style={{ color: 'var(--text-light)' }}>—</span>
                  )}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

// ─── Order Tab (shared for Giao & Nhan) ──────────────────────────────────────
function OrderTab({ type }) {
  const [rows, setRows] = useState([])
  const [fileName, setFileName] = useState('')
  const [loading, setLoading] = useState(false)
  const [search, setSearch] = useState('')
  const [trangThai, setTrangThai] = useState('')
  const [duAn, setDuAn] = useState('')

  const isGiao = type === 'giao'
  const label = isGiao ? 'Đơn Giao' : 'Đơn Nhận'
  const uploadLabel = isGiao
    ? 'Tải lên file Report_Orders_Đơn giao'
    : 'Tải lên file Report_Orders_Đơn nhận'

  const handleFile = useCallback((data, name) => {
    setLoading(true)
    setTimeout(() => {
      const parsed = parseXlsxToRows(data)
      setRows(parsed)
      setFileName(name)
      setLoading(false)
    }, 80)
  }, [])

  const trangThaiOptions = useMemo(() =>
    [...new Set(rows.map(r => r.trangThai).filter(Boolean))].sort(), [rows])
  const duAnOptions = useMemo(() =>
    [...new Set(rows.map(r => r.duAn).filter(Boolean))].sort(), [rows])

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
    if (duAn) r = r.filter(row => row.duAn === duAn)
    return r
  }, [rows, search, trangThai, duAn])

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
            <h2 style={{ fontSize: 17, fontWeight: 800, color: 'var(--text)', marginBottom: 6 }}>
              Tab {label}
            </h2>
            <p style={{ color: 'var(--text-muted)', fontSize: 13 }}>
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
            <div style={{ fontSize: 12, color: 'var(--primary-dark)' }}>
              <strong>Cấu trúc file yêu cầu:</strong> File Excel có sheet "Báo cáo theo dõi giao nhận" với 36 cột từ Ngày xuất nhập đến Nhãn hiệu. Dòng 1 là nhóm cột, dòng 2 là sub-header, dữ liệu từ dòng 3.
            </div>
          </div>

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
                <span style={{ fontSize: 12, fontWeight: 600, color: 'var(--primary)', maxWidth: 220, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                  {fileName}
                </span>
              </div>
              <span style={{ fontSize: 12, color: 'var(--text-muted)' }}>
                {rows.length.toLocaleString()} dòng dữ liệu · {filtered.length.toLocaleString()} hiển thị
              </span>
            </div>
            <div style={{ display: 'flex', gap: 8 }}>
              <button className="btn btn-outline btn-sm" onClick={() => { setRows([]); setFileName(''); setSearch(''); setTrangThai(''); setDuAn('') }}>
                <X size={12} /> Xóa file
              </button>
              <label className="btn btn-primary btn-sm" style={{ cursor: 'pointer' }}>
                <Upload size={12} /> Đổi file
                <input type="file" accept=".xlsx,.xls" style={{ display: 'none' }}
                  onChange={e => { if (e.target.files[0]) handleFile(e.target.files[0], e.target.files[0].name); e.target.value = '' }} />
              </label>
            </div>
          </div>

          <StatsBar rows={rows} />

          <FilterBar
            search={search} setSearch={setSearch}
            trangThai={trangThai} setTrangThai={setTrangThai}
            trangThaiOptions={trangThaiOptions}
            duAn={duAn} setDuAn={setDuAn}
            duAnOptions={duAnOptions}
            onClear={() => { setSearch(''); setTrangThai(''); setDuAn('') }}
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
          fontSize: 12,
          color: '#92400e',
          fontWeight: 600,
        }}>
          🚧 Tính năng đang được phát triển
        </div>
      </div>
    </div>
  )
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [tab, setTab] = useState('giao')

  const tabs = [
    { id: 'giao', label: 'Đơn Giao', icon: <Truck size={15} /> },
    { id: 'nhan', label: 'Đơn Nhận', icon: <PackageCheck size={15} /> },
    { id: 'config', label: 'Cấu hình dữ liệu', icon: <Settings size={15} /> },
    { id: 'summary', label: 'Tổng hợp', icon: <BarChart3 size={15} /> },
  ]

  return (
    <div style={{ minHeight: '100vh', display: 'flex', flexDirection: 'column' }}>
      <Header />
      <TabBar tabs={tabs} active={tab} onChange={setTab} />

      <div style={{ flex: 1, overflow: 'hidden' }}>
        {tab === 'giao' && <OrderTab type="giao" />}
        {tab === 'nhan' && <OrderTab type="nhan" />}
        {tab === 'config' && (
          <PlaceholderTab
            icon={<Settings />}
            title="Cấu hình dữ liệu"
            desc="Tính năng cấu hình dữ liệu sẽ được triển khai trong phiên bản tiếp theo."
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

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
      `}</style>
    </div>
  )
}
