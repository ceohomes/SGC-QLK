import React, { useState, useMemo, useCallback, useRef, useEffect } from 'react'
import {
  Building2,
  Check,
  CheckSquare,
  Square,
  Search,
  Plus,
  Trash2,
  Edit2,
  Save,
  Copy,
  Download,
  RefreshCw,
  AlertCircle,
  CheckCircle2,
  Layers,
  ArrowRight,
  Filter,
  Tag,
  ChevronDown,
  FileSpreadsheet,
  X,
  HelpCircle,
  Sparkles,
  ArrowUpDown,
  SlidersHorizontal,
  Info,
  FolderPlus,
  Unlink,
  ExternalLink
} from 'lucide-react'
import * as XLSXStyleRaw from 'xlsx-js-style'
const XLSX = XLSXStyleRaw.default || XLSXStyleRaw
import { supabase, isSupabaseConfigured } from '../supabaseClient'

export const SQL_SETUP_SCRIPT_BCH = `-- =========================================================================
-- SCRIPT TẠO BẢNG CHUẨN HÓA TÊN BAN CHỈ HUY (KHO BCH) TRÊN SUPABASE
-- Chức năng: Lưu trữ quy tắc chuyển đổi tên cũ / biến thể -> tên BCH chuẩn
-- =========================================================================

-- 1. Tạo bảng chuan_hoa_bch nếu chưa tồn tại
CREATE TABLE IF NOT EXISTS public.chuan_hoa_bch (
    id bigint GENERATED ALWAYS AS IDENTITY PRIMARY KEY,
    ten_cu text NOT NULL UNIQUE,          -- Tên cũ/biến thể xuất hiện trong Excel
    ten_chuan text NOT NULL,              -- Tên BCH chuẩn hóa thống nhất
    ghi_chu text,                         -- Ghi chú bổ sung (tùy chọn)
    created_at timestamp with time zone DEFAULT timezone('utc'::text, now())
);

-- 2. Đặt comment cho bảng và các cột
COMMENT ON TABLE public.chuan_hoa_bch IS 'Bảng quy tắc chuẩn hóa và gộp tên Ban Chỉ Huy (Kho BCH)';
COMMENT ON COLUMN public.chuan_hoa_bch.ten_cu IS 'Tên cũ hoặc tên biến thể trong file Excel';
COMMENT ON COLUMN public.chuan_hoa_bch.ten_chuan IS 'Tên BCH chuẩn hóa hiển thị đồng bộ trên hệ thống';

-- 3. Bật Row Level Security (RLS)
ALTER TABLE public.chuan_hoa_bch ENABLE ROW LEVEL SECURITY;

-- 4. Tạo các Policy cho phép đọc/ghi công khai (hoặc theo tài khoản)
DROP POLICY IF EXISTS "Allow public read chuan_hoa_bch" ON public.chuan_hoa_bch;
CREATE POLICY "Allow public read chuan_hoa_bch" ON public.chuan_hoa_bch 
    FOR SELECT USING (true);

DROP POLICY IF EXISTS "Allow public insert chuan_hoa_bch" ON public.chuan_hoa_bch;
CREATE POLICY "Allow public insert chuan_hoa_bch" ON public.chuan_hoa_bch 
    FOR INSERT WITH CHECK (true);

DROP POLICY IF EXISTS "Allow public update chuan_hoa_bch" ON public.chuan_hoa_bch;
CREATE POLICY "Allow public update chuan_hoa_bch" ON public.chuan_hoa_bch 
    FOR UPDATE USING (true);

DROP POLICY IF EXISTS "Allow public delete chuan_hoa_bch" ON public.chuan_hoa_bch;
CREATE POLICY "Allow public delete chuan_hoa_bch" ON public.chuan_hoa_bch 
    FOR DELETE USING (true);
`

function getUnitCategoryLocal(name) {
  if (!name) return 'chuaphanbo'
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
  if (
    s.startsWith('cty') ||
    s.startsWith('công ty') ||
    s.includes('cổ phần') ||
    s.includes('tnhh') ||
    s.includes('doanh nghiệp') ||
    s.includes('nhà máy') ||
    s.includes('xí nghiệp') ||
    s.includes('đại lý')
  ) {
    return 'ncc'
  }
  if (
    s.startsWith('tổ') ||
    s.startsWith('đội') ||
    s.includes('tổ đội') ||
    s.includes('ông ') ||
    s.includes('bà ') ||
    s.includes('anh ') ||
    s.includes('chị ')
  ) {
    return 'todoi'
  }
  return 'chuaphanbo'
}

// ─── Trình bày dùng chung trong tab Chuẩn hóa BCH (bút sửa, dấu X kiểu iPhone, nút Đóng Esc) ───
export const CHB_CSS = `
.chb-pen{display:inline-grid;place-items:center;width:20px;height:20px;padding:0;border:0;background:transparent;color:#0A2486;border-radius:6px;cursor:pointer;flex:none;vertical-align:middle}
.chb-pen:hover{background:#EEF1FA}
.chb-x{display:inline-grid;place-items:center;width:28px;height:28px;padding:0;border:0;border-radius:50%;background:#FFE7E5;color:#FF3B30;cursor:pointer;flex:none;box-shadow:none}
.chb-x:hover{background:#FF3B30;color:#fff}
.chb-x.sm{width:20px;height:20px}
.chb-close{display:inline-flex;align-items:center;gap:7px;height:36px;padding:0 8px 0 14px;border-radius:99px;border:0;background:#FFE7E5;color:#D92D3A;font-size:15px;font-weight:600;cursor:pointer;white-space:nowrap;flex:none}
.chb-close:hover{background:#FF3B30;color:#fff}
.chb-close kbd{font-family:inherit;font-size:12px;font-weight:700;line-height:20px;padding:0 7px;border-radius:99px;background:rgba(217,45,58,.12);color:inherit}
.chb-close:hover kbd{background:rgba(255,255,255,.25)}
.chb-search{position:relative;display:flex;align-items:center;height:36px;flex:1 1 240px;min-width:200px;max-width:420px;background:#fff;border-radius:99px;box-shadow:inset 0 0 0 1px #D5DCEB}
.chb-search:focus-within{box-shadow:inset 0 0 0 1.5px #0F58A7,0 0 0 3px rgba(15,88,167,.16)}
.chb-search .chb-lead{position:absolute;left:13px;top:50%;transform:translateY(-50%);color:#8C8E97;pointer-events:none}
html body .chb-search input.chb-in{flex:1;min-width:0;height:34px;border:0 !important;background:transparent;padding:0 6px 0 38px;font-size:15px;outline:none;box-shadow:none !important;border-radius:99px;color:#16171C}
html body .chb-search input.chb-in:focus:not([type=checkbox]):not([type=radio]){box-shadow:none !important;border:0 !important}
.chb-search .chb-x{margin-right:7px}
.chb-tb{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.chb-tb .sg-btn{flex:none}
.chb-tb .sg-seg{padding:2px}
.chb-tb .chb-sh{display:none}
@media (max-width:1500px){
  .chb-tb .chb-lg{display:none}
  .chb-tb .chb-sh{display:inline}
  .chb-tb .chb-ico-only .chb-lb{display:none}
  .chb-tb .chb-ico-only{padding:0;width:36px}
}
.chb-card{background:rgba(255,255,255,.82);border-radius:24px;box-shadow:0 10px 30px -18px rgba(10,36,80,.35),inset 0 0 0 1px rgba(255,255,255,.9);padding:12px 14px 10px;display:flex;flex-direction:column;gap:10px;min-width:0}
.chb-scroll{overflow:auto;border-radius:14px;position:relative}
.chb-scroll table{width:100%}
.chb-scroll thead th{position:sticky;top:0;z-index:5}
.chb-foot{font-size:12px;font-weight:600;color:#4A4C55;padding:0 4px}
.chb-chip{display:inline-flex;align-items:center;gap:6px;height:28px;padding:0 4px 0 12px;border-radius:99px;background:#fff;box-shadow:inset 0 0 0 1px #C9D3E6;font-size:12px;font-weight:600;color:#2A2C33;max-width:100%;min-width:0;box-sizing:border-box}
.chb-chip.main{background:#E4EEFA;box-shadow:inset 0 0 0 1px #B9CFEE;color:#0A3D73;font-weight:700}
.chb-chip.solo{padding-right:12px}
.chb-chip .chb-nm{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0;max-width:340px}
.chb-chip .chb-ct{flex:none;font-family:var(--fn);font-size:12px;font-weight:700;color:#4A4C55;background:rgba(10,61,115,.08);padding:0 7px;border-radius:99px;line-height:18px;white-space:nowrap}
.chb-add{display:inline-flex;align-items:center;gap:5px;height:28px;padding:0 12px;border-radius:99px;border:0;background:#E4EEFA;color:#0F58A7;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap;box-shadow:none}
.chb-add:hover{background:#D3E3F7}
.chb-tag{display:inline-flex;align-items:center;gap:4px;height:22px;padding:0 9px;border-radius:99px;font-size:12px;font-weight:700;white-space:nowrap}
.chb-tag.multi{background:#ECE7FB;color:#4B2FA0}
.chb-tag.solo{background:#EEF0F4;color:#4A4C55}
.chb-tag.warn{background:#FDEBDD;color:#A0420B}
.chb-edit{width:100%;box-sizing:border-box;height:34px;padding:0 10px;border-radius:10px;border:1.5px solid #0F58A7 !important;box-shadow:0 0 0 3px rgba(15,88,167,.16) !important;background:#fff;font-size:15px;font-weight:400;outline:none;color:#16171C}
.chb-mini{display:inline-flex;align-items:center;gap:4px;height:28px;padding:0 12px;border-radius:99px;border:0;font-size:12px;font-weight:700;cursor:pointer;white-space:nowrap}
.chb-mini.ok{background:linear-gradient(135deg,#12A150,#3CCB7F);color:#fff}
.chb-mini.no{background:#fff;color:#4A4C55;box-shadow:inset 0 0 0 1px #D5DCEB}
.chb-hint{font-size:12px;font-weight:600;color:#8C8E97;white-space:nowrap}
.chb-ov{position:fixed;inset:0;background:rgba(15,23,42,.55);backdrop-filter:blur(3px);display:flex;align-items:center;justify-content:center;z-index:99999;padding:20px}
.chb-dlg{background:#fff;border-radius:24px;width:100%;max-height:90vh;display:flex;flex-direction:column;overflow:hidden;box-shadow:0 30px 60px -20px rgba(10,36,80,.45)}
.chb-dlg-hd{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 16px 12px 22px;border-bottom:1px solid #E3E6EE}
.chb-dlg-ttl{display:flex;align-items:center;gap:10px;min-width:0}
.chb-dlg-ic{flex:none;width:36px;height:36px;border-radius:50%;display:grid;place-items:center;color:#fff;background:linear-gradient(135deg,#0F58A7,#2F86E0)}
.chb-dlg-ttl h3{margin:0;font-size:15px;font-weight:700;color:#0A3D73;line-height:1.3}
.chb-dlg-ttl .chb-dlg-sub{font-size:12px;font-weight:600;color:#4A4C55;margin-top:1px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.chb-dlg-bd{padding:16px 22px;overflow-y:auto;flex:1;display:flex;flex-direction:column;gap:14px}
.chb-dlg-ft{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:12px 22px;border-top:1px solid #E3E6EE;background:#F7F8FB}
.chb-list{max-height:260px;overflow-y:auto;border-radius:14px;box-shadow:inset 0 0 0 1px #D5DCEB;background:#fff;padding:4px}
.chb-row{display:flex;align-items:center;gap:10px;padding:7px 10px;border-radius:10px;cursor:pointer;font-size:15px}
.chb-row:hover{background:#F3F7FD}
.chb-row.on{background:#E4EEFA}
.chb-row .chb-nm{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#16171C}
.chb-row.on .chb-nm{font-weight:700;color:#0A3D73}
.chb-lbl{font-size:15px;font-weight:700;color:#0A3D73;display:flex;align-items:center;gap:6px}
.chb-fld{width:100%;box-sizing:border-box;height:36px;padding:0 12px;border-radius:12px;border:1.5px solid #D5DCEB;background:#fff;font-size:15px;outline:none}
`

export function PenIcon() {
  return (
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="M12 20h8" />
      <path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L8 18l-4 1 1-4z" />
    </svg>
  )
}

export function XMark({ size = 12, w = 2 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" aria-hidden="true" style={{ flex: 'none', display: 'block' }}>
      <path d="M5 5L19 19M19 5L5 19" strokeWidth={w} vectorEffect="non-scaling-stroke" />
    </svg>
  )
}

export function CloseEscButton({ onClick }) {
  return (
    <button type="button" className="chb-close" onClick={onClick} title="Đóng (Esc)">
      <XMark size={12} w={2.3} /> Đóng <kbd>Esc</kbd>
    </button>
  )
}

/** Esc đóng hộp thoại (bắt ở pha capture để không lan ra phím Esc thoát toàn màn hình của app) */
export function useEscClose(active, onClose) {
  React.useEffect(() => {
    if (!active) return undefined
    const h = (e) => {
      if (e.key === 'Escape') {
        e.stopPropagation()
        onClose && onClose()
      }
    }
    window.addEventListener('keydown', h, true)
    return () => window.removeEventListener('keydown', h, true)
  }, [active, onClose])
}

/**
 * Modal Dialog Gán Kho Gốc vào một nhóm BCH Chuẩn Hóa
 * Hiển thị toàn màn hình trung tâm, không bị che/cắt bởi overflow của bảng
 */
function AssignKhoToGroupModal({
  isOpen,
  onClose,
  targetGroup,
  allOriginalKhos = [],
  onConfirmAdd
}) {
  const [search, setSearch] = useState('')
  const [filterType, setFilterType] = useState('all') // 'all' | 'unmerged' | 'other'
  const [selectedKhos, setSelectedKhos] = useState(new Set())
  const inputRef = useRef(null)

  const norm = (s) => (s || '').trim().normalize('NFC').replace(/\s+/g, ' ')
  const normKey = (s) => norm(s).toLowerCase()

  const currentMemberNames = useMemo(() => {
    if (!targetGroup) return []
    return (targetGroup.members || []).map(m => m.name)
  }, [targetGroup])

  const currentMemberSet = useMemo(() => {
    return new Set(currentMemberNames.map(name => normKey(name)))
  }, [currentMemberNames])

  useEffect(() => {
    if (isOpen) {
      setSelectedKhos(new Set())
      setSearch('')
      setFilterType('all')
      setTimeout(() => {
        if (inputRef.current) inputRef.current.focus()
      }, 60)
    }
  }, [isOpen])

  // Danh sách kho có thể gán
  const availableKhos = useMemo(() => {
    if (!targetGroup) return []
    const stdKey = normKey(targetGroup.standardName)

    return allOriginalKhos
      .filter(k => !currentMemberSet.has(normKey(k.name)))
      .map(k => {
        const kKey = normKey(k.name)
        const currentStd = k.currentStandardName ? norm(k.currentStandardName) : k.name
        const isAssignedOther = k.isMapped && normKey(currentStd) !== stdKey && normKey(currentStd) !== kKey
        const isStandalone = !k.isMapped || normKey(currentStd) === kKey

        return {
          ...k,
          isAssignedOther,
          isStandalone,
          currentStd
        }
      })
  }, [allOriginalKhos, currentMemberSet, targetGroup])

  // Lọc theo search và filterType
  const filteredList = useMemo(() => {
    const q = search.toLowerCase().trim()
    return availableKhos.filter(k => {
      // Filter tab
      if (filterType === 'unmerged' && !k.isStandalone) return false
      if (filterType === 'other' && !k.isAssignedOther) return false

      // Search keyword
      if (!q) return true
      return k.name.toLowerCase().includes(q) || (k.currentStd && k.currentStd.toLowerCase().includes(q))
    })
  }, [availableKhos, search, filterType])

  const toggleSelect = (name) => {
    setSelectedKhos(prev => {
      const next = new Set(prev)
      if (next.has(name)) next.delete(name)
      else next.add(name)
      return next
    })
  }

  const handleSelectAllFiltered = () => {
    if (selectedKhos.size === filteredList.length && filteredList.length > 0) {
      setSelectedKhos(new Set())
    } else {
      setSelectedKhos(new Set(filteredList.map(k => k.name)))
    }
  }

  useEscClose(Boolean(isOpen && targetGroup), onClose)

  if (!isOpen || !targetGroup) return null

  const unmergedCount = availableKhos.filter(k => k.isStandalone).length
  const otherCount = availableKhos.filter(k => k.isAssignedOther).length

  return (
    <div className="chb-ov">
      <style>{CHB_CSS}</style>
      <div className="chb-dlg" style={{ maxWidth: '760px' }}>
        {/* Đầu hộp thoại */}
        <div className="chb-dlg-hd">
          <div className="chb-dlg-ttl">
            <span className="chb-dlg-ic"><FolderPlus size={18} /></span>
            <div style={{ minWidth: 0 }}>
              <h3>Gán kho gốc vào Ban Chỉ Huy</h3>
              <div className="chb-dlg-sub" title={targetGroup.standardName}>{targetGroup.standardName}</div>
            </div>
          </div>
          <CloseEscButton onClick={onClose} />
        </div>

        {/* Thân hộp thoại */}
        <div className="chb-dlg-bd">
          {/* Các kho đã thuộc nhóm */}
          <div>
            <div className="chb-lbl" style={{ fontSize: '12px', marginBottom: '8px' }}>
              <Layers size={14} />
              <span>Kho gốc đang thuộc Ban Chỉ Huy này ({currentMemberNames.length} kho)</span>
            </div>
            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', maxHeight: '96px', overflowY: 'auto' }}>
              {targetGroup.members && targetGroup.members.map(m => (
                <span key={m.name} className="chb-chip main solo" title={m.name}>
                  <span className="chb-nm">{m.name}</span>
                  <span className="chb-ct">{(m.totalCount || 0).toLocaleString('vi-VN')} phiếu</span>
                </span>
              ))}
            </div>
          </div>

          {/* Chọn thêm kho gốc */}
          <div>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '10px', flexWrap: 'wrap', gap: '8px' }}>
              <div className="chb-lbl">
                <Plus size={16} />
                <span>Chọn kho gốc cần gộp thêm</span>
              </div>
              <div className="sg-seg">
                <button type="button" className={filterType === 'all' ? 'on' : ''} onClick={() => setFilterType('all')}>
                  Tất cả ({availableKhos.length})
                </button>
                <button type="button" className={filterType === 'unmerged' ? 'on' : ''} onClick={() => setFilterType('unmerged')}>
                  Chưa gộp ({unmergedCount})
                </button>
                <button type="button" className={filterType === 'other' ? 'on' : ''} onClick={() => setFilterType('other')}>
                  Thuộc BCH khác ({otherCount})
                </button>
              </div>
            </div>

            <div style={{ display: 'flex', alignItems: 'center', gap: '10px', marginBottom: '10px' }}>
              <div className="chb-search" style={{ maxWidth: 'none' }}>
                <Search size={16} className="chb-lead" />
                <input
                  ref={inputRef}
                  className="chb-in"
                  type="text"
                  placeholder="Tìm theo tên kho gốc..."
                  value={search}
                  onChange={(e) => setSearch(e.target.value)}
                />
                {search && (
                  <button type="button" className="chb-x sm" onClick={() => setSearch('')} title="Xoá ô tìm">
                    <XMark size={11} w={2.3} />
                  </button>
                )}
              </div>
              {filteredList.length > 0 && (
                <button type="button" className="sg-btn ghost" onClick={handleSelectAllFiltered}>
                  {selectedKhos.size === filteredList.length ? 'Bỏ chọn tất cả' : `Chọn tất cả (${filteredList.length})`}
                </button>
              )}
            </div>

            <div className="chb-list">
              {filteredList.length === 0 ? (
                <div style={{ padding: '32px 16px', textAlign: 'center', color: '#8C8E97', fontSize: '15px' }}>
                  Không có kho gốc nào khớp với từ khóa tìm kiếm
                </div>
              ) : (
                filteredList.map((k) => {
                  const isChecked = selectedKhos.has(k.name)
                  return (
                    <div key={k.name} className={'chb-row' + (isChecked ? ' on' : '')} onClick={() => toggleSelect(k.name)}>
                      {isChecked ? <CheckSquare size={18} color="#0F58A7" /> : <Square size={18} color="#B4BBC9" />}
                      <span className="chb-nm" title={k.name}>{k.name}</span>
                      {k.isAssignedOther ? (
                        <span className="chb-tag multi" style={{ maxWidth: '220px' }} title={`Đang thuộc: ${k.currentStd}`}>
                          <span style={{ minWidth: 0, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>Đang thuộc: {k.currentStd}</span>
                        </span>
                      ) : (
                        <span className="chb-tag warn">Chưa gộp</span>
                      )}
                      <span className="chb-tag solo sg-num" style={{ minWidth: '72px', justifyContent: 'flex-end' }}>
                        {(k.totalCount || 0).toLocaleString('vi-VN')} phiếu
                      </span>
                    </div>
                  )
                })
              )}
            </div>
            <div className="chb-hint" style={{ marginTop: '6px' }}>
              Tìm thấy {filteredList.length} kho gốc có thể gán
            </div>
          </div>
        </div>

        {/* Chân hộp thoại */}
        <div className="chb-dlg-ft">
          <div style={{ fontSize: '15px', fontWeight: 600, color: '#4A4C55' }}>
            Đã chọn <strong style={{ color: '#0F58A7' }}>{selectedKhos.size}</strong> kho gốc
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
            <button type="button" className="sg-btn ghost" onClick={onClose}>Hủy bỏ</button>
            <button
              type="button"
              className="sg-btn"
              onClick={() => {
                if (selectedKhos.size === 0) return
                onConfirmAdd(targetGroup.standardName, Array.from(selectedKhos))
                onClose()
              }}
              disabled={selectedKhos.size === 0}
            >
              <Check size={16} /> Gán vào BCH ({selectedKhos.size})
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}

export default function ChuanHoaBchTab({
  chungRows = [],
  customCategoryMap = {},
  dbCategoryMap = {},
  bchAliasRules = [],
  setBchAliasRules,
  bchAliasMap = {},
  onApplyRulesToCurrentData,
  onReloadRulesFromSupabase
}) {
  const [searchTerm, setSearchTerm] = useState('')
  const [statusFilter, setStatusFilter] = useState('all') // 'all' | 'merged_multi' | 'standalone'
  
  // Create new standard BCH modal
  const [showCreateModal, setShowCreateModal] = useState(false)
  const [newBchStandardName, setNewBchStandardName] = useState('')
  const [newBchSelectedKhos, setNewBchSelectedKhos] = useState(new Set())
  const [newBchSearchKho, setNewBchSearchKho] = useState('')
  const [newBchNote, setNewBchNote] = useState('')

  // Inline edit standard name
  const [editingStandardKey, setEditingStandardKey] = useState(null)
  const [editingStandardValue, setEditingStandardValue] = useState('')
  const [editingNoteValue, setEditingNoteValue] = useState('')

  // Modal Gán Kho Gốc vào một nhóm BCH chuẩn
  const [assignModalGroup, setAssignModalGroup] = useState(null)

  // Status message
  const [saveStatus, setSaveStatus] = useState('idle') // 'idle' | 'saving' | 'success' | 'error'
  const [statusMessage, setStatusMessage] = useState('')

  const norm = (s) => (s || '').trim().normalize('NFC').replace(/\s+/g, ' ')
  const normKey = (s) => norm(s).toLowerCase()

  // 1. Trích xuất danh sách tất cả các Kho BCH gốc từ chungRows & bchAliasRules
  const allOriginalKhos = useMemo(() => {
    const counts = {}
    chungRows.forEach(r => {
      const gRaw = norm(r.donViGiao)
      const nRaw = norm(r.donViNhan)
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

    // Bao gồm các tên cũ đã có trong rule
    bchAliasRules.forEach(r => {
      if (r.ten_cu) {
        const key = r.ten_cu.toLowerCase()
        if (!counts[key]) {
          counts[key] = { name: r.ten_cu, giaoCount: 0, nhanCount: 0 }
        }
      }
    })

    const normCustom = {}
    Object.entries(customCategoryMap || {}).forEach(([k, v]) => {
      if (k) normCustom[normKey(k)] = v
    })

    const normDb = {}
    Object.entries(dbCategoryMap || {}).forEach(([k, v]) => {
      if (k) normDb[normKey(k)] = v
    })

    const result = Object.values(counts)
      .map(item => {
        const k = normKey(item.name)
        const cat = normCustom[k] || normDb[k] || getUnitCategoryLocal(item.name)
        const totalCount = item.giaoCount + item.nhanCount
        
        // Tìm tên chuẩn hiện tại
        const rule = bchAliasRules.find(r => normKey(r.ten_cu) === k)
        const currentStandardName = rule ? norm(rule.ten_chuan) : item.name
        const isMapped = !!rule && normKey(rule.ten_chuan) !== k

        return {
          name: item.name,
          category: cat,
          giaoCount: item.giaoCount,
          nhanCount: item.nhanCount,
          totalCount,
          currentStandardName,
          isMapped,
          ruleNote: rule?.ghi_chu || ''
        }
      })
      .filter(item => item.category === 'kho' || item.isMapped || bchAliasRules.some(r => normKey(r.ten_chuan) === normKey(item.name)))
      .sort((a, b) => a.name.localeCompare(b.name, 'vi', { sensitivity: 'base' }))

    return result
  }, [chungRows, customCategoryMap, dbCategoryMap, bchAliasRules])

  // 2. Gom nhóm: MỖI DÒNG LÀ 1 BAN CHỈ HUY CHUẨN HÓA
  const standardizedBchGroups = useMemo(() => {
    const groupsMap = {}

    // Map nhanh tìm stats của 1 kho gốc
    const khoStatsMap = {}
    allOriginalKhos.forEach(k => {
      khoStatsMap[normKey(k.name)] = k
    })

    // Bước A: Tạo nhóm từ tất cả quy tắc bchAliasRules
    bchAliasRules.forEach(r => {
      if (!r.ten_chuan) return
      const stdClean = norm(r.ten_chuan)
      const stdKey = normKey(stdClean)

      if (!groupsMap[stdKey]) {
        groupsMap[stdKey] = {
          key: stdKey,
          standardName: stdClean,
          members: new Map(), // name -> info
          note: r.ghi_chu || '',
          hasRules: true
        }
      }

      const oldClean = norm(r.ten_cu)
      if (oldClean) {
        const stats = khoStatsMap[normKey(oldClean)] || { giaoCount: 0, nhanCount: 0, totalCount: 0 }
        groupsMap[stdKey].members.set(normKey(oldClean), {
          name: oldClean,
          isAlias: normKey(oldClean) !== stdKey,
          giaoCount: stats.giaoCount,
          nhanCount: stats.nhanCount,
          totalCount: stats.totalCount,
          note: r.ghi_chu || ''
        })
      }
    })

    // Bước B: Với các nhóm tên chuẩn, nếu chính tên chuẩn đó cũng là 1 Kho gốc trong dữ liệu, thêm nó vào nhóm
    Object.values(groupsMap).forEach(group => {
      const stats = khoStatsMap[group.key]
      if (stats && !group.members.has(group.key)) {
        // Tên chuẩn là 1 kho gốc thực tế
        group.members.set(group.key, {
          name: stats.name,
          isAlias: false,
          giaoCount: stats.giaoCount,
          nhanCount: stats.nhanCount,
          totalCount: stats.totalCount,
          note: ''
        })
      }
    })

    // Bước C: Duyệt qua tất cả các kho gốc còn lại chưa thuộc bất kỳ nhóm chuẩn nào -> Tạo nhóm đơn lẻ (1 kho = 1 BCH)
    allOriginalKhos.forEach(kho => {
      const kKey = normKey(kho.name)
      // Kiểm tra kho này đã nằm trong members của nhóm nào chưa
      let isAlreadyMember = false
      for (const group of Object.values(groupsMap)) {
        if (group.members.has(kKey)) {
          isAlreadyMember = true
          break
        }
      }

      if (!isAlreadyMember) {
        // Kho này độc lập, chưa gộp
        groupsMap[kKey] = {
          key: kKey,
          standardName: kho.name,
          members: new Map([
            [kKey, {
              name: kho.name,
              isAlias: false,
              giaoCount: kho.giaoCount,
              nhanCount: kho.nhanCount,
              totalCount: kho.totalCount,
              note: ''
            }]
          ]),
          note: '',
          hasRules: false
        }
      }
    })

    // Chuyển sang Array & tính tổng số phiếu phát sinh
    const list = Object.values(groupsMap).map(group => {
      const memberList = Array.from(group.members.values())
      const totalPhieu = memberList.reduce((acc, m) => acc + (m.totalCount || 0), 0)
      const totalGiao = memberList.reduce((acc, m) => acc + (m.giaoCount || 0), 0)
      const totalNhan = memberList.reduce((acc, m) => acc + (m.nhanCount || 0), 0)
      const isMultiMerged = memberList.length > 1

      return {
        key: group.key,
        standardName: group.standardName,
        members: memberList,
        memberCount: memberList.length,
        isMultiMerged,
        totalPhieu,
        totalGiao,
        totalNhan,
        note: group.note,
        hasRules: group.hasRules
      }
    })

    // Sắp xếp: Các nhóm đã gộp nhiều kho lên trước, sau đó theo tên A-Z
    list.sort((a, b) => {
      if (a.isMultiMerged && !b.isMultiMerged) return -1
      if (!a.isMultiMerged && b.isMultiMerged) return 1
      return a.standardName.localeCompare(b.standardName, 'vi', { sensitivity: 'base' })
    })

    return list
  }, [allOriginalKhos, bchAliasRules])

  // Thống kê
  const multiMergedGroupsCount = useMemo(() => {
    return standardizedBchGroups.filter(g => g.isMultiMerged).length
  }, [standardizedBchGroups])

  const totalMergedOldKhoCount = useMemo(() => {
    return allOriginalKhos.filter(k => k.isMapped).length
  }, [allOriginalKhos])

  const standaloneGroupsCount = useMemo(() => {
    return standardizedBchGroups.filter(g => !g.isMultiMerged).length
  }, [standardizedBchGroups])

  // Lọc danh sách BCH chuẩn theo tìm kiếm & bộ lọc
  const filteredStandardGroups = useMemo(() => {
    return standardizedBchGroups.filter(group => {
      // 1. Tìm kiếm
      if (searchTerm) {
        const q = searchTerm.toLowerCase().trim()
        const matchStd = group.standardName.toLowerCase().includes(q)
        const matchMembers = group.members.some(m => m.name.toLowerCase().includes(q))
        const matchNote = (group.note || '').toLowerCase().includes(q)
        if (!matchStd && !matchMembers && !matchNote) return false
      }

      // 2. Bộ lọc trạng thái
      if (statusFilter === 'merged_multi') {
        if (!group.isMultiMerged) return false
      } else if (statusFilter === 'standalone') {
        if (group.isMultiMerged) return false
      }

      return true
    })
  }, [standardizedBchGroups, searchTerm, statusFilter])

  // ─── HANDLERS ─────────────────────────────────────────────────────────────

  // Gỡ 1 kho gốc ra khỏi nhóm chuẩn (trả về tên gốc độc lập)
  const handleRemoveKhoFromGroup = (khoName) => {
    const cleanOld = norm(khoName)
    const newRules = bchAliasRules.filter(r => normKey(r.ten_cu) !== normKey(cleanOld))
    setBchAliasRules(newRules)
    try {
      localStorage.setItem('sgc_chuan_hoa_bch', JSON.stringify(newRules))
    } catch (e) {}

    if (onApplyRulesToCurrentData) {
      onApplyRulesToCurrentData(newRules)
    }
  }

  // Gán thêm danh sách các kho gốc vào 1 nhóm BCH chuẩn
  const handleAddKhosToGroup = (targetStandardName, khoNamesToAdd = []) => {
    const stdName = norm(targetStandardName)
    if (!stdName || khoNamesToAdd.length === 0) return

    const newRules = [...bchAliasRules]
    khoNamesToAdd.forEach(oldName => {
      const cleanOld = norm(oldName)
      if (!cleanOld) return

      // Nếu tên kho trùng với tên chuẩn, không cần rule
      if (normKey(cleanOld) === normKey(stdName)) {
        const idx = newRules.findIndex(r => normKey(r.ten_cu) === normKey(cleanOld))
        if (idx >= 0) newRules.splice(idx, 1)
        return
      }

      const idx = newRules.findIndex(r => normKey(r.ten_cu) === normKey(cleanOld))
      if (idx >= 0) {
        newRules[idx] = {
          ...newRules[idx],
          ten_chuan: stdName
        }
      } else {
        newRules.push({
          ten_cu: cleanOld,
          ten_chuan: stdName,
          ghi_chu: ''
        })
      }
    })

    setBchAliasRules(newRules)
    try {
      localStorage.setItem('sgc_chuan_hoa_bch', JSON.stringify(newRules))
    } catch (e) {}

    if (onApplyRulesToCurrentData) {
      onApplyRulesToCurrentData(newRules)
    }

    setStatusMessage(`Đã thêm ${khoNamesToAdd.length} kho gốc vào BCH "${stdName}". Bấm "Lưu lên Supabase" để đồng bộ!`)
    setSaveStatus('success')
    setTimeout(() => {
      setStatusMessage('')
      setSaveStatus('idle')
    }, 4000)
  }

  // Xóa toàn bộ nhóm chuẩn (hủy gộp tất cả các kho trong nhóm)
  const handleDeleteGroup = (group) => {
    if (!confirm(`Bạn có chắc chắn muốn hủy gộp toàn bộ các kho trong nhóm "${group.standardName}" không?`)) {
      return
    }

    const memberKeys = new Set(group.members.map(m => normKey(m.name)))
    const newRules = bchAliasRules.filter(r => {
      // Xóa các rule có ten_chuan là nhóm này HOẶC ten_cu nằm trong nhóm này
      if (normKey(r.ten_chuan) === group.key) return false
      if (memberKeys.has(normKey(r.ten_cu))) return false
      return true
    })

    setBchAliasRules(newRules)
    try {
      localStorage.setItem('sgc_chuan_hoa_bch', JSON.stringify(newRules))
    } catch (e) {}

    if (onApplyRulesToCurrentData) {
      onApplyRulesToCurrentData(newRules)
    }
  }

  // Đổi tên BCH Chuẩn Hóa
  const handleSaveRenameGroup = (group) => {
    const newName = norm(editingStandardValue)
    if (!newName) {
      alert('Tên BCH chuẩn hóa không được để trống!')
      return
    }

    const oldStdKey = group.key
    const newStdKey = normKey(newName)

    if (oldStdKey === newStdKey) {
      setEditingStandardKey(null)
      return
    }

    const newRules = [...bchAliasRules]

    // Cập nhật tất cả các rule đang trỏ tới oldStdKey
    newRules.forEach((r, idx) => {
      if (normKey(r.ten_chuan) === oldStdKey) {
        newRules[idx] = {
          ...r,
          ten_chuan: newName,
          ghi_chu: editingNoteValue || r.ghi_chu || ''
        }
      }
    })

    // Nếu tên cũ của nhóm là 1 kho gốc thực tế và khác tên mới, tạo rule cho chính kho đó
    const wasOriginalKho = allOriginalKhos.some(k => normKey(k.name) === oldStdKey)
    if (wasOriginalKho && oldStdKey !== newStdKey) {
      const existingRule = newRules.find(r => normKey(r.ten_cu) === oldStdKey)
      if (existingRule) {
        existingRule.ten_chuan = newName
      } else {
        newRules.push({
          ten_cu: group.standardName,
          ten_chuan: newName,
          ghi_chu: editingNoteValue || ''
        })
      }
    }

    setBchAliasRules(newRules)
    try {
      localStorage.setItem('sgc_chuan_hoa_bch', JSON.stringify(newRules))
    } catch (e) {}

    if (onApplyRulesToCurrentData) {
      onApplyRulesToCurrentData(newRules)
    }

    setEditingStandardKey(null)
    setEditingStandardValue('')
    setEditingNoteValue('')
  }

  // Tạo nhóm BCH Chuẩn Hóa Mới từ Modal
  const handleCreateNewGroup = () => {
    const stdName = norm(newBchStandardName)
    if (!stdName) {
      alert('Vui lòng nhập Tên Ban Chỉ Huy Chuẩn Hóa!')
      return
    }

    if (newBchSelectedKhos.size === 0) {
      alert('Vui lòng chọn ít nhất 1 Kho BCH gốc để gộp vào nhóm này!')
      return
    }

    const newRules = [...bchAliasRules]
    let count = 0

    newBchSelectedKhos.forEach(oldName => {
      const cleanOld = norm(oldName)
      if (!cleanOld) return

      if (normKey(cleanOld) === normKey(stdName)) {
        const idx = newRules.findIndex(r => normKey(r.ten_cu) === normKey(cleanOld))
        if (idx >= 0) newRules.splice(idx, 1)
        return
      }

      const idx = newRules.findIndex(r => normKey(r.ten_cu) === normKey(cleanOld))
      if (idx >= 0) {
        newRules[idx] = {
          ...newRules[idx],
          ten_chuan: stdName,
          ghi_chu: newBchNote || newRules[idx].ghi_chu || ''
        }
      } else {
        newRules.push({
          ten_cu: cleanOld,
          ten_chuan: stdName,
          ghi_chu: newBchNote || ''
        })
      }
      count++
    })

    setBchAliasRules(newRules)
    try {
      localStorage.setItem('sgc_chuan_hoa_bch', JSON.stringify(newRules))
    } catch (e) {}

    if (onApplyRulesToCurrentData) {
      onApplyRulesToCurrentData(newRules)
    }

    setShowCreateModal(false)
    setNewBchStandardName('')
    setNewBchSelectedKhos(new Set())
    setNewBchSearchKho('')
    setNewBchNote('')

    setStatusMessage(`Đã tạo nhóm chuẩn "${stdName}" với ${count} kho gốc! Bấm "Lưu lên Supabase" để đồng bộ vĩnh viễn.`)
    setSaveStatus('success')
    setTimeout(() => {
      setStatusMessage('')
      setSaveStatus('idle')
    }, 5000)
  }

  // Lưu lên Supabase
  const handleSaveToSupabase = async () => {
    if (!isSupabaseConfigured) {
      alert('Chưa cấu hình Supabase!')
      return
    }

    setSaveStatus('saving')
    setStatusMessage('Đang lưu quy tắc chuẩn hóa lên Supabase...')

    try {
      // Check table
      const { error: testErr } = await supabase
        .from('chuan_hoa_bch')
        .select('id')
        .limit(1)

      if (testErr) {
        if (testErr.code === '42P01' || testErr.message?.includes('does not exist')) {
          throw new Error('Bảng "chuan_hoa_bch" chưa tồn tại trên Supabase.')
        }
      }

      // Clean old records
      const { error: delErr } = await supabase
        .from('chuan_hoa_bch')
        .delete()
        .neq('id', -999999)

      if (delErr) {
        console.warn('Lỗi khi làm sạch dữ liệu cũ:', delErr)
      }

      // Insert new rules
      if (bchAliasRules.length > 0) {
        const payload = bchAliasRules.map(r => ({
          ten_cu: norm(r.ten_cu),
          ten_chuan: norm(r.ten_chuan),
          ghi_chu: r.ghi_chu || ''
        }))

        const chunkSize = 100
        for (let i = 0; i < payload.length; i += chunkSize) {
          const chunk = payload.slice(i, i + chunkSize)
          const { error: insErr } = await supabase
            .from('chuan_hoa_bch')
            .insert(chunk)
          if (insErr) throw insErr
        }
      }

      setSaveStatus('success')
      setStatusMessage(`Đã lưu thành công ${bchAliasRules.length} quy tắc chuẩn hóa BCH lên Supabase!`)
      setTimeout(() => {
        setStatusMessage('')
        setSaveStatus('idle')
      }, 5000)

      if (onApplyRulesToCurrentData) {
        onApplyRulesToCurrentData(bchAliasRules)
      }
    } catch (err) {
      console.error(err)
      setSaveStatus('error')
      setStatusMessage(`Lỗi lưu Supabase: ${err.message || 'Kiểm tra bảng dữ liệu hoặc kết nối mạng'}`)
    }
  }

  // Xuất file Excel (định dạng đẹp: tiêu đề, màu header, căn lề, độ rộng cột hợp lý)
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new()
    const today = new Date().toLocaleDateString('vi-VN')

    // ── Style dùng chung ──────────────────────────────────────────────────
    const borderThin = (rgb) => ({
      top: { style: 'thin', color: { rgb } },
      bottom: { style: 'thin', color: { rgb } },
      left: { style: 'thin', color: { rgb } },
      right: { style: 'thin', color: { rgb } }
    })
    const titleStyle = {
      font: { name: 'Segoe UI', sz: 15, bold: true, color: { rgb: '0F172A' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    }
    const subtitleStyle = {
      font: { name: 'Segoe UI', sz: 10, italic: true, color: { rgb: '64748B' } },
      alignment: { horizontal: 'left', vertical: 'center' }
    }
    const makeHeaderStyle = (rgb, borderRgb) => ({
      fill: { patternType: 'solid', fgColor: { rgb } },
      font: { name: 'Segoe UI', sz: 10.5, bold: true, color: { rgb: 'FFFFFF' } },
      alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
      border: borderThin(borderRgb)
    })
    const makeDataStyle = (align = 'left', zebra = false) => ({
      font: { name: 'Segoe UI', sz: 10, color: { rgb: '1E293B' } },
      ...(zebra ? { fill: { patternType: 'solid', fgColor: { rgb: 'F8FAFC' } } } : {}),
      alignment: { horizontal: align, vertical: 'center', wrapText: true },
      border: borderThin('E2E8F0')
    })

    // Ước lượng số dòng văn bản sẽ tự xuống hàng (wrap) trong 1 ô, dựa theo
    // độ rộng cột (ký tự) để tính chiều cao dòng vừa đủ, không bị cắt chữ.
    const estimateWrappedLines = (text, widthChars) => {
      if (text === null || text === undefined || text === '') return 1
      const charsPerLine = Math.max(6, Math.round(widthChars * 0.95))
      return String(text)
        .split('\n')
        .reduce((total, line) => total + Math.max(1, Math.ceil(line.length / charsPerLine)), 0)
    }

    // Dựng 1 sheet hoàn chỉnh: dòng tiêu đề + phụ đề + bảng dữ liệu canh chỉnh đẹp
    const buildSheet = ({ title, subtitle, columns, rows, headerColor, headerBorderColor }) => {
      const ws = {}
      const colCount = columns.length
      const lastColLetter = XLSX.utils.encode_col(colCount - 1)

      // Dòng 1: Tiêu đề lớn (merge toàn bộ chiều rộng bảng)
      ws['A1'] = { v: title, t: 's', s: titleStyle }
      ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } }]
      // Dòng 2: Phụ đề (ngày xuất, số lượng)
      ws['A2'] = { v: subtitle, t: 's', s: subtitleStyle }
      ws['!merges'].push({ s: { r: 1, c: 0 }, e: { r: 1, c: colCount - 1 } })

      // Dòng 4 (chừa 1 dòng trống ở dòng 3): tiêu đề cột
      const headerRow = 3
      const headerStyle = makeHeaderStyle(headerColor, headerBorderColor)
      columns.forEach((col, cIdx) => {
        const addr = XLSX.utils.encode_cell({ r: headerRow, c: cIdx })
        ws[addr] = { v: col.label, t: 's', s: headerStyle }
      })

      // Dữ liệu (căn lề theo từng cột, tô sọc xen kẽ để dễ đọc)
      rows.forEach((row, rIdx) => {
        const excelRow = headerRow + 1 + rIdx
        const zebra = rIdx % 2 === 1
        columns.forEach((col, cIdx) => {
          const addr = XLSX.utils.encode_cell({ r: excelRow, c: cIdx })
          const rawVal = row[col.key]
          const isNumber = typeof rawVal === 'number'
          let cellStyle = makeDataStyle(col.align || (isNumber ? 'center' : 'left'), zebra)
          if (col.colorize) {
            const extra = col.colorize(rawVal, row)
            if (extra) cellStyle = { ...cellStyle, font: { ...cellStyle.font, ...extra } }
          }
          ws[addr] = { v: rawVal ?? '', t: isNumber ? 'n' : 's', s: cellStyle }
        })
      })

      const totalRows = headerRow + rows.length
      ws['!ref'] = `A1:${lastColLetter}${totalRows + 1}`
      ws['!cols'] = columns.map(c => ({ wch: c.width }))

      // Chiều cao dòng tiêu đề cột: đủ chỗ cho nhãn cột dài nhất bị wrap
      const headerLines = Math.max(1, ...columns.map(c => estimateWrappedLines(c.label, c.width)))
      const headerHeight = Math.max(24, headerLines * 14 + 10)

      // Các dòng dữ liệu: KHÔNG đặt chiều cao cố định — để Excel tự autofit
      // chiều cao theo nội dung wrap text thật sự (chính xác hơn nhiều so với
      // việc tự ước lượng số dòng). Excel chỉ autofit khi dòng không có
      // customHeight cố định, nên ta chỉ set chiều cao cho các dòng tiêu đề.
      ws['!rows'] = [
        { hpt: 26 }, { hpt: 18 }, { hpt: 8 }, { hpt: headerHeight }
      ]
      return ws
    }

    // ── Sheet 1: Tổng hợp danh sách BCH sau chuẩn hóa (1 dòng = 1 BCH) ──────
    // Sắp xếp theo thứ tự bảng chữ cái A→Z theo Tên BCH Chuẩn hóa (không phân biệt hoa/thường, có dấu)
    const viCollator = new Intl.Collator('vi', { sensitivity: 'base', numeric: true })
    const sortedBchGroups = [...standardizedBchGroups].sort((a, b) => viCollator.compare(a.standardName, b.standardName))
    const sheet1Rows = sortedBchGroups.map((g, idx) => ({
      stt: idx + 1,
      ten: g.standardName,
      loai: g.isMultiMerged ? 'Đã gộp nhiều kho' : 'BCH độc lập (1 kho)',
      soKho: g.memberCount,
      danhSachKho: g.members.map(m => m.name).join('; '),
      tongPhieu: g.totalPhieu,
      giao: g.totalGiao,
      nhan: g.totalNhan,
      ghiChu: g.note || ''
    }))
    const ws1 = buildSheet({
      title: 'DANH SÁCH BAN CHỈ HUY SAU CHUẨN HÓA',
      subtitle: `Xuất ngày ${today}  •  Tổng số: ${sheet1Rows.length} Ban Chỉ Huy  •  Từ ${allOriginalKhos.length} kho gốc`,
      headerColor: '0F766E',
      headerBorderColor: '0D9488',
      columns: [
        { key: 'stt', label: 'STT', width: 6, align: 'center' },
        { key: 'ten', label: 'Tên BCH Chuẩn hóa', width: 34, align: 'left', colorize: () => ({ bold: true }) },
        {
          key: 'loai', label: 'Loại', width: 20, align: 'center',
          colorize: (v) => v === 'Đã gộp nhiều kho' ? { bold: true, color: { rgb: '7C3AED' } } : { color: { rgb: '64748B' } }
        },
        { key: 'soKho', label: 'Số kho gốc', width: 11, align: 'center' },
        { key: 'danhSachKho', label: 'Danh sách Kho gốc', width: 60, align: 'left' },
        { key: 'tongPhieu', label: 'Tổng số phiếu phát sinh', width: 16, align: 'center', colorize: () => ({ bold: true }) },
        { key: 'giao', label: 'Số phiếu Giao', width: 13, align: 'center' },
        { key: 'nhan', label: 'Số phiếu Nhận', width: 13, align: 'center' },
        { key: 'ghiChu', label: 'Ghi chú', width: 22, align: 'left' }
      ],
      rows: sheet1Rows
    })
    XLSX.utils.book_append_sheet(wb, ws1, 'Danh_Sach_BCH_Chuan_Hoa')

    // ── Sheet 2: Chi tiết từng Kho gốc & Tên chuẩn tương ứng ────────────────
    // Sắp xếp theo thứ tự bảng chữ cái A→Z theo Tên Kho BCH (Gốc)
    const sortedOriginalKhos = [...allOriginalKhos].sort((a, b) => viCollator.compare(a.name, b.name))
    const sheet2Rows = sortedOriginalKhos.map((item, idx) => ({
      stt: idx + 1,
      tenKho: item.name,
      trangThai: item.isMapped ? 'Đã gộp vào tên chuẩn' : 'Tên gốc (Độc lập)',
      tenChuan: item.currentStandardName,
      tongPhieu: item.totalCount,
      giao: item.giaoCount,
      nhan: item.nhanCount
    }))
    const ws2 = buildSheet({
      title: 'CHI TIẾT KHO GỐC & TÊN CHUẨN TƯƠNG ỨNG',
      subtitle: `Xuất ngày ${today}  •  Tổng số: ${sheet2Rows.length} kho gốc`,
      headerColor: '1E40AF',
      headerBorderColor: '2563EB',
      columns: [
        { key: 'stt', label: 'STT', width: 6, align: 'center' },
        { key: 'tenKho', label: 'Tên Kho BCH (Gốc)', width: 36, align: 'left' },
        {
          key: 'trangThai', label: 'Trạng thái', width: 24, align: 'center',
          colorize: (v) => v === 'Đã gộp vào tên chuẩn' ? { bold: true, color: { rgb: '15803D' } } : { color: { rgb: '64748B' } }
        },
        { key: 'tenChuan', label: 'Tên BCH Chuẩn hóa', width: 34, align: 'left', colorize: () => ({ bold: true, color: { rgb: '1E40AF' } }) },
        { key: 'tongPhieu', label: 'Tổng phiếu phát sinh', width: 16, align: 'center' },
        { key: 'giao', label: 'Số phiếu Giao', width: 13, align: 'center' },
        { key: 'nhan', label: 'Số phiếu Nhận', width: 13, align: 'center' }
      ],
      rows: sheet2Rows
    })
    XLSX.utils.book_append_sheet(wb, ws2, 'Chi_Tiet_Kho_Goc')

    // ── Sheet 3: Quy tắc chuẩn hóa (dùng để đồng bộ Supabase) ───────────────
    // Sắp xếp theo thứ tự bảng chữ cái A→Z theo Tên cũ / Biến thể
    const sortedAliasRules = [...bchAliasRules].sort((a, b) => viCollator.compare(a.ten_cu, b.ten_cu))
    const sheet3Rows = sortedAliasRules.map((r, idx) => ({
      stt: idx + 1,
      tenCu: r.ten_cu,
      tenChuan: r.ten_chuan,
      ghiChu: r.ghi_chu || ''
    }))
    const ws3 = buildSheet({
      title: 'QUY TẮC CHUẨN HÓA (DỮ LIỆU SUPABASE)',
      subtitle: `Xuất ngày ${today}  •  Tổng số: ${sheet3Rows.length} quy tắc`,
      headerColor: '475569',
      headerBorderColor: '64748B',
      columns: [
        { key: 'stt', label: 'STT', width: 6, align: 'center' },
        { key: 'tenCu', label: 'Tên cũ / Biến thể (ten_cu)', width: 36, align: 'left' },
        { key: 'tenChuan', label: 'Tên BCH Chuẩn (ten_chuan)', width: 36, align: 'left', colorize: () => ({ bold: true, color: { rgb: '0F766E' } }) },
        { key: 'ghiChu', label: 'Ghi chú', width: 30, align: 'left' }
      ],
      rows: sheet3Rows
    })
    XLSX.utils.book_append_sheet(wb, ws3, 'Quy_Tac_Supabase')

    XLSX.writeFile(wb, `SGC_Danh_Sach_BCH_Chuan_Hoa_${new Date().toISOString().slice(0, 10)}.xlsx`)
  }

  // Esc đóng hộp thoại "Tạo nhóm BCH chuẩn mới"
  useEscClose(showCreateModal, () => setShowCreateModal(false))

  const fmt = (n) => (Number(n) || 0).toLocaleString('vi-VN')
  const hasFilter = Boolean(searchTerm) || statusFilter !== 'all'

  return (
    <div
      style={{ padding: '10px 16px 12px', maxWidth: '1600px', margin: '0 auto', display: 'flex', flexDirection: 'column', gap: '10px' }}

    >
      <style>{CHB_CSS}</style>

      {/* ─── DÃY SỐ LIỆU GỌN ───────────────────────────────────────────────── */}
      <div className="sg-kpis">
        <div className="sg-kpi sm sg-c-blue" title={`Mỗi dòng là 1 tên BCH chuẩn hóa, gộp các tên kho biến thể về một Ban Chỉ Huy. Giảm từ ${fmt(allOriginalKhos.length)} kho gốc ban đầu.`}>
          <span className="sg-kpi-ic"><Building2 size={16} /></span>
          <span className="sg-kpi-bd">
            <span className="sg-kpi-lb">BCH sau chuẩn hóa</span>
            <span className="sg-kpi-v">{fmt(standardizedBchGroups.length)}<small>BCH</small></span>
          </span>
        </div>
        <div className="sg-kpi sm sg-c-violet" title={`Đã gộp ${fmt(totalMergedOldKhoCount)} kho gốc vào các nhóm này`}>
          <span className="sg-kpi-ic"><Layers size={16} /></span>
          <span className="sg-kpi-bd">
            <span className="sg-kpi-lb">Đã gộp nhiều kho</span>
            <span className="sg-kpi-v">{fmt(multiMergedGroupsCount)}<small>nhóm</small></span>
          </span>
        </div>
        <div className="sg-kpi sm sg-c-orange" title="1 kho = 1 BCH, chưa gộp thêm kho khác">
          <span className="sg-kpi-ic"><Info size={16} /></span>
          <span className="sg-kpi-bd">
            <span className="sg-kpi-lb">BCH độc lập</span>
            <span className="sg-kpi-v">{fmt(standaloneGroupsCount)}<small>BCH</small></span>
          </span>
        </div>
        <div className="sg-kpi sm sg-c-green" title={`${fmt(bchAliasRules.length)} quy tắc chuẩn hóa đang áp dụng`}>
          <span className="sg-kpi-ic"><Tag size={16} /></span>
          <span className="sg-kpi-bd">
            <span className="sg-kpi-lb">Kho BCH gốc trong dữ liệu</span>
            <span className="sg-kpi-v">{fmt(allOriginalKhos.length)}<small>kho</small></span>
          </span>
        </div>
      </div>

      {/* ─── THÔNG BÁO TRẠNG THÁI ─────────────────────────────────────────── */}
      {statusMessage && (
        <div style={{
          padding: '8px 14px',
          borderRadius: '14px',
          display: 'flex',
          alignItems: 'center',
          gap: '8px',
          fontSize: '15px',
          fontWeight: 600,
          background: saveStatus === 'error' ? '#FFE7E5' : '#DDF5E6',
          color: saveStatus === 'error' ? '#B42318' : '#0D7A3C'
        }}>
          {saveStatus === 'error' ? <AlertCircle size={18} /> : <CheckCircle2 size={18} />}
          <span>{statusMessage}</span>
        </div>
      )}

      {/* ─── KHUNG BẢNG: HÀNG CÔNG CỤ + BẢNG ───────────────────────────────── */}
      <div className="chb-card">
        <div className="chb-tb">
          <div className="chb-search">
            <Search size={16} className="chb-lead" />
            <input
              className="chb-in"
              type="text"
              placeholder="Tìm tên BCH chuẩn hóa hoặc tên kho gốc..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
            />
            {searchTerm && (
              <button type="button" className="chb-x sm" onClick={() => setSearchTerm('')} title="Xoá ô tìm">
                <XMark size={11} w={2.3} />
              </button>
            )}
          </div>

          <div className="sg-seg" title="Lọc hiển thị">
            <button type="button" className={statusFilter === 'all' ? 'on' : ''} onClick={() => setStatusFilter('all')}>
              Tất cả ({fmt(standardizedBchGroups.length)})
            </button>
            <button type="button" className={statusFilter === 'merged_multi' ? 'on' : ''} onClick={() => setStatusFilter('merged_multi')}>
              Đã gộp nhiều kho ({fmt(multiMergedGroupsCount)})
            </button>
            <button type="button" className={statusFilter === 'standalone' ? 'on' : ''} onClick={() => setStatusFilter('standalone')}>
              BCH độc lập ({fmt(standaloneGroupsCount)})
            </button>
          </div>

          <button
            type="button"
            className="sg-btn"
            onClick={() => {
              setNewBchStandardName('')
              setNewBchSelectedKhos(new Set())
              setNewBchSearchKho('')
              setNewBchNote('')
              setShowCreateModal(true)
            }}
            title="Tạo nhóm BCH chuẩn mới và chọn các kho gốc gộp vào"
          >
            <Plus size={16} /><span className="chb-lg">Tạo nhóm BCH</span><span className="chb-sh">Tạo nhóm</span>
          </button>

          {isSupabaseConfigured && (
            <button
              type="button"
              className="sg-btn green"
              onClick={handleSaveToSupabase}
              disabled={saveStatus === 'saving'}
              title="Lưu toàn bộ quy tắc chuẩn hóa lên Supabase"
            >
              <Save size={15} />
              {saveStatus === 'saving'
                ? 'Đang lưu...'
                : (<><span className="chb-lg">Lưu lên Supabase</span><span className="chb-sh">Lưu</span></>)}
            </button>
          )}

          {onReloadRulesFromSupabase && (
            <button
              type="button"
              className="sg-btn ghost chb-ico-only"
              onClick={onReloadRulesFromSupabase}
              title="Tải lại quy tắc từ cơ sở dữ liệu Supabase"
            >
              <RefreshCw size={15} /><span className="chb-lb">Tải lại</span>
            </button>
          )}

          <button type="button" onClick={handleExportExcel} style={{ marginLeft: 'auto' }} title="Xuất danh sách BCH chuẩn hóa ra Excel">
            <FileSpreadsheet size={15} /> Xuất Excel
          </button>
        </div>

        {/* Bảng: mỗi dòng là 1 tên BCH chuẩn hóa */}
        <div className="chb-scroll" style={{ maxHeight: 'calc(100vh - 282px)', minHeight: '320px' }}>
          <table style={{ fontSize: '15px' }}>
            <colgroup>
              <col style={{ width: '56px' }} />
              <col style={{ width: '27%' }} />
              <col />
              <col style={{ width: '96px' }} />
              <col style={{ width: '150px' }} />
              <col style={{ width: '13%' }} />
              <col style={{ width: '84px' }} />
            </colgroup>
            <thead>
              <tr>
                <th>STT</th>
                <th>Tên BCH chuẩn hóa</th>
                <th>Kho BCH gốc đã gộp vào</th>
                <th>Số kho gộp</th>
                <th>Tổng phát sinh</th>
                <th>Ghi chú</th>
                <th>Thao tác</th>
              </tr>
            </thead>
            <tbody>
              {filteredStandardGroups.length === 0 ? (
                <tr>
                  <td colSpan={7} style={{ padding: '40px 20px', textAlign: 'center', color: '#8C8E97' }}>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px' }}>
                      <Building2 size={32} color="#B4BBC9" />
                      <div style={{ fontSize: '15px', fontWeight: 600, color: '#4A4C55' }}>
                        Không có Ban Chỉ Huy nào khớp với bộ lọc / từ khóa tìm
                      </div>
                      {hasFilter && (
                        <button
                          type="button"
                          className="sg-btn danger"
                          onClick={() => { setSearchTerm(''); setStatusFilter('all') }}
                        >
                          <XMark size={12} w={2.3} /> Bỏ lọc
                        </button>
                      )}
                    </div>
                  </td>
                </tr>
              ) : (
                filteredStandardGroups.map((group, index) => {
                  const isEditing = editingStandardKey === group.key

                  return (
                    <tr key={group.key}>
                      {/* STT */}
                      <td className="sg-num" style={{ textAlign: 'center', color: '#4A4C55', verticalAlign: 'top', paddingTop: '10px' }}>
                        {index + 1}
                      </td>

                      {/* TÊN BCH CHUẨN HÓA */}
                      <td style={{ verticalAlign: 'top' }}>
                        {isEditing ? (
                          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            <input
                              type="text"
                              className="chb-edit"
                              value={editingStandardValue}
                              onChange={(e) => setEditingStandardValue(e.target.value)}
                              placeholder="Nhập tên BCH chuẩn hóa mới..."
                              autoFocus
                              onFocus={(e) => { const l = e.target.value.length; e.target.setSelectionRange(l, l) }}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') handleSaveRenameGroup(group)
                                if (e.key === 'Escape') { e.stopPropagation(); setEditingStandardKey(null) }
                              }}
                            />
                            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                              <button type="button" className="chb-mini ok" onClick={() => handleSaveRenameGroup(group)}>
                                <Check size={13} /> Lưu
                              </button>
                              <button type="button" className="chb-mini no" onClick={() => setEditingStandardKey(null)}>
                                Hủy
                              </button>
                              <span className="chb-hint">Enter lưu · Esc huỷ</span>
                            </div>
                          </div>
                        ) : (
                          <div
                            onDoubleClick={() => {
                              setEditingStandardKey(group.key)
                              setEditingStandardValue(group.standardName)
                              setEditingNoteValue(group.note || '')
                            }}
                          >
                            <div style={{
                              fontWeight: 700,
                              color: group.isMultiMerged ? '#0A3D73' : '#16171C',
                              fontSize: '15px',
                              lineHeight: 1.4
                            }}>
                              <span>{group.standardName}</span>
                              <button
                                type="button"
                                className="chb-pen"
                                style={{ marginLeft: '4px', marginTop: '-2px' }}
                                onClick={() => {
                                  setEditingStandardKey(group.key)
                                  setEditingStandardValue(group.standardName)
                                  setEditingNoteValue(group.note || '')
                                }}
                                title="Đổi tên Ban Chỉ Huy chuẩn hóa (hoặc bấm đúp vào tên)"
                              >
                                <PenIcon />
                              </button>
                            </div>
                            <div style={{ marginTop: '4px' }}>
                              {group.isMultiMerged ? (
                                <span className="chb-tag multi"><Layers size={12} /> Đã gộp {group.memberCount} kho gốc</span>
                              ) : (
                                <span className="chb-tag solo">BCH độc lập (1 kho)</span>
                              )}
                            </div>
                          </div>
                        )}
                      </td>

                      {/* CÁC KHO BCH GỐC ĐÃ GỘP */}
                      <td style={{ verticalAlign: 'top' }}>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', alignItems: 'center', minWidth: 0 }}>
                          {group.members.map((member) => {
                            const isExactSameName = normKey(member.name) === normKey(group.standardName)
                            const canRemove = !isExactSameName || group.members.length > 1

                            return (
                              <span
                                key={member.name}
                                className={'chb-chip' + (isExactSameName ? ' main' : '') + (canRemove ? '' : ' solo')}
                                title={`${member.name} (${fmt(member.totalCount)} phiếu)${isExactSameName ? ' – trùng tên chuẩn' : ''}`}
                              >
                                <span className="chb-nm">{member.name}</span>
                                <span className="chb-ct">{fmt(member.totalCount)} phiếu</span>
                                {canRemove && (
                                  <button
                                    type="button"
                                    className="chb-x sm"
                                    onClick={() => handleRemoveKhoFromGroup(member.name)}
                                    title={`Gỡ "${member.name}" ra khỏi nhóm này`}
                                  >
                                    <XMark size={11} w={2.3} />
                                  </button>
                                )}
                              </span>
                            )
                          })}

                          <button
                            type="button"
                            className="chb-add"
                            onClick={() => setAssignModalGroup(group)}
                            title={`Chọn thêm kho gốc để gộp vào Ban Chỉ Huy "${group.standardName}"`}
                          >
                            <Plus size={13} /> Gán thêm kho
                          </button>
                        </div>
                      </td>

                      {/* SỐ KHO GỘP */}
                      <td style={{ textAlign: 'center', verticalAlign: 'top' }}>
                        <span className={'chb-tag ' + (group.isMultiMerged ? 'multi' : 'solo')} style={{ height: '28px', padding: '0 12px' }}>
                          {group.memberCount} kho
                        </span>
                      </td>

                      {/* TỔNG PHÁT SINH */}
                      <td style={{ textAlign: 'center', verticalAlign: 'top' }}>
                        <div style={{ fontWeight: 700, color: '#16171C', fontSize: '15px', lineHeight: 1.4 }}>
                          <span className="sg-num">{fmt(group.totalPhieu)}</span> <span style={{ fontSize: '12px', color: '#4A4C55' }}>phiếu</span>
                        </div>
                        <div style={{ fontSize: '12px', color: '#8C8E97', fontWeight: 600, marginTop: '2px', whiteSpace: 'nowrap' }}>
                          {fmt(group.totalGiao)} giao / {fmt(group.totalNhan)} nhận
                        </div>
                      </td>

                      {/* GHI CHÚ */}
                      <td style={{ color: '#4A4C55', fontSize: '12px', verticalAlign: 'top', paddingTop: '10px' }}>
                        {group.note || '-'}
                      </td>

                      {/* THAO TÁC */}
                      <td style={{ textAlign: 'center', verticalAlign: 'top' }}>
                        {group.hasRules ? (
                          <button
                            type="button"
                            className="chb-x"
                            onClick={() => handleDeleteGroup(group)}
                            title="Hủy gộp nhóm này (trả tất cả về kho gốc ban đầu)"
                          >
                            <XMark size={13} w={2.3} />
                          </button>
                        ) : (
                          <span style={{ color: '#8C8E97' }}>-</span>
                        )}
                      </td>
                    </tr>
                  )
                })
              )}
            </tbody>
          </table>
        </div>

        <div className="chb-foot">
          Hiển thị {fmt(filteredStandardGroups.length)} / {fmt(standardizedBchGroups.length)} Ban Chỉ Huy chuẩn hóa · tổng hợp từ {fmt(allOriginalKhos.length)} kho BCH gốc
        </div>
      </div>

      {/* ─── HỘP THOẠI TẠO NHÓM BCH CHUẨN HÓA MỚI ─────────────────────────── */}
      {showCreateModal && (
        <div className="chb-ov">
          <div className="chb-dlg" style={{ maxWidth: '700px' }}>
            <div className="chb-dlg-hd">
              <div className="chb-dlg-ttl">
                <span className="chb-dlg-ic"><Plus size={18} /></span>
                <h3>Tạo nhóm Ban Chỉ Huy chuẩn hóa mới</h3>
              </div>
              <CloseEscButton onClick={() => setShowCreateModal(false)} />
            </div>

            <div className="chb-dlg-bd">
              <div>
                <label className="chb-lbl" style={{ marginBottom: '6px' }}>
                  1. Tên Ban Chỉ Huy chuẩn hóa thống nhất <span style={{ color: '#D92D3A' }}>*</span>
                </label>
                <input
                  type="text"
                  className="chb-fld"
                  placeholder="Ví dụ: BCH Cọc Khoan Nhồi - Trống Đồng..."
                  value={newBchStandardName}
                  onChange={(e) => setNewBchStandardName(e.target.value)}
                />
              </div>

              <div>
                <label className="chb-lbl" style={{ marginBottom: '6px' }}>
                  2. Chọn các kho BCH gốc để gộp vào nhóm ({newBchSelectedKhos.size} đã chọn) <span style={{ color: '#D92D3A' }}>*</span>
                </label>
                <div className="chb-search" style={{ maxWidth: 'none', marginBottom: '8px' }}>
                  <Search size={16} className="chb-lead" />
                  <input
                    className="chb-in"
                    type="text"
                    placeholder="Tìm nhanh tên kho gốc..."
                    value={newBchSearchKho}
                    onChange={(e) => setNewBchSearchKho(e.target.value)}
                  />
                  {newBchSearchKho && (
                    <button type="button" className="chb-x sm" onClick={() => setNewBchSearchKho('')} title="Xoá ô tìm">
                      <XMark size={11} w={2.3} />
                    </button>
                  )}
                </div>

                <div className="chb-list" style={{ maxHeight: '240px' }}>
                  {allOriginalKhos
                    .filter(k => !newBchSearchKho || k.name.toLowerCase().includes(newBchSearchKho.toLowerCase().trim()))
                    .map(k => {
                      const isChecked = newBchSelectedKhos.has(k.name)
                      return (
                        <div
                          key={k.name}
                          className={'chb-row' + (isChecked ? ' on' : '')}
                          onClick={() => {
                            setNewBchSelectedKhos(prev => {
                              const next = new Set(prev)
                              if (next.has(k.name)) next.delete(k.name)
                              else next.add(k.name)
                              return next
                            })
                          }}
                        >
                          {isChecked ? <CheckSquare size={18} color="#0F58A7" /> : <Square size={18} color="#B4BBC9" />}
                          <span className="chb-nm" title={k.name}>{k.name}</span>
                          <span className="chb-tag solo sg-num">{fmt(k.totalCount)} phiếu</span>
                        </div>
                      )
                    })}
                </div>
              </div>

              <div>
                <label className="chb-lbl" style={{ marginBottom: '6px', color: '#4A4C55' }}>
                  3. Ghi chú (tùy chọn)
                </label>
                <input
                  type="text"
                  className="chb-fld"
                  placeholder="Ghi chú thêm về nhóm BCH này..."
                  value={newBchNote}
                  onChange={(e) => setNewBchNote(e.target.value)}
                />
              </div>
            </div>

            <div className="chb-dlg-ft" style={{ justifyContent: 'flex-end' }}>
              <button type="button" className="sg-btn ghost" onClick={() => setShowCreateModal(false)}>
                Hủy bỏ
              </button>
              <button
                type="button"
                className="sg-btn"
                onClick={handleCreateNewGroup}
                disabled={!newBchStandardName || newBchSelectedKhos.size === 0}
              >
                <Check size={16} /> Tạo nhóm chuẩn ({newBchSelectedKhos.size} kho)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ─── HỘP THOẠI GÁN KHO GỐC VÀO BCH CHUẨN HÓA ──────────────────────── */}
      <AssignKhoToGroupModal
        isOpen={Boolean(assignModalGroup)}
        onClose={() => setAssignModalGroup(null)}
        targetGroup={assignModalGroup}
        allOriginalKhos={allOriginalKhos}
        onConfirmAdd={(targetStandardName, khos) => handleAddKhosToGroup(targetStandardName, khos)}
      />
    </div>
  )
}
