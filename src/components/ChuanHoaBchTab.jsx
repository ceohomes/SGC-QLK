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

  if (!isOpen || !targetGroup) return null

  return (
    <div style={{
      position: 'fixed',
      top: 0,
      left: 0,
      right: 0,
      bottom: 0,
      background: 'rgba(15, 23, 42, 0.65)',
      backdropFilter: 'blur(3px)',
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      zIndex: 99999,
      padding: '20px'
    }}>
      <div style={{
        background: '#ffffff',
        borderRadius: '14px',
        maxWidth: '750px',
        width: '100%',
        maxHeight: '90vh',
        display: 'flex',
        flexDirection: 'column',
        boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)',
        overflow: 'hidden',
        border: '1px solid #cbd5e1'
      }}>
        {/* Header Modal */}
        <div style={{
          padding: '16px 22px',
          background: 'linear-gradient(135deg, #00529C 0%, #1e40af 100%)',
          color: '#ffffff',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          borderBottom: '1px solid rgba(255, 255, 255, 0.15)'
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <div style={{
              width: 36,
              height: 36,
              borderRadius: '8px',
              background: 'rgba(255, 255, 255, 0.2)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center'
            }}>
              <FolderPlus size={20} color="#ffffff" />
            </div>
            <div>
              <h3 style={{ fontSize: '16px', fontWeight: 800, margin: 0, color: '#ffffff' }}>
                Gán Kho Gốc vào Ban Chỉ Huy
              </h3>
              <div style={{ fontSize: '13px', color: '#bfdbfe', marginTop: '2px', fontWeight: 600 }}>
                {targetGroup.standardName}
              </div>
            </div>
          </div>

          <button
            onClick={onClose}
            style={{
              background: 'rgba(255, 255, 255, 0.1)',
              border: 'none',
              borderRadius: '6px',
              cursor: 'pointer',
              color: '#ffffff',
              padding: '6px',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              transition: 'background 0.15s'
            }}
          >
            <X size={18} />
          </button>
        </div>

        {/* Body Modal */}
        <div style={{ padding: '18px 22px', overflowY: 'auto', flex: 1, display: 'flex', flexDirection: 'column', gap: '14px' }}>
          
          {/* Section 1: Hiện các kho đã thuộc nhóm này */}
          <div style={{
            background: '#f8fafc',
            border: '1px solid #e2e8f0',
            borderRadius: '8px',
            padding: '12px 14px'
          }}>
            <div style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569', marginBottom: '8px', display: 'flex', alignItems: 'center', gap: '6px' }}>
              <Layers size={14} color="#2563eb" />
              <span>Các Kho Gốc hiện đang thuộc Ban Chỉ Huy này ({currentMemberNames.length} kho):</span>
            </div>

            <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', maxHeight: '90px', overflowY: 'auto' }}>
              {targetGroup.members && targetGroup.members.map(m => (
                <span
                  key={m.name}
                  style={{
                    fontSize: '12px',
                    padding: '3px 9px',
                    background: '#eff6ff',
                    color: '#1e40af',
                    borderRadius: '6px',
                    border: '1px solid #bfdbfe',
                    fontWeight: 600,
                    display: 'inline-flex',
                    alignItems: 'center',
                    gap: '5px'
                  }}
                >
                  <span>{m.name}</span>
                  <span style={{ fontSize: '11px', color: '#60a5fa', fontWeight: 400 }}>({m.totalCount || 0}p)</span>
                </span>
              ))}
            </div>
          </div>

          {/* Section 2: Chọn thêm kho gốc để gộp vào */}
          <div>
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '8px', flexWrap: 'wrap', gap: '8px' }}>
              <label style={{ fontSize: '13px', fontWeight: 700, color: '#0f172a', display: 'flex', alignItems: 'center', gap: '6px' }}>
                <Plus size={15} color="#16a34a" />
                <span>Chọn các Kho Gốc cần gộp thêm vào nhóm này:</span>
              </label>

              {/* Filter Tabs */}
              <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                <button
                  onClick={() => setFilterType('all')}
                  style={{
                    padding: '4px 10px',
                    borderRadius: '14px',
                    fontSize: '11.5px',
                    fontWeight: 600,
                    cursor: 'pointer',
                    border: filterType === 'all' ? '1px solid #2563eb' : '1px solid #e2e8f0',
                    background: filterType === 'all' ? '#eff6ff' : '#ffffff',
                    color: filterType === 'all' ? '#2563eb' : '#64748b'
                  }}
                >
                  Tất cả ({availableKhos.length})
                </button>
                <button
                  onClick={() => setFilterType('unmerged')}
                  style={{
                    padding: '4px 10px',
                    borderRadius: '14px',
                    fontSize: '11.5px',
                    fontWeight: 600,
                    cursor: 'pointer',
                    border: filterType === 'unmerged' ? '1px solid #d97706' : '1px solid #e2e8f0',
                    background: filterType === 'unmerged' ? '#fffbeb' : '#ffffff',
                    color: filterType === 'unmerged' ? '#d97706' : '#64748b'
                  }}
                >
                  Chưa gộp ({availableKhos.filter(k => k.isStandalone).length})
                </button>
                <button
                  onClick={() => setFilterType('other')}
                  style={{
                    padding: '4px 10px',
                    borderRadius: '14px',
                    fontSize: '11.5px',
                    fontWeight: 600,
                    cursor: 'pointer',
                    border: filterType === 'other' ? '1px solid #7c3aed' : '1px solid #e2e8f0',
                    background: filterType === 'other' ? '#f5f3ff' : '#ffffff',
                    color: filterType === 'other' ? '#7c3aed' : '#64748b'
                  }}
                >
                  Thuộc BCH khác ({availableKhos.filter(k => k.isAssignedOther).length})
                </button>
              </div>
            </div>

            {/* Search Input Box */}
            <div style={{
              display: 'flex',
              alignItems: 'center',
              gap: '8px',
              background: '#ffffff',
              border: '1.5px solid #cbd5e1',
              borderRadius: '8px',
              padding: '8px 12px',
              marginBottom: '10px'
            }}>
              <Search size={16} color="#64748b" />
              <input
                ref={inputRef}
                type="text"
                placeholder="🔍 Gõ tìm nhanh theo tên kho gốc..."
                value={search}
                onChange={(e) => setSearch(e.target.value)}
                style={{
                  flex: 1,
                  border: 'none',
                  background: 'transparent',
                  fontSize: '13.5px',
                  outline: 'none',
                  color: '#0f172a'
                }}
              />
              {search && (
                <button
                  onClick={() => setSearch('')}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8', padding: 0 }}
                >
                  <X size={14} />
                </button>
              )}
            </div>

            {/* Quick Action Bar (Select all) */}
            <div style={{
              padding: '8px 12px',
              background: '#f1f5f9',
              borderRadius: '6px 6px 0 0',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              fontSize: '12px',
              color: '#475569',
              border: '1px solid #e2e8f0',
              borderBottom: 'none'
            }}>
              <span>
                Tìm thấy <strong>{filteredList.length}</strong> kho gốc có thể gán
              </span>
              {filteredList.length > 0 && (
                <button
                  onClick={handleSelectAllFiltered}
                  style={{
                    background: 'none',
                    border: 'none',
                    color: '#2563eb',
                    fontWeight: 700,
                    cursor: 'pointer',
                    fontSize: '12px'
                  }}
                >
                  {selectedKhos.size === filteredList.length ? 'Bỏ chọn tất cả' : `Chọn tất cả (${filteredList.length})`}
                </button>
              )}
            </div>

            {/* Scrollable Checklist of available original khos */}
            <div style={{
              maxHeight: '260px',
              overflowY: 'auto',
              border: '1px solid #e2e8f0',
              borderRadius: '0 0 8px 8px',
              background: '#ffffff'
            }}>
              {filteredList.length === 0 ? (
                <div style={{ padding: '36px 16px', textAlign: 'center', color: '#94a3b8', fontSize: '13px' }}>
                  Không tìm thấy kho gốc nào phù hợp với từ khóa tìm kiếm
                </div>
              ) : (
                filteredList.map((k, idx) => {
                  const isChecked = selectedKhos.has(k.name)

                  return (
                    <div
                      key={k.name}
                      onClick={() => toggleSelect(k.name)}
                      style={{
                        padding: '9px 14px',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        gap: '12px',
                        cursor: 'pointer',
                        background: isChecked ? '#eff6ff' : (idx % 2 === 0 ? '#ffffff' : '#fafafa'),
                        borderBottom: '1px solid #f1f5f9',
                        transition: 'background 0.1s'
                      }}
                      onMouseEnter={(e) => { if (!isChecked) e.currentTarget.style.background = '#f8fafc' }}
                      onMouseLeave={(e) => { if (!isChecked) e.currentTarget.style.background = idx % 2 === 0 ? '#ffffff' : '#fafafa' }}
                    >
                      <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flex: 1, minWidth: 0 }}>
                        <div style={{ display: 'flex', alignItems: 'center' }}>
                          {isChecked ? (
                            <CheckSquare size={18} color="#2563eb" />
                          ) : (
                            <Square size={18} color="#cbd5e1" />
                          )}
                        </div>

                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{
                            fontSize: '13px',
                            fontWeight: isChecked ? 700 : 500,
                            color: isChecked ? '#1d4ed8' : '#1e293b',
                            overflow: 'hidden',
                            textOverflow: 'ellipsis',
                            whiteSpace: 'nowrap'
                          }}>
                            {k.name}
                          </div>
                        </div>
                      </div>

                      {/* Info Badges */}
                      <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexShrink: 0 }}>
                        {k.isAssignedOther ? (
                          <span style={{
                            fontSize: '11px',
                            color: '#7c3aed',
                            background: '#f5f3ff',
                            border: '1px solid #ddd6fe',
                            padding: '2px 7px',
                            borderRadius: '4px',
                            maxWidth: '180px',
                            overflow: 'hidden',
                            textOverflow: 'ellipsis',
                            whiteSpace: 'nowrap'
                          }}>
                            Đang thuộc: {k.currentStd}
                          </span>
                        ) : (
                          <span style={{
                            fontSize: '11px',
                            color: '#d97706',
                            background: '#fffbeb',
                            border: '1px solid #fef3c7',
                            padding: '2px 7px',
                            borderRadius: '4px'
                          }}>
                            Chưa gộp
                          </span>
                        )}

                        <span style={{
                          fontSize: '11.5px',
                          color: '#64748b',
                          fontWeight: 600,
                          background: '#f1f5f9',
                          padding: '2px 7px',
                          borderRadius: '4px',
                          minWidth: '55px',
                          textAlign: 'right'
                        }}>
                          {k.totalCount || 0} phiếu
                        </span>
                      </div>
                    </div>
                  )
                })
              )}
            </div>
          </div>
        </div>

        {/* Footer Actions */}
        <div style={{
          padding: '14px 22px',
          borderTop: '1px solid #e2e8f0',
          background: '#f8fafc',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          gap: '12px'
        }}>
          <div style={{ fontSize: '13px', color: '#475569' }}>
            Đã chọn: <strong style={{ color: '#2563eb', fontSize: '14px' }}>{selectedKhos.size}</strong> kho gốc để gán thêm
          </div>

          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <button
              onClick={onClose}
              style={{
                padding: '8px 16px',
                borderRadius: '8px',
                border: '1px solid #cbd5e1',
                background: '#ffffff',
                color: '#475569',
                fontSize: '13px',
                fontWeight: 600,
                cursor: 'pointer'
              }}
            >
              Hủy bỏ
            </button>

            <button
              onClick={() => {
                if (selectedKhos.size === 0) return
                onConfirmAdd(targetGroup.standardName, Array.from(selectedKhos))
                onClose()
              }}
              disabled={selectedKhos.size === 0}
              style={{
                padding: '8px 20px',
                borderRadius: '8px',
                border: 'none',
                background: selectedKhos.size > 0 ? '#2563eb' : '#94a3b8',
                color: '#ffffff',
                fontSize: '13px',
                fontWeight: 700,
                cursor: selectedKhos.size > 0 ? 'pointer' : 'not-allowed',
                display: 'flex',
                alignItems: 'center',
                gap: '6px',
                boxShadow: selectedKhos.size > 0 ? '0 2px 6px rgba(37, 99, 235, 0.25)' : 'none'
              }}
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
    const sheet1Rows = standardizedBchGroups.map((g, idx) => ({
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
    const sheet2Rows = allOriginalKhos.map((item, idx) => ({
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
    const sheet3Rows = bchAliasRules.map((r, idx) => ({
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

  return (
    <div style={{ padding: '24px', maxWidth: '1600px', margin: '0 auto', fontFamily: 'system-ui, -apple-system, sans-serif' }}>
      
      {/* ─── HEADER ────────────────────────────────────────────────────────── */}
      <div style={{
        display: 'flex',
        flexWrap: 'wrap',
        alignItems: 'center',
        justifyContent: 'space-between',
        gap: '16px',
        marginBottom: '20px',
        paddingBottom: '16px',
        borderBottom: '1px solid #e2e8f0'
      }}>
        <div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
            <div style={{
              width: 38,
              height: 38,
              borderRadius: 8,
              background: 'linear-gradient(135deg, #1e40af 0%, #3b82f6 100%)',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              color: '#ffffff',
              boxShadow: '0 2px 6px rgba(30, 64, 175, 0.25)'
            }}>
              <Building2 size={22} />
            </div>
            <div>
              <h1 style={{ fontSize: '20px', fontWeight: 800, color: '#0f172a', margin: 0 }}>
                Chuẩn hóa tên Ban Chỉ Huy (Kho BCH)
              </h1>
              <p style={{ fontSize: '13px', color: '#64748b', margin: '2px 0 0 0' }}>
                Mỗi dòng là 1 Tên BCH chuẩn hóa, gộp các tên kho biến thể/ngăn kho về duy nhất một Ban Chỉ Huy.
              </p>
            </div>
          </div>
        </div>

        {/* Action Buttons */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
          <button
            onClick={() => {
              setNewBchStandardName('')
              setNewBchSelectedKhos(new Set())
              setNewBchSearchKho('')
              setNewBchNote('')
              setShowCreateModal(true)
            }}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              padding: '8px 16px',
              background: '#2563eb',
              color: '#ffffff',
              border: 'none',
              borderRadius: '8px',
              fontSize: '13px',
              fontWeight: 700,
              cursor: 'pointer',
              boxShadow: '0 2px 6px rgba(37, 99, 235, 0.25)',
              transition: 'all 0.15s ease'
            }}
          >
            <Plus size={16} /> Tạo Nhóm BCH Chuẩn Mới
          </button>

          {isSupabaseConfigured && (
            <button
              onClick={handleSaveToSupabase}
              disabled={saveStatus === 'saving'}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '6px',
                padding: '8px 16px',
                background: saveStatus === 'saving' ? '#94a3b8' : 'linear-gradient(135deg, #16a34a 0%, #15803d 100%)',
                color: '#ffffff',
                border: 'none',
                borderRadius: '8px',
                fontSize: '13px',
                fontWeight: 600,
                cursor: saveStatus === 'saving' ? 'not-allowed' : 'pointer',
                boxShadow: '0 2px 6px rgba(22, 163, 74, 0.25)',
                transition: 'all 0.15s ease'
              }}
            >
              <Save size={15} />
              {saveStatus === 'saving' ? 'Đang lưu...' : 'Lưu lên Supabase'}
            </button>
          )}

          {onReloadRulesFromSupabase && (
            <button
              onClick={onReloadRulesFromSupabase}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: '6px',
                padding: '8px 14px',
                background: '#f8fafc',
                color: '#334155',
                border: '1px solid #cbd5e1',
                borderRadius: '8px',
                fontSize: '13px',
                fontWeight: 600,
                cursor: 'pointer'
              }}
              title="Tải lại quy tắc từ cơ sở dữ liệu Supabase"
            >
              <RefreshCw size={14} /> Tải lại DB
            </button>
          )}

          <button
            onClick={handleExportExcel}
            style={{
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              padding: '8px 14px',
              background: '#ffffff',
              color: '#0f766e',
              border: '1px solid #99f6e4',
              borderRadius: '8px',
              fontSize: '13px',
              fontWeight: 600,
              cursor: 'pointer'
            }}
          >
            <FileSpreadsheet size={15} /> Xuất Excel
          </button>
        </div>
      </div>

      {/* ─── STATUS MESSAGE NOTIFICATION ────────────────────────────────────── */}
      {statusMessage && (
        <div style={{
          padding: '12px 16px',
          borderRadius: '8px',
          marginBottom: '16px',
          display: 'flex',
          alignItems: 'center',
          gap: '10px',
          fontSize: '13.5px',
          fontWeight: 600,
          background: saveStatus === 'error' ? '#fef2f2' : '#f0fdf4',
          color: saveStatus === 'error' ? '#b91c1c' : '#15803d',
          border: `1px solid ${saveStatus === 'error' ? '#fecaca' : '#bbf7d0'}`
        }}>
          {saveStatus === 'error' ? <AlertCircle size={18} /> : <CheckCircle2 size={18} />}
          <span>{statusMessage}</span>
        </div>
      )}

      {/* ─── STATS CARDS: TỔNG QUAN TỔNG DANH SÁCH BCH CHUẨN HÓA ────────────── */}
      <div style={{
        display: 'grid',
        gridTemplateColumns: 'repeat(auto-fit, minmax(240px, 1fr))',
        gap: '14px',
        marginBottom: '20px'
      }}>
        {/* Card 1: Tổng số BCH sau chuẩn hóa */}
        <div style={{
          background: 'linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%)',
          padding: '16px 20px',
          borderRadius: '12px',
          border: '1.5px solid #bfdbfe',
          boxShadow: '0 2px 6px rgba(37, 99, 235, 0.08)',
          display: 'flex',
          alignItems: 'center',
          gap: '14px'
        }}>
          <div style={{
            width: 44,
            height: 44,
            borderRadius: 10,
            background: '#2563eb',
            color: '#ffffff',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            boxShadow: '0 2px 6px rgba(37, 99, 235, 0.3)'
          }}>
            <Building2 size={24} />
          </div>
          <div>
            <div style={{ fontSize: '12.5px', color: '#1e40af', fontWeight: 700, textTransform: 'uppercase', letterSpacing: '0.3px' }}>
              Tổng số BCH (Sau chuẩn hóa)
            </div>
            <div style={{ fontSize: '24px', fontWeight: 900, color: '#1e3a8a', lineHeight: 1.2 }}>
              {standardizedBchGroups.length} <span style={{ fontSize: '14px', fontWeight: 600, color: '#3b82f6' }}>Ban Chỉ Huy</span>
            </div>
            <div style={{ fontSize: '11.5px', color: '#60a5fa', marginTop: '2px' }}>
              Giảm từ {allOriginalKhos.length} kho gốc ban đầu
            </div>
          </div>
        </div>

        {/* Card 2: BCH đã gộp nhiều kho */}
        <div style={{
          background: '#ffffff',
          padding: '16px 20px',
          borderRadius: '12px',
          border: '1px solid #e2e8f0',
          boxShadow: '0 1px 4px rgba(0,0,0,0.04)',
          display: 'flex',
          alignItems: 'center',
          gap: '14px'
        }}>
          <div style={{
            width: 44,
            height: 44,
            borderRadius: 10,
            background: '#f5f3ff',
            color: '#7c3aed',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
          }}>
            <Layers size={24} />
          </div>
          <div>
            <div style={{ fontSize: '12.5px', color: '#64748b', fontWeight: 600 }}>
              BCH đã gộp nhiều kho
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#7c3aed' }}>
              {multiMergedGroupsCount} <span style={{ fontSize: '13px', fontWeight: 500, color: '#94a3b8' }}>nhóm</span>
            </div>
            <div style={{ fontSize: '11.5px', color: '#8b5cf6', marginTop: '2px' }}>
              (Đã gộp {totalMergedOldKhoCount} kho gốc)
            </div>
          </div>
        </div>

        {/* Card 3: BCH độc lập (1 kho) */}
        <div style={{
          background: '#ffffff',
          padding: '16px 20px',
          borderRadius: '12px',
          border: '1px solid #e2e8f0',
          boxShadow: '0 1px 4px rgba(0,0,0,0.04)',
          display: 'flex',
          alignItems: 'center',
          gap: '14px'
        }}>
          <div style={{
            width: 44,
            height: 44,
            borderRadius: 10,
            background: '#fffbeb',
            color: '#d97706',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
          }}>
            <Info size={24} />
          </div>
          <div>
            <div style={{ fontSize: '12.5px', color: '#64748b', fontWeight: 600 }}>
              BCH độc lập (1 kho = 1 BCH)
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#d97706' }}>
              {standaloneGroupsCount} <span style={{ fontSize: '13px', fontWeight: 500, color: '#94a3b8' }}>BCH</span>
            </div>
            <div style={{ fontSize: '11.5px', color: '#b45309', marginTop: '2px' }}>
              Chưa cần gộp thêm kho khác
            </div>
          </div>
        </div>

        {/* Card 4: Tổng kho gốc ban đầu */}
        <div style={{
          background: '#ffffff',
          padding: '16px 20px',
          borderRadius: '12px',
          border: '1px solid #e2e8f0',
          boxShadow: '0 1px 4px rgba(0,0,0,0.04)',
          display: 'flex',
          alignItems: 'center',
          gap: '14px'
        }}>
          <div style={{
            width: 44,
            height: 44,
            borderRadius: 10,
            background: '#f0fdf4',
            color: '#16a34a',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center'
          }}>
            <Tag size={24} />
          </div>
          <div>
            <div style={{ fontSize: '12.5px', color: '#64748b', fontWeight: 600 }}>
              Tổng Kho BCH gốc trong dữ liệu
            </div>
            <div style={{ fontSize: '22px', fontWeight: 800, color: '#16a34a' }}>
              {allOriginalKhos.length} <span style={{ fontSize: '13px', fontWeight: 500, color: '#94a3b8' }}>kho gốc</span>
            </div>
            <div style={{ fontSize: '11.5px', color: '#15803d', marginTop: '2px' }}>
              {bchAliasRules.length} quy tắc chuẩn hóa đang áp dụng
            </div>
          </div>
        </div>
      </div>

      {/* ─── CONTROLS & FILTER BAR ─────────────────────────────────────────── */}
      <div style={{
        background: '#ffffff',
        padding: '14px 18px',
        borderRadius: '10px',
        border: '1px solid #e2e8f0',
        marginBottom: '16px',
        display: 'flex',
        flexWrap: 'wrap',
        alignItems: 'center',
        justifyContent: 'space-between',
        gap: '12px'
      }}>
        {/* Search */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flex: 1, minWidth: '280px', maxWidth: '500px' }}>
          <div style={{ position: 'relative', width: '100%' }}>
            <Search size={16} style={{ position: 'absolute', left: 12, top: '50%', transform: 'translateY(-50%)', color: '#94a3b8' }} />
            <input
              type="text"
              placeholder="🔍 Tìm theo Tên BCH chuẩn hóa hoặc tên Kho gốc..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              style={{
                width: '100%',
                padding: '8px 12px 8px 36px',
                borderRadius: '8px',
                border: '1px solid #cbd5e1',
                fontSize: '13.5px',
                outline: 'none'
              }}
            />
            {searchTerm && (
              <button
                onClick={() => setSearchTerm('')}
                style={{
                  position: 'absolute',
                  right: 10,
                  top: '50%',
                  transform: 'translateY(-50%)',
                  background: 'none',
                  border: 'none',
                  color: '#94a3b8',
                  cursor: 'pointer'
                }}
              >
                <X size={14} />
              </button>
            )}
          </div>
        </div>

        {/* Filter Pills */}
        <div style={{ display: 'flex', alignItems: 'center', gap: '8px', flexWrap: 'wrap' }}>
          <span style={{ fontSize: '13px', color: '#64748b', fontWeight: 600 }}>
            Lọc hiển thị:
          </span>

          <button
            onClick={() => setStatusFilter('all')}
            style={{
              padding: '6px 14px',
              borderRadius: '20px',
              fontSize: '12.5px',
              fontWeight: 600,
              cursor: 'pointer',
              border: statusFilter === 'all' ? '1.5px solid #2563eb' : '1px solid #e2e8f0',
              background: statusFilter === 'all' ? '#eff6ff' : '#f8fafc',
              color: statusFilter === 'all' ? '#2563eb' : '#475569',
              transition: 'all 0.15s'
            }}
          >
            Tất cả BCH ({standardizedBchGroups.length})
          </button>

          <button
            onClick={() => setStatusFilter('merged_multi')}
            style={{
              padding: '6px 14px',
              borderRadius: '20px',
              fontSize: '12.5px',
              fontWeight: 600,
              cursor: 'pointer',
              border: statusFilter === 'merged_multi' ? '1.5px solid #7c3aed' : '1px solid #e2e8f0',
              background: statusFilter === 'merged_multi' ? '#f5f3ff' : '#f8fafc',
              color: statusFilter === 'merged_multi' ? '#7c3aed' : '#475569',
              transition: 'all 0.15s'
            }}
          >
            Đã gộp nhiều kho ({multiMergedGroupsCount})
          </button>

          <button
            onClick={() => setStatusFilter('standalone')}
            style={{
              padding: '6px 14px',
              borderRadius: '20px',
              fontSize: '12.5px',
              fontWeight: 600,
              cursor: 'pointer',
              border: statusFilter === 'standalone' ? '1.5px solid #d97706' : '1px solid #e2e8f0',
              background: statusFilter === 'standalone' ? '#fffbeb' : '#f8fafc',
              color: statusFilter === 'standalone' ? '#d97706' : '#475569',
              transition: 'all 0.15s'
            }}
          >
            BCH độc lập ({standaloneGroupsCount})
          </button>
        </div>
      </div>

      {/* ─── MAIN TABLE: 1 DÒNG LÀ 1 TÊN BCH CHUẨN HÓA ────────────────────── */}
      <div style={{
        background: '#ffffff',
        borderRadius: '10px',
        border: '1px solid #e2e8f0',
        boxShadow: '0 2px 8px rgba(0,0,0,0.04)',
        overflow: 'hidden',
        display: 'flex',
        flexDirection: 'column'
      }}>
        {/* Scrollable container with sticky header */}
        <div style={{
          maxHeight: '660px',
          minHeight: '400px',
          overflowY: 'auto',
          overflowX: 'auto',
          position: 'relative'
        }}>
          <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', fontSize: '13px' }}>
            <thead style={{
              position: 'sticky',
              top: 0,
              zIndex: 30,
              background: '#00529C',
              boxShadow: '0 2px 6px rgba(0, 0, 0, 0.15)'
            }}>
              <tr style={{ background: '#00529C', color: '#ffffff' }}>
                <th style={{ padding: '12px 10px', width: '50px', textAlign: 'center', fontWeight: 700, background: '#00529C', color: '#ffffff', borderRight: '1px solid rgba(255,255,255,0.2)' }}>
                  STT
                </th>
                <th style={{ padding: '12px 16px', fontWeight: 700, width: '320px', minWidth: '280px', background: '#00529C', color: '#ffffff', borderRight: '1px solid rgba(255,255,255,0.2)', textAlign: 'left' }}>
                  Tên Ban Chỉ Huy (Chuẩn Hóa)
                </th>
                <th style={{ padding: '12px 16px', fontWeight: 700, minWidth: '460px', background: '#00529C', color: '#ffffff', borderRight: '1px solid rgba(255,255,255,0.2)', textAlign: 'left' }}>
                  Các Kho BCH Gốc Đã Gộp Vào (Chọn & Quản lý Kho)
                </th>
                <th style={{ padding: '12px 12px', width: '100px', textAlign: 'center', fontWeight: 700, background: '#00529C', color: '#ffffff', borderRight: '1px solid rgba(255,255,255,0.2)' }}>
                  Số Kho Gộp
                </th>
                <th style={{ padding: '12px 14px', width: '120px', textAlign: 'center', fontWeight: 700, background: '#00529C', color: '#ffffff', borderRight: '1px solid rgba(255,255,255,0.2)' }}>
                  Tổng Phát Sinh
                </th>
                <th style={{ padding: '12px 14px', width: '150px', minWidth: '120px', fontWeight: 700, background: '#00529C', color: '#ffffff', borderRight: '1px solid rgba(255,255,255,0.2)', textAlign: 'left' }}>
                  Ghi chú
                </th>
                <th style={{ padding: '12px 14px', width: '140px', textAlign: 'center', fontWeight: 700, background: '#00529C', color: '#ffffff' }}>
                  Thao tác
                </th>
              </tr>
            </thead>
            <tbody>
              {filteredStandardGroups.length === 0 ? (
                <tr>
                  <td colSpan={7} style={{ padding: '48px 20px', textAlign: 'center', color: '#94a3b8' }}>
                    <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '8px' }}>
                      <Building2 size={36} color="#cbd5e1" />
                      <div style={{ fontSize: '14px', fontWeight: 600, color: '#475569' }}>
                        Không tìm thấy Ban Chỉ Huy nào phù hợp với bộ lọc/tìm kiếm
                      </div>
                      <div style={{ fontSize: '12px' }}>Thử xóa từ khóa tìm kiếm hoặc chọn lọc "Tất cả BCH"</div>
                    </div>
                  </td>
                </tr>
              ) : (
                filteredStandardGroups.map((group, index) => {
                  const isEditing = editingStandardKey === group.key

                  return (
                    <tr
                      key={group.key}
                      style={{
                        borderBottom: '1px solid #f1f5f9',
                        background: group.isMultiMerged ? (index % 2 === 0 ? '#ffffff' : '#fcfaff') : (index % 2 === 0 ? '#ffffff' : '#fafafa'),
                        transition: 'background-color 0.15s ease'
                      }}
                    >
                      {/* STT */}
                      <td style={{ padding: '12px 10px', textAlign: 'center', color: '#64748b', fontWeight: 600, fontSize: '12.5px' }}>
                        {index + 1}
                      </td>

                      {/* TÊN BCH CHUẨN HÓA */}
                      <td style={{ padding: '12px 16px', verticalAlign: 'top' }}>
                        {isEditing ? (
                          <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                            <input
                              type="text"
                              value={editingStandardValue}
                              onChange={(e) => setEditingStandardValue(e.target.value)}
                              placeholder="Nhập tên BCH chuẩn hóa mới..."
                              autoFocus
                              style={{
                                width: '100%',
                                padding: '6px 10px',
                                borderRadius: '6px',
                                border: '1.5px solid #2563eb',
                                fontSize: '13px',
                                outline: 'none'
                              }}
                              onKeyDown={(e) => {
                                if (e.key === 'Enter') handleSaveRenameGroup(group)
                                if (e.key === 'Escape') setEditingStandardKey(null)
                              }}
                            />
                            <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                              <button
                                onClick={() => handleSaveRenameGroup(group)}
                                style={{
                                  padding: '4px 10px',
                                  background: '#16a34a',
                                  color: '#ffffff',
                                  border: 'none',
                                  borderRadius: '4px',
                                  fontSize: '12px',
                                  fontWeight: 600,
                                  cursor: 'pointer',
                                  display: 'flex',
                                  alignItems: 'center',
                                  gap: '4px'
                                }}
                              >
                                <Check size={12} /> Lưu
                              </button>
                              <button
                                onClick={() => setEditingStandardKey(null)}
                                style={{
                                  padding: '4px 10px',
                                  background: '#e2e8f0',
                                  color: '#475569',
                                  border: 'none',
                                  borderRadius: '4px',
                                  fontSize: '12px',
                                  cursor: 'pointer'
                                }}
                              >
                                Hủy
                              </button>
                            </div>
                          </div>
                        ) : (
                          <div>
                            <div style={{
                              fontWeight: 700,
                              color: group.isMultiMerged ? '#1e40af' : '#1e293b',
                              fontSize: '14px',
                              lineHeight: 1.4,
                              display: 'flex',
                              alignItems: 'flex-start',
                              gap: '6px'
                            }}>
                              <Building2 size={16} color={group.isMultiMerged ? '#2563eb' : '#64748b'} style={{ flexShrink: 0, marginTop: '3px' }} />
                              <span>{group.standardName}</span>
                            </div>

                            <div style={{ marginTop: '4px', display: 'flex', alignItems: 'center', gap: '6px' }}>
                              {group.isMultiMerged ? (
                                <span style={{
                                  fontSize: '11px',
                                  fontWeight: 700,
                                  color: '#7c3aed',
                                  background: '#f5f3ff',
                                  border: '1px solid #ddd6fe',
                                  padding: '2px 8px',
                                  borderRadius: '12px',
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '3px'
                                }}>
                                  <Layers size={11} /> Đã gộp {group.memberCount} kho gốc
                                </span>
                              ) : (
                                <span style={{
                                  fontSize: '11px',
                                  color: '#64748b',
                                  background: '#f1f5f9',
                                  border: '1px solid #e2e8f0',
                                  padding: '2px 8px',
                                  borderRadius: '12px',
                                  display: 'inline-block'
                                }}>
                                  BCH độc lập (1 kho)
                                </span>
                              )}
                            </div>
                          </div>
                        )}
                      </td>

                      {/* CÁC KHO BCH GỐC ĐÃ GỘP (MULTI-SELECT CHIPS + ADD BUTTON) */}
                      <td style={{ padding: '12px 16px', verticalAlign: 'top' }}>
                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: '6px', alignItems: 'center', position: 'relative' }}>
                          {/* Danh sách các Kho gốc */}
                          {group.members.map((member) => {
                            const isExactSameName = normKey(member.name) === normKey(group.standardName)

                            return (
                              <div
                                key={member.name}
                                style={{
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: '6px',
                                  padding: '4px 10px',
                                  background: isExactSameName ? '#eff6ff' : '#ffffff',
                                  border: isExactSameName ? '1px solid #bfdbfe' : '1px solid #cbd5e1',
                                  borderRadius: '6px',
                                  fontSize: '12.5px',
                                  color: isExactSameName ? '#1e40af' : '#334155',
                                  fontWeight: isExactSameName ? 600 : 500,
                                  boxShadow: '0 1px 2px rgba(0,0,0,0.03)'
                                }}
                                title={`${member.name} (${member.totalCount || 0} phiếu)`}
                              >
                                <span style={{ maxWidth: '320px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                                  {isExactSameName ? '⭐ ' : ''}{member.name}
                                </span>

                                <span style={{ fontSize: '11px', color: '#94a3b8', background: 'rgba(0,0,0,0.05)', padding: '1px 5px', borderRadius: '4px' }}>
                                  {member.totalCount || 0}p
                                </span>

                                {/* Nút gỡ kho ra khỏi nhóm nếu không phải kho duy nhất hoặc là alias */}
                                {(!isExactSameName || group.members.length > 1) && (
                                  <button
                                    onClick={() => handleRemoveKhoFromGroup(member.name)}
                                    style={{
                                      background: 'none',
                                      border: 'none',
                                      color: '#ef4444',
                                      cursor: 'pointer',
                                      padding: '0 2px',
                                      display: 'flex',
                                      alignItems: 'center',
                                      borderRadius: '3px'
                                    }}
                                    title={`Gỡ bỏ "${member.name}" ra khỏi nhóm này`}
                                  >
                                    <X size={13} />
                                  </button>
                                )}
                              </div>
                            )
                          })}

                          {/* Nút "+ Gán thêm kho gốc" */}
                          <div>
                            <button
                              onClick={() => setAssignModalGroup(group)}
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: '5px',
                                padding: '5px 12px',
                                background: '#eff6ff',
                                color: '#1d4ed8',
                                border: '1px solid #93c5fd',
                                borderRadius: '6px',
                                fontSize: '12px',
                                fontWeight: 700,
                                cursor: 'pointer',
                                transition: 'all 0.15s ease',
                                boxShadow: '0 1px 2px rgba(37, 99, 235, 0.08)'
                              }}
                              title={`Chọn thêm kho gốc để gộp vào Ban Chỉ Huy "${group.standardName}"`}
                            >
                              <Plus size={13} color="#2563eb" /> Gán thêm kho
                            </button>
                          </div>
                        </div>
                      </td>

                      {/* SỐ LƯỢNG KHO GỐC */}
                      <td style={{ padding: '12px 12px', textAlign: 'center', verticalAlign: 'top' }}>
                        <span style={{
                          display: 'inline-block',
                          padding: '3px 10px',
                          background: group.isMultiMerged ? '#f5f3ff' : '#f8fafc',
                          color: group.isMultiMerged ? '#7c3aed' : '#475569',
                          border: group.isMultiMerged ? '1px solid #ddd6fe' : '1px solid #e2e8f0',
                          borderRadius: '12px',
                          fontSize: '12px',
                          fontWeight: 700
                        }}>
                          {group.memberCount} kho
                        </span>
                      </td>

                      {/* TỔNG PHÁT SINH */}
                      <td style={{ padding: '12px 14px', textAlign: 'center', verticalAlign: 'top' }}>
                        <div style={{ fontWeight: 700, color: '#1e293b', fontSize: '13px' }}>
                          {group.totalPhieu.toLocaleString()} <span style={{ fontSize: '11px', fontWeight: 500, color: '#64748b' }}>phiếu</span>
                        </div>
                        <div style={{ fontSize: '11px', color: '#94a3b8', marginTop: '2px' }}>
                          {group.totalGiao} giao / {group.totalNhan} nhận
                        </div>
                      </td>

                      {/* GHI CHÚ */}
                      <td style={{ padding: '12px 14px', color: '#64748b', fontSize: '12px', verticalAlign: 'top' }}>
                        {group.note || '-'}
                      </td>

                      {/* THAO TÁC */}
                      <td style={{ padding: '12px 14px', textAlign: 'center', verticalAlign: 'top' }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
                          <button
                            onClick={() => {
                              setEditingStandardKey(group.key)
                              setEditingStandardValue(group.standardName)
                              setEditingNoteValue(group.note || '')
                            }}
                            style={{
                              display: 'flex',
                              alignItems: 'center',
                              gap: '4px',
                              padding: '5px 10px',
                              background: '#eff6ff',
                              color: '#2563eb',
                              border: '1px solid #bfdbfe',
                              borderRadius: '6px',
                              fontSize: '12px',
                              fontWeight: 600,
                              cursor: 'pointer'
                            }}
                            title="Đổi tên Ban Chỉ Huy chuẩn hóa"
                          >
                            <Edit2 size={12} /> Đổi tên
                          </button>

                          {group.hasRules && (
                            <button
                              onClick={() => handleDeleteGroup(group)}
                              style={{
                                display: 'flex',
                                alignItems: 'center',
                                padding: '5px 8px',
                                background: '#fef2f2',
                                color: '#dc2626',
                                border: '1px solid #fecaca',
                                borderRadius: '6px',
                                fontSize: '12px',
                                cursor: 'pointer'
                              }}
                              title="Hủy gộp nhóm này (Trả tất cả về kho gốc ban đầu)"
                            >
                              <Trash2 size={12} />
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  )
                })
              )}
            </tbody>
          </table>
        </div>

        {/* Table Footer Summary */}
        <div style={{
          padding: '12px 18px',
          background: '#f8fafc',
          borderTop: '1px solid #e2e8f0',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          fontSize: '13px',
          color: '#64748b'
        }}>
          <div>
            Hiển thị <strong>{filteredStandardGroups.length}</strong> / <strong>{standardizedBchGroups.length}</strong> Ban Chỉ Huy chuẩn hóa
            (Tổng hợp từ <strong>{allOriginalKhos.length}</strong> Kho BCH gốc)
          </div>
          <div style={{ display: 'flex', alignItems: 'center', gap: '6px', color: '#64748b', fontSize: '12px' }}>
            <span>💡 Dùng con lăn chuột để cuộn xem danh sách</span>
          </div>
        </div>
      </div>

      {/* ─── MODAL TẠO NHÓM BCH CHUẨN HÓA MỚI ───────────────────────────────── */}
      {showCreateModal && (
        <div style={{
          position: 'fixed',
          top: 0,
          left: 0,
          right: 0,
          bottom: 0,
          background: 'rgba(0, 0, 0, 0.5)',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'center',
          zIndex: 9999,
          padding: '20px'
        }}>
          <div style={{
            background: '#ffffff',
            borderRadius: '12px',
            maxWidth: '680px',
            width: '100%',
            maxHeight: '90vh',
            display: 'flex',
            flexDirection: 'column',
            boxShadow: '0 20px 25px -5px rgba(0, 0, 0, 0.15)',
            overflow: 'hidden'
          }}>
            {/* Modal Header */}
            <div style={{
              padding: '16px 20px',
              borderBottom: '1px solid #e2e8f0',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              background: '#00529C',
              color: '#ffffff'
            }}>
              <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <Plus size={18} />
                <h3 style={{ fontSize: '16px', fontWeight: 700, margin: 0, color: '#ffffff' }}>
                  Tạo Nhóm Ban Chỉ Huy Chuẩn Hóa Mới
                </h3>
              </div>
              <button
                onClick={() => setShowCreateModal(false)}
                style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#ffffff' }}
              >
                <X size={18} />
              </button>
            </div>

            {/* Modal Body */}
            <div style={{ padding: '20px', overflowY: 'auto', flex: 1, display: 'flex', flexDirection: 'column', gap: '16px' }}>
              {/* Tên BCH Chuẩn */}
              <div>
                <label style={{ display: 'block', fontSize: '13px', fontWeight: 700, color: '#1e293b', marginBottom: '6px' }}>
                  1. Tên Ban Chỉ Huy Chuẩn Hóa thống nhất <span style={{ color: '#ef4444' }}>*</span>:
                </label>
                <input
                  type="text"
                  placeholder="Ví dụ: BCH Cọc Khoan Nhồi - Trống Đồng..."
                  value={newBchStandardName}
                  onChange={(e) => setNewBchStandardName(e.target.value)}
                  style={{
                    width: '100%',
                    padding: '9px 12px',
                    borderRadius: '6px',
                    border: '1.5px solid #cbd5e1',
                    fontSize: '13.5px',
                    outline: 'none',
                    boxSizing: 'border-box'
                  }}
                />
              </div>

              {/* Chọn các Kho BCH gốc */}
              <div>
                <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', marginBottom: '6px' }}>
                  <label style={{ fontSize: '13px', fontWeight: 700, color: '#1e293b' }}>
                    2. Chọn các Kho BCH gốc để gộp vào nhóm này ({newBchSelectedKhos.size} đã chọn) <span style={{ color: '#ef4444' }}>*</span>:
                  </label>
                </div>

                {/* Ô tìm kiếm kho */}
                <div style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                  background: '#f8fafc',
                  border: '1px solid #cbd5e1',
                  borderRadius: '6px',
                  padding: '6px 10px',
                  marginBottom: '8px'
                }}>
                  <Search size={14} color="#94a3b8" />
                  <input
                    type="text"
                    placeholder="🔍 Gõ tìm nhanh tên kho gốc..."
                    value={newBchSearchKho}
                    onChange={(e) => setNewBchSearchKho(e.target.value)}
                    style={{
                      flex: 1,
                      border: 'none',
                      background: 'transparent',
                      fontSize: '13px',
                      outline: 'none'
                    }}
                  />
                  {newBchSearchKho && (
                    <button
                      onClick={() => setNewBchSearchKho('')}
                      style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#94a3b8', padding: 0 }}
                    >
                      <X size={12} />
                    </button>
                  )}
                </div>

                {/* Danh sách checkbox kho */}
                <div style={{
                  maxHeight: '220px',
                  overflowY: 'auto',
                  border: '1px solid #e2e8f0',
                  borderRadius: '8px',
                  padding: '4px 0',
                  background: '#ffffff'
                }}>
                  {allOriginalKhos
                    .filter(k => !newBchSearchKho || k.name.toLowerCase().includes(newBchSearchKho.toLowerCase().trim()))
                    .map(k => {
                      const isChecked = newBchSelectedKhos.has(k.name)

                      return (
                        <div
                          key={k.name}
                          onClick={() => {
                            setNewBchSelectedKhos(prev => {
                              const next = new Set(prev)
                              if (next.has(k.name)) next.delete(k.name)
                              else next.add(k.name)
                              return next
                            })
                          }}
                          style={{
                            padding: '8px 12px',
                            display: 'flex',
                            alignItems: 'center',
                            gap: '8px',
                            cursor: 'pointer',
                            background: isChecked ? '#eff6ff' : 'transparent',
                            borderBottom: '1px solid #f8fafc'
                          }}
                        >
                          {isChecked ? (
                            <CheckSquare size={16} color="#2563eb" />
                          ) : (
                            <Square size={16} color="#cbd5e1" />
                          )}
                          <div style={{ flex: 1, fontSize: '13px', color: isChecked ? '#1d4ed8' : '#1e293b', fontWeight: isChecked ? 600 : 400 }}>
                            {k.name}
                          </div>
                          <span style={{ fontSize: '11.5px', color: '#94a3b8' }}>
                            {k.totalCount} phiếu
                          </span>
                        </div>
                      )
                    })}
                </div>
              </div>

              {/* Ghi chú */}
              <div>
                <label style={{ display: 'block', fontSize: '13px', fontWeight: 600, color: '#475569', marginBottom: '4px' }}>
                  3. Ghi chú (tùy chọn):
                </label>
                <input
                  type="text"
                  placeholder="Ghi chú thêm về nhóm BCH này..."
                  value={newBchNote}
                  onChange={(e) => setNewBchNote(e.target.value)}
                  style={{
                    width: '100%',
                    padding: '8px 12px',
                    borderRadius: '6px',
                    border: '1px solid #cbd5e1',
                    fontSize: '13px',
                    outline: 'none',
                    boxSizing: 'border-box'
                  }}
                />
              </div>
            </div>

            {/* Modal Footer */}
            <div style={{
              padding: '14px 20px',
              borderTop: '1px solid #e2e8f0',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'flex-end',
              gap: '10px',
              background: '#f8fafc'
            }}>
              <button
                onClick={() => setShowCreateModal(false)}
                style={{
                  padding: '8px 16px',
                  background: '#e2e8f0',
                  color: '#475569',
                  border: 'none',
                  borderRadius: '6px',
                  fontSize: '13px',
                  fontWeight: 600,
                  cursor: 'pointer'
                }}
              >
                Hủy bỏ
              </button>

              <button
                onClick={handleCreateNewGroup}
                disabled={!newBchStandardName || newBchSelectedKhos.size === 0}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  gap: '6px',
                  padding: '8px 18px',
                  background: (!newBchStandardName || newBchSelectedKhos.size === 0) ? '#94a3b8' : '#2563eb',
                  color: '#ffffff',
                  border: 'none',
                  borderRadius: '6px',
                  fontSize: '13px',
                  fontWeight: 700,
                  cursor: (!newBchStandardName || newBchSelectedKhos.size === 0) ? 'not-allowed' : 'pointer'
                }}
              >
                <Check size={16} /> Tạo Nhóm Chuẩn ({newBchSelectedKhos.size} kho)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ─── MODAL GÁN KHO GỐC VÀO BCH CHUẨN HÓA (KHÔNG BỊ CHE BỞI BẢNG) ──── */}
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
