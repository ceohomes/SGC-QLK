import React, { useState, useMemo } from 'react'
import {
  X,
  Sparkles,
  Layers,
  Search,
  Check,
  Trash2,
  Plus,
  ArrowRight,
  Database,
  Copy,
  AlertCircle,
  CheckCircle2,
  RefreshCw,
  Edit2,
  Building,
  HelpCircle,
  Zap,
  Info
} from 'lucide-react'
import { supabase, isSupabaseConfigured } from '../supabaseClient'
import { CHB_CSS, PenIcon, XMark, CloseEscButton, useEscClose } from './ChuanHoaBchTab.jsx'

export const SQL_SETUP_SCRIPT = `-- =========================================================================
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

-- Hoàn tất! Bảng đã sẵn sàng kết nối.
`

export default function ChuanHoaBchModal({
  isOpen,
  onClose,
  allUnitNames = [],
  bchAliasRules = [], // array of { ten_cu, ten_chuan, ghi_chu, id? }
  onSaveRules,
  onApplyRulesToCurrentData
}) {
  const [activeTab, setActiveTab] = useState('rules') // 'rules' | 'add' | 'suggest' | 'sql'
  const [rules, setRules] = useState([])
  const [searchTerm, setSearchTerm] = useState('')
  const [copiedSql, setCopiedSql] = useState(false)
  const [saving, setSaving] = useState(false)
  const [saveSuccessMsg, setSaveSuccessMsg] = useState('')
  const [saveErrorMsg, setSaveErrorMsg] = useState('')

  // Form state for adding new rule
  const [selectedStandard, setSelectedStandard] = useState('')
  const [customStandardInput, setCustomStandardInput] = useState('')
  const [selectedOldNames, setSelectedOldNames] = useState([])
  const [unitSearch, setUnitSearch] = useState('')
  const [ruleNote, setRuleNote] = useState('')
  const [editingRuleIndex, setEditingRuleIndex] = useState(null)
  const [editStandardValue, setEditStandardValue] = useState('')

  // Sync rules from props when modal opens
  React.useEffect(() => {
    if (isOpen) {
      setRules([...bchAliasRules])
      setSaveSuccessMsg('')
      setSaveErrorMsg('')
      setEditingRuleIndex(null)
    }
  }, [isOpen, bchAliasRules])

  // Clean and normalize helper
  const norm = (s) => (s || '').trim().normalize('NFC').replace(/\s+/g, ' ')
  const normKey = (s) => norm(s).toLowerCase()

  // Standard BCH options from all unique units that look like BCH or currently used as standard
  const standardOptions = useMemo(() => {
    const set = new Set()
    allUnitNames.forEach((name) => {
      const n = norm(name)
      if (n) set.add(n)
    })
    rules.forEach((r) => {
      const n = norm(r.ten_chuan)
      if (n) set.add(n)
    })
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'vi'))
  }, [allUnitNames, rules])

  // Filtered unit names for selection (excluding already assigned standard name if same)
  const filteredUnitsForSelection = useMemo(() => {
    const q = unitSearch.trim().toLowerCase()
    return allUnitNames.filter((name) => {
      const n = norm(name)
      if (!n) return false
      if (q && !n.toLowerCase().includes(q)) return false
      return true
    })
  }, [allUnitNames, unitSearch])

  // Filtered rules in Tab 1
  const filteredRules = useMemo(() => {
    const q = searchTerm.trim().toLowerCase()
    if (!q) return rules
    return rules.filter(
      (r) =>
        r.ten_cu.toLowerCase().includes(q) ||
        r.ten_chuan.toLowerCase().includes(q) ||
        (r.ghi_chu && r.ghi_chu.toLowerCase().includes(q))
    )
  }, [rules, searchTerm])

  // Smart suggestions generator
  const smartSuggestions = useMemo(() => {
    const suggestions = []
    const cleanForCompare = (s) =>
      norm(s)
        .toLowerCase()
        .replace(/^[•\-\*\s]+/, '')
        .replace(/^(bch|bch\s+ckn|bch\s+cọc\s+khoan\s+nhồi|ban\s+chỉ\s+huy)\s*[\-\:]*\s*/i, '')
        .replace(/\s+/g, '')
        .replace(/0[1-9]$/, '')

    const processed = new Set()

    for (let i = 0; i < allUnitNames.length; i++) {
      const u1 = norm(allUnitNames[i])
      if (!u1 || processed.has(u1)) continue
      const c1 = cleanForCompare(u1)
      if (!c1 || c1.length < 3) continue

      const group = [u1]
      for (let j = i + 1; j < allUnitNames.length; j++) {
        const u2 = norm(allUnitNames[j])
        if (!u2 || processed.has(u2) || u1.toLowerCase() === u2.toLowerCase()) continue
        const c2 = cleanForCompare(u2)
        if (c1 === c2 || c1.includes(c2) || c2.includes(c1)) {
          group.push(u2)
        }
      }

      if (group.length > 1) {
        group.forEach((item) => processed.add(item))
        // Choose best standard name (longest or cleanest)
        const sortedGroup = [...group].sort((a, b) => {
          const aHasDot = a.startsWith('•') ? 1 : 0
          const bHasDot = b.startsWith('•') ? 1 : 0
          if (aHasDot !== bHasDot) return aHasDot - bHasDot
          return b.length - a.length
        })
        const suggestedStandard = sortedGroup[0].replace(/^[•\-\*\s]+/, '').trim()
        const oldItems = group.filter((g) => g !== suggestedStandard)

        // Check if any rule already exists
        const unmappedOldItems = oldItems.filter(
          (old) => !rules.some((r) => normKey(r.ten_cu) === normKey(old))
        )

        if (unmappedOldItems.length > 0) {
          suggestions.push({
            id: `sug_${i}`,
            suggestedStandard,
            oldItems: unmappedOldItems,
            allGroup: group
          })
        }
      }
    }
    return suggestions
  }, [allUnitNames, rules])

  // Handle adding rule from Form
  const handleAddRule = () => {
    const standardName = norm(customStandardInput || selectedStandard)
    if (!standardName) {
      alert('Vui lòng chọn hoặc nhập Tên BCH chuẩn!')
      return
    }
    if (selectedOldNames.length === 0) {
      alert('Vui lòng chọn ít nhất một Tên cũ / Biến thể cần chuẩn hóa!')
      return
    }

    const newRules = [...rules]
    let addedCount = 0

    selectedOldNames.forEach((oldName) => {
      const cleanOld = norm(oldName)
      if (!cleanOld || cleanOld.toLowerCase() === standardName.toLowerCase()) return

      const existingIndex = newRules.findIndex(
        (r) => normKey(r.ten_cu) === normKey(cleanOld)
      )
      if (existingIndex >= 0) {
        newRules[existingIndex] = {
          ...newRules[existingIndex],
          ten_chuan: standardName,
          ghi_chu: ruleNote || newRules[existingIndex].ghi_chu || ''
        }
      } else {
        newRules.push({
          ten_cu: cleanOld,
          ten_chuan: standardName,
          ghi_chu: ruleNote || ''
        })
      }
      addedCount++
    })

    setRules(newRules)
    setSelectedOldNames([])
    setRuleNote('')
    setActiveTab('rules')
    setSaveSuccessMsg(`Đã thêm ${addedCount} quy tắc chuẩn hóa! Nhấn "Lưu & Áp dụng" bên dưới để hoàn tất.`)
  }

  // Apply a smart suggestion
  const handleApplySuggestion = (sug) => {
    const newRules = [...rules]
    sug.oldItems.forEach((oldName) => {
      const cleanOld = norm(oldName)
      const existingIndex = newRules.findIndex(
        (r) => normKey(r.ten_cu) === normKey(cleanOld)
      )
      if (existingIndex >= 0) {
        newRules[existingIndex] = {
          ...newRules[existingIndex],
          ten_chuan: sug.suggestedStandard,
          ghi_chu: 'Tự động gộp thông minh'
        }
      } else {
        newRules.push({
          ten_cu: cleanOld,
          ten_chuan: sug.suggestedStandard,
          ghi_chu: 'Tự động gộp thông minh'
        })
      }
    })
    setRules(newRules)
    setSaveSuccessMsg(`Đã gộp ${sug.oldItems.length} tên cũ về "${sug.suggestedStandard}". Nhấn "Lưu & Áp dụng" để đồng bộ!`)
  }

  // Delete rule
  const handleDeleteRule = (indexToDelete) => {
    const newRules = rules.filter((_, idx) => idx !== indexToDelete)
    setRules(newRules)
  }

  // Save rules to Supabase & apply globally
  const handleSaveAndApply = async () => {
    setSaving(true)
    setSaveErrorMsg('')
    setSaveSuccessMsg('')

    try {
      if (isSupabaseConfigured) {
        // 1. Delete all and insert new rules (or upsert)
        // Check if table exists
        const { error: testErr } = await supabase
          .from('chuan_hoa_bch')
          .select('id')
          .limit(1)

        if (testErr) {
          // If table does not exist or RLS issue
          console.warn('Supabase table chuan_hoa_bch check warning:', testErr)
          if (testErr.code === '42P01' || testErr.message?.includes('does not exist')) {
            throw new Error(
              'Bảng "chuan_hoa_bch" chưa được tạo trên Supabase. Vui lòng vào tab "Câu lệnh SQL Supabase", copy mã SQL và chạy trên SQL Editor của Supabase!'
            )
          }
        }

        // Delete all old rules in database
        const { error: delErr } = await supabase
          .from('chuan_hoa_bch')
          .delete()
          .neq('id', -999999) // delete all

        if (delErr) {
          console.warn('Delete error, trying insert:', delErr)
        }

        // Chunk insert rules
        if (rules.length > 0) {
          const insertData = rules.map((r) => ({
            ten_cu: norm(r.ten_cu),
            ten_chuan: norm(r.ten_chuan),
            ghi_chu: r.ghi_chu || ''
          }))

          const chunkSize = 100
          for (let i = 0; i < insertData.length; i += chunkSize) {
            const chunk = insertData.slice(i, i + chunkSize)
            const { error: insErr } = await supabase
              .from('chuan_hoa_bch')
              .insert(chunk)
            if (insErr) throw insErr
          }
        }
      }

      // 2. Save to localStorage fallback
      try {
        localStorage.setItem('sgc_chuan_hoa_bch', JSON.stringify(rules))
      } catch (e) {
        console.error('LocalStorage write error:', e)
      }

      // 3. Callback to parent App to update active alias map and state
      if (onSaveRules) {
        onSaveRules(rules)
      }

      // 4. Apply to currently loaded rows
      if (onApplyRulesToCurrentData) {
        onApplyRulesToCurrentData(rules)
      }

      setSaveSuccessMsg('Đã lưu quy tắc chuẩn hóa lên hệ thống và tự động cập nhật toàn bộ dữ liệu!')
      setTimeout(() => {
        setSaveSuccessMsg('')
      }, 4000)
    } catch (err) {
      console.error('Save rules error:', err)
      setSaveErrorMsg(err.message || String(err))
      if (err.message && err.message.includes('chuan_hoa_bch')) {
        setActiveTab('sql')
      }
    } finally {
      setSaving(false)
    }
  }

  // Esc đóng hộp thoại
  useEscClose(Boolean(isOpen), onClose)

  if (!isOpen) return null

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'rgba(15, 23, 42, 0.55)',
        backdropFilter: 'blur(4px)',
        zIndex: 99999,
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: '16px'
      }}
      onClick={(e) => {
        if (e.target === e.currentTarget) onClose()
      }}
    >
      <div
        style={{
          background: '#ffffff',
          width: '100%',
          maxWidth: '960px',
          maxHeight: '92vh',
          borderRadius: '24px',
          boxShadow: '0 30px 60px -20px rgba(10, 36, 80, 0.45)',
          display: 'flex',
          flexDirection: 'column',
          overflow: 'hidden'
        }}
      >
        <style>{CHB_CSS}</style>
        {/* Đầu hộp thoại */}
        <div className="chb-dlg-hd" style={{ flexShrink: 0 }}>
          <div className="chb-dlg-ttl">
            <span className="chb-dlg-ic"><Sparkles size={18} /></span>
            <div style={{ minWidth: 0 }}>
              <h3 style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                Chuẩn hóa và hợp nhất tên Ban Chỉ Huy (Kho BCH)
                <span className="chb-tag multi">{rules.length} quy tắc</span>
              </h3>
              <div className="chb-dlg-sub" title="Tự động gom và quy đổi các tên cũ/biến thể trong Excel về một tên BCH chuẩn hóa duy nhất khi tải đơn lên">
                Gom các tên cũ / biến thể trong Excel về một tên BCH chuẩn duy nhất khi tải đơn lên
              </div>
            </div>
          </div>
          <CloseEscButton onClick={onClose} />
        </div>

        {/* Chọn mục */}
        <div style={{ padding: '10px 22px', borderBottom: '1px solid #E3E6EE', background: '#F7F8FB', flexShrink: 0, overflowX: 'auto' }}>
          <div className="sg-seg">
            <button type="button" className={activeTab === 'rules' ? 'on' : ''} onClick={() => setActiveTab('rules')}>
              <Layers size={15} /> Danh sách quy tắc ({rules.length})
            </button>
            <button type="button" className={activeTab === 'add' ? 'on' : ''} onClick={() => setActiveTab('add')}>
              <Plus size={15} /> Thêm quy tắc / Gộp tên
            </button>
            <button type="button" className={activeTab === 'suggest' ? 'on' : ''} onClick={() => setActiveTab('suggest')}>
              <Zap size={15} /> Gợi ý gộp
              {smartSuggestions.length > 0 && (
                <span className="sg-num" style={{ fontSize: 12, fontWeight: 700, padding: '0 7px', borderRadius: 99, lineHeight: '18px', background: activeTab === 'suggest' ? 'rgba(255,255,255,.25)' : '#DDF5E6', color: activeTab === 'suggest' ? '#fff' : '#0D7A3C' }}>
                  {smartSuggestions.length}
                </span>
              )}
            </button>
            <button type="button" className={activeTab === 'sql' ? 'on' : ''} onClick={() => setActiveTab('sql')}>
              <Database size={15} /> Câu lệnh SQL Supabase
            </button>
          </div>
        </div>
        {/* Alert Messages */}
        {saveSuccessMsg && (
          <div
            style={{
              margin: '12px 16px 0',
              padding: '10px 14px',
              background: '#f0fdf4',
              border: '1px solid #bbf7d0',
              borderRadius: 8,
              color: '#166534',
              fontSize: 15,
              display: 'flex',
              alignItems: 'center',
              gap: 8
            }}
          >
            <CheckCircle2 size={16} color="#16a34a" />
            <span>{saveSuccessMsg}</span>
          </div>
        )}

        {saveErrorMsg && (
          <div
            style={{
              margin: '12px 16px 0',
              padding: '10px 14px',
              background: '#fef2f2',
              border: '1px solid #fecaca',
              borderRadius: 8,
              color: '#991b1b',
              fontSize: 15,
              display: 'flex',
              alignItems: 'center',
              gap: 8
            }}
          >
            <AlertCircle size={16} color="#dc2626" />
            <span style={{ flex: 1 }}>{saveErrorMsg}</span>
            {saveErrorMsg.includes('chuan_hoa_bch') && (
              <button
                onClick={() => setActiveTab('sql')}
                style={{
                  background: '#dc2626',
                  color: '#fff',
                  border: 'none',
                  borderRadius: 4,
                  padding: '4px 8px',
                  fontSize: 12,
                  cursor: 'pointer',
                  fontWeight: 600
                }}
              >
                Xem câu lệnh SQL tạo bảng
              </button>
            )}
          </div>
        )}

        {/* Body Content */}
        <div style={{ flex: 1, overflowY: 'auto', padding: '16px 20px' }}>
          {/* TAB 1: RULES LIST */}
          {activeTab === 'rules' && (
            <div>
              {/* Hàng công cụ: tìm + nút */}
              <div className="chb-tb" style={{ marginBottom: 12 }}>
                <div className="chb-search">
                  <Search size={16} className="chb-lead" />
                  <input
                    className="chb-in"
                    type="text"
                    placeholder="Tìm tên cũ hoặc tên chuẩn..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                  />
                  {searchTerm && (
                    <button type="button" className="chb-x sm" onClick={() => setSearchTerm('')} title="Xoá ô tìm">
                      <XMark size={11} w={2.3} />
                    </button>
                  )}
                </div>
                <button type="button" className="sg-btn" onClick={() => setActiveTab('add')}>
                  <Plus size={15} /> Thêm quy tắc
                </button>
                {rules.length > 0 && (
                  <button
                    type="button"
                    className="sg-btn danger"
                    style={{ marginLeft: 'auto' }}
                    onClick={() => {
                      if (window.confirm('Bạn có chắc chắn muốn xóa toàn bộ quy tắc chuẩn hóa?')) {
                        setRules([])
                      }
                    }}
                  >
                    <XMark size={12} w={2.3} /> Xóa hết
                  </button>
                )}
              </div>

              {/* Rules Table */}
              {filteredRules.length === 0 ? (
                <div
                  style={{
                    padding: '36px 20px',
                    textAlign: 'center',
                    background: '#F7F8FB',
                    borderRadius: 18,
                    border: '1px dashed #C9D3E6'
                  }}
                >
                  <Layers size={32} color="#8C8E97" style={{ margin: '0 auto 10px' }} />
                  <div style={{ margin: '0 0 6px', fontSize: 15, color: '#16171C', fontWeight: 700 }}>
                    {searchTerm ? 'Không tìm thấy quy tắc phù hợp' : 'Chưa có quy tắc chuẩn hóa tên BCH nào'}
                  </div>
                  <p style={{ margin: '0 0 14px', fontSize: 15, color: '#4A4C55', maxWidth: 480, marginInline: 'auto' }}>
                    Bấm "Thêm quy tắc / Gộp tên" hoặc dùng "Gợi ý gộp" để gom các tên BCH cũ và mới về cùng một tên chuẩn.
                  </p>
                  <button type="button" className="sg-btn" onClick={() => setActiveTab('add')}>
                    <Plus size={15} /> Tạo quy tắc đầu tiên
                  </button>
                </div>
              ) : (
                <div className="chb-scroll">
                  <table style={{ fontSize: 15 }}>
                    <colgroup>
                      <col style={{ width: 56 }} />
                      <col style={{ width: '40%' }} />
                      <col style={{ width: 40 }} />
                      <col />
                      <col style={{ width: 84 }} />
                    </colgroup>
                    <thead>
                      <tr>
                        <th>STT</th>
                        <th>Tên cũ / Biến thể trong Excel</th>
                        <th></th>
                        <th>Tên BCH chuẩn hóa (hệ thống)</th>
                        <th>Thao tác</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRules.map((rule, idx) => (
                        <tr key={idx}>
                          <td className="sg-num" style={{ textAlign: 'center', color: '#4A4C55' }}>
                            {idx + 1}
                          </td>
                          <td>
                            <div style={{ fontWeight: 600, color: '#2A2C33', wordBreak: 'break-word' }}>{rule.ten_cu}</div>
                            {rule.ghi_chu && (
                              <div style={{ fontSize: 12, color: '#4A4C55', fontWeight: 600, marginTop: 2 }}>
                                Ghi chú: {rule.ghi_chu}
                              </div>
                            )}
                          </td>
                          <td style={{ textAlign: 'center' }}>
                            <ArrowRight size={16} color="#0F58A7" />
                          </td>
                          <td>
                            {editingRuleIndex === idx ? (
                              <div style={{ display: 'flex', gap: 6, alignItems: 'center', flexWrap: 'wrap' }}>
                                <input
                                  type="text"
                                  className="chb-edit"
                                  style={{ flex: '1 1 200px', width: 'auto' }}
                                  value={editStandardValue}
                                  onChange={(e) => setEditStandardValue(e.target.value)}
                                  autoFocus
                                  onFocus={(e) => { const l = e.target.value.length; e.target.setSelectionRange(l, l) }}
                                />
                                <button
                                  type="button"
                                  className="chb-mini ok"
                                  onClick={() => {
                                    if (!editStandardValue.trim()) return
                                    const next = [...rules]
                                    next[idx].ten_chuan = norm(editStandardValue)
                                    setRules(next)
                                    setEditingRuleIndex(null)
                                  }}
                                >
                                  <Check size={13} /> Lưu
                                </button>
                                <button type="button" className="chb-mini no" onClick={() => setEditingRuleIndex(null)}>
                                  Hủy
                                </button>
                              </div>
                            ) : (
                              <div style={{ fontWeight: 700, color: '#0A3D73', wordBreak: 'break-word' }}>
                                <span>{rule.ten_chuan}</span>
                                <button
                                  type="button"
                                  className="chb-pen"
                                  style={{ marginLeft: 4, marginTop: -2 }}
                                  onClick={() => {
                                    setEditingRuleIndex(idx)
                                    setEditStandardValue(rule.ten_chuan)
                                  }}
                                  title="Đổi tên chuẩn"
                                >
                                  <PenIcon />
                                </button>
                              </div>
                            )}
                          </td>
                          <td style={{ textAlign: 'center' }}>
                            <button
                              type="button"
                              className="chb-x"
                              onClick={() => handleDeleteRule(idx)}
                              title="Xóa quy tắc"
                            >
                              <XMark size={13} w={2.3} />
                            </button>
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              )}
            </div>
          )}

          {/* TAB 2: ADD / MERGE RULE */}
          {activeTab === 'add' && (
            <div style={{ display: 'flex', flexDirection: 'column', gap: 16 }}>
              {/* Step 1: Destination Standard Name */}
              <div
                style={{
                  background: '#f8fafc',
                  border: '1px solid #e2e8f0',
                  borderRadius: 12,
                  padding: '14px 16px'
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 15, color: '#1e293b', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
                  <span
                    style={{
                      width: 22,
                      height: 22,
                      borderRadius: '50%',
                      background: '#2563eb',
                      color: '#fff',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      fontSize: 12
                    }}
                  >
                    1
                  </span>
                  Chọn hoặc Nhập Tên Ban Chỉ Huy Chuẩn (Đích đến thống nhất)
                </div>

                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
                  <div>
                    <label style={{ fontSize: 12, fontWeight: 600, color: '#64748b', display: 'block', marginBottom: 4 }}>
                      Chọn từ danh sách BCH hiện có:
                    </label>
                    <select
                      value={selectedStandard}
                      onChange={(e) => {
                        setSelectedStandard(e.target.value)
                        if (e.target.value) setCustomStandardInput('')
                      }}
                      style={{
                        width: '100%',
                        padding: '8px 10px',
                        borderRadius: 8,
                        border: '1px solid #cbd5e1',
                        fontSize: 15,
                        outline: 'none',
                        background: '#ffffff'
                      }}
                    >
                      <option value="">-- Chọn tên chuẩn có sẵn --</option>
                      {standardOptions.map((opt, i) => (
                        <option key={i} value={opt}>
                          {opt}
                        </option>
                      ))}
                    </select>
                  </div>

                  <div>
                    <label style={{ fontSize: 12, fontWeight: 600, color: '#64748b', display: 'block', marginBottom: 4 }}>
                      Hoặc tự nhập Tên BCH Chuẩn Mới:
                    </label>
                    <input
                      type="text"
                      placeholder="Ví dụ: BCH Cọc Khoan Nhồi - Trống Đồng"
                      value={customStandardInput}
                      onChange={(e) => {
                        setCustomStandardInput(e.target.value)
                        if (e.target.value) setSelectedStandard('')
                      }}
                      style={{
                        width: '100%',
                        padding: '8px 10px',
                        borderRadius: 8,
                        border: '1px solid #cbd5e1',
                        fontSize: 15,
                        outline: 'none',
                        boxSizing: 'border-box'
                      }}
                    />
                  </div>
                </div>

                {(customStandardInput || selectedStandard) && (
                  <div style={{ marginTop: 10, fontSize: 12, color: '#1e40af', display: 'flex', alignItems: 'center', gap: 6 }}>
                    <Check size={14} color="#16a34a" />
                    Tên chuẩn sẽ áp dụng: <strong>{customStandardInput || selectedStandard}</strong>
                  </div>
                )}
              </div>

              {/* Step 2: Select Old / Variant Names */}
              <div
                style={{
                  background: '#f8fafc',
                  border: '1px solid #e2e8f0',
                  borderRadius: 12,
                  padding: '14px 16px'
                }}
              >
                <div style={{ fontWeight: 700, fontSize: 15, color: '#1e293b', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
                  <span
                    style={{
                      width: 22,
                      height: 22,
                      borderRadius: '50%',
                      background: '#2563eb',
                      color: '#fff',
                      display: 'flex',
                      alignItems: 'center',
                      justifyContent: 'center',
                      fontSize: 12
                    }}
                  >
                    2
                  </span>
                  Chọn Các Tên Cũ / Biến Thể Cần Quy Về Tên Chuẩn Ở Trên
                  <span style={{ fontSize: 12, fontWeight: 500, color: '#64748b' }}>
                    (Đã chọn {selectedOldNames.length} tên)
                  </span>
                </div>

                {/* Filter and Quick Select */}
                <div style={{ display: 'flex', gap: 10, marginBottom: 10, alignItems: 'center' }}>
                  <div style={{ position: 'relative', flex: 1 }}>
                    <Search
                      size={14}
                      color="#94a3b8"
                      style={{ position: 'absolute', left: 8, top: '50%', transform: 'translateY(-50%)' }}
                    />
                    <input
                      type="text"
                      placeholder="Lọc nhanh tên đơn vị..."
                      value={unitSearch}
                      onChange={(e) => setUnitSearch(e.target.value)}
                      style={{
                        width: '100%',
                        padding: '6px 10px 6px 28px',
                        borderRadius: 6,
                        border: '1px solid #cbd5e1',
                        fontSize: 12,
                        outline: 'none',
                        boxSizing: 'border-box'
                      }}
                    />
                  </div>

                  <button
                    onClick={() => {
                      const allVisible = filteredUnitsForSelection.map((u) => norm(u))
                      setSelectedOldNames(Array.from(new Set([...selectedOldNames, ...allVisible])))
                    }}
                    style={{
                      padding: '6px 10px',
                      background: '#eff6ff',
                      color: '#2563eb',
                      border: '1px solid #bfdbfe',
                      borderRadius: 6,
                      fontSize: 12,
                      fontWeight: 600,
                      cursor: 'pointer'
                    }}
                  >
                    Chọn tất cả đang hiện
                  </button>

                  <button
                    onClick={() => setSelectedOldNames([])}
                    style={{
                      padding: '6px 10px',
                      background: '#ffffff',
                      color: '#64748b',
                      border: '1px solid #cbd5e1',
                      borderRadius: 6,
                      fontSize: 12,
                      cursor: 'pointer'
                    }}
                  >
                    Bỏ chọn
                  </button>
                </div>

                {/* Scrollable Checkbox List */}
                <div
                  style={{
                    maxHeight: '220px',
                    overflowY: 'auto',
                    border: '1px solid #cbd5e1',
                    borderRadius: 8,
                    background: '#ffffff',
                    padding: '6px 8px'
                  }}
                >
                  {filteredUnitsForSelection.length === 0 ? (
                    <div style={{ padding: 16, textAlign: 'center', color: '#94a3b8', fontSize: 12 }}>
                      Không có tên nào khớp với tìm kiếm
                    </div>
                  ) : (
                    filteredUnitsForSelection.map((unitName, i) => {
                      const isSelected = selectedOldNames.includes(norm(unitName))
                      const existingRule = rules.find((r) => normKey(r.ten_cu) === normKey(unitName))
                      const isCurrentStandard =
                        normKey(unitName) === normKey(customStandardInput || selectedStandard)

                      return (
                        <label
                          key={i}
                          style={{
                            display: 'flex',
                            alignItems: 'center',
                            gap: 8,
                            padding: '6px 8px',
                            borderRadius: 6,
                            cursor: isCurrentStandard ? 'not-allowed' : 'pointer',
                            background: isSelected ? '#eff6ff' : 'transparent',
                            opacity: isCurrentStandard ? 0.4 : 1,
                            fontSize: 12,
                            borderBottom: '1px solid #f1f5f9'
                          }}
                        >
                          <input
                            type="checkbox"
                            checked={isSelected}
                            disabled={isCurrentStandard}
                            onChange={(e) => {
                              const n = norm(unitName)
                              if (e.target.checked) {
                                setSelectedOldNames((prev) => [...prev, n])
                              } else {
                                setSelectedOldNames((prev) => prev.filter((item) => item !== n))
                              }
                            }}
                          />
                          <span style={{ flex: 1, fontWeight: isSelected ? 600 : 400, color: '#1e293b' }}>
                            {unitName}
                          </span>
                          {existingRule && (
                            <span
                              style={{
                                fontSize: 12,
                                color: '#2563eb',
                                background: '#dbeafe',
                                padding: '1px 6px',
                                borderRadius: 4
                              }}
                            >
                              Đang quy về: {existingRule.ten_chuan}
                            </span>
                          )}
                        </label>
                      )
                    })
                  )}
                </div>
              </div>

              {/* Step 3: Note (Optional) & Submit */}
              <div style={{ display: 'flex', gap: 12, alignItems: 'flex-end' }}>
                <div style={{ flex: 1 }}>
                  <label style={{ fontSize: 12, fontWeight: 600, color: '#64748b', display: 'block', marginBottom: 4 }}>
                    Ghi chú quy tắc (Tùy chọn):
                  </label>
                  <input
                    type="text"
                    placeholder="Ví dụ: Tên viết tắt dự án 2024..."
                    value={ruleNote}
                    onChange={(e) => setRuleNote(e.target.value)}
                    style={{
                      width: '100%',
                      padding: '8px 10px',
                      borderRadius: 8,
                      border: '1px solid #cbd5e1',
                      fontSize: 15,
                      outline: 'none',
                      boxSizing: 'border-box'
                    }}
                  />
                </div>

                <button
                  onClick={handleAddRule}
                  style={{
                    background: 'linear-gradient(135deg, #2563eb 0%, #3b82f6 100%)',
                    color: '#ffffff',
                    border: 'none',
                    borderRadius: 8,
                    padding: '9px 20px',
                    fontSize: 15,
                    fontWeight: 700,
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    gap: 6,
                    boxShadow: '0 2px 4px rgba(37,99,235,0.2)'
                  }}
                >
                  <Plus size={16} /> Thêm Quy Tắc
                </button>
              </div>
            </div>
          )}

          {/* TAB 3: SMART SUGGESTIONS */}
          {activeTab === 'suggest' && (
            <div>
              <div
                style={{
                  background: '#f0fdf4',
                  border: '1px solid #bbf7d0',
                  borderRadius: 10,
                  padding: '12px 14px',
                  marginBottom: 14,
                  display: 'flex',
                  gap: 10,
                  alignItems: 'center'
                }}
              >
                <Zap size={20} color="#16a34a" style={{ flexShrink: 0 }} />
                <div style={{ fontSize: 15, color: '#166534', lineHeight: 1.4 }}>
                  Hệ thống tự động phát hiện các tên BCH gần giống nhau (như có ký tự <code>•</code>, khác số đuôi <code>01, 02</code>, hoặc viết tắt <code>CKN</code> vs <code>Cọc Khoan Nhồi</code>). Bạn có thể bấm <strong>"Gộp tên"</strong> để áp dụng ngay!
                </div>
              </div>

              {smartSuggestions.length === 0 ? (
                <div
                  style={{
                    padding: '40px 20px',
                    textAlign: 'center',
                    background: '#f8fafc',
                    borderRadius: 12,
                    border: '1px dashed #cbd5e1'
                  }}
                >
                  <CheckCircle2 size={36} color="#16a34a" style={{ margin: '0 auto 12px' }} />
                  <h4 style={{ margin: '0 0 6px', fontSize: 15, color: '#334155', fontWeight: 600 }}>
                    Tất cả các tên BCH hiện tại đã rất đồng nhất!
                  </h4>
                  <p style={{ margin: 0, fontSize: 15, color: '#64748b' }}>
                    Không phát hiện thêm các biến thể tên bị phân mảnh trong tệp hiện tại.
                  </p>
                </div>
              ) : (
                <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
                  {smartSuggestions.map((sug) => (
                    <div
                      key={sug.id}
                      style={{
                        background: '#ffffff',
                        border: '1px solid #e2e8f0',
                        borderRadius: 10,
                        padding: '14px 16px',
                        boxShadow: '0 1px 3px rgba(0,0,0,0.02)',
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'space-between',
                        gap: 16
                      }}
                    >
                      <div style={{ flex: 1 }}>
                        <div style={{ fontSize: 12, color: '#64748b', marginBottom: 4, fontWeight: 600 }}>
                          Gợi ý gom {sug.oldItems.length} tên biến thể về:
                        </div>
                        <div
                          style={{
                            fontSize: 15,
                            fontWeight: 700,
                            color: '#1e40af',
                            display: 'flex',
                            alignItems: 'center',
                            gap: 6,
                            marginBottom: 8
                          }}
                        >
                          <Building size={16} color="#2563eb" />
                          <span>{sug.suggestedStandard}</span>
                        </div>

                        <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                          {sug.oldItems.map((old, oIdx) => (
                            <span
                              key={oIdx}
                              style={{
                                background: '#fef2f2',
                                border: '1px solid #fecaca',
                                color: '#991b1b',
                                fontSize: 12,
                                padding: '2px 8px',
                                borderRadius: 4,
                                textDecoration: 'line-through'
                              }}
                            >
                              {old}
                            </span>
                          ))}
                        </div>
                      </div>

                      <button
                        onClick={() => handleApplySuggestion(sug)}
                        style={{
                          background: 'linear-gradient(135deg, #059669 0%, #10b981 100%)',
                          color: '#ffffff',
                          border: 'none',
                          borderRadius: 8,
                          padding: '8px 16px',
                          fontSize: 15,
                          fontWeight: 700,
                          cursor: 'pointer',
                          display: 'flex',
                          alignItems: 'center',
                          gap: 6,
                          flexShrink: 0,
                          boxShadow: '0 2px 4px rgba(16,185,129,0.2)'
                        }}
                      >
                        <Zap size={14} /> Gộp Về Tên Này
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}

          {/* TAB 4: SQL SCRIPT & SUPABASE SETUP */}
          {activeTab === 'sql' && (
            <div>
              <div
                style={{
                  background: '#fffbeb',
                  border: '1px solid #fef3c7',
                  borderRadius: 10,
                  padding: '12px 14px',
                  marginBottom: 14,
                  display: 'flex',
                  gap: 10,
                  alignItems: 'center'
                }}
              >
                <Info size={20} color="#d97706" style={{ flexShrink: 0 }} />
                <div style={{ fontSize: 15, color: '#92400e', lineHeight: 1.5 }}>
                  Để lưu trữ vĩnh viễn các quy tắc chuẩn hóa tên BCH trên Supabase và đồng bộ tự động cho tất cả người dùng, hãy chạy đoạn mã SQL dưới đây trong mục <strong>SQL Editor</strong> của Supabase.
                </div>
              </div>

              {/* Instructions */}
              <div
                style={{
                  background: '#f8fafc',
                  border: '1px solid #e2e8f0',
                  borderRadius: 10,
                  padding: '12px 16px',
                  marginBottom: 14,
                  fontSize: 15,
                  color: '#334155'
                }}
              >
                <strong style={{ display: 'block', marginBottom: 6, color: '#0f172a' }}>
                  4 Bước Thực Hiện Trên Supabase:
                </strong>
                <ol style={{ margin: 0, paddingLeft: 18, lineHeight: 1.6 }}>
                  <li>
                    Bấm nút <strong>"Copy toàn bộ mã SQL"</strong> ở góc dưới bên phải.
                  </li>
                  <li>
                    Truy cập vào{' '}
                    <a
                      href="https://supabase.com/dashboard"
                      target="_blank"
                      rel="noreferrer"
                      style={{ color: '#2563eb', fontWeight: 600, textDecoration: 'underline' }}
                    >
                      Supabase Dashboard
                    </a>
                    , chọn dự án của bạn.
                  </li>
                  <li>
                    Nhấn vào <strong>SQL Editor</strong> (biểu tượng <code>&gt;_</code> ở thanh bên trái), chọn{' '}
                    <strong>New query</strong>.
                  </li>
                  <li>
                    Dán (Paste) toàn bộ mã vừa copy rồi bấm nút <strong>RUN</strong> (hoặc nhấn phím <code>Ctrl + Enter</code>).
                  </li>
                </ol>
              </div>

              {/* Code Box */}
              <div style={{ position: 'relative' }}>
                <pre
                  style={{
                    background: '#0f172a',
                    color: '#e2e8f0',
                    padding: '16px',
                    borderRadius: 10,
                    fontSize: 12,
                    lineHeight: 1.5,
                    fontFamily: 'Consolas, Monaco, "Courier New", monospace',
                    overflowX: 'auto',
                    margin: 0,
                    maxHeight: '260px'
                  }}
                >
                  {SQL_SETUP_SCRIPT}
                </pre>

                <button
                  onClick={() => {
                    navigator.clipboard.writeText(SQL_SETUP_SCRIPT)
                    setCopiedSql(true)
                    setTimeout(() => setCopiedSql(false), 2500)
                  }}
                  style={{
                    position: 'absolute',
                    top: 10,
                    right: 10,
                    background: copiedSql ? '#16a34a' : 'rgba(255,255,255,0.15)',
                    color: '#ffffff',
                    border: 'none',
                    borderRadius: 6,
                    padding: '6px 12px',
                    fontSize: 12,
                    fontWeight: 600,
                    cursor: 'pointer',
                    display: 'flex',
                    alignItems: 'center',
                    gap: 6,
                    backdropFilter: 'blur(4px)'
                  }}
                >
                  {copiedSql ? <Check size={14} /> : <Copy size={14} />}
                  {copiedSql ? 'Đã copy mã SQL!' : 'Copy toàn bộ mã SQL'}
                </button>
              </div>
            </div>
          )}
        </div>

        {/* Chân hộp thoại */}
        <div className="chb-dlg-ft" style={{ flexShrink: 0 }}>
          <span className={'sg-conn ' + (isSupabaseConfigured ? 'ok' : 'off')}>
            <i />
            {isSupabaseConfigured ? 'Đã kết nối Supabase' : 'Lưu trên máy (chưa kết nối Supabase)'}
          </span>

          <div style={{ display: 'flex', gap: 8 }}>
            <button type="button" className="sg-btn ghost" onClick={onClose}>
              Đóng
            </button>
            <button type="button" className="sg-btn" onClick={handleSaveAndApply} disabled={saving}>
              <RefreshCw size={15} className={saving ? 'animate-spin' : ''} />
              {saving ? 'Đang lưu...' : 'Lưu lên Supabase và áp dụng'}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
