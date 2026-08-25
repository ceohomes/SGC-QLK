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

  if (!isOpen) return null

  return (
    <div
      style={{
        position: 'fixed',
        top: 0,
        left: 0,
        right: 0,
        bottom: 0,
        backgroundColor: 'rgba(15, 23, 42, 0.65)',
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
          borderRadius: '16px',
          boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)',
          display: 'flex',
          flexDirection: 'column',
          overflow: 'hidden',
          border: '1px solid #e2e8f0'
        }}
      >
        {/* Header */}
        <div
          style={{
            padding: '16px 22px',
            background: 'linear-gradient(135deg, #0f172a 0%, #1e293b 100%)',
            color: '#ffffff',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flexShrink: 0
          }}
        >
          <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
            <div
              style={{
                width: 38,
                height: 38,
                borderRadius: 10,
                background: 'linear-gradient(135deg, #2563eb 0%, #3b82f6 100%)',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                boxShadow: '0 4px 10px rgba(37,99,235,0.3)'
              }}
            >
              <Sparkles size={20} color="#ffffff" />
            </div>
            <div>
              <h2 style={{ fontSize: 17, fontWeight: 700, margin: 0, display: 'flex', alignItems: 'center', gap: 8 }}>
                Chuẩn Hóa & Hợp Nhất Tên Ban Chỉ Huy (Kho BCH)
                <span
                  style={{
                    fontSize: 12,
                    fontWeight: 600,
                    padding: '2px 8px',
                    borderRadius: 12,
                    background: 'rgba(59, 130, 246, 0.25)',
                    border: '1px solid rgba(147, 197, 253, 0.3)',
                    color: '#93c5fd'
                  }}
                >
                  {rules.length} quy tắc
                </span>
              </h2>
              <p style={{ margin: '2px 0 0', fontSize: 12.5, color: '#94a3b8' }}>
                Tự động gom và quy đổi các tên cũ/biến thể trong Excel về một tên BCH chuẩn hóa duy nhất khi up đơn
              </p>
            </div>
          </div>

          <button
            onClick={onClose}
            style={{
              background: 'rgba(255,255,255,0.1)',
              border: 'none',
              borderRadius: 8,
              padding: 6,
              color: '#cbd5e1',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'center',
              transition: 'all 0.15s'
            }}
            title="Đóng modal"
          >
            <X size={18} />
          </button>
        </div>

        {/* Tab Navigation */}
        <div
          style={{
            display: 'flex',
            borderBottom: '1px solid #e2e8f0',
            background: '#f8fafc',
            padding: '0 16px',
            flexShrink: 0,
            gap: 8
          }}
        >
          <button
            onClick={() => setActiveTab('rules')}
            style={{
              padding: '12px 16px',
              fontSize: 13.5,
              fontWeight: activeTab === 'rules' ? 700 : 500,
              color: activeTab === 'rules' ? '#2563eb' : '#64748b',
              borderBottom: activeTab === 'rules' ? '2.5px solid #2563eb' : '2.5px solid transparent',
              background: 'transparent',
              borderTop: 'none',
              borderLeft: 'none',
              borderRight: 'none',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: 6
            }}
          >
            <Layers size={16} />
            Danh sách Quy tắc ({rules.length})
          </button>

          <button
            onClick={() => setActiveTab('add')}
            style={{
              padding: '12px 16px',
              fontSize: 13.5,
              fontWeight: activeTab === 'add' ? 700 : 500,
              color: activeTab === 'add' ? '#2563eb' : '#64748b',
              borderBottom: activeTab === 'add' ? '2.5px solid #2563eb' : '2.5px solid transparent',
              background: 'transparent',
              borderTop: 'none',
              borderLeft: 'none',
              borderRight: 'none',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: 6
            }}
          >
            <Plus size={16} />
            Thêm quy tắc / Gộp tên mới
          </button>

          <button
            onClick={() => setActiveTab('suggest')}
            style={{
              padding: '12px 16px',
              fontSize: 13.5,
              fontWeight: activeTab === 'suggest' ? 700 : 500,
              color: activeTab === 'suggest' ? '#059669' : '#64748b',
              borderBottom: activeTab === 'suggest' ? '2.5px solid #059669' : '2.5px solid transparent',
              background: 'transparent',
              borderTop: 'none',
              borderLeft: 'none',
              borderRight: 'none',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: 6
            }}
          >
            <Zap size={16} color={activeTab === 'suggest' ? '#059669' : '#64748b'} />
            Gợi ý gộp thông minh
            {smartSuggestions.length > 0 && (
              <span
                style={{
                  background: '#dcfce7',
                  color: '#15803d',
                  fontSize: 11,
                  fontWeight: 700,
                  padding: '1px 6px',
                  borderRadius: 10
                }}
              >
                {smartSuggestions.length}
              </span>
            )}
          </button>

          <button
            onClick={() => setActiveTab('sql')}
            style={{
              padding: '12px 16px',
              fontSize: 13.5,
              fontWeight: activeTab === 'sql' ? 700 : 500,
              color: activeTab === 'sql' ? '#d97706' : '#64748b',
              borderBottom: activeTab === 'sql' ? '2.5px solid #d97706' : '2.5px solid transparent',
              background: 'transparent',
              borderTop: 'none',
              borderLeft: 'none',
              borderRight: 'none',
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: 6
            }}
          >
            <Database size={16} color={activeTab === 'sql' ? '#d97706' : '#64748b'} />
            Câu lệnh SQL Supabase
          </button>
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
              fontSize: 13,
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
              fontSize: 13,
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
                  fontSize: 11.5,
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
              {/* Search & Actions Bar */}
              <div
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  marginBottom: 14,
                  gap: 12
                }}
              >
                <div
                  style={{
                    position: 'relative',
                    flex: 1,
                    maxWidth: 420
                  }}
                >
                  <Search
                    size={15}
                    color="#94a3b8"
                    style={{ position: 'absolute', left: 10, top: '50%', transform: 'translateY(-50%)' }}
                  />
                  <input
                    type="text"
                    placeholder="Tìm kiếm tên cũ hoặc tên chuẩn..."
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                    style={{
                      width: '100%',
                      padding: '8px 12px 8px 32px',
                      borderRadius: 8,
                      border: '1px solid #cbd5e1',
                      fontSize: 13,
                      outline: 'none',
                      boxSizing: 'border-box'
                    }}
                  />
                </div>

                <div style={{ display: 'flex', gap: 8 }}>
                  <button
                    onClick={() => setActiveTab('add')}
                    style={{
                      background: 'linear-gradient(135deg, #2563eb 0%, #3b82f6 100%)',
                      color: '#ffffff',
                      border: 'none',
                      borderRadius: 8,
                      padding: '8px 14px',
                      fontSize: 13,
                      fontWeight: 600,
                      cursor: 'pointer',
                      display: 'flex',
                      alignItems: 'center',
                      gap: 6
                    }}
                  >
                    <Plus size={15} /> Thêm quy tắc
                  </button>

                  {rules.length > 0 && (
                    <button
                      onClick={() => {
                        if (window.confirm('Bạn có chắc chắn muốn xóa toàn bộ quy tắc chuẩn hóa?')) {
                          setRules([])
                        }
                      }}
                      style={{
                        background: '#f8fafc',
                        color: '#64748b',
                        border: '1px solid #cbd5e1',
                        borderRadius: 8,
                        padding: '8px 12px',
                        fontSize: 13,
                        fontWeight: 500,
                        cursor: 'pointer',
                        display: 'flex',
                        alignItems: 'center',
                        gap: 6
                      }}
                    >
                      <Trash2 size={14} /> Xóa hết
                    </button>
                  )}
                </div>
              </div>

              {/* Rules Table */}
              {filteredRules.length === 0 ? (
                <div
                  style={{
                    padding: '40px 20px',
                    textAlign: 'center',
                    background: '#f8fafc',
                    borderRadius: 12,
                    border: '1px dashed #cbd5e1'
                  }}
                >
                  <Layers size={36} color="#94a3b8" style={{ margin: '0 auto 12px' }} />
                  <h4 style={{ margin: '0 0 6px', fontSize: 15, color: '#334155', fontWeight: 600 }}>
                    {searchTerm ? 'Không tìm thấy quy tắc phù hợp' : 'Chưa có quy tắc chuẩn hóa tên BCH nào'}
                  </h4>
                  <p style={{ margin: '0 0 16px', fontSize: 13, color: '#64748b', maxWidth: 460, marginInline: 'auto' }}>
                    Hãy bấm nút "Thêm quy tắc / Gộp tên mới" hoặc sử dụng tính năng "Gợi ý gộp thông minh" để gom các tên BCH cũ và mới về cùng 1 tên chuẩn!
                  </p>
                  <button
                    onClick={() => setActiveTab('add')}
                    style={{
                      background: '#2563eb',
                      color: '#fff',
                      border: 'none',
                      borderRadius: 8,
                      padding: '8px 16px',
                      fontSize: 13,
                      fontWeight: 600,
                      cursor: 'pointer'
                    }}
                  >
                    + Tạo quy tắc đầu tiên
                  </button>
                </div>
              ) : (
                <div
                  style={{
                    border: '1px solid #e2e8f0',
                    borderRadius: 10,
                    overflow: 'hidden'
                  }}
                >
                  <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                    <thead>
                      <tr style={{ background: '#f1f5f9', borderBottom: '1px solid #e2e8f0', color: '#475569' }}>
                        <th style={{ padding: '10px 12px', textAlign: 'center', width: 45, fontWeight: 700 }}>STT</th>
                        <th style={{ padding: '10px 12px', textAlign: 'left', fontWeight: 700, width: '38%' }}>
                          Tên cũ / Biến thể trong Excel
                        </th>
                        <th style={{ padding: '10px 12px', textAlign: 'center', width: 30 }}></th>
                        <th style={{ padding: '10px 12px', textAlign: 'left', fontWeight: 700, width: '38%' }}>
                          Tên BCH chuẩn hóa (Hệ thống)
                        </th>
                        <th style={{ padding: '10px 12px', textAlign: 'center', width: 80, fontWeight: 700 }}>Thao tác</th>
                      </tr>
                    </thead>
                    <tbody>
                      {filteredRules.map((rule, idx) => (
                        <tr
                          key={idx}
                          style={{
                            borderBottom: '1px solid #f1f5f9',
                            background: idx % 2 === 1 ? '#fafafa' : '#ffffff'
                          }}
                        >
                          <td style={{ padding: '10px 12px', textAlign: 'center', color: '#94a3b8', fontWeight: 600 }}>
                            {idx + 1}
                          </td>
                          <td style={{ padding: '10px 12px' }}>
                            <div
                              style={{
                                display: 'inline-flex',
                                alignItems: 'center',
                                gap: 6,
                                background: '#fef2f2',
                                border: '1px solid #fecaca',
                                padding: '4px 10px',
                                borderRadius: 6,
                                color: '#991b1b',
                                fontWeight: 600,
                                fontSize: 12.5,
                                wordBreak: 'break-word'
                              }}
                            >
                              <span>{rule.ten_cu}</span>
                            </div>
                            {rule.ghi_chu && (
                              <div style={{ fontSize: 11, color: '#64748b', marginTop: 3 }}>
                                <em>Ghi chú: {rule.ghi_chu}</em>
                              </div>
                            )}
                          </td>
                          <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                            <ArrowRight size={16} color="#3b82f6" />
                          </td>
                          <td style={{ padding: '10px 12px' }}>
                            {editingRuleIndex === idx ? (
                              <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                                <input
                                  type="text"
                                  value={editStandardValue}
                                  onChange={(e) => setEditStandardValue(e.target.value)}
                                  style={{
                                    flex: 1,
                                    padding: '4px 8px',
                                    borderRadius: 6,
                                    border: '1px solid #2563eb',
                                    fontSize: 12.5
                                  }}
                                  autoFocus
                                />
                                <button
                                  onClick={() => {
                                    if (!editStandardValue.trim()) return
                                    const next = [...rules]
                                    next[idx].ten_chuan = norm(editStandardValue)
                                    setRules(next)
                                    setEditingRuleIndex(null)
                                  }}
                                  style={{
                                    background: '#16a34a',
                                    color: '#fff',
                                    border: 'none',
                                    borderRadius: 4,
                                    padding: '4px 8px',
                                    fontSize: 11.5,
                                    cursor: 'pointer'
                                  }}
                                >
                                  Lưu
                                </button>
                                <button
                                  onClick={() => setEditingRuleIndex(null)}
                                  style={{
                                    background: '#94a3b8',
                                    color: '#fff',
                                    border: 'none',
                                    borderRadius: 4,
                                    padding: '4px 8px',
                                    fontSize: 11.5,
                                    cursor: 'pointer'
                                  }}
                                >
                                  Hủy
                                </button>
                              </div>
                            ) : (
                              <div
                                style={{
                                  display: 'inline-flex',
                                  alignItems: 'center',
                                  gap: 6,
                                  background: '#eff6ff',
                                  border: '1px solid #bfdbfe',
                                  padding: '4px 10px',
                                  borderRadius: 6,
                                  color: '#1e40af',
                                  fontWeight: 700,
                                  fontSize: 12.5,
                                  wordBreak: 'break-word'
                                }}
                              >
                                <Building size={13} color="#2563eb" />
                                <span>{rule.ten_chuan}</span>
                              </div>
                            )}
                          </td>
                          <td style={{ padding: '10px 12px', textAlign: 'center' }}>
                            <div style={{ display: 'flex', gap: 6, justifyContent: 'center' }}>
                              <button
                                onClick={() => {
                                  setEditingRuleIndex(idx)
                                  setEditStandardValue(rule.ten_chuan)
                                }}
                                style={{
                                  background: 'transparent',
                                  border: 'none',
                                  color: '#2563eb',
                                  cursor: 'pointer',
                                  padding: 4
                                }}
                                title="Đổi tên chuẩn"
                              >
                                <Edit2 size={14} />
                              </button>
                              <button
                                onClick={() => handleDeleteRule(idx)}
                                style={{
                                  background: 'transparent',
                                  border: 'none',
                                  color: '#ef4444',
                                  cursor: 'pointer',
                                  padding: 4
                                }}
                                title="Xóa quy tắc"
                              >
                                <Trash2 size={14} />
                              </button>
                            </div>
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
                <div style={{ fontWeight: 700, fontSize: 14, color: '#1e293b', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
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
                        fontSize: 13,
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
                        fontSize: 13,
                        outline: 'none',
                        boxSizing: 'border-box'
                      }}
                    />
                  </div>
                </div>

                {(customStandardInput || selectedStandard) && (
                  <div style={{ marginTop: 10, fontSize: 12.5, color: '#1e40af', display: 'flex', alignItems: 'center', gap: 6 }}>
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
                <div style={{ fontWeight: 700, fontSize: 14, color: '#1e293b', marginBottom: 10, display: 'flex', alignItems: 'center', gap: 8 }}>
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
                        fontSize: 12.5,
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
                    <div style={{ padding: 16, textAlign: 'center', color: '#94a3b8', fontSize: 12.5 }}>
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
                            fontSize: 12.5,
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
                                fontSize: 11,
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
                      fontSize: 13,
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
                    fontSize: 13.5,
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
                <div style={{ fontSize: 13, color: '#166534', lineHeight: 1.4 }}>
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
                  <p style={{ margin: 0, fontSize: 13, color: '#64748b' }}>
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
                            fontSize: 14,
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
                          fontSize: 13,
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
                <div style={{ fontSize: 13, color: '#92400e', lineHeight: 1.5 }}>
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
                  fontSize: 13,
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
                    fontSize: 12.5,
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

        {/* Modal Footer */}
        <div
          style={{
            padding: '14px 20px',
            background: '#f8fafc',
            borderTop: '1px solid #e2e8f0',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            flexShrink: 0
          }}
        >
          <div style={{ fontSize: 12.5, color: '#64748b', display: 'flex', alignItems: 'center', gap: 6 }}>
            <span
              style={{
                width: 8,
                height: 8,
                borderRadius: '50%',
                background: isSupabaseConfigured ? '#22c55e' : '#f59e0b'
              }}
            />
            <span>{isSupabaseConfigured ? 'Supabase Connected' : 'Chế độ lưu trữ Offline (Local)'}</span>
          </div>

          <div style={{ display: 'flex', gap: 10 }}>
            <button
              onClick={onClose}
              style={{
                background: '#ffffff',
                border: '1px solid #cbd5e1',
                color: '#475569',
                padding: '8px 16px',
                borderRadius: 8,
                fontSize: 13.5,
                fontWeight: 600,
                cursor: 'pointer'
              }}
            >
              Đóng
            </button>

            <button
              onClick={handleSaveAndApply}
              disabled={saving}
              style={{
                background: 'linear-gradient(135deg, #2563eb 0%, #3b82f6 100%)',
                color: '#ffffff',
                border: 'none',
                padding: '8px 20px',
                borderRadius: 8,
                fontSize: 13.5,
                fontWeight: 700,
                cursor: saving ? 'not-allowed' : 'pointer',
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                boxShadow: '0 2px 6px rgba(37,99,235,0.25)',
                opacity: saving ? 0.7 : 1
              }}
            >
              <RefreshCw size={15} className={saving ? 'animate-spin' : ''} />
              {saving ? 'Đang lưu...' : '💾 Lưu Lên Supabase & Áp Dụng Ngay'}
            </button>
          </div>
        </div>
      </div>
    </div>
  )
}
