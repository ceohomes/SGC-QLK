import React, { useState, useEffect } from 'react'
import { motion, AnimatePresence } from 'motion/react'
import { 
  Search, Users, UserCheck, Shield, Plus, Pencil, Trash2, X, Lock, 
  Copy, Check, HelpCircle, AlertTriangle, ChevronDown, User, Mail, Briefcase, Building,
  Eye, EyeOff
} from 'lucide-react'
import { getAccounts, createAccount, updateAccount, deleteAccount } from '../utils/authStore.js'
import { isSupabaseConfigured } from '../supabaseClient.js'

const formatDate = (dateStr) => {
  if (!dateStr) return '—'
  try {
    const d = new Date(dateStr)
    if (isNaN(d.getTime())) return '—'
    const day = String(d.getDate()).padStart(2, '0')
    const month = String(d.getMonth() + 1).padStart(2, '0')
    const year = d.getFullYear()
    return `${day}/${month}/${year}`
  } catch (e) {
    return '—'
  }
}

export default function QuanLyTaiKhoanTab() {
  const [accounts, setAccounts] = useState([])
  const [loading, setLoading] = useState(true)
  
  // Search and filter states
  const [searchQuery, setSearchQuery] = useState('')
  const [selectedRole, setSelectedRole] = useState('all')
  const [selectedTitle, setSelectedTitle] = useState('all')
  const [selectedDept, setSelectedDept] = useState('all')

  // Modals state
  const [showCreateModal, setShowCreateModal] = useState(false)
  const [showEditModal, setShowEditModal] = useState(false)
  const [showDeleteModal, setShowDeleteModal] = useState(false)
  const [selectedAccount, setSelectedAccount] = useState(null)
  
  // Form fields
  const [hoTen, setHoTen] = useState('')
  const [tenDangNhap, setTenDangNhap] = useState('')
  const [matKhau, setMatKhau] = useState('')
  const [email, setEmail] = useState('')
  const [phongBan, setPhongBan] = useState('Phòng Vật tư')
  const [quyen, setQuyen] = useState('User')
  const [chucDanh, setChucDanh] = useState('Chuyên viên')
  const [trangThai, setTrangThai] = useState('Hoạt động')
  
  const [formError, setFormError] = useState('')
  const [copiedSql, setCopiedSql] = useState(false)
  const [currentUser, setCurrentUser] = useState(null)
  const [showPasswords, setShowPasswords] = useState({})

  useEffect(() => {
    try {
      const saved = localStorage.getItem('sgc_logged_in_user')
      if (saved) {
        setCurrentUser(JSON.parse(saved))
      }
    } catch (e) {
      console.error(e)
    }
  }, [])

  const togglePasswordVisibility = (userId) => {
    setShowPasswords(prev => ({
      ...prev,
      [userId]: !prev[userId]
    }))
  }

  // Load accounts on mount
  const loadAccounts = async () => {
    setLoading(true)
    try {
      const data = await getAccounts()
      setAccounts(data)
    } catch (err) {
      console.error('Error loading accounts:', err)
    } finally {
      setLoading(false)
    }
  }

  useEffect(() => {
    loadAccounts()
  }, [])

  // Calculate statistics
  const totalAccounts = accounts.length
  const activeAccounts = accounts.filter(a => a.trang_thai === 'Hoạt động' || a.trang_thai === 'Đang hoạt động').length
  const adminAccounts = accounts.filter(a => a.quyen === 'Admin').length

  // Filter accounts based on queries
  const filteredAccounts = accounts.filter(acc => {
    const matchesSearch = 
      acc.ho_ten.toLowerCase().includes(searchQuery.toLowerCase()) ||
      acc.ten_dang_nhap.toLowerCase().includes(searchQuery.toLowerCase())

    const matchesRole = selectedRole === 'all' || acc.quyen === selectedRole

    return matchesSearch && matchesRole
  })

  // Handle open create modal
  const openCreateModal = () => {
    setHoTen('')
    setTenDangNhap('')
    setMatKhau('')
    setEmail('')
    setPhongBan('Phòng Vật tư')
    setQuyen('User')
    setChucDanh('Chuyên viên')
    setTrangThai('Hoạt động')
    setFormError('')
    setShowCreateModal(true)
  }

  // Handle open edit modal
  const openEditModal = (acc) => {
    setSelectedAccount(acc)
    setHoTen(acc.ho_ten)
    setTenDangNhap(acc.ten_dang_nhap)
    setMatKhau(acc.mat_khau || '')
    setEmail(acc.email === '—' ? '' : acc.email)
    setPhongBan(acc.phong_ban === '—' ? 'Phòng Vật tư' : acc.phong_ban)
    setQuyen(acc.quyen)
    setChucDanh(acc.chuc_danh === '—' ? 'Chuyên viên' : acc.chuc_danh)
    setTrangThai(acc.trang_thai)
    setFormError('')
    setShowEditModal(true)
  }

  // Handle open delete confirmation
  const openDeleteModal = (acc) => {
    if (acc.ten_dang_nhap === 'Docongchung') {
      alert('Tài khoản Admin mặc định của hệ thống không thể xóa!')
      return
    }
    setSelectedAccount(acc)
    setShowDeleteModal(true)
  }

  // Submit Create Account
  const handleCreateSubmit = async (e) => {
    e.preventDefault()
    setFormError('')

    if (!hoTen.trim() || !tenDangNhap.trim() || !matKhau.trim()) {
      setFormError('Vui lòng điền đầy đủ các thông tin bắt buộc!')
      return
    }

    const payload = {
      ho_ten: hoTen.trim(),
      ten_dang_nhap: tenDangNhap.trim(),
      mat_khau: matKhau,
      email: email.trim() || '—',
      phong_ban: phongBan,
      quyen: quyen,
      chuc_danh: chucDanh,
      trang_thai: trangThai
    }

    const result = await createAccount(payload)
    if (result.success) {
      setShowCreateModal(false)
      loadAccounts()
    } else {
      setFormError(result.error || 'Đã xảy ra lỗi khi tạo tài khoản!')
    }
  }

  // Submit Edit Account
  const handleEditSubmit = async (e) => {
    e.preventDefault()
    setFormError('')

    if (!hoTen.trim() || !tenDangNhap.trim() || !matKhau.trim()) {
      setFormError('Vui lòng điền đầy đủ các thông tin bắt buộc!')
      return
    }

    const payload = {
      ho_ten: hoTen.trim(),
      ten_dang_nhap: tenDangNhap.trim(),
      mat_khau: matKhau,
      email: email.trim() || '—',
      phong_ban: phongBan,
      quyen: quyen,
      chuc_danh: chucDanh,
      trang_thai: trangThai
    }

    const result = await updateAccount(selectedAccount.id, payload)
    if (result.success) {
      setShowEditModal(false)
      loadAccounts()
    } else {
      setFormError(result.error || 'Đã xảy ra lỗi khi cập nhật tài khoản!')
    }
  }

  // Confirm Delete Account
  const handleConfirmDelete = async () => {
    const result = await deleteAccount(selectedAccount.id, selectedAccount.ten_dang_nhap)
    if (result.success) {
      setShowDeleteModal(false)
      loadAccounts()
    } else {
      alert(result.error || 'Đã xảy ra lỗi khi xóa tài khoản!')
    }
  }

  // SQL Script text to create table in Supabase
  const sqlScript = `-- 1. THÊM CỘT trang_thai NẾU CHƯA CÓ TRÊN BẢNG quan_ly_tai_khoan
ALTER TABLE public.quan_ly_tai_khoan ADD COLUMN IF NOT EXISTS trang_thai TEXT DEFAULT 'Hoạt động';

-- 2. TẠO UNIQUE INDEX ĐỂ ĐẢM BẢO TÊN ĐĂNG NHẬP LÀ DUY NHẤT (SẠCH, KHÔNG GÂY LỖI TRÙNG LẶP VÀ ĐỂ SỬ DỤNG CHO ON CONFLICT)
CREATE UNIQUE INDEX IF NOT EXISTS quan_ly_tai_khoan_ten_dang_nhap_idx ON public.quan_ly_tai_khoan (ten_dang_nhap);

-- 3. BẬT ROW LEVEL SECURITY (RLS) ĐỂ CHO PHÉP TRUY VẤN
ALTER TABLE public.quan_ly_tai_khoan ENABLE ROW LEVEL SECURITY;

-- 4. TẠO CÁC CHÍNH SÁCH BẢO MẬT (POLICIES) CHO PHÉP WEB APP TRUY CẬP ĐẦY ĐỦ QUYỀN (SELECT, INSERT, UPDATE, DELETE)
DROP POLICY IF EXISTS "Allow public read" ON public.quan_ly_tai_khoan;
CREATE POLICY "Allow public read" ON public.quan_ly_tai_khoan FOR SELECT USING (true);

DROP POLICY IF EXISTS "Allow public insert" ON public.quan_ly_tai_khoan;
CREATE POLICY "Allow public insert" ON public.quan_ly_tai_khoan FOR INSERT WITH CHECK (true);

DROP POLICY IF EXISTS "Allow public update" ON public.quan_ly_tai_khoan;
CREATE POLICY "Allow public update" ON public.quan_ly_tai_khoan FOR UPDATE USING (true);

DROP POLICY IF EXISTS "Allow public delete" ON public.quan_ly_tai_khoan;
CREATE POLICY "Allow public delete" ON public.quan_ly_tai_khoan FOR DELETE USING (true);

-- 5. CHÈN TÀI KHOẢN ADMIN MẶC ĐỊNH HOẶC CẬP NHẬT NẾU ĐÃ TỒN TẠI
INSERT INTO public.quan_ly_tai_khoan (ho_ten, ten_dang_nhap, mat_khau, quyen, trang_thai)
VALUES (
  'Đỗ Công Chung', 
  'Docongchung', 
  'Chung10x7', 
  'Admin',
  'Hoạt động'
)
ON CONFLICT (ten_dang_nhap) DO UPDATE 
SET ho_ten = EXCLUDED.ho_ten, 
    mat_khau = EXCLUDED.mat_khau, 
    quyen = EXCLUDED.quyen, 
    trang_thai = EXCLUDED.trang_thai;`

  const handleCopySql = () => {
    navigator.clipboard.writeText(sqlScript)
    setCopiedSql(true)
    setTimeout(() => setCopiedSql(false), 2000)
  }

  // Get avatar background based on character code
  const getAvatarBg = (name) => {
    const firstChar = name ? name.charAt(0).toUpperCase() : 'U'
    const colors = [
      '#ef4444', '#f97316', '#f59e0b', '#10b981', '#06b6d4', 
      '#3b82f6', '#6366f1', '#8b5cf6', '#d946ef', '#ec4899'
    ]
    const idx = firstChar.charCodeAt(0) % colors.length
    return colors[idx]
  }

  // Phím Esc đóng hộp thoại đang mở (trên cùng trước)
  useEffect(() => {
    if (!showCreateModal && !showEditModal && !showDeleteModal) return
    const onKey = (e) => {
      if (e.key !== 'Escape') return
      e.preventDefault()
      if (showDeleteModal) setShowDeleteModal(false)
      else if (showEditModal) setShowEditModal(false)
      else if (showCreateModal) setShowCreateModal(false)
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [showCreateModal, showEditModal, showDeleteModal])

  const isActiveStatus = (s) => s === 'Hoạt động' || s === 'Đang hoạt động'

  // Màu cố định theo vai trò
  const roleTone = (role) => ROLE_TONES[role] || ROLE_TONES._other

  const closeBtn = (onClose) => (
    <button type="button" className="sgx-close" onClick={onClose} title="Đóng (Esc)">
      <IosX />
      <span>Đóng</span>
    </button>
  )

  return (
    <div className="sgx-page">
      <style>{QLTK_CSS}</style>

      {/* ─── HÀNG CÔNG CỤ: số liệu + tìm + lọc + nút ─── */}
      <div className="sgx-bar">
        <div className="sg-kpis" style={{ flexWrap: 'nowrap' }}>
          <div className="sg-kpi sm sg-c-blue" title="Tổng số tài khoản trong hệ thống">
            <span className="sg-kpi-ic"><Users size={16} /></span>
            <span className="sg-kpi-bd"><span className="sg-kpi-lb">Tài khoản</span><span className="sg-kpi-v">{totalAccounts}</span></span>
          </div>
          <div className="sg-kpi sm sg-c-green" title="Số tài khoản đang hoạt động">
            <span className="sg-kpi-ic"><UserCheck size={16} /></span>
            <span className="sg-kpi-bd"><span className="sg-kpi-lb">Đang hoạt động</span><span className="sg-kpi-v">{activeAccounts}</span></span>
          </div>
          <div className="sg-kpi sm sg-c-violet" title="Số tài khoản có quyền Admin">
            <span className="sg-kpi-ic"><Shield size={16} /></span>
            <span className="sg-kpi-bd"><span className="sg-kpi-lb">Admin</span><span className="sg-kpi-v">{adminAccounts}</span></span>
          </div>
        </div>

        <div className="sgx-search">
          <span className="sgx-search-ic"><Search size={16} /></span>
          <input
            id="acc_search_input"
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Tìm họ tên, tên đăng nhập..."
          />
          {searchQuery && (
            <button type="button" className="sgx-x" onClick={() => setSearchQuery('')} title="Xoá tìm kiếm" aria-label="Xoá tìm kiếm">
              <IosX />
            </button>
          )}
        </div>

        <select
          id="acc_role_filter"
          className="sgx-select"
          value={selectedRole}
          onChange={(e) => setSelectedRole(e.target.value)}
        >
          <option value="all">Tất cả vai trò</option>
          <option value="Admin">Vai trò: Admin</option>
          <option value="User">Vai trò: User</option>
        </select>

        <span className="sgx-note" title="Tài khoản Admin mặc định của hệ thống không thể xoá">
          <AlertTriangle size={14} />
          Admin mặc định không thể xoá
        </span>

        <span className="sgx-spacer" />

        <button id="acc_create_btn" type="button" className="sg-btn" onClick={openCreateModal}>
          <Plus size={16} />
          Tạo tài khoản
        </button>
      </div>

      {/* ─── BẢNG ─── */}
      <div className="sgx-tablebox">
        {loading ? (
          <div className="sgx-empty">
            <svg className="sgx-spin" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#0F58A7" strokeWidth="3">
              <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/>
            </svg>
            <span>Đang tải danh sách tài khoản...</span>
          </div>
        ) : filteredAccounts.length === 0 ? (
          accounts.length === 0 && isSupabaseConfigured ? (
            <div className="sgx-card sgx-warn">
              <div style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', textAlign: 'center', gap: '10px' }}>
                <AlertTriangle size={40} color="#E07A10" />
                <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: '#A85607' }}>
                  Chưa thiết lập bảng hoặc bị chặn bởi chính sách bảo mật RLS trên Supabase
                </h3>
                <p style={{ margin: '0 max(20px, 10%)', fontSize: '15px', color: '#4A4C55', lineHeight: '1.6' }}>
                  Mặc dù ứng dụng đã kết nối thành công tới Supabase, nhưng truy vấn trả về 0 dòng dữ liệu. Điều này xảy ra do chính sách <strong>Row Level Security (RLS)</strong> trên bảng <code style={{ background: '#FDF3D6', padding: '2px 6px', borderRadius: '6px', fontWeight: 600 }}>quan_ly_tai_khoan</code> đang chặn quyền đọc dữ liệu của công chúng (anonymous role), hoặc bảng chưa được tạo đúng cấu trúc.
                </p>
              </div>

              <div style={{ background: '#fff', borderRadius: '16px', padding: '16px', marginTop: '16px', boxShadow: 'inset 0 0 0 1px #D5DCEB' }}>
                <h4 style={{ margin: '0 0 10px 0', fontSize: '15px', fontWeight: 700, color: '#16171C' }}>Các bước khắc phục nhanh:</h4>
                <ol style={{ margin: 0, paddingLeft: '20px', fontSize: '15px', color: '#4A4C55', display: 'flex', flexDirection: 'column', gap: '8px', lineHeight: '1.6' }}>
                  <li>Nhấn nút <strong>Copy mã SQL cấu hình</strong> ở bên dưới.</li>
                  <li>Mở <a href="https://supabase.com/dashboard" target="_blank" rel="noreferrer" style={{ color: '#0F58A7', fontWeight: 600, textDecoration: 'underline' }}>Supabase Dashboard</a>, truy cập vào mục <strong>SQL Editor</strong> ở thanh menu bên trái.</li>
                  <li>Nhấn <strong>New Query</strong>, dán toàn bộ mã SQL vừa copy vào và nhấn nút <strong>RUN</strong>.</li>
                  <li>Sau khi thấy kết quả hiển thị thành công, vui lòng tải lại trang web này để hiển thị danh sách tài khoản.</li>
                </ol>

                <div style={{ display: 'flex', gap: '10px', marginTop: '16px', flexWrap: 'wrap' }}>
                  <button type="button" onClick={handleCopySql} className={copiedSql ? 'sg-btn green' : 'sg-btn'}>
                    {copiedSql ? <Check size={16} /> : <Copy size={16} />}
                    {copiedSql ? 'Đã copy mã SQL!' : 'Copy mã SQL cấu hình RLS'}
                  </button>
                  <a href="https://supabase.com/dashboard" target="_blank" rel="noreferrer" className="sg-btn ghost" style={{ textDecoration: 'none' }}>
                    <HelpCircle size={16} />
                    Mở Supabase Dashboard
                  </a>
                </div>
              </div>
            </div>
          ) : (
            <div className="sgx-empty">
              <Users size={44} color="#9AA0AD" />
              <h3 style={{ margin: 0, fontSize: '15px', fontWeight: 700, color: '#4A4C55' }}>Không tìm thấy tài khoản nào</h3>
              <p style={{ margin: 0, fontSize: '15px', color: '#8C8E97' }}>Hãy điều chỉnh từ khoá tìm kiếm hoặc tạo một tài khoản mới.</p>
            </div>
          )
        ) : (
          <div className="sgx-scroll">
            <table className="sgx-table">
              <thead>
                <tr>
                  <th style={{ width: '64px' }}>STT</th>
                  <th>Họ và tên</th>
                  <th style={{ width: '200px' }}>Tên đăng nhập</th>
                  <th style={{ width: '180px' }}>Mật khẩu</th>
                  <th style={{ width: '130px' }}>Vai trò</th>
                  <th style={{ width: '130px' }}>Ngày tạo</th>
                  <th style={{ width: '220px' }}>Trạng thái</th>
                  <th style={{ width: '110px' }}>Thao tác</th>
                </tr>
              </thead>
              <tbody>
                {filteredAccounts.map((acc, index) => {
                  const avatarLetter = acc.ho_ten ? acc.ho_ten.split(' ').pop().charAt(0).toUpperCase() : 'U'
                  const isSystemAdmin = acc.ten_dang_nhap === 'Docongchung'
                  const active = isActiveStatus(acc.trang_thai)
                  const tone = roleTone(acc.quyen)

                  return (
                    <tr
                      key={acc.id}
                      style={{ cursor: 'pointer' }}
                      title="Bấm đúp để sửa tài khoản"
                      onDoubleClick={() => openEditModal(acc)}
                    >
                      {/* 1. STT */}
                      <td style={{ textAlign: 'center', color: '#4A4C55' }}>{index + 1}</td>

                      {/* 2. HỌ VÀ TÊN */}
                      <td>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                          <div className="sgx-ava" style={{ backgroundColor: getAvatarBg(acc.ho_ten) }}>{avatarLetter}</div>
                          <span style={{ fontWeight: 600, color: '#16171C' }}>{acc.ho_ten}</span>
                        </div>
                      </td>

                      {/* 3. TÊN ĐĂNG NHẬP */}
                      <td style={{ textAlign: 'center', color: '#0A3D73', fontWeight: 600 }}>{acc.ten_dang_nhap}</td>

                      {/* 4. MẬT KHẨU */}
                      <td style={{ textAlign: 'center', color: '#4A4C55' }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '6px' }}>
                          <span>
                            {currentUser?.quyen === 'Admin' && showPasswords[acc.id]
                              ? acc.mat_khau
                              : '••••••'}
                          </span>
                          {currentUser?.quyen === 'Admin' && (
                            <button
                              type="button"
                              className="sgx-ib"
                              title={showPasswords[acc.id] ? 'Ẩn mật khẩu' : 'Hiện mật khẩu'}
                              onClick={(e) => {
                                e.stopPropagation() // Không kích hoạt bấm đúp của dòng
                                togglePasswordVisibility(acc.id)
                              }}
                            >
                              {showPasswords[acc.id] ? <EyeOff size={15} /> : <Eye size={15} />}
                            </button>
                          )}
                        </div>
                      </td>

                      {/* 5. VAI TRÒ */}
                      <td style={{ textAlign: 'center' }}>
                        <span className="sgx-role" style={{ background: tone.bg, color: tone.ink }}>
                          <i style={{ background: tone.dot }} />
                          {acc.quyen}
                        </span>
                      </td>

                      {/* 6. NGÀY TẠO */}
                      <td className="sg-num" style={{ textAlign: 'center', color: '#4A4C55' }}>{formatDate(acc.created_at)}</td>

                      {/* 7. TRẠNG THÁI (gạt) */}
                      <td style={{ textAlign: 'center' }}>
                        <div
                          style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '10px' }}
                          onClick={(e) => e.stopPropagation()}
                        >
                          <label
                            className={'sgx-tg' + (active ? ' on' : '') + (isSystemAdmin ? ' lock' : '')}
                            title={isSystemAdmin ? 'Tài khoản Admin mặc định luôn hoạt động' : (active ? 'Bấm để ngừng hoạt động' : 'Bấm để kích hoạt')}
                          >
                            <input
                              type="checkbox"
                              checked={active}
                              disabled={isSystemAdmin}
                              onChange={async () => {
                                if (isSystemAdmin) return
                                const currentActive = acc.trang_thai === 'Hoạt động' || acc.trang_thai === 'Đang hoạt động'
                                const newStatus = currentActive ? 'Ngừng hoạt động' : 'Đang hoạt động'

                                const res = await updateAccount(acc.id, {
                                  ho_ten: acc.ho_ten,
                                  ten_dang_nhap: acc.ten_dang_nhap,
                                  mat_khau: acc.mat_khau,
                                  quyen: acc.quyen,
                                  trang_thai: newStatus
                                })

                                if (res.success) {
                                  loadAccounts()
                                } else {
                                  alert(res.error || 'Không thể cập nhật trạng thái!')
                                }
                              }}
                            />
                            <span />
                          </label>
                          <span style={{ fontSize: '12px', fontWeight: 700, color: active ? '#0D7A3C' : '#D92D3A', minWidth: '112px', textAlign: 'left' }}>
                            {active ? 'Đang hoạt động' : 'Ngừng hoạt động'}
                          </span>
                        </div>
                      </td>

                      {/* 8. THAO TÁC */}
                      <td style={{ textAlign: 'center' }}>
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }} onDoubleClick={(e) => e.stopPropagation()}>
                          <button type="button" className="sgx-pen" title="Sửa tài khoản" aria-label="Sửa tài khoản" onClick={() => openEditModal(acc)}>
                            <PenIcon />
                          </button>
                          {isSystemAdmin ? (
                            <span className="sgx-x-ph" title="Tài khoản Admin mặc định không thể xoá" />
                          ) : (
                            <button type="button" className="sgx-x" title="Xoá tài khoản" aria-label="Xoá tài khoản" onClick={() => openDeleteModal(acc)}>
                              <IosX />
                            </button>
                          )}
                        </div>
                      </td>
                    </tr>
                  )
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* ─── HỘP THOẠI ─── */}

      {/* 1. TẠO TÀI KHOẢN */}
      <AnimatePresence>
        {showCreateModal && (
          <div className="sgx-ov">
            <motion.div
              className="sgx-dlg"
              initial={{ scale: 0.94, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.94, opacity: 0 }}
            >
              <div className="sgx-dlg-hd">
                <h3>Tạo tài khoản mới</h3>
                {closeBtn(() => setShowCreateModal(false))}
              </div>

              <form onSubmit={handleCreateSubmit} className="sgx-form">
                {formError && (
                  <div className="sgx-err">
                    <AlertTriangle size={16} />
                    <span>{formError}</span>
                  </div>
                )}

                <div className="sgx-field">
                  <label>Họ và tên <b>*</b></label>
                  <input
                    type="text"
                    required
                    value={hoTen}
                    onChange={(e) => setHoTen(e.target.value)}
                    placeholder="Nhập họ và tên (ví dụ: Nguyễn Văn A)..."
                  />
                </div>

                <div className="sgx-row2">
                  <div className="sgx-field">
                    <label>Tên đăng nhập <b>*</b></label>
                    <input
                      type="text"
                      required
                      value={tenDangNhap}
                      onChange={(e) => setTenDangNhap(e.target.value)}
                      placeholder="Docongchung..."
                    />
                  </div>
                  <div className="sgx-field">
                    <label>Mật khẩu <b>*</b></label>
                    <input
                      type="text"
                      required
                      value={matKhau}
                      onChange={(e) => setMatKhau(e.target.value)}
                      placeholder="Nhập mật khẩu..."
                    />
                  </div>
                </div>

                <div className="sgx-row2">
                  <div className="sgx-field">
                    <label>Vai trò</label>
                    <select value={quyen} onChange={(e) => setQuyen(e.target.value)} className="sgx-roleselect" style={{ backgroundColor: roleTone(quyen).bg, color: roleTone(quyen).ink }}>
                      <option value="User">User (Nhân viên)</option>
                      <option value="Admin">Admin (Quản trị viên)</option>
                    </select>
                  </div>
                  <div className="sgx-field">
                    <label>Trạng thái</label>
                    <select value={trangThai} onChange={(e) => setTrangThai(e.target.value)}>
                      <option value="Hoạt động">Đang hoạt động</option>
                      <option value="Ngừng hoạt động">Ngừng hoạt động</option>
                    </select>
                  </div>
                </div>

                <div className="sgx-ft">
                  <button type="button" className="sg-btn ghost" onClick={() => setShowCreateModal(false)}>Huỷ bỏ</button>
                  <button type="submit" className="sg-btn">Tạo tài khoản</button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 2. SỬA TÀI KHOẢN */}
      <AnimatePresence>
        {showEditModal && (
          <div className="sgx-ov">
            <motion.div
              className="sgx-dlg"
              initial={{ scale: 0.94, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.94, opacity: 0 }}
            >
              <div className="sgx-dlg-hd">
                <h3>Cập nhật tài khoản</h3>
                {closeBtn(() => setShowEditModal(false))}
              </div>

              <form onSubmit={handleEditSubmit} className="sgx-form">
                {formError && (
                  <div className="sgx-err">
                    <AlertTriangle size={16} />
                    <span>{formError}</span>
                  </div>
                )}

                <div className="sgx-field">
                  <label>Họ và tên <b>*</b></label>
                  <input
                    type="text"
                    required
                    value={hoTen}
                    onChange={(e) => setHoTen(e.target.value)}
                  />
                </div>

                <div className="sgx-row2">
                  <div className="sgx-field">
                    <label>Tên đăng nhập <b>*</b></label>
                    <input
                      type="text"
                      required
                      disabled={selectedAccount?.ten_dang_nhap === 'Docongchung'}
                      value={tenDangNhap}
                      onChange={(e) => setTenDangNhap(e.target.value)}
                    />
                  </div>
                  <div className="sgx-field">
                    <label>Mật khẩu <b>*</b></label>
                    <input
                      type="text"
                      required
                      value={matKhau}
                      onChange={(e) => setMatKhau(e.target.value)}
                    />
                  </div>
                </div>

                <div className="sgx-row2">
                  <div className="sgx-field">
                    <label>Vai trò</label>
                    <select
                      value={quyen}
                      disabled={selectedAccount?.ten_dang_nhap === 'Docongchung'}
                      onChange={(e) => setQuyen(e.target.value)}
                      className="sgx-roleselect"
                      style={{ backgroundColor: roleTone(quyen).bg, color: roleTone(quyen).ink }}
                    >
                      <option value="User">User (Nhân viên)</option>
                      <option value="Admin">Admin (Quản trị viên)</option>
                    </select>
                  </div>
                  <div className="sgx-field">
                    <label>Trạng thái</label>
                    <select
                      value={trangThai}
                      disabled={selectedAccount?.ten_dang_nhap === 'Docongchung'}
                      onChange={(e) => setTrangThai(e.target.value)}
                    >
                      <option value="Hoạt động">Đang hoạt động</option>
                      <option value="Ngừng hoạt động">Ngừng hoạt động</option>
                    </select>
                  </div>
                </div>

                <div className="sgx-ft">
                  {selectedAccount?.ten_dang_nhap !== 'Docongchung' && (
                    <button
                      type="button"
                      className="sg-btn danger"
                      style={{ marginRight: 'auto' }}
                      onClick={() => {
                        setShowEditModal(false)
                        openDeleteModal(selectedAccount)
                      }}
                    >
                      <Trash2 size={15} />
                      Xoá tài khoản
                    </button>
                  )}
                  <button type="button" className="sg-btn ghost" onClick={() => setShowEditModal(false)}>Huỷ bỏ</button>
                  <button type="submit" className="sg-btn">Cập nhật</button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 3. XÁC NHẬN XOÁ */}
      <AnimatePresence>
        {showDeleteModal && (
          <div className="sgx-ov">
            <motion.div
              className="sgx-dlg"
              style={{ maxWidth: '420px' }}
              initial={{ scale: 0.94, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.94, opacity: 0 }}
            >
              <div className="sgx-dlg-hd">
                <h3>Xác nhận xoá tài khoản</h3>
                {closeBtn(() => setShowDeleteModal(false))}
              </div>
              <div style={{ padding: '20px 24px 22px', textAlign: 'center' }}>
                <div style={{ width: '48px', height: '48px', borderRadius: '50%', background: '#FFE7E5', color: '#FF3B30', display: 'grid', placeItems: 'center', margin: '0 auto 14px' }}>
                  <AlertTriangle size={24} />
                </div>
                <p style={{ margin: '0 0 20px 0', fontSize: '15px', color: '#4A4C55', lineHeight: '1.5' }}>
                  Bạn có chắc chắn muốn xoá vĩnh viễn tài khoản của <strong>{selectedAccount?.ho_ten}</strong> (tên đăng nhập: <strong>{selectedAccount?.ten_dang_nhap}</strong>) khỏi hệ thống? Hành động này không thể hoàn tác.
                </p>
                <div style={{ display: 'flex', gap: '10px', justifyContent: 'center' }}>
                  <button type="button" className="sg-btn ghost" style={{ flex: 1 }} onClick={() => setShowDeleteModal(false)}>Bỏ qua</button>
                  <button type="button" className="sg-btn" style={{ flex: 1, background: 'linear-gradient(135deg, #D92D3A, #F2646E)', boxShadow: '0 8px 16px -10px rgba(217,45,58,.9)' }} onClick={handleConfirmDelete}>Xác nhận xoá</button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </div>
  )
}

// ─── Màu cố định theo vai trò ───
const ROLE_TONES = {
  Admin: { bg: '#ECE7FB', ink: '#4B2FA0', dot: '#6D4AFF' },
  User: { bg: '#E4EEFA', ink: '#0A3D73', dot: '#2F86E0' },
  _other: { bg: '#DDF4EC', ink: '#0B6B50', dot: '#0E8F80' }
}

// Dấu X kiểu iPhone (SVG, nét đều đầu tròn)
function IosX() {
  return (
    <svg width="11" height="11" viewBox="0 0 12 12" fill="none" aria-hidden="true">
      <path d="M2.5 2.5l7 7M9.5 2.5l-7 7" stroke="currentColor" strokeWidth="2.3" strokeLinecap="round" vectorEffect="non-scaling-stroke" />
    </svg>
  )
}

// Bút có gạch chân
function PenIcon() {
  return (
    <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
      <path d="M15.2 4.3l3.5 3.5L9 17.5l-4.3.9.9-4.3z" />
      <path d="M13.2 6.3l3.5 3.5" />
      <path d="M4 21.5h16" />
    </svg>
  )
}

const QLTK_CSS = `
.sgx-page { height: 100%; display: flex; flex-direction: column; gap: 12px; padding: 14px 20px 16px; box-sizing: border-box; overflow: hidden; text-align: left; }
.sgx-bar { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; flex: none; }
.sgx-spacer { flex: 1 1 0; }
.sgx-search { position: relative; flex: 0 1 300px; min-width: 200px; }
.sgx-search-ic { position: absolute; left: 13px; top: 50%; transform: translateY(-50%); color: #6B7AA0; display: flex; pointer-events: none; }
html body .sgx-search input { width: 100%; height: 36px; box-sizing: border-box; padding: 0 38px 0 38px; font-size: 15px; color: #16171C; background: #fff; border: 1.5px solid #D5DCEB; border-radius: 99px; outline: none; }
.sgx-search .sgx-x { position: absolute; right: 7px; top: 50%; transform: translateY(-50%); }
html body .sgx-search .sgx-x:active { transform: translateY(-50%) scale(.94); }
html body .sgx-select { height: 36px; box-sizing: border-box; border: 1.5px solid #D5DCEB; padding-left: 14px; font-size: 15px; font-weight: 600; color: #0A3D73; }
.sgx-note { display: inline-flex; align-items: center; gap: 6px; height: 30px; padding: 0 12px; border-radius: 99px; background: #FFE7E5; color: #B4232C; font-size: 12px; font-weight: 700; white-space: nowrap; }
.sgx-x { flex: none; width: 24px; height: 24px; padding: 0; border: 0; border-radius: 50%; display: inline-grid; place-items: center; background: #FFE7E5; color: #FF3B30; cursor: pointer; transition: background .15s, color .15s; }
.sgx-x:hover { background: #FF3B30; color: #fff; }
.sgx-x-ph { display: inline-block; width: 24px; height: 24px; }
.sgx-pen { flex: none; width: 30px; height: 30px; padding: 0; border: 0; border-radius: 8px; display: inline-grid; place-items: center; background: transparent; color: #0A2486; cursor: pointer; }
.sgx-pen:hover { background: #EEF1FA; }
.sgx-ib { width: 28px; height: 28px; padding: 0; border: 0; border-radius: 8px; display: inline-grid; place-items: center; background: transparent; color: #6B7AA0; cursor: pointer; }
.sgx-ib:hover { background: #E4EEFA; color: #0F58A7; }
.sgx-tablebox { flex: 1 1 auto; min-height: 0; display: flex; flex-direction: column; }
.sgx-scroll { flex: 1 1 auto; min-height: 0; overflow: auto; border-radius: 14px; }
.sgx-table { width: 100%; min-width: 1100px; background: #fff; }
.sgx-table > thead > tr > th { position: sticky; top: 0; z-index: 5; }
.sgx-ava { flex: none; width: 30px; height: 30px; border-radius: 50%; color: #fff; display: grid; place-items: center; font-weight: 700; font-size: 15px; }
.sgx-role { display: inline-flex; align-items: center; gap: 6px; height: 26px; padding: 0 12px; border-radius: 99px; font-size: 12px; font-weight: 700; white-space: nowrap; }
.sgx-role i { width: 7px; height: 7px; border-radius: 50%; }
.sgx-tg { position: relative; display: inline-block; width: 44px; height: 24px; cursor: pointer; flex: none; }
.sgx-tg input { position: absolute; opacity: 0; width: 0; height: 0; }
.sgx-tg > span { position: absolute; inset: 0; border-radius: 99px; background: #C9CDD8; transition: background .2s; }
.sgx-tg > span::after { content: ""; position: absolute; top: 3px; left: 3px; width: 18px; height: 18px; border-radius: 50%; background: #fff; box-shadow: 0 1px 3px rgba(0,0,0,.25); transition: left .2s; }
.sgx-tg.on > span { background: linear-gradient(135deg, #12A150, #3CCB7F); }
.sgx-tg.on > span::after { left: 23px; }
.sgx-tg.lock { cursor: not-allowed; opacity: .6; }
.sgx-empty { flex: 1; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 10px; padding: 48px 24px; text-align: center; color: #4A4C55; background: rgba(255,255,255,.72); border-radius: 24px; }
.sgx-spin { animation: sgxSpin 1s linear infinite; }
@keyframes sgxSpin { to { transform: rotate(360deg); } }
.sgx-card { background: rgba(255,255,255,.78); border-radius: 24px; padding: 24px; overflow: auto; }
.sgx-warn { box-shadow: inset 0 0 0 1px #F5D9A8; }
.sgx-ov { position: fixed; inset: 0; z-index: 10000; background: rgba(22, 23, 28, .45); backdrop-filter: blur(4px); display: flex; align-items: center; justify-content: center; padding: 20px; }
.sgx-dlg { width: 100%; max-width: 540px; background: #fff; border-radius: 24px; box-shadow: 0 30px 60px -20px rgba(10, 36, 80, .45); overflow: hidden; color: #16171C; }
.sgx-dlg-hd { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 16px 16px 12px 24px; }
.sgx-dlg-hd h3 { margin: 0; font-size: 22px; font-weight: 700; color: #0A3D73; }
.sgx-close { display: inline-flex; align-items: center; gap: 6px; height: 32px; padding: 0 12px 0 6px; border: 0; border-radius: 99px; background: #FFE7E5; color: #FF3B30; font-size: 15px; font-weight: 600; cursor: pointer; }
.sgx-close svg { width: 20px; height: 20px; padding: 4.5px; box-sizing: border-box; border-radius: 50%; background: rgba(255,255,255,.7); }
.sgx-close:hover { background: #FF3B30; color: #fff; }
.sgx-close:hover svg { background: rgba(255,255,255,.2); }
.sgx-form { padding: 6px 24px 22px; display: flex; flex-direction: column; gap: 14px; }
.sgx-row2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.sgx-field { display: flex; flex-direction: column; gap: 6px; min-width: 0; }
.sgx-field > label { font-size: 15px; font-weight: 600; color: #0A3D73; }
.sgx-field > label b { color: #FF3B30; }
html body .sgx-field input, html body .sgx-field select { height: 40px; box-sizing: border-box; padding: 0 12px; border: 1.5px solid #D5DCEB; border-radius: 12px; font-size: 15px; color: #16171C; outline: none; width: 100%; }
html body .sgx-field input:disabled, html body .sgx-field select:disabled { background-color: #F5F6F8; color: #8C8E97; cursor: not-allowed; }
html body .sgx-field select.sgx-roleselect { font-weight: 700; border-radius: 99px; padding-left: 16px; }
.sgx-err { display: flex; align-items: center; gap: 8px; padding: 10px 14px; border-radius: 14px; background: #FFE7E5; color: #B4232C; font-size: 15px; font-weight: 600; }
.sgx-ft { display: flex; justify-content: flex-end; gap: 10px; margin-top: 6px; padding-top: 14px; border-top: 1px solid #E6E9F0; }
@media (max-width: 760px) {
  .sgx-page { padding: 10px 12px; }
  .sgx-row2 { grid-template-columns: 1fr; }
  .sgx-note { display: none; }
}
`
