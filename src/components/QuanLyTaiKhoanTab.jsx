import React, { useState, useEffect } from 'react'
import { motion, AnimatePresence } from 'motion/react'
import { 
  Search, Users, UserCheck, Shield, Plus, Pencil, Trash2, X, Lock, 
  Copy, Check, HelpCircle, AlertTriangle, ChevronDown, User, Mail, Briefcase, Building,
  Eye, EyeOff
} from 'lucide-react'
import { getAccounts, createAccount, updateAccount, deleteAccount } from '../utils/authStore.js'

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
  const sqlScript = `-- 1. TẠO BẢNG LƯU TRỮ TÀI KHOẢN TRÊN SUPABASE SQL EDITOR
CREATE TABLE IF NOT EXISTS quan_ly_tai_khoan (
  id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
  ho_ten TEXT NOT NULL,
  ten_dang_nhap TEXT NOT NULL UNIQUE,
  mat_khau TEXT NOT NULL,
  quyen TEXT NOT NULL DEFAULT 'User', -- 'Admin' hoặc 'User'
  created_at TIMESTAMPTZ DEFAULT NOW()
);

-- 2. CHÈN TÀI KHOẢN ADMIN MẶC ĐỊNH
INSERT INTO quan_ly_tai_khoan (ho_ten, ten_dang_nhap, mat_khau, quyen)
VALUES (
  'Đỗ Công Chung', 
  'Docongchung', 
  'Chung10x7', 
  'Admin'
)
ON CONFLICT (ten_dang_nhap) DO NOTHING;`

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

  return (
    <div style={{
      height: '100%',
      display: 'flex',
      flexDirection: 'column',
      background: '#f8fafc',
      overflow: 'hidden',
      fontFamily: 'var(--font-sans)',
      textAlign: 'left'
    }}>
      
      {/* ─── SCROLLABLE CONTENT BODY ─── */}
      <div style={{
        flex: 1,
        overflowY: 'auto',
        padding: '20px 24px',
        display: 'flex',
        flexDirection: 'column',
        gap: '20px'
      }}>

        {/* ─── FILTERS & CONTROLS ─── */}
        <div style={{
          background: '#ffffff',
          borderRadius: '16px',
          border: '1px solid #e2e8f0',
          padding: '16px',
          display: 'flex',
          flexWrap: 'wrap',
          gap: '12px',
          alignItems: 'center',
          boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.05)',
          justifyContent: 'space-between'
        }}>
          {/* Left inputs */}
          <div style={{ display: 'flex', flexWrap: 'wrap', gap: '12px', flex: 1, minWidth: '320px' }}>
            {/* Search Input */}
            <div style={{ position: 'relative', width: '260px' }}>
              <span style={{ position: 'absolute', left: '12px', top: '50%', transform: 'translateY(-50%)', color: '#64748b', display: 'flex', alignItems: 'center' }}>
                <Search size={16} />
              </span>
              <input
                id="acc_search_input"
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Tìm tên, username, email..."
                style={{
                  width: '100%',
                  padding: '9px 12px 9px 36px',
                  border: '1px solid #cbd5e1',
                  borderRadius: '10px',
                  fontSize: '13.5px',
                  outline: 'none',
                  color: '#1e293b',
                  boxSizing: 'border-box'
                }}
              />
            </div>

            {/* Quyền Dropdown */}
            <select
              id="acc_role_filter"
              value={selectedRole}
              onChange={(e) => setSelectedRole(e.target.value)}
              style={{
                padding: '9px 12px',
                border: '1px solid #cbd5e1',
                borderRadius: '10px',
                fontSize: '13px',
                fontWeight: 600,
                color: '#475569',
                background: '#ffffff',
                outline: 'none',
                cursor: 'pointer'
              }}
            >
              <option value="all">Tất cả quyền</option>
              <option value="Admin">Quyền: Admin</option>
              <option value="User">Quyền: User</option>
            </select>
          </div>

          {/* Right Action Button */}
          <button
            id="acc_create_btn"
            onClick={openCreateModal}
            style={{
              background: '#2563eb',
              color: '#ffffff',
              border: 'none',
              borderRadius: '10px',
              padding: '10px 18px',
              fontSize: '13.5px',
              fontWeight: 700,
              cursor: 'pointer',
              display: 'flex',
              alignItems: 'center',
              gap: '6px',
              boxShadow: '0 4px 10px rgba(37, 99, 235, 0.15)',
              transition: 'all 0.15s ease'
            }}
            onMouseOver={(e) => {
              e.currentTarget.style.background = '#1d4ed8'
              e.currentTarget.style.transform = 'translateY(-1px)'
            }}
            onMouseOut={(e) => {
              e.currentTarget.style.background = '#2563eb'
              e.currentTarget.style.transform = 'none'
            }}
          >
            <Plus size={16} />
            Tạo tài khoản
          </button>
        </div>

        {/* ─── DATA TABLE ─── */}
        <div style={{
          background: '#ffffff',
          borderRadius: '16px',
          border: '1px solid #e2e8f0',
          boxShadow: '0 4px 6px -1px rgba(0, 0, 0, 0.05)',
          overflow: 'hidden'
        }}>
          {loading ? (
            <div style={{ padding: '60px', textAlign: 'center', color: '#64748b' }}>
              <svg className="animate-spin" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="#2563eb" strokeWidth="3" style={{ margin: '0 auto 12px auto', animation: 'spin 1s linear infinite' }}>
                <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/>
              </svg>
              <span>Đang tải danh sách tài khoản...</span>
            </div>
          ) : filteredAccounts.length === 0 ? (
            <div style={{ padding: '80px 24px', textAlign: 'center', color: '#64748b' }}>
              <Users size={48} color="#94a3b8" style={{ margin: '0 auto 16px auto' }} />
              <h3 style={{ margin: '0 0 6px 0', fontSize: '15px', fontWeight: 700, color: '#475569' }}>Không tìm thấy tài khoản nào</h3>
              <p style={{ margin: 0, fontSize: '13px', color: '#94a3b8' }}>Hãy điều chỉnh từ khóa tìm kiếm hoặc tạo một tài khoản mới.</p>
            </div>
          ) : (
            <div style={{ overflowX: 'auto' }}>
              <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13.5px' }}>
                <thead>
                  <tr style={{ background: '#f8fafc', borderBottom: '1px solid #e2e8f0' }}>
                    <th style={{ padding: '14px 16px', fontWeight: 700, color: '#475569', textAlign: 'center', width: '80px' }}>STT</th>
                    <th style={{ padding: '14px 16px', fontWeight: 700, color: '#475569', textAlign: 'center' }}>Họ & tên</th>
                    <th style={{ padding: '14px 16px', fontWeight: 700, color: '#475569', textAlign: 'center' }}>Tên đăng nhập</th>
                    <th style={{ padding: '14px 16px', fontWeight: 700, color: '#475569', textAlign: 'center' }}>Mật khẩu</th>
                    <th style={{ padding: '14px 16px', fontWeight: 700, color: '#475569', textAlign: 'center' }}>Vai trò</th>
                    <th style={{ padding: '14px 16px', fontWeight: 700, color: '#475569', textAlign: 'center' }}>Ngày tạo</th>
                    <th style={{ padding: '14px 16px', fontWeight: 700, color: '#475569', textAlign: 'center', width: '180px' }}>Trạng thái</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredAccounts.map((acc, index) => {
                    const avatarLetter = acc.ho_ten ? acc.ho_ten.split(' ').pop().charAt(0).toUpperCase() : 'U'
                    const isSystemAdmin = acc.ten_dang_nhap === 'Docongchung'

                    return (
                      <tr 
                        key={acc.id} 
                        style={{ 
                          borderBottom: '1px solid #f1f5f9',
                          transition: 'background 0.15s ease',
                          cursor: 'pointer'
                        }}
                        title="Click đúp chuột để chỉnh sửa và xóa tài khoản"
                        onMouseOver={(e) => { e.currentTarget.style.background = '#f8fafc' }}
                        onMouseOut={(e) => { e.currentTarget.style.background = 'transparent' }}
                        onDoubleClick={() => openEditModal(acc)}
                      >
                        {/* 1. STT */}
                        <td style={{ padding: '12px 16px', textAlign: 'center', color: '#64748b', fontWeight: 600 }}>
                          {index + 1}
                        </td>

                        {/* 2. HỌ & TÊN WITH AVATAR */}
                        <td style={{ padding: '12px 16px' }}>
                          <div style={{ display: 'flex', alignItems: 'center', gap: '10px' }}>
                            <div style={{
                              width: '32px',
                              height: '32px',
                              borderRadius: '50%',
                              backgroundColor: getAvatarBg(acc.ho_ten),
                              color: '#ffffff',
                              display: 'flex',
                              alignItems: 'center',
                              justifyContent: 'center',
                              fontWeight: 700,
                              fontSize: '13px',
                              boxShadow: '0 2px 4px rgba(0,0,0,0.05)'
                            }}>
                              {avatarLetter}
                            </div>
                            <span style={{ fontWeight: 700, color: '#0f172a' }}>{acc.ho_ten}</span>
                          </div>
                        </td>

                        {/* 3. TÊN ĐĂNG NHẬP */}
                        <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                          <span style={{
                            background: '#eff6ff',
                            color: '#2563eb',
                            padding: '3px 10px',
                            borderRadius: '20px',
                            fontSize: '12px',
                            fontWeight: 700,
                            border: '1px solid #dbeafe'
                          }}>
                            {acc.ten_dang_nhap}
                          </span>
                        </td>

                        {/* 4. MẬT KHẨU */}
                        <td style={{ padding: '12px 16px', fontFamily: 'var(--font-mono)', color: '#475569', fontWeight: 500, textAlign: 'center' }}>
                          <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}>
                            <span>
                              {currentUser?.quyen === 'Admin' && showPasswords[acc.id]
                                ? acc.mat_khau
                                : '••••••'}
                            </span>
                            {currentUser?.quyen === 'Admin' && (
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation() // Prevent tr doubleClick trigger
                                  togglePasswordVisibility(acc.id)
                                }}
                                style={{
                                  background: 'none',
                                  border: 'none',
                                  cursor: 'pointer',
                                  color: '#64748b',
                                  padding: '4px',
                                  borderRadius: '4px',
                                  display: 'flex',
                                  alignItems: 'center',
                                  justifyContent: 'center'
                                }}
                                onMouseOver={(e) => { e.currentTarget.style.backgroundColor = '#f1f5f9' }}
                                onMouseOut={(e) => { e.currentTarget.style.backgroundColor = 'transparent' }}
                              >
                                {showPasswords[acc.id] ? <EyeOff size={14} /> : <Eye size={14} />}
                              </button>
                            )}
                          </div>
                        </td>

                        {/* 5. VAI TRÒ */}
                        <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                          {acc.quyen === 'Admin' ? (
                            <span style={{ background: '#fce7f3', color: '#be185d', padding: '4px 10px', borderRadius: '12px', fontSize: '11px', fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
                              👑 Admin
                            </span>
                          ) : (
                            <span style={{ background: '#eff6ff', color: '#1d4ed8', padding: '4px 10px', borderRadius: '12px', fontSize: '11px', fontWeight: 700, display: 'inline-flex', alignItems: 'center', gap: '4px' }}>
                              👤 User
                            </span>
                          )}
                        </td>

                        {/* 6. NGÀY TẠO */}
                        <td style={{ padding: '12px 16px', color: '#475569', fontWeight: 500, textAlign: 'center' }}>
                          {formatDate(acc.created_at)}
                        </td>

                        {/* 7. TRẠNG THÁI (Gạt toggle) */}
                        <td style={{ padding: '12px 16px', textAlign: 'center' }}>
                          <div 
                            style={{ display: 'flex', alignItems: 'center', justifyContent: 'center', gap: '8px' }}
                            onClick={(e) => e.stopPropagation()} // Prevent doubleClick edit modal
                          >
                            <label style={{
                              position: 'relative',
                              display: 'inline-block',
                              width: '44px',
                              height: '22px',
                              cursor: isSystemAdmin ? 'not-allowed' : 'pointer',
                              opacity: isSystemAdmin ? 0.6 : 1
                            }}>
                              <input 
                                type="checkbox" 
                                checked={acc.trang_thai === 'Hoạt động' || acc.trang_thai === 'Đang hoạt động'}
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
                                style={{ opacity: 0, width: 0, height: 0 }}
                              />
                              <span style={{
                                position: 'absolute',
                                cursor: isSystemAdmin ? 'not-allowed' : 'pointer',
                                top: 0, left: 0, right: 0, bottom: 0,
                                backgroundColor: (acc.trang_thai === 'Hoạt động' || acc.trang_thai === 'Đang hoạt động') ? '#22c55e' : '#cbd5e1',
                                transition: '0.2s',
                                borderRadius: '34px'
                              }}>
                                <span style={{
                                  position: 'absolute',
                                  content: '""',
                                  height: '16px',
                                  width: '16px',
                                  left: (acc.trang_thai === 'Hoạt động' || acc.trang_thai === 'Đang hoạt động') ? '24px' : '4px',
                                  bottom: '3px',
                                  backgroundColor: 'white',
                                  transition: '0.2s',
                                  borderRadius: '50%',
                                  boxShadow: '0 1px 3px rgba(0,0,0,0.2)'
                                }} />
                              </span>
                            </label>
                            <span style={{ 
                              fontSize: '11px', 
                              fontWeight: 700, 
                              color: (acc.trang_thai === 'Hoạt động' || acc.trang_thai === 'Đang hoạt động') ? '#16a34a' : '#ef4444',
                              minWidth: '100px',
                              textAlign: 'left'
                            }}>
                              {(acc.trang_thai === 'Hoạt động' || acc.trang_thai === 'Đang hoạt động') ? 'Đang hoạt động' : 'Ngừng hoạt động'}
                            </span>
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

        {/* Bottom Warning Message matching bottom right text of image */}
        <div style={{
          display: 'flex',
          justifyContent: 'flex-end',
          alignItems: 'center',
          gap: '8px',
          color: '#e11d48',
          fontSize: '12px',
          fontWeight: 700,
          marginTop: '4px'
        }}>
          <AlertTriangle size={14} />
          <span>Tài khoản Admin hệ thống mặc định không thể xóa</span>
        </div>

      </div>

      {/* ─── MODALS ─── */}

      {/* 1. CREATE MODAL */}
      <AnimatePresence>
        {showCreateModal && (
          <div style={{
            position: 'fixed',
            inset: 0,
            zIndex: 10000,
            background: 'rgba(15, 23, 42, 0.6)',
            backdropFilter: 'blur(4px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '20px'
          }}>
            <motion.div
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              style={{
                width: '100%',
                maxWidth: '520px',
                background: '#ffffff',
                borderRadius: '16px',
                boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)',
                overflow: 'hidden',
                color: '#1e293b'
              }}
            >
              {/* Header */}
              <div style={{
                background: '#f8fafc',
                padding: '16px 24px',
                borderBottom: '1px solid #e2e8f0',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between'
              }}>
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 800, color: '#0f172a' }}>TẠO TÀI KHOẢN MỚI</h3>
                <button 
                  onClick={() => setShowCreateModal(false)}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#64748b', padding: '4px', borderRadius: '50%' }}
                >
                  <X size={18} />
                </button>
              </div>

              {/* Form Content */}
              <form onSubmit={handleCreateSubmit} style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                {formError && (
                  <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', padding: '10px 14px', borderRadius: '8px', color: '#991b1b', fontSize: '13px', display: 'flex', gap: '8px', alignItems: 'center' }}>
                    <AlertTriangle size={16} />
                    <span>{formError}</span>
                  </div>
                )}

                {/* Họ tên */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  <label style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569' }}>Họ và tên <span style={{ color: '#ef4444' }}>*</span></label>
                  <input
                    type="text"
                    required
                    value={hoTen}
                    onChange={(e) => setHoTen(e.target.value)}
                    placeholder="Nhập họ và tên (Ví dụ: Nguyễn Văn A)..."
                    style={{ padding: '10px 12px', border: '1px solid #cbd5e1', borderRadius: '8px', fontSize: '14px', outline: 'none' }}
                  />
                </div>

                {/* Username & Pass Row */}
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                    <label style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569' }}>Tên đăng nhập <span style={{ color: '#ef4444' }}>*</span></label>
                    <input
                      type="text"
                      required
                      value={tenDangNhap}
                      onChange={(e) => setTenDangNhap(e.target.value)}
                      placeholder="Docongchung..."
                      style={{ padding: '10px 12px', border: '1px solid #cbd5e1', borderRadius: '8px', fontSize: '14px', outline: 'none' }}
                    />
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                    <label style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569' }}>Mật khẩu <span style={{ color: '#ef4444' }}>*</span></label>
                    <input
                      type="text"
                      required
                      value={matKhau}
                      onChange={(e) => setMatKhau(e.target.value)}
                      placeholder="Nhập mật khẩu..."
                      style={{ padding: '10px 12px', border: '1px solid #cbd5e1', borderRadius: '8px', fontSize: '14px', outline: 'none' }}
                    />
                  </div>
                </div>

                {/* Quyền hạn & Trạng thái Row */}
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                    <label style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569' }}>Quyền hạn</label>
                    <select
                      value={quyen}
                      onChange={(e) => setQuyen(e.target.value)}
                      style={{ padding: '10px 12px', border: '1px solid #cbd5e1', borderRadius: '8px', fontSize: '14px', background: '#ffffff', outline: 'none' }}
                    >
                      <option value="User">User (Nhân viên)</option>
                      <option value="Admin">Admin (Quản trị viên)</option>
                    </select>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                    <label style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569' }}>Trạng thái</label>
                    <select
                      value={trangThai}
                      onChange={(e) => setTrangThai(e.target.value)}
                      style={{ padding: '10px 12px', border: '1px solid #cbd5e1', borderRadius: '8px', fontSize: '14px', background: '#ffffff', outline: 'none' }}
                    >
                      <option value="Hoạt động">Đang hoạt động</option>
                      <option value="Ngừng hoạt động">Ngừng hoạt động</option>
                    </select>
                  </div>
                </div>

                {/* Footer Buttons */}
                <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px', marginTop: '12px', borderTop: '1px solid #e2e8f0', paddingTop: '16px' }}>
                  <button
                    type="button"
                    onClick={() => setShowCreateModal(false)}
                    style={{ padding: '10px 18px', background: '#f1f5f9', border: 'none', borderRadius: '8px', color: '#475569', fontWeight: 600, cursor: 'pointer' }}
                  >
                    Hủy bỏ
                  </button>
                  <button
                    type="submit"
                    style={{ padding: '10px 18px', background: '#2563eb', border: 'none', borderRadius: '8px', color: '#ffffff', fontWeight: 700, cursor: 'pointer' }}
                  >
                    Tạo tài khoản
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 2. EDIT MODAL */}
      <AnimatePresence>
        {showEditModal && (
          <div style={{
            position: 'fixed',
            inset: 0,
            zIndex: 10000,
            background: 'rgba(15, 23, 42, 0.6)',
            backdropFilter: 'blur(4px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '20px'
          }}>
            <motion.div
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              style={{
                width: '100%',
                maxWidth: '520px',
                background: '#ffffff',
                borderRadius: '16px',
                boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)',
                overflow: 'hidden',
                color: '#1e293b'
              }}
            >
              {/* Header */}
              <div style={{
                background: '#f8fafc',
                padding: '16px 24px',
                borderBottom: '1px solid #e2e8f0',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between'
              }}>
                <h3 style={{ margin: 0, fontSize: '16px', fontWeight: 800, color: '#0f172a' }}>CẬP NHẬT TÀI KHOẢN</h3>
                <button 
                  onClick={() => setShowEditModal(false)}
                  style={{ background: 'none', border: 'none', cursor: 'pointer', color: '#64748b', padding: '4px', borderRadius: '50%' }}
                >
                  <X size={18} />
                </button>
              </div>

              {/* Form Content */}
              <form onSubmit={handleEditSubmit} style={{ padding: '24px', display: 'flex', flexDirection: 'column', gap: '16px' }}>
                {formError && (
                  <div style={{ background: '#fef2f2', border: '1px solid #fca5a5', padding: '10px 14px', borderRadius: '8px', color: '#991b1b', fontSize: '13px', display: 'flex', gap: '8px', alignItems: 'center' }}>
                    <AlertTriangle size={16} />
                    <span>{formError}</span>
                  </div>
                )}

                {/* Họ tên */}
                <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                  <label style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569' }}>Họ và tên <span style={{ color: '#ef4444' }}>*</span></label>
                  <input
                    type="text"
                    required
                    value={hoTen}
                    onChange={(e) => setHoTen(e.target.value)}
                    style={{ padding: '10px 12px', border: '1px solid #cbd5e1', borderRadius: '8px', fontSize: '14px', outline: 'none' }}
                  />
                </div>

                {/* Username & Pass Row */}
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                    <label style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569' }}>Tên đăng nhập <span style={{ color: '#ef4444' }}>*</span></label>
                    <input
                      type="text"
                      required
                      disabled={selectedAccount?.ten_dang_nhap === 'Docongchung'}
                      value={tenDangNhap}
                      onChange={(e) => setTenDangNhap(e.target.value)}
                      style={{ padding: '10px 12px', border: '1px solid #cbd5e1', borderRadius: '8px', fontSize: '14px', outline: 'none', background: selectedAccount?.ten_dang_nhap === 'Docongchung' ? '#f1f5f9' : '#ffffff' }}
                    />
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                    <label style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569' }}>Mật khẩu <span style={{ color: '#ef4444' }}>*</span></label>
                    <input
                      type="text"
                      required
                      value={matKhau}
                      onChange={(e) => setMatKhau(e.target.value)}
                      style={{ padding: '10px 12px', border: '1px solid #cbd5e1', borderRadius: '8px', fontSize: '14px', outline: 'none' }}
                    />
                  </div>
                </div>

                {/* Quyền hạn & Trạng thái Row */}
                <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '12px' }}>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                    <label style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569' }}>Quyền hạn</label>
                    <select
                      value={quyen}
                      disabled={selectedAccount?.ten_dang_nhap === 'Docongchung'}
                      onChange={(e) => setQuyen(e.target.value)}
                      style={{ padding: '10px 12px', border: '1px solid #cbd5e1', borderRadius: '8px', fontSize: '14px', background: selectedAccount?.ten_dang_nhap === 'Docongchung' ? '#f1f5f9' : '#ffffff', outline: 'none' }}
                    >
                      <option value="User">User (Nhân viên)</option>
                      <option value="Admin">Admin (Quản trị viên)</option>
                    </select>
                  </div>
                  <div style={{ display: 'flex', flexDirection: 'column', gap: '6px' }}>
                    <label style={{ fontSize: '12.5px', fontWeight: 700, color: '#475569' }}>Trạng thái</label>
                    <select
                      value={trangThai}
                      disabled={selectedAccount?.ten_dang_nhap === 'Docongchung'}
                      onChange={(e) => setTrangThai(e.target.value)}
                      style={{ padding: '10px 12px', border: '1px solid #cbd5e1', borderRadius: '8px', fontSize: '14px', background: selectedAccount?.ten_dang_nhap === 'Docongchung' ? '#f1f5f9' : '#ffffff', outline: 'none' }}
                    >
                      <option value="Hoạt động">Đang hoạt động</option>
                      <option value="Ngừng hoạt động">Ngừng hoạt động</option>
                    </select>
                  </div>
                </div>

                {/* Footer Buttons */}
                <div style={{ display: 'flex', justifyContent: 'flex-end', gap: '10px', marginTop: '12px', borderTop: '1px solid #e2e8f0', paddingTop: '16px' }}>
                  {selectedAccount?.ten_dang_nhap !== 'Docongchung' && (
                    <button
                      type="button"
                      onClick={() => {
                        setShowEditModal(false)
                        openDeleteModal(selectedAccount)
                      }}
                      style={{ 
                        padding: '10px 18px', 
                        background: '#fee2e2', 
                        border: '1px solid #fecdd3', 
                        borderRadius: '8px', 
                        color: '#ef4444', 
                        fontWeight: 600, 
                        cursor: 'pointer',
                        marginRight: 'auto',
                        display: 'flex',
                        alignItems: 'center',
                        gap: '4px'
                      }}
                      onMouseOver={(e) => { e.currentTarget.style.background = '#fecdd3' }}
                      onMouseOut={(e) => { e.currentTarget.style.background = '#fee2e2' }}
                    >
                      <Trash2 size={14} />
                      Xóa tài khoản
                    </button>
                  )}
                  <button
                    type="button"
                    onClick={() => setShowEditModal(false)}
                    style={{ padding: '10px 18px', background: '#f1f5f9', border: 'none', borderRadius: '8px', color: '#475569', fontWeight: 600, cursor: 'pointer' }}
                  >
                    Hủy bỏ
                  </button>
                  <button
                    type="submit"
                    style={{ padding: '10px 18px', background: '#2563eb', border: 'none', borderRadius: '8px', color: '#ffffff', fontWeight: 700, cursor: 'pointer' }}
                  >
                    Cập nhật
                  </button>
                </div>
              </form>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* 3. DELETE CONFIRMATION MODAL */}
      <AnimatePresence>
        {showDeleteModal && (
          <div style={{
            position: 'fixed',
            inset: 0,
            zIndex: 10000,
            background: 'rgba(15, 23, 42, 0.6)',
            backdropFilter: 'blur(4px)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            padding: '20px'
          }}>
            <motion.div
              initial={{ scale: 0.9, opacity: 0 }}
              animate={{ scale: 1, opacity: 1 }}
              exit={{ scale: 0.9, opacity: 0 }}
              style={{
                width: '100%',
                maxWidth: '400px',
                background: '#ffffff',
                borderRadius: '16px',
                boxShadow: '0 25px 50px -12px rgba(0, 0, 0, 0.25)',
                padding: '24px',
                color: '#1e293b'
              }}
            >
              <div style={{ textAlign: 'center' }}>
                <div style={{
                  width: '48px',
                  height: '48px',
                  borderRadius: '50%',
                  background: '#fee2e2',
                  color: '#ef4444',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  margin: '0 auto 16px auto'
                }}>
                  <AlertTriangle size={24} />
                </div>
                <h3 style={{ margin: '0 0 8px 0', fontSize: '16px', fontWeight: 800, color: '#0f172a' }}>XÁC NHẬN XÓA TÀI KHOẢN</h3>
                <p style={{ margin: '0 0 24px 0', fontSize: '13.5px', color: '#475569', lineHeight: '1.5' }}>
                  Bạn có chắc chắn muốn xóa vĩnh viễn tài khoản của <strong>{selectedAccount?.ho_ten}</strong> (Tên đăng nhập: <strong>{selectedAccount?.ten_dang_nhap}</strong>) khỏi hệ thống? Hành động này không thể hoàn tác!
                </p>
                <div style={{ display: 'flex', gap: '10px', justifyContent: 'center' }}>
                  <button
                    onClick={() => setShowDeleteModal(false)}
                    style={{ padding: '10px 18px', background: '#f1f5f9', border: 'none', borderRadius: '8px', color: '#475569', fontWeight: 600, cursor: 'pointer', flex: 1 }}
                  >
                    Bỏ qua
                  </button>
                  <button
                    onClick={handleConfirmDelete}
                    style={{ padding: '10px 18px', background: '#ef4444', border: 'none', borderRadius: '8px', color: '#ffffff', fontWeight: 700, cursor: 'pointer', flex: 1 }}
                  >
                    Xác nhận xóa
                  </button>
                </div>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

    </div>
  )
}
