import { supabase, isSupabaseConfigured } from '../supabaseClient.js'

export const DEFAULT_USERS = [
  {
    id: 1,
    ho_ten: 'Đỗ Công Chung',
    ten_dang_nhap: 'Docongchung',
    mat_khau: 'Chung10x7',
    email: 'chung.do@sgc.vn',
    phong_ban: 'Phòng Vật tư',
    quyen: 'Admin',
    chuc_danh: 'Trưởng nhóm',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-10T12:00:00Z'
  },
  {
    id: 2,
    ho_ten: 'Đỗ Hồng Quang',
    ten_dang_nhap: 'Đỗ Hồng Quang',
    mat_khau: '123456',
    email: '—',
    phong_ban: 'Phòng Vật tư',
    quyen: 'User',
    chuc_danh: 'Trưởng nhóm',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-11T12:00:00Z'
  },
  {
    id: 3,
    ho_ten: 'Chử Thanh Hải',
    ten_dang_nhap: 'Chử Thanh Hải',
    mat_khau: '123456',
    email: '—',
    phong_ban: 'Phòng Vật tư',
    quyen: 'User',
    chuc_danh: 'Chuyên viên',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-16T12:00:00Z'
  },
  {
    id: 4,
    ho_ten: 'Nguyễn Thanh Hương',
    ten_dang_nhap: 'Nguyễn Thanh Hương',
    mat_khau: '123456',
    email: '—',
    phong_ban: 'Phòng Vật tư',
    quyen: 'User',
    chuc_danh: 'Chuyên viên',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-16T12:00:00Z'
  },
  {
    id: 5,
    ho_ten: 'Nguyễn Thị Thương',
    ten_dang_nhap: 'Nguyễn Thị Thương',
    mat_khau: '123456',
    email: '—',
    phong_ban: 'Phòng Vật tư',
    quyen: 'User',
    chuc_danh: 'Chuyên viên',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-16T12:00:00Z'
  },
  {
    id: 6,
    ho_ten: 'Tăng Trung Hiếu',
    ten_dang_nhap: 'Tăng Trung Hiếu',
    mat_khau: '123456',
    email: '—',
    phong_ban: 'Phòng Vật tư',
    quyen: 'User',
    chuc_danh: 'Chuyên viên',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-16T12:00:00Z'
  },
  {
    id: 7,
    ho_ten: 'Trần Đức Trung',
    ten_dang_nhap: 'Trần Đức Trung',
    mat_khau: '123456',
    email: '—',
    phong_ban: 'Phòng Vật tư',
    quyen: 'User',
    chuc_danh: 'Chuyên viên',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-13T12:00:00Z'
  },
  {
    id: 8,
    ho_ten: 'Trần Phương Thảo',
    ten_dang_nhap: 'Trần Phương Thảo',
    mat_khau: '123456',
    email: '—',
    phong_ban: 'Phòng Vật tư',
    quyen: 'User',
    chuc_danh: 'Chuyên viên',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-16T12:00:00Z'
  },
  {
    id: 9,
    ho_ten: 'Hoàng Văn Triệu',
    ten_dang_nhap: 'Hoàng Văn Triệu',
    mat_khau: '123456',
    email: '—',
    phong_ban: 'Phòng MMTB',
    quyen: 'User',
    chuc_danh: 'Chuyên viên',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-16T12:00:00Z'
  },
  {
    id: 10,
    ho_ten: 'Lương Quang Khánh',
    ten_dang_nhap: 'Lương Quang Khánh',
    mat_khau: '123456',
    email: '—',
    phong_ban: 'Phòng MMTB',
    quyen: 'User',
    chuc_danh: 'Chuyên viên',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-16T12:00:00Z'
  },
  {
    id: 11,
    ho_ten: 'Nguyễn Anh Quyn',
    ten_dang_nhap: 'Nguyễn Anh Quyn',
    mat_khau: '123456',
    email: '—',
    phong_ban: 'Phòng MMTB',
    quyen: 'User',
    chuc_danh: 'Chuyên viên',
    trang_thai: 'Hoạt động',
    created_at: '2026-05-16T12:00:00Z'
  }
]

// Initialize local fallback store
function getLocalUsers() {
  try {
    const saved = localStorage.getItem('sgc_local_users')
    if (saved) return JSON.parse(saved)
    
    // Save defaults first time
    localStorage.setItem('sgc_local_users', JSON.stringify(DEFAULT_USERS))
    return DEFAULT_USERS
  } catch (e) {
    return DEFAULT_USERS
  }
}

function saveLocalUsers(users) {
  try {
    localStorage.setItem('sgc_local_users', JSON.stringify(users))
  } catch (e) {
    console.warn('Notice: Issue saving local users:', e)
  }
}

// Check if table exists by doing a quick count query
export async function checkSupabaseTableExists() {
  if (!isSupabaseConfigured || !supabase) return false
  try {
    const { error } = await supabase.from('quan_ly_tai_khoan').select('id', { count: 'exact', head: true })
    if (error) {
      // 42P01 means table does not exist
      if (error.code === '42P01' || error.message?.includes('does not exist')) {
        return false
      }
      // If other authorization or network error, let's also fallback safely
      return false
    }
    return true
  } catch (err) {
    return false
  }
}

// Get all accounts
export async function getAccounts() {
  const hasDbTable = await checkSupabaseTableExists()
  if (hasDbTable) {
    try {
      const { data, error } = await supabase
        .from('quan_ly_tai_khoan')
        .select('*')
        .order('id', { ascending: true })
      
      if (!error && data) {
        // Retrieve locally cached states in case column doesn't exist on Supabase table yet
        const extraStatuses = JSON.parse(localStorage.getItem('sgc_supabase_trang_thai') || '{}')
        return data.map(item => {
          let status = item.trang_thai !== undefined && item.trang_thai !== null ? item.trang_thai : 'Hoạt động'
          if (item.trang_thai === undefined || item.trang_thai === null) {
            if (extraStatuses[item.id]) {
              status = extraStatuses[item.id]
            }
          }
          return {
            id: item.id,
            ho_ten: item.ho_ten,
            ten_dang_nhap: item.ten_dang_nhap,
            mat_khau: item.mat_khau,
            email: '—',
            phong_ban: '—',
            quyen: item.quyen || 'User',
            chuc_danh: '—',
            trang_thai: status,
            created_at: item.created_at
          }
        })
      }
      console.log('Notice: Fallback to local accounts storage')
    } catch (err) {
      console.log('Notice: Fallback exception to local accounts storage')
    }
  }
  
  return getLocalUsers()
}

// Create new account
export async function createAccount(accountData) {
  const hasDbTable = await checkSupabaseTableExists()
  if (hasDbTable) {
    try {
      const { data, error } = await supabase
        .from('quan_ly_tai_khoan')
        .insert([{
          ho_ten: accountData.ho_ten,
          ten_dang_nhap: accountData.ten_dang_nhap,
          mat_khau: accountData.mat_khau,
          quyen: accountData.quyen
        }])
        .select()
      
      if (!error) return { success: true, data }
      return { success: false, error: error.message }
    } catch (err) {
      return { success: false, error: err.message }
    }
  }

  // Local storage fallback
  const users = getLocalUsers()
  // Check duplicate
  if (users.some(u => u.ten_dang_nhap.toLowerCase() === accountData.ten_dang_nhap.toLowerCase())) {
    return { success: false, error: 'Tên đăng nhập đã tồn tại trong hệ thống!' }
  }
  
  const newId = users.length > 0 ? Math.max(...users.map(u => u.id)) + 1 : 1
  const newUser = {
    ...accountData,
    id: newId,
    created_at: new Date().toISOString()
  }
  users.push(newUser)
  saveLocalUsers(users)
  return { success: true, data: newUser }
}

// Update existing account
export async function updateAccount(id, accountData) {
  const hasDbTable = await checkSupabaseTableExists()
  if (hasDbTable) {
    try {
      const payload = {
        ho_ten: accountData.ho_ten,
        ten_dang_nhap: accountData.ten_dang_nhap,
        mat_khau: accountData.mat_khau,
        quyen: accountData.quyen
      }
      if (accountData.trang_thai !== undefined) {
        payload.trang_thai = accountData.trang_thai
      }

      const { data, error } = await supabase
        .from('quan_ly_tai_khoan')
        .update(payload)
        .eq('id', id)
        .select()
      
      if (!error) return { success: true, data }

      // Handle missing 'trang_thai' column safely by retrying without it and using local storage mapping
      if (error && (error.code === '42703' || error.message?.includes('trang_thai') || error.message?.includes('does not exist'))) {
        delete payload.trang_thai
        const { data: dataRetry, error: errorRetry } = await supabase
          .from('quan_ly_tai_khoan')
          .update(payload)
          .eq('id', id)
          .select()

        if (!errorRetry) {
          if (accountData.trang_thai) {
            const extraStatuses = JSON.parse(localStorage.getItem('sgc_supabase_trang_thai') || '{}')
            extraStatuses[id] = accountData.trang_thai
            localStorage.setItem('sgc_supabase_trang_thai', JSON.stringify(extraStatuses))
          }
          return { success: true, data: dataRetry }
        }
        return { success: false, error: errorRetry.message }
      }

      return { success: false, error: error.message }
    } catch (err) {
      return { success: false, error: err.message }
    }
  }

  // Local storage fallback
  const users = getLocalUsers()
  const idx = users.findIndex(u => u.id === id)
  if (idx === -1) return { success: false, error: 'Không tìm thấy tài khoản!' }
  
  // Check duplicate username if changed
  if (users.some(u => u.id !== id && u.ten_dang_nhap.toLowerCase() === accountData.ten_dang_nhap.toLowerCase())) {
    return { success: false, error: 'Tên đăng nhập đã tồn tại trong hệ thống!' }
  }

  users[idx] = {
    ...users[idx],
    ...accountData
  }
  saveLocalUsers(users)
  return { success: true, data: users[idx] }
}

// Delete account
export async function deleteAccount(id, username) {
  if (username === 'Docongchung') {
    return { success: false, error: 'Không thể xóa tài khoản Admin hệ thống mặc định!' }
  }

  const hasDbTable = await checkSupabaseTableExists()
  if (hasDbTable) {
    try {
      const { error } = await supabase
        .from('quan_ly_tai_khoan')
        .delete()
        .eq('id', id)
      
      if (!error) return { success: true }
      return { success: false, error: error.message }
    } catch (err) {
      return { success: false, error: err.message }
    }
  }

  // Local storage fallback
  const users = getLocalUsers()
  const filtered = users.filter(u => u.id !== id)
  saveLocalUsers(filtered)
  return { success: true }
}

// Authenticate user login
export async function loginUser(username, password) {
  // First check super-admin hardcoded to ensure constant access
  if (username === 'Docongchung' && password === 'Chung10x7') {
    return {
      success: true,
      user: {
        id: 1,
        ho_ten: 'Đỗ Công Chung',
        ten_dang_nhap: 'Docongchung',
        email: 'chung.do@sgc.vn',
        phong_ban: 'Phòng Vật tư',
        quyen: 'Admin',
        chuc_danh: 'Trưởng nhóm',
        trang_thai: 'Hoạt động'
      }
    }
  }

  const hasDbTable = await checkSupabaseTableExists()
  if (hasDbTable) {
    try {
      const { data, error } = await supabase
        .from('quan_ly_tai_khoan')
        .select('*')
        .eq('ten_dang_nhap', username)
        .maybeSingle()
      
      if (!error && data) {
        if (data.mat_khau === password) {
          let status = data.trang_thai !== undefined && data.trang_thai !== null ? data.trang_thai : 'Hoạt động'
          if (data.trang_thai === undefined || data.trang_thai === null) {
            const extraStatuses = JSON.parse(localStorage.getItem('sgc_supabase_trang_thai') || '{}')
            if (extraStatuses[data.id]) {
              status = extraStatuses[data.id]
            }
          }

          const isActive = (status === 'Hoạt động' || status === 'Đang hoạt động')
          if (!isActive) {
            return { success: false, error: 'Tài khoản đã bị tạm khóa!' }
          }

          return {
            success: true,
            user: {
              id: data.id,
              ho_ten: data.ho_ten,
              ten_dang_nhap: data.ten_dang_nhap,
              email: '—',
              phong_ban: '—',
              quyen: data.quyen || 'User',
              chuc_danh: '—',
              trang_thai: status
            }
          }
        }
      }
    } catch (err) {
      console.log('Notice: Fallback auth check due to exception')
    }
  }

  // Fallback check against local users
  const localUsers = getLocalUsers()
  const found = localUsers.find(u => u.ten_dang_nhap.toLowerCase() === username.toLowerCase())
  if (found) {
    if (found.mat_khau === password) {
      const status = found.trang_thai || 'Hoạt động'
      const isActive = (status === 'Hoạt động' || status === 'Đang hoạt động')
      if (!isActive) {
        return { success: false, error: 'Tài khoản đã bị tạm khóa!' }
      }
      return {
        success: true,
        user: {
          id: found.id,
          ho_ten: found.ho_ten,
          ten_dang_nhap: found.ten_dang_nhap,
          email: found.email || '—',
          phong_ban: found.phong_ban || '—',
          quyen: found.quyen || 'User',
          chuc_danh: found.chuc_danh || '—',
          trang_thai: status
        }
      }
    }
  }

  return { success: false, error: 'Tên đăng nhập hoặc mật khẩu không chính xác!' }
}
