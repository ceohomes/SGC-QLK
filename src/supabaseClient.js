import { NeonPostgrestClient } from '@neondatabase/postgrest-js'

// ============================================================================
// KẾT NỐI NEON DATA API (thay cho Supabase)
// ============================================================================
// Neon Data API tương thích PostgREST, dùng cùng dạng URL /rest/v1 và cùng
// cú pháp .from().select()/.insert()/.delete()/.eq()... như Supabase trước đây,
// nên hầu hết code trong App.jsx không cần sửa gì thêm.
//
// LƯU Ý BẢO MẬT: KHÔNG được dán connection string Postgres (dạng
// postgresql://user:password@...) vào đây. Đó là quyền admin toàn bộ DB.
// Ở đây chỉ dùng Data API URL (public, dạng https://...apirest...neon.tech/...)
// kết hợp với cơ chế GRANT quyền cho role "anonymous" trong Postgres (xem
// hướng dẫn SQL đã gửi kèm) để mô phỏng đúng cách Supabase anon key + RLS
// hoạt động trước đây.
// ============================================================================

// Ưu tiên: localStorage (người dùng tự đổi trong Cài đặt) > biến môi trường > mặc định
const storedUrl = typeof window !== 'undefined' ? localStorage.getItem('sgc_supabase_url') : null
const storedKey = typeof window !== 'undefined' ? localStorage.getItem('sgc_supabase_key') : null

// Neon Data API URL của bạn (đã bao gồm sẵn /rest/v1)
const DEFAULT_SUPABASE_URL = 'https://ep-small-pine-aopqcbty.apirest.c-2.ap-southeast-1.aws.neon.tech/neondb/rest/v1'
// Data API dùng vai trò "anonymous" khi KHÔNG có Authorization header.
// Để trống mặc định (giống cách bạn đang mở toàn quyền cho anon role trên Supabase).
// Nếu sau này bạn cấu hình JWT/Auth provider cho Data API, dán token vào ô "anon key" trong Cài đặt.
const DEFAULT_SUPABASE_ANON_KEY = ''

export const supabaseUrl = storedUrl || import.meta.env.VITE_SUPABASE_URL || DEFAULT_SUPABASE_URL
export const supabaseAnonKey = storedKey || import.meta.env.VITE_SUPABASE_ANON_KEY || DEFAULT_SUPABASE_ANON_KEY

export const isSupabaseConfigured = !!supabaseUrl

const postgrestClient = isSupabaseConfigured
  ? new NeonPostgrestClient({
      dataApiUrl: supabaseUrl,
      options: {
        db: { schema: 'public' },
        // Chỉ gắn Authorization khi có key/token — để trống thì Data API tự
        // dùng role "anonymous", KHÔNG được gửi header rỗng vì sẽ bị hiểu
        // nhầm là có token và bị từ chối (401).
        global: supabaseAnonKey ? { headers: { Authorization: `Bearer ${supabaseAnonKey}` } } : {},
      },
    })
  : null

// ============================================================================
// Mô phỏng Supabase Realtime bằng Polling
// ============================================================================
// Neon Data API (bản beta hiện tại) chưa hỗ trợ Realtime/websocket như Supabase.
// Đây là lớp giả lập giữ nguyên API .channel().on('postgres_changes', ...).subscribe()
// để các đoạn code cũ trong App.jsx không cần sửa, chỉ đổi cơ chế bên trong
// thành polling định kỳ (mặc định 15s) — phù hợp vì các bảng đang lắng nghe
// Realtime trong app đều là bảng nhỏ (vài đến vài chục dòng).
// ============================================================================
const POLL_INTERVAL_MS = 15000

function createPollingChannel() {
  const listeners = []
  let timer = null
  let stopped = false

  const channelObj = {
    on(_event, filter, callback) {
      listeners.push({ filter, callback })
      return channelObj
    },
    subscribe(cb) {
      listeners.forEach(({ filter }) => startPolling(filter))
      if (typeof cb === 'function') cb('SUBSCRIBED')
      return channelObj
    },
    _stop() {
      stopped = true
      if (timer) clearInterval(timer)
    },
  }

  async function startPolling(filter) {
    const table = filter?.table
    if (!table || !postgrestClient) return
    let prevRows = null

    const poll = async () => {
      if (stopped) return
      try {
        const { data, error } = await postgrestClient.from(table).select('*')
        if (error || !data) return

        if (prevRows) {
          const prevMap = new Map(prevRows.map((r) => [r.id, r]))
          const currMap = new Map(data.map((r) => [r.id, r]))

          currMap.forEach((row, id) => {
            if (!prevMap.has(id)) {
              listeners.forEach((l) => l.callback({ eventType: 'INSERT', new: row, old: null }))
            } else if (JSON.stringify(prevMap.get(id)) !== JSON.stringify(row)) {
              listeners.forEach((l) => l.callback({ eventType: 'UPDATE', new: row, old: prevMap.get(id) }))
            }
          })
          prevMap.forEach((row, id) => {
            if (!currMap.has(id)) {
              listeners.forEach((l) => l.callback({ eventType: 'DELETE', new: null, old: row }))
            }
          })
        }
        prevRows = data
      } catch (e) {
        console.error('[Polling] Lỗi khi kiểm tra thay đổi bảng', table, e)
      }
    }

    poll()
    timer = setInterval(poll, POLL_INTERVAL_MS)
  }

  return channelObj
}

// Giữ nguyên "hình dạng" API giống supabase-js để không phải sửa các chỗ gọi
// supabase.from(...), supabase.rpc(...), supabase.channel(...), supabase.removeChannel(...)
// ở khắp App.jsx.
export const supabase = postgrestClient
  ? {
      from: (table) => postgrestClient.from(table),
      rpc: (fn, params) => postgrestClient.rpc(fn, params),
      channel: () => createPollingChannel(),
      removeChannel: (channel) => {
        if (channel && typeof channel._stop === 'function') channel._stop()
      },
    }
  : null
