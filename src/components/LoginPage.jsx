import React, { useState } from 'react'
import { motion } from 'motion/react'
import { Lock, User, Eye, EyeOff, AlertCircle } from 'lucide-react'
import { loginUser } from '../utils/authStore.js'

// Kiểu riêng của màn đăng nhập (tông sáng, thẻ kính, logo bật lên)
const LOGIN_CSS = `
.sgl-wrap { min-height: 100vh; min-height: 100dvh; display: flex; align-items: center; justify-content: center; padding: 20px 16px; position: relative; overflow: hidden; box-sizing: border-box;
  background-color: #ECECEE;
  background-image:
    radial-gradient(620px 420px at 12% 18%, rgba(255, 143, 177, .30) 0%, transparent 70%),
    radial-gradient(680px 460px at 88% 22%, rgba(150, 120, 255, .26) 0%, transparent 70%),
    radial-gradient(760px 520px at 55% 105%, rgba(47, 134, 224, .30) 0%, transparent 70%); }
.sgl-card { position: relative; z-index: 2; width: 100%; max-width: 420px; box-sizing: border-box; padding: 36px 32px 30px; border-radius: 30px;
  background: rgba(255, 255, 255, .72); backdrop-filter: blur(18px) saturate(1.2); -webkit-backdrop-filter: blur(18px) saturate(1.2);
  box-shadow: 0 30px 60px -30px rgba(10, 61, 115, .45), inset 0 1px 0 rgba(255, 255, 255, .9), 0 0 0 1px rgba(255, 255, 255, .7); }
.sgl-head { text-align: center; margin-bottom: 26px; }
.sgl-logo { width: 64px; height: 64px; margin: 0 auto 16px; border-radius: 12px; display: grid; place-items: center; position: relative; overflow: hidden;
  background: linear-gradient(135deg, #0A3D73 0%, #0F58A7 55%, #2F86E0 100%); color: #fff; font-size: 22px; font-weight: 700; letter-spacing: .02em;
  box-shadow: 0 14px 28px -12px rgba(15, 88, 167, .85), inset 0 1px 0 rgba(255, 255, 255, .35);
  animation: sglPop .7s cubic-bezier(.34, 1.56, .64, 1) both; }
.sgl-logo::after { content: ""; position: absolute; top: -30%; bottom: -30%; left: -70%; width: 45%; transform: skewX(-20deg);
  background: linear-gradient(90deg, transparent, rgba(255, 255, 255, .55), transparent); animation: sglShine 1s .55s ease-out both; }
@keyframes sglPop { 0% { opacity: 0; transform: translateY(18px) scale(.6); } 60% { opacity: 1; } 100% { opacity: 1; transform: none; } }
@keyframes sglShine { from { left: -70%; } to { left: 130%; } }
.sgl-title { margin: 0 0 6px; font-size: 22px; font-weight: 700; color: #16171C; line-height: 1.3; }
.sgl-sub { margin: 0; font-size: 15px; color: #4A4C55; font-weight: 400; }
.sgl-form { display: flex; flex-direction: column; gap: 16px; }
.sgl-err { display: flex; align-items: center; gap: 10px; padding: 10px 14px; border-radius: 14px; background: #FFE7E5; color: #B4232C; font-size: 15px; font-weight: 600; line-height: 1.4; }
.sgl-field { display: flex; flex-direction: column; gap: 6px; }
.sgl-label { font-size: 15px; font-weight: 600; color: #0A3D73; text-transform: none; letter-spacing: 0; }
.sgl-box { position: relative; }
.sgl-ic { position: absolute; left: 14px; top: 50%; transform: translateY(-50%); color: #6B7AA0; display: flex; pointer-events: none; }
html body .sgl-input { width: 100%; height: 44px; box-sizing: border-box; padding: 0 14px 0 42px; font-size: 16px !important; color: #16171C;
  background: #fff; border: 1.5px solid #D5DCEB; border-radius: 12px; outline: none; transition: border-color .2s, box-shadow .2s; }
html body .sgl-input.pw { padding-right: 46px; }
html body .sgl-input::placeholder { color: #9AA0AD; }
.sgl-eye { position: absolute; right: 6px; top: 50%; transform: translateY(-50%); width: 34px; height: 34px; border: 0; border-radius: 10px; background: transparent; color: #6B7AA0; cursor: pointer; display: grid; place-items: center; padding: 0; }
.sgl-eye:hover { background: #E4EEFA; color: #0F58A7; }
html body .sgl-eye:active:not(:disabled) { transform: translateY(-50%) scale(.96); }
.sgl-submit { margin-top: 6px; width: 100%; height: 46px; border: 0; border-radius: 99px; color: #fff; font-size: 15px; font-weight: 700; cursor: pointer;
  display: flex; align-items: center; justify-content: center; gap: 10px;
  background: linear-gradient(135deg, #0A3D73 0%, #0F58A7 50%, #2F86E0 100%); box-shadow: 0 12px 24px -12px rgba(15, 88, 167, .9), inset 0 1px 0 rgba(255, 255, 255, .3);
  transition: filter .2s, box-shadow .2s; }
.sgl-submit:hover:not(:disabled) { filter: brightness(1.07); box-shadow: 0 14px 28px -12px rgba(15, 88, 167, 1), inset 0 1px 0 rgba(255, 255, 255, .3); }
.sgl-submit:disabled { opacity: .75; cursor: progress; }
.sgl-spin { animation: sglSpin 1s linear infinite; }
@keyframes sglSpin { to { transform: rotate(360deg); } }
@media (max-width: 480px) { .sgl-card { padding: 30px 20px 24px; border-radius: 26px; } }
@media (prefers-reduced-motion: reduce) { .sgl-logo, .sgl-logo::after, .sgl-spin { animation: none !important; } .sgl-logo::after { display: none; } }
`

export default function LoginPage({ onLoginSuccess }) {
  const [username, setUsername] = useState('')
  const [password, setPassword] = useState('')
  const [showPassword, setShowPassword] = useState(false)
  const [error, setError] = useState('')
  const [isLoading, setIsLoading] = useState(false)

  const handleSubmit = async (e) => {
    e.preventDefault()
    if (!username.trim() || !password.trim()) {
      setError('Vui lòng điền đầy đủ Tên đăng nhập và Mật khẩu!')
      return
    }

    setIsLoading(true)
    setError('')

    try {
      const result = await loginUser(username.trim(), password)
      if (result.success) {
        onLoginSuccess(result.user)
      } else {
        setError(result.error)
      }
    } catch (err) {
      setError('Đã xảy ra lỗi hệ thống khi đăng nhập!')
      console.error(err)
    } finally {
      setIsLoading(false)
    }
  }

  return (
    <div className="sgl-wrap">
      <style>{LOGIN_CSS}</style>

      <motion.div
        className="sgl-card"
        initial={{ opacity: 0, y: 24 }}
        animate={{ opacity: 1, y: 0 }}
        transition={{ duration: 0.5, ease: 'easeOut' }}
      >
        {/* Logo và tiêu đề */}
        <div className="sgl-head">
          <div className="sgl-logo" aria-hidden="true">SGC</div>
          <h2 className="sgl-title">Hệ thống báo cáo giao nhận</h2>
          <p className="sgl-sub">Đăng nhập để truy cập ứng dụng</p>
        </div>

        {/* Form đăng nhập */}
        <form onSubmit={handleSubmit} className="sgl-form">
          {error && (
            <motion.div
              className="sgl-err"
              initial={{ opacity: 0, scale: 0.95 }}
              animate={{ opacity: 1, scale: 1 }}
            >
              <AlertCircle size={18} style={{ flexShrink: 0, color: '#FF3B30' }} />
              <span>{error}</span>
            </motion.div>
          )}

          {/* Tên đăng nhập */}
          <div className="sgl-field">
            <label className="sgl-label" htmlFor="login_username_input">Tên đăng nhập</label>
            <div className="sgl-box">
              <span className="sgl-ic"><User size={18} /></span>
              <input
                id="login_username_input"
                className="sgl-input"
                type="text"
                required
                autoComplete="username"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder="Nhập tên đăng nhập..."
              />
            </div>
          </div>

          {/* Mật khẩu */}
          <div className="sgl-field">
            <label className="sgl-label" htmlFor="login_password_input">Mật khẩu</label>
            <div className="sgl-box">
              <span className="sgl-ic"><Lock size={18} /></span>
              <input
                id="login_password_input"
                className="sgl-input pw"
                type={showPassword ? 'text' : 'password'}
                required
                autoComplete="current-password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Nhập mật khẩu..."
              />
              <button
                type="button"
                className="sgl-eye"
                onClick={() => setShowPassword(!showPassword)}
                title={showPassword ? 'Ẩn mật khẩu' : 'Hiện mật khẩu'}
                aria-label={showPassword ? 'Ẩn mật khẩu' : 'Hiện mật khẩu'}
              >
                {showPassword ? <EyeOff size={18} /> : <Eye size={18} />}
              </button>
            </div>
          </div>

          {/* Nút đăng nhập */}
          <button
            id="login_submit_btn"
            type="submit"
            className="sgl-submit"
            disabled={isLoading}
          >
            {isLoading ? (
              <span style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                <svg className="sgl-spin" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="3">
                  <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/>
                </svg>
                Đang đăng nhập...
              </span>
            ) : (
              'Đăng nhập hệ thống'
            )}
          </button>
        </form>
      </motion.div>
    </div>
  )
}
