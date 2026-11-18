import React from 'react'
import { Upload } from 'lucide-react'

// Khối tải file lên kiểu "gradient nổi bật" (mẫu 4 anh Chung chọn 18/11/2026):
// khối xanh gradient có tiêu đề + các ô số, bên trong là vùng thả file nền trắng kính.
// Chỉ lo phần trình bày; việc đọc file do component cha truyền vào qua các hàm xử lý.
export default function UploadHero({
  title,
  sub,
  stats = [],
  dropTitle = 'Kéo thả file vào khung này',
  dropNote,
  buttonLabel = 'Chọn file Excel',
  dragging = false,
  disabled = false,
  disabledNote,
  onPick,
  onDragOver,
  onDragLeave,
  onDrop,
}) {
  return (
    <div className={'sg-hero' + (disabled ? ' off' : '')}>
      <div className="sg-hero-top">
        <div className="sg-hero-txt">
          <h3>{title}</h3>
          {sub && <p>{sub}</p>}
        </div>
        {stats.length > 0 && (
          <div className="sg-hero-stats">
            {stats.map(s => (
              <div key={s.l} className="sg-hero-k"><b>{s.v}</b><small>{s.l}</small></div>
            ))}
          </div>
        )}
      </div>
      <div
        className={'sg-hero-drop' + (dragging && !disabled ? ' over' : '')}
        onClick={() => { if (!disabled && onPick) onPick() }}
        onDragOver={onDragOver}
        onDragLeave={onDragLeave}
        onDrop={onDrop}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => { if (!disabled && onPick && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); onPick() } }}
      >
        <span className="sg-hero-ic"><Upload size={26} strokeWidth={2.3} /></span>
        <span className="sg-hero-dt">
          <b>{disabled ? 'Chưa tải lên được' : (dragging ? 'Thả file vào đây' : dropTitle)}</b>
          <span>{disabled ? disabledNote : dropNote}</span>
        </span>
        <span className="sg-btn" aria-hidden="true" style={{ pointerEvents: 'none', opacity: disabled ? 0.5 : 1 }}>
          <Upload size={15} />{buttonLabel}
        </span>
      </div>
    </div>
  )
}
