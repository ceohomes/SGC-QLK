// Chỉ chỉnh trình bày (không đụng logic):
// 1) Ô tiêu đề bảng có màu đậm riêng (nhóm Nhập kho xanh lá, Xuất kho cam…) → gắn data-tone
//    để theme.css tô nền nhạt cùng tông thay vì khối màu đậm chữ trắng.
// 2) Ô chỉ chứa số (số to trên thẻ, số trong bảng) → gắn lớp sg-num để dùng phông Roboto.

function toneOf(color) {
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(color || '')
  if (!m) return null
  const r = +m[1] / 255, g = +m[2] / 255, b = +m[3] / 255
  const max = Math.max(r, g, b), min = Math.min(r, g, b)
  const l = (max + min) / 2
  if (l > 0.8) return null // đã là nền nhạt
  const d = max - min
  if (d < 0.08) return null
  let h
  if (max === r) h = ((g - b) / d) % 6
  else if (max === g) h = (b - r) / d + 2
  else h = (r - g) / d + 4
  h = (h * 60 + 360) % 360
  if (h < 15 || h >= 345) return 'red'
  if (h < 40) return 'orange'
  if (h < 65) return 'amber'
  if (h < 190) return 'green'
  if (h < 250) return l < 0.32 ? 'navy' : null
  if (h < 345) return 'violet'
  return null
}

const NUM_RE = /^[\s+\-−]*[\d.,]+\s*(%|đ|tỷ|tr)?\s*$/

function fixTh(th) {
  const bg = th.style.background || th.style.backgroundColor
  const tone = bg ? toneOf(th.style.backgroundColor || bg) : null
  if (tone) th.setAttribute('data-tone', tone)
  else th.removeAttribute('data-tone')
}

function fixNum(el) {
  if (el.childElementCount === 0 && NUM_RE.test(el.textContent || '')) el.classList.add('sg-num')
}

const XL_RE = /^\s*(Xuất\s+(file\s+)?Excel|Tải\s+Excel)/i
function fixBtn(b) {
  if (XL_RE.test(b.textContent || '')) b.classList.add('sg-xl')
  // Nút chỉ có dấu X (không chữ) → nút tròn đỏ nhạt kiểu iPhone
  const onlyX = b.childElementCount === 1 && b.firstElementChild.matches('svg.lucide-x') && !(b.textContent || '').trim()
  b.classList.toggle('sg-xbtn', onlyX)
}

function scan(root) {
  if (!(root instanceof Element)) return
  if (root.tagName === 'BUTTON') fixBtn(root)
  root.querySelectorAll('button').forEach(fixBtn)
  if (root.tagName === 'TH') fixTh(root)
  root.querySelectorAll('thead th').forEach(fixTh)
  if (root.childElementCount === 0) fixNum(root)
  root.querySelectorAll('td, b, strong, span, div').forEach(el => {
    if (el.childElementCount === 0) fixNum(el)
  })
}

let queued = new Set()
let raf = 0
function flush() {
  raf = 0
  const items = queued
  queued = new Set()
  items.forEach(scan)
}

export function startThemeFix() {
  if (typeof window === 'undefined') return
  scan(document.body)
  const obs = new MutationObserver(list => {
    for (const m of list) {
      if (m.type === 'attributes') {
        if (m.target.tagName === 'TH') queued.add(m.target)
        continue
      }
      if (m.type === 'characterData') {
        const p = m.target.parentElement
        if (p && p.childElementCount === 0) { p.classList.remove('sg-num'); queued.add(p) }
        continue
      }
      m.addedNodes.forEach(n => {
        if (n.nodeType === 1) queued.add(n)
        else if (n.nodeType === 3 && n.parentElement) { n.parentElement.classList.remove('sg-num'); queued.add(n.parentElement) }
      })
    }
    if (!raf && queued.size) raf = requestAnimationFrame(flush)
  })
  obs.observe(document.body, { childList: true, subtree: true, characterData: true, attributes: true, attributeFilter: ['style'] })
}
