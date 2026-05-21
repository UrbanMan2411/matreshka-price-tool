import { PDFDocument, rgb } from 'pdf-lib'
import fontkit from '@pdf-lib/fontkit'

const MM = 2.83465 // pt per mm
const mm = (v) => v * MM

// Brand palette — MATRЁSHKA (red / gold / warm cream folk theme).
// Variable names kept from the Green Panda base; only values changed.
const GREEN9 = rgb(0x7a / 255, 0x1b / 255, 0x22 / 255) // dark bordeaux — header bar, title
const GREEN7 = rgb(0xa8 / 255, 0x26 / 255, 0x2f / 255) // red — section headers, price
const GREEN5 = rgb(0xc8 / 255, 0x97 / 255, 0x2b / 255) // gold — divider accents
const INK = rgb(0x2a / 255, 0x1a / 255, 0x14 / 255) // dark brown ink
const INK7 = rgb(0x4a / 255, 0x3a / 255, 0x30 / 255)
const MUTED = rgb(0x8a / 255, 0x7a / 255, 0x6a / 255)
const LINE = rgb(0xe6 / 255, 0xd9 / 255, 0xc2 / 255)
const PAPER = rgb(0xfb / 255, 0xf4 / 255, 0xe6 / 255) // warm cream
const WHITE = rgb(1, 1, 1)

const today = () => {
  const d = new Date()
  const p = (n) => String(n).padStart(2, '0')
  return `${p(d.getDate())}.${p(d.getMonth() + 1)}.${d.getFullYear()}`
}

const fmtPrice = (p) =>
  p > 0 ? p.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : '—'

async function fetchBytes(url) {
  const r = await fetch(url)
  if (!r.ok) throw new Error(`fetch ${url}: ${r.status}`)
  return new Uint8Array(await r.arrayBuffer())
}

// word-wrap text to a max width; returns array of lines
function wrapText(text, font, size, maxW) {
  const words = String(text).split(/\s+/)
  const lines = []
  let cur = ''
  for (const w of words) {
    const cand = cur ? cur + ' ' + w : w
    if (font.widthOfTextAtSize(cand, size) > maxW && cur) {
      lines.push(cur)
      cur = w
    } else {
      cur = cand
    }
  }
  if (cur) lines.push(cur)
  return lines
}

export async function buildPriceListPdf(rows, options = {}) {
  const doc = await PDFDocument.create()
  doc.registerFontkit(fontkit)

  const [regBytes, boldBytes] = await Promise.all([
    fetchBytes('/brand/Manrope-Regular.ttf'),
    fetchBytes('/brand/Manrope-Bold.ttf'),
  ])
  const reg = await doc.embedFont(regBytes, { subset: true })
  const bold = await doc.embedFont(boldBytes, { subset: true })
  const logo = null // MATRЁSHKA uses a text wordmark, no logo image

  // Background watermark: 'default' (bundled), 'none', or a custom data URL.
  const bgOpt = options.bg ?? 'default'
  const bgOpacity = options.bgOpacity ?? 0.1
  let bg = null
  try {
    if (bgOpt === 'default') {
      bg = await doc.embedJpg(await fetchBytes('/brand/bg.jpg'))
    } else if (typeof bgOpt === 'string' && bgOpt.startsWith('data:')) {
      const isPng = bgOpt.startsWith('data:image/png')
      const bytes = Uint8Array.from(atob(bgOpt.split(',')[1]), (c) => c.charCodeAt(0))
      bg = isPng ? await doc.embedPng(bytes) : await doc.embedJpg(bytes)
    }
  } catch {
    bg = null
  }

  // embed product images once (dedupe by dataUrl)
  const imgCache = new Map()
  async function embedImg(dataUrl) {
    if (!dataUrl) return null
    if (imgCache.has(dataUrl)) return imgCache.get(dataUrl)
    const isPng = dataUrl.startsWith('data:image/png')
    const b64 = dataUrl.split(',')[1]
    const bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))
    let img = null
    try {
      img = isPng ? await doc.embedPng(bytes) : await doc.embedJpg(bytes)
    } catch {
      img = null
    }
    imgCache.set(dataUrl, img)
    return img
  }
  for (const r of rows) r._img = await embedImg(r.image)

  // A4 landscape
  const PW = mm(297), PH = mm(210)
  // Columns (mm)
  const C = { photo: 10, name: 46, vol: 150, sku: 178, price: 230 }
  const RIGHT = 287
  const nameW = mm(C.vol - C.name - 3)
  const ROW_MIN = mm(16)

  // Precompute row layouts (wrapped name + height)
  const NAME_SIZE = 8.6, NAME_LEAD = 10.6
  for (const r of rows) {
    r._nameLines = wrapText(r.name, bold, NAME_SIZE, nameW)
    const textH = (r._nameLines.length - 1) * NAME_LEAD + NAME_SIZE
    r._h = Math.max(ROW_MIN, textH + mm(4)) // padding top+bottom
  }

  // Paginate
  const top0 = PH - mm(30)
  const bottomLim = mm(13)
  // Section header sits close to its own rows (small below), with clear
  // separation from the previous group (larger above).
  const SEC_ABOVE = mm(5), SEC_BELOW = mm(1.5)
  const pages = []
  let cur = []
  let y = top0 - mm(8)
  let sec = null
  for (const r of rows) {
    const need = r._h + (r.section !== sec ? SEC_ABOVE + SEC_BELOW : 0)
    if (y - need < bottomLim) {
      pages.push(cur)
      cur = []
      y = top0 - mm(8)
      sec = null
    }
    if (r.section !== sec) {
      sec = r.section
      y -= SEC_ABOVE
      cur.push({ kind: 'sec', y, text: r.section })
      y -= SEC_BELOW
    }
    cur.push({ kind: 'row', y, r })
    y -= r._h
  }
  if (cur.length) pages.push(cur)

  pages.forEach((items, pi) => {
    const page = doc.addPage([PW, PH])
    drawBg(page, PW, PH, bg, bgOpacity)
    drawHeader(page, PW, PH, logo, reg, bold)
    drawTHead(page, top0, C, RIGHT, reg, bold)

    for (const it of items) {
      if (it.kind === 'sec') {
        drawSection(page, it.y, it.text, C, RIGHT, bold)
        continue
      }
      const r = it.r
      const rb = it.y - r._h
      const cy = rb + r._h / 2 // vertical centre of the row band
      // row underline
      page.drawLine({
        start: { x: mm(8), y: rb },
        end: { x: mm(RIGHT), y: rb },
        thickness: 0.3,
        color: LINE,
      })
      // photo — centred in its cell
      if (r._img) {
        const cw = mm(C.name - C.photo - 3)
        const ch = r._h - mm(4)
        const scale = Math.min(cw / r._img.width, ch / r._img.height)
        const w = r._img.width * scale
        const h = r._img.height * scale
        page.drawImage(r._img, {
          x: mm(C.photo) + (cw - w) / 2,
          y: cy - h / 2,
          width: w,
          height: h,
        })
      }
      // name (wrapped) — block vertically centred
      const n = r._nameLines.length
      const firstBaseline = cy + ((n - 1) * NAME_LEAD) / 2 - NAME_SIZE * 0.25
      r._nameLines.forEach((ln, k) => {
        page.drawText(ln, {
          x: mm(C.name) + mm(1),
          y: firstBaseline - k * NAME_LEAD,
          size: NAME_SIZE,
          font: bold,
          color: INK,
        })
      })
      // vol / sku — centre-aligned within their columns
      const baseS = cy - 8.4 * 0.35
      const volCenter = mm((C.vol + C.sku) / 2)
      const skuCenter = mm((C.sku + C.price) / 2)
      const priceCenter = mm((C.price + RIGHT) / 2)
      page.drawText(r.volume, { x: volCenter - reg.widthOfTextAtSize(r.volume, 8.4) / 2, y: baseS, size: 8.4, font: reg, color: INK7 })
      const skuS = String(r.sku)
      page.drawText(skuS, { x: skuCenter - reg.widthOfTextAtSize(skuS, 8.4) / 2, y: baseS, size: 8.4, font: reg, color: INK7 })
      // price — centre-aligned within its column
      const priceTxt = fmtPrice(r.price)
      page.drawText(priceTxt, {
        x: priceCenter - bold.widthOfTextAtSize(priceTxt, 12) / 2,
        y: cy - 12 * 0.35,
        size: 12,
        font: bold,
        color: GREEN7,
      })
    }

    drawFooter(page, PW, pi + 1, pages.length, reg)
  })

  return await doc.save()
}

function drawBg(page, W, H, bg, opacity = 0.1) {
  page.drawRectangle({ x: 0, y: 0, width: W, height: H, color: PAPER })
  // faint watermark, cover (skipped when bg is null = 'none')
  if (bg) {
    const asp = bg.width / bg.height
    let tw = W, th = tw / asp
    if (th < H) { th = H; tw = th * asp }
    page.drawImage(bg, { x: (W - tw) / 2, y: (H - th) / 2, width: tw, height: th, opacity })
  }
  // top + bottom green ribbons
  page.drawRectangle({ x: 0, y: H - mm(3.5), width: W, height: mm(3.5), color: GREEN9 })
  page.drawRectangle({ x: 0, y: 0, width: W, height: mm(3.5), color: GREEN9 })
}

function drawHeader(page, W, H, _logo, reg, bold) {
  const y = H - mm(11)
  // Text wordmark (no logo image for MATRЁSHKA)
  page.drawText('MATRЁSHKA', { x: mm(10), y: y - mm(6), size: 21, font: bold, color: GREEN9 })
  page.drawText('Чистота с юга России', { x: mm(10), y: y - mm(10.5), size: 7.5, font: reg, color: MUTED })
  const title = 'Прайс-лист MATRЁSHKA · оптовые цены'
  page.drawText(title, {
    x: W / 2 - bold.widthOfTextAtSize(title, 15) / 2,
    y: y - mm(3),
    size: 15,
    font: bold,
    color: GREEN9,
  })
  const sub = `ООО «КубаньБытХим» · ИНН 2315984520 · г. Новороссийск · ред. от ${today()}`
  page.drawText(sub, {
    x: W / 2 - reg.widthOfTextAtSize(sub, 8) / 2,
    y: y - mm(8),
    size: 8,
    font: reg,
    color: MUTED,
  })
  const phone = '+7 (8617) 60-00-88'
  page.drawText(phone, { x: W - mm(10) - bold.widthOfTextAtSize(phone, 9.5), y: y - mm(1), size: 9.5, font: bold, color: INK })
  const mail = 'info@kubanbithim.ru'
  page.drawText(mail, { x: W - mm(10) - reg.widthOfTextAtSize(mail, 8), y: y - mm(5.5), size: 8, font: reg, color: MUTED })
}

function drawTHead(page, ymm, C, RIGHT, reg, bold) {
  page.drawRectangle({ x: mm(8), y: ymm - mm(6), width: mm(RIGHT - 8), height: mm(7), color: GREEN9 })
  const ty = ymm - mm(4)
  const sz = 8.2
  // 'Наименование' stays left-aligned (text is left-aligned)
  page.drawText('Наименование', { x: mm(C.name) + mm(1), y: ty, size: sz, font: bold, color: WHITE })
  // centred headers over their column centres (values are centred too)
  const center = (a, b) => mm((a + b) / 2)
  const drawCentered = (text, cx, size = sz) =>
    page.drawText(text, { x: cx - bold.widthOfTextAtSize(text, size) / 2, y: ty, size, font: bold, color: WHITE })
  drawCentered('Фото', center(C.photo, C.name - 3))
  drawCentered('Объём', center(C.vol, C.sku))
  drawCentered('Артикул', center(C.sku, C.price))
  // price label (Matreshka price column = «Цена в руб, Дистрибьютор»)
  drawCentered('Цена ₽, дистрибьютор', center(C.price, RIGHT), 7.6)
}

function drawSection(page, y, title, C, RIGHT, bold) {
  const t = '› ' + (title || '').toUpperCase()
  page.drawText(t, { x: mm(10), y, size: 9.5, font: bold, color: GREEN7 })
  const tx = mm(10) + bold.widthOfTextAtSize(t, 9.5) + mm(3)
  page.drawLine({ start: { x: tx, y: y + mm(1.4) }, end: { x: mm(RIGHT), y: y + mm(1.4) }, thickness: 0.5, color: GREEN5 })
}

function drawFooter(page, W, pi, total, reg) {
  const txt = 'ООО «КубаньБытХим» · ТМ MATRЁSHKA · г. Новороссийск, ул. Кутузовская, 117 · +7 (8617) 60-00-88 · info@kubanbithim.ru'
  page.drawText(txt, { x: mm(10), y: mm(7), size: 7.5, font: reg, color: MUTED })
  const pg = `Стр. ${pi} из ${total}`
  page.drawText(pg, { x: W - mm(10) - reg.widthOfTextAtSize(pg, 7.5), y: mm(7), size: 7.5, font: reg, color: MUTED })
}
