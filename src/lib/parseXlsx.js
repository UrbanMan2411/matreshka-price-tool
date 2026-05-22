import * as XLSX from 'xlsx'
import JSZip from 'jszip'

// SKU prefixes used only as a hint for fallback labels; photos come from xlsx.
const norm = (v) => (v == null ? '' : String(v)).replace(/\s+/g, ' ').trim()

/**
 * Parse the GREEN PANDA price xlsx.
 * Returns { rows: [{section, name, volume, sku, price, imageIndex|null}], images: [{dataUrl}] }
 * Layout (1-based cols): 1 Фото | 2 наименование | 3 описание | 4 объём | 5 артикул |
 *                        6 штрих-код | 7 паллет | 8 в коробе | 9 цена
 * Section headers: a row where col1 has text but col2 & col5 are empty.
 */
export async function parsePriceXlsx(file) {
  const buf = await file.arrayBuffer()

  // --- 1. Cells via SheetJS ---
  const wb = XLSX.read(buf, { type: 'array' })
  const wsName = wb.SheetNames[0]
  const ws = wb.Sheets[wsName]
  const grid = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null })

  // --- 2. Embedded image anchors via JSZip: [{from, to, url}] sorted by row ---
  const anchors = await extractImageAnchors(buf)

  // --- 3. Walk rows, build product list (remember each product's grid row) ---
  const rows = []
  let section = ''
  // find header row (the one containing 'наименование')
  let headerIdx = grid.findIndex(
    (r) => r && r.some((c) => norm(c).toLowerCase() === 'наименование')
  )
  if (headerIdx < 0) headerIdx = 3 // fallback

  for (let i = headerIdx + 1; i < grid.length; i++) {
    const r = grid[i] || []
    const c1 = norm(r[0]), name = norm(r[1]), sku = r[4]
    const skuStr = norm(sku)

    // section header row
    if (c1 && !name && !skuStr) { section = c1; continue }
    if (!name || !skuStr) continue

    const vol = typeof r[3] === 'number' ? `${r[3]} л` : norm(r[3])
    const priceRaw = r[8]
    const price = typeof priceRaw === 'number'
      ? Math.round(priceRaw * 100) / 100
      : parseFloat(norm(priceRaw).replace(',', '.')) || 0

    rows.push({ rowIdx: i, section, name, volume: vol, sku: skuStr, price, image: null })
  }

  // --- 4. Assign images to products by row span (handles twoCellAnchor
  //        offsets where the image top-edge sits a row off the data row). ---
  assignImages(rows, anchors)
  for (const r of rows) delete r.rowIdx

  return { rows, sheetName: wsName }
}

/**
 * Greedily match image anchors to product rows. Each product takes the best
 * still-unused anchor: (1) exact from-row, (2) the anchor whose from..to span
 * contains the row and starts closest above it, (3) nearest by centre (±2 rows).
 * Header/decoration images (above the first product row) match nothing.
 */
function assignImages(products, anchors) {
  const used = new Array(anchors.length).fill(false)
  const pick = (rowIdx) => {
    // tier 1 — exact from-row
    for (let k = 0; k < anchors.length; k++) {
      if (!used[k] && anchors[k].from === rowIdx) return k
    }
    // tier 2 — span contains the row; pick the closest from above
    let best = -1, bestFrom = -Infinity
    for (let k = 0; k < anchors.length; k++) {
      if (used[k]) continue
      const a = anchors[k]
      if (a.from <= rowIdx && rowIdx <= a.to && a.from > bestFrom) { best = k; bestFrom = a.from }
    }
    if (best >= 0) return best
    // tier 2b — span ±1 row (twoCellAnchor top-edge can sit a row off)
    best = -1; bestFrom = -Infinity
    for (let k = 0; k < anchors.length; k++) {
      if (used[k]) continue
      const a = anchors[k]
      if (a.from - 1 <= rowIdx && rowIdx <= a.to + 1 && a.from > bestFrom) { best = k; bestFrom = a.from }
    }
    if (best >= 0) return best
    // tier 3 — nearest anchor centre within 2 rows
    let bd = 2.5, bk = -1
    for (let k = 0; k < anchors.length; k++) {
      if (used[k]) continue
      const c = (anchors[k].from + anchors[k].to) / 2
      const d = Math.abs(c - rowIdx)
      if (d < bd) { bd = d; bk = k }
    }
    return bk
  }
  for (const p of products) {
    const k = pick(p.rowIdx)
    if (k >= 0) { p.image = anchors[k].url; used[k] = true }
  }
}

/**
 * Unzip xlsx, read drawing anchors → array of {from, to, url} (0-based rows),
 * sorted by from-row. `to` = bottom row the image spans (twoCellAnchor); for
 * oneCellAnchor `to` falls back to `from`.
 */
async function extractImageAnchors(buf) {
  const out = []
  let zip
  try {
    zip = await JSZip.loadAsync(buf)
  } catch {
    return out
  }

  // drawing rels: rId → media path
  const drawingRelsPath = 'xl/drawings/_rels/drawing1.xml.rels'
  const drawingPath = 'xl/drawings/drawing1.xml'
  const relsFile = zip.file(drawingRelsPath)
  const drawFile = zip.file(drawingPath)
  if (!relsFile || !drawFile) return out

  const relsXml = await relsFile.async('string')
  const drawXml = await drawFile.async('string')

  // rId -> target media
  const ridToMedia = {}
  for (const m of relsXml.matchAll(/Id="(rId\d+)"[^>]*Target="([^"]+)"/g)) {
    const target = m[2].replace('../', 'xl/')
    ridToMedia[m[1]] = target
  }

  // cache media → dataUrl
  const mediaCache = {}
  async function mediaDataUrl(path) {
    if (mediaCache[path]) return mediaCache[path]
    const f = zip.file(path)
    if (!f) return null
    const b64 = await f.async('base64')
    const ext = path.split('.').pop().toLowerCase()
    const mime = ext === 'png' ? 'image/png' : 'image/jpeg'
    const url = `data:${mime};base64,${b64}`
    mediaCache[path] = url
    return url
  }

  // Each anchor block: <xdr:from><xdr:row>R</xdr:row>…<xdr:to><xdr:row>R2</xdr:row>
  // … up to a blip embed="rIdN". Split by anchor tags to keep blocks isolated.
  const blocks = drawXml.split(/<xdr:(?:oneCellAnchor|twoCellAnchor)/).slice(1)
  for (const blk of blocks) {
    const fromM = blk.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/)
    const toM = blk.match(/<xdr:to>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/)
    const ridM = blk.match(/embed="(rId\d+)"/) // any ns prefix (r:embed etc.)
    if (!fromM || !ridM) continue
    const from = parseInt(fromM[1], 10) // 0-based worksheet row
    const to = toM ? parseInt(toM[1], 10) : from
    const media = ridToMedia[ridM[1]]
    if (!media) continue
    const url = await mediaDataUrl(media)
    if (url) out.push({ from, to: Math.max(from, to), url })
  }
  out.sort((a, b) => a.from - b.from || a.to - b.to)
  return out
}
