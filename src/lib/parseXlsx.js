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

  // --- 2. Embedded images + their row anchors via JSZip ---
  const rowImage = await extractRowImages(buf) // Map<rowIndex0, dataUrl>

  // --- 3. Walk rows, build product list ---
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

    rows.push({
      section,
      name,
      volume: vol,
      sku: skuStr,
      price,
      image: rowImage.get(i) || null, // dataUrl or null
    })
  }

  return { rows, sheetName: wsName }
}

/**
 * Unzip xlsx, read drawing anchors → map worksheet row → image dataUrl.
 */
async function extractRowImages(buf) {
  const map = new Map()
  let zip
  try {
    zip = await JSZip.loadAsync(buf)
  } catch {
    return map
  }

  // drawing rels: rId → media path
  const drawingRelsPath = 'xl/drawings/_rels/drawing1.xml.rels'
  const drawingPath = 'xl/drawings/drawing1.xml'
  const relsFile = zip.file(drawingRelsPath)
  const drawFile = zip.file(drawingPath)
  if (!relsFile || !drawFile) return map

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

  // Each anchor block: <xdr:from>...<xdr:row>R</xdr:row>... up to blip r:embed="rIdN"
  // Split by anchor tags to keep each block self-contained.
  const blocks = drawXml.split(/<xdr:(?:oneCellAnchor|twoCellAnchor)/).slice(1)
  for (const blk of blocks) {
    const rowM = blk.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/)
    const ridM = blk.match(/r:embed="(rId\d+)"/)
    if (!rowM || !ridM) continue
    const rowIdx = parseInt(rowM[1], 10) // 0-based worksheet row
    const media = ridToMedia[ridM[1]]
    if (!media) continue
    const url = await mediaDataUrl(media)
    if (url && !map.has(rowIdx)) map.set(rowIdx, url)
  }
  return map
}
