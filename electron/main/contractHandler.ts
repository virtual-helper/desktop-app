import { ipcMain, app, shell } from 'electron'
import { writeFileSync, mkdirSync } from 'node:fs'
import path from 'node:path'
import { createRequire } from 'node:module'

const require = createRequire(import.meta.url)

export interface SheetTable {
  headers: string[]
  rows: Record<string, string>[]
}

export interface ParentHeader {
  text: string
  span: number
}

// One table placeholder and how to fill it. The marker is {{#key}}/【#key】 by
// default, or a plain {{key}}/【key】 when promoted from a text placeholder in UI.
export interface TableMapping {
  key: string                 // placeholder name (the part after # if any)
  headers: string[]
  parentHeaders?: ParentHeader[]  // optional two-level header row above the column names
  rows: Record<string, string>[]
  requireHash?: boolean       // false => marker has no # (promoted text placeholder)
  mergeColumns?: string[]     // columns whose consecutive equal values are vertically merged
  landscape?: boolean         // wrap the inserted table in its own landscape section
  autoLocate?: boolean        // no marker: replace the static appendix table in the template
  filterMainColumn?: string   // main-row column used to filter the table
  filterSheetColumn?: string  // table column matched against the main-row value
}

export interface GeneratePayload {
  templateBuffer: number[]
  rows: Record<string, string>[]      // main-sheet rows, one document per row
  textMap: Record<string, string>     // placeholder name -> main-sheet column
  tables: TableMapping[]
  fileNamePattern: string             // supports {{列名}} tokens
  outputDirName: string
}

function escapeXml(s: string): string {
  return String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function escapeRegExp(s: string): string {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

// Supported placeholder delimiters: {{name}} and 【name】 (tables: {{#x}} / 【#x】).
// Each entry is [openEscaped, closeEscaped, innerNegatedCharClass].
const DELIMS: ReadonlyArray<readonly [string, string, string]> = [
  ['\\{\\{', '\\}\\}', '{}'],
  ['【', '】', '【】'],
]

// Matches any opening delimiter; used to skip parts with no placeholders.
const HAS_PLACEHOLDER = new RegExp(DELIMS.map(([o]) => o).join('|'))

// Build a regex source that matches the placeholder `key` in any delimiter form.
// When `table` is true the `#` table prefix is required.
function keyPattern(key: string, table: boolean): string {
  const hash = table ? '#\\s*' : ''
  return DELIMS.map(([o, c]) => `${o}\\s*${hash}${escapeRegExp(key)}\\s*${c}`).join('|')
}

// Word frequently splits a placeholder across several runs (e.g. `{{`, name, `}}`).
// This collapses everything between the delimiters into a single run so plain
// text matching works. It keeps the first run's formatting and drops empty ones.
function repairPlaceholders(xml: string): string {
  const merge = (m: string) => m.replace(/<\/w:t>[\s\S]*?<w:t[^>]*>/g, '')
  let out = xml
  for (const [o, c] of DELIMS) {
    // A real placeholder always lives inside a single paragraph. Forbidding the
    // match from crossing a </w:p> stops a stray delimiter (or one whose closer
    // is paragraphs away) from collapsing paragraphs and silently deleting
    // structure such as section breaks (<w:sectPr>, e.g. the landscape appendix).
    out = out.replace(new RegExp(`${o}(?:(?!</w:p>)[\\s\\S])*?${c}`, 'g'), merge)
  }
  return out
}

// vMerge: 'restart' starts a vertically-merged group, 'continue' joins it upward.
// widthPct is the cell width in 50ths of a percent (so 5000 == 100%).
function cellXml(text: string, bold: boolean, widthPct: number, vMerge?: 'restart' | 'continue', gridSpan?: number): string {
  const rpr = `<w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体" w:hint="eastAsia"/>${bold ? '<w:b/>' : '<w:bCs/>'}<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr>`
  const vMergeXml = vMerge === 'restart' ? '<w:vMerge w:val="restart"/>' : vMerge === 'continue' ? '<w:vMerge/>' : ''
  const gridSpanXml = gridSpan && gridSpan > 1 ? `<w:gridSpan w:val="${gridSpan}"/>` : ''
  // A continued cell carries no text; its content shows from the restart cell above.
  const body =
    vMerge === 'continue'
      ? '<w:p/>'
      : `<w:p><w:pPr><w:jc w:val="center"/><w:rPr>${bold ? '<w:b/>' : '<w:bCs/>'}</w:rPr></w:pPr>` +
        `<w:r>${rpr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`
  return (
    '<w:tc>' +
    `<w:tcPr><w:tcW w:w="${widthPct}" w:type="pct"/>${gridSpanXml}${vMergeXml}<w:vAlign w:val="center"/></w:tcPr>` +
    body +
    '</w:tc>'
  )
}

function buildTableXml(
  headers: string[],
  rows: Record<string, string>[],
  mergeColumns: string[] = [],
  parentHeaders?: ParentHeader[],
): string {
  const borders =
    '<w:tblBorders>' +
    ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
      .map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
      .join('') +
    '</w:tblBorders>'
  // Fixed column widths matching the template: all cols 312 pct, last col 313 pct.
  const n = Math.max(headers.length, 1)
  const colPct = (i: number) => i === n - 1 ? 313 : 312
  const grid = '<w:tblGrid>' + Array.from({ length: n }, (_, i) => `<w:gridCol w:w="${i === n - 1 ? 1004 : 1001}"/>`).join('') + '</w:tblGrid>'
  let parentRow = ''
  let headerRow = ''
  if (parentHeaders && parentHeaders.length > 0) {
    let colOffset = 0
    const parentCells: string[] = []
    const childCells: string[] = []
    for (const ph of parentHeaders) {
      if (ph.span <= 1) {
        // No parent text → no vertical merge, header text in child row
        parentCells.push(cellXml(ph.text, true, colPct(colOffset)))
        childCells.push(cellXml(headers[colOffset] ?? '', true, colPct(colOffset)))
      } else {
        const spanWidth = ph.span > 1 && colOffset + ph.span - 1 === n - 1
          ? 312 * (ph.span - 1) + 313
          : 312 * ph.span
        parentCells.push(cellXml(ph.text, true, spanWidth, undefined, ph.span))
        for (let k = 0; k < ph.span; k++) {
          childCells.push(cellXml(headers[colOffset + k] ?? '', true, colPct(colOffset + k)))
        }
      }
      colOffset += ph.span
    }
    parentRow = '<w:tr><w:trPr><w:trHeight w:val="280"/></w:trPr>' + parentCells.join('') + '</w:tr>'
    headerRow = '<w:tr><w:trPr><w:trHeight w:val="280"/></w:trPr>' + childCells.join('') + '</w:tr>'
  } else {
    headerRow =
      '<w:tr><w:trPr><w:trHeight w:val="280"/></w:trPr>' +
      headers.map((h, i) => cellXml(h, true, colPct(i))).join('') +
      '</w:tr>'
  }
  const mergeSet = new Set(mergeColumns)
  const bodyRows = rows
    .map((r, i) => {
      const cells = headers.map((h, hi) => {
        const val = String(r[h] ?? '')
        let vMerge: 'restart' | 'continue' | undefined
        if (mergeSet.has(h)) {
          const prev = i > 0 ? String(rows[i - 1][h] ?? '') : ''
          vMerge = i > 0 && val.trim() !== '' && val === prev ? 'continue' : 'restart'
        }
        return cellXml(val, false, colPct(hi), vMerge)
      })
      return '<w:tr><w:trPr><w:trHeight w:val="280"/></w:trPr>' + cells.join('') + '</w:tr>'
    })
    .join('')
  return (
    '<w:tbl>' +
    '<w:tblPr><w:tblW w:w="5853" w:type="pct"/><w:tblInd w:w="-714" w:type="dxa"/>' +
    borders +
    '<w:tblLayout w:type="fixed"/>' +
    '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>' +
    '</w:tblPr>' +
    grid +
    parentRow +
    headerRow +
    bodyRows +
    '</w:tbl>'
  )
}

// Half-width display units: CJK/fullwidth = 2, everything else = 1.
function isWide(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x115f) ||
    (cp >= 0x2e80 && cp <= 0x303e) ||
    (cp >= 0x3041 && cp <= 0x33ff) ||
    (cp >= 0x3400 && cp <= 0x4dbf) ||
    (cp >= 0x4e00 && cp <= 0x9fff) ||
    (cp >= 0xa000 && cp <= 0xa4cf) ||
    (cp >= 0xac00 && cp <= 0xd7a3) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xfe30 && cp <= 0xfe4f) ||
    (cp >= 0xff00 && cp <= 0xff60) ||
    (cp >= 0xffe0 && cp <= 0xffe6) ||
    (cp >= 0x20000 && cp <= 0x3fffd)
  )
}

function displayWidth(s: string): number {
  let w = 0
  for (const ch of s) {
    w += isWide(ch.codePointAt(0)!) ? 2 : 1
  }
  return w
}

// Trim at most `n` half-width space units from the start of `text`.
// Regular space = 1 unit, ideographic space (U+3000) = 2 units.
// Returns [trimmedText, unitsConsumed].
function trimLeadingSpaces(text: string, n: number): [string, number] {
  let rem = n, i = 0
  while (i < text.length && rem > 0) {
    if (text[i] === ' ') { rem--; i++ }
    else if (text[i] === '　' && rem >= 2) { rem -= 2; i++ }
    else break
  }
  return [text.slice(i), n - rem]
}

// Replace every {{name}} / 【name】 (allowing inner whitespace) with the value.
// When the replacement is wider than the placeholder (display-width comparison),
// the same number of space units is trimmed from the text that immediately follows,
// so alignment spaces in the template don't push the rest of the line off-screen.
export function replaceText(xml: string, values: Record<string, string>): string {
  let out = repairPlaceholders(xml)
  for (const [key, value] of Object.entries(values)) {
    const pattern = keyPattern(key, false)
    if (!new RegExp(pattern).test(out)) continue
    const valueXml = escapeXml(value)
    const valueWidth = displayWidth(value)
    // Replace in runs; embed \x00TRIM:N\x00 after the value when it is wider.
    out = out.replace(/<w:r\b[\s\S]*?<\/w:r>/g, (run) => {
      const m = new RegExp(pattern).exec(run)
      if (!m) return run
      const delta = valueWidth - displayWidth(m[0])
      const marker = delta > 0 ? `\x00TRIM:${delta}\x00` : ''
      return run
        .replace(new RegExp(pattern, 'g'), valueXml + marker)
        .replace(/<w:highlight\b[^>]*\/>/g, '')
        .replace(/<w:shd\b[^>]*\/>/g, '')
    })
    out = out.replace(new RegExp(pattern, 'g'), valueXml)
    // Pass 1: trim spaces that immediately follow the marker in the same text node.
    // If nothing is consumed (no spaces there), the marker is re-emitted for pass 2.
    out = out.replace(/\x00TRIM:(\d+)\x00([ 　]*)/g, (_, nStr, spaces) => {
      const [trimmed, consumed] = trimLeadingSpaces(spaces, Number(nStr))
      const remaining = Number(nStr) - consumed
      return (remaining > 0 ? `\x00TRIM:${remaining}\x00` : '') + trimmed
    })
    // Pass 2: marker is at the end of a text node; trim spaces from the next text node.
    out = out.replace(/\x00TRIM:(\d+)\x00(<\/w:t>[\s\S]*?<w:t[^>]*>)([ 　]*)/g,
      (_, nStr, between, spaces) => between + trimLeadingSpaces(spaces, Number(nStr))[0],
    )
    out = out.replace(/\x00TRIM:\d+\x00/g, '')
  }
  return out
}

// Find the start offset of the <w:p ...> element that encloses `idx`.
function paragraphStart(xml: string, idx: number): number {
  const re = /<w:p(?=[ >/])/g
  let last = -1
  let m: RegExpExecArray | null
  while ((m = re.exec(xml)) && m.index < idx) last = m.index
  return last
}

function paragraphEnd(xml: string, idx: number): number {
  const closeIdx = xml.indexOf('</w:p>', idx)
  return closeIdx === -1 ? idx : closeIdx + '</w:p>'.length
}

function previousParagraphStart(xml: string, idx: number): number {
  const before = xml.slice(0, idx)
  const lastClose = before.lastIndexOf('</w:p>')
  if (lastClose === -1) return -1
  return paragraphStart(xml, lastClose)
}

function visibleText(xml: string): string {
  return [...xml.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)]
    .map((m) => m[1])
    .join('')
    .replace(/\s+/g, '')
}

function removePageBreaks(xml: string): string {
  return xml
    .replace(/<w:br w:type="page"\/>/g, '')
    .replace(/<w:lastRenderedPageBreak\/>/g, '')
    .replace(/<w:pageBreakBefore\/>/g, '')
}

function skipFollowingBlankParagraphs(xml: string, idx: number): number {
  let cur = idx
  while (xml.startsWith('<w:p', cur)) {
    const end = paragraphEnd(xml, cur)
    const para = xml.slice(cur, end)
    if (visibleText(para) !== '') break
    if (/<w:sectPr[\s\S]*?<\/w:sectPr>/.test(para)) break
    cur = end
  }
  return cur
}

// The document's primary section properties (the body-final <w:sectPr>), used as
// the template for the page size / margins / headers of generated sections.
function primarySectPr(xml: string): string | undefined {
  const all = xml.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)
  return all ? all[all.length - 1] : undefined
}

// True if the section that encloses position `fromIdx` is already landscape,
// i.e. the next <w:sectPr> at/after that point is landscape. When the template
// already puts the appendix in a landscape section we must not add our own
// breaks (that would split the section and push the table onto an extra page).
function enclosingSectIsLandscape(xml: string, fromIdx: number): boolean {
  const m = xml.slice(fromIdx).match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)
  return m ? /w:orient="landscape"/.test(m[0]) : false
}

// Turn a <w:pgSz .../> into portrait (short side as width) or landscape.
function orientPgSz(pgSz: string, landscape: boolean): string {
  const w = Number((pgSz.match(/w:w="(\d+)"/) || [])[1])
  const h = Number((pgSz.match(/w:h="(\d+)"/) || [])[1])
  if (!w || !h) {
    return landscape
      ? '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>'
      : '<w:pgSz w:w="11906" w:h="16838"/>'
  }
  const long = Math.max(w, h)
  const short = Math.min(w, h)
  return landscape
    ? `<w:pgSz w:w="${long}" w:h="${short}" w:orient="landscape"/>`
    : `<w:pgSz w:w="${short}" w:h="${long}"/>`
}

// Build a <w:sectPr> in the given orientation, cloning page margins / header &
// footer references from the document's primary section so the rest of the page
// setup stays consistent. Element order follows the OOXML schema.
function buildSectPr(base: string | undefined, landscape: boolean): string {
  let pgSz = orientPgSz('', landscape)
  let pgMar = '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
  let refs = ''
  let cols = ''
  let pgNumType = ''
  if (base) {
    const sz = base.match(/<w:pgSz[^>]*\/>/)
    if (sz) pgSz = orientPgSz(sz[0], landscape)
    const mar = base.match(/<w:pgMar[^>]*\/>/)
    if (mar) pgMar = mar[0]
    refs = (base.match(/<w:(?:header|footer)Reference[^>]*\/>/g) || []).join('')
    const pgn = base.match(/<w:pgNumType[^>]*\/>/)
    if (pgn) pgNumType = pgn[0]
    const c = base.match(/<w:cols[^>]*\/>/)
    if (c) cols = c[0]
  }
  if (landscape) pgMar = setAppendixLandscapeMargins(pgMar)
  return `<w:sectPr>${refs}<w:type w:val="nextPage"/>${pgSz}${pgMar}${pgNumType}${cols}</w:sectPr>`
}

function setAppendixLandscapeMargins(pgMar: string): string {
  // Word stores margins in twips: 1 cm ~= 567 twips.
  // Appendix A landscape pages use left 1.75cm and right 3.8cm.
  return pgMar
    .replace(/\bw:left="[^"]*"/, 'w:left="992"')
    .replace(/\bw:right="[^"]*"/, 'w:right="2155"')
}

function normalizeLandscapeMargins(xml: string): string {
  return xml.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/g, (sect) => {
    if (!/w:orient="landscape"/.test(sect)) return sect
    if (/<w:pgMar[^>]*\/>/.test(sect)) {
      return sect.replace(/<w:pgMar[^>]*\/>/, (pgMar) => setAppendixLandscapeMargins(pgMar))
    }
    return sect.replace(/(<w:pgSz[^>]*\/>)/, '$1<w:pgMar w:top="1440" w:right="2155" w:bottom="1440" w:left="992" w:header="720" w:footer="720" w:gutter="0"/>')
  })
}

function replaceAutoLocatedAppendixTable(xml: string, key: string, tableXml: string): string {
  const sects = [...xml.matchAll(/<w:sectPr[\s\S]*?<\/w:sectPr>/g)]
  const fallbackCandidates: Array<{ start: number; end: number }> = []
  const allAppendixCandidates: Array<{ start: number; end: number }> = []
  const normalizedKey = key.replace(/\s+/g, '')
  const keyIdx = visibleText(xml).lastIndexOf(normalizedKey)
  for (let i = 0; i < sects.length; i++) {
    const sect = sects[i]
    if (!/w:orient="landscape"/.test(sect[0])) continue

    const prev = i > 0 ? sects[i - 1] : undefined
    const segStart = prev ? paragraphEnd(xml, prev.index) : 0
    const segEnd = sect.index
    const segment = xml.slice(segStart, segEnd)
    const tables = [...segment.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)]
    for (const table of tables) {
      const colCount = (table[0].match(/<w:gridCol/g) || []).length
      const looksLikeAppendixTable = /<w:vMerge/.test(table[0]) || colCount >= 8
      if (!looksLikeAppendixTable) continue

      const start = segStart + table.index
      const end = start + table[0].length
      fallbackCandidates.push({ start, end })

      const beforeTableText = visibleText(segment.slice(0, table.index))
      if (!beforeTableText.includes(normalizedKey)) continue

      return xml.slice(0, start) + tableXml + xml.slice(end)
    }
  }
  if (fallbackCandidates.length === 1) {
    const [{ start, end }] = fallbackCandidates
    return xml.slice(0, start) + tableXml + xml.slice(end)
  }

  // Last-resort fallback for templates whose title/section boundary was edited:
  // replace the first appendix-like table that appears after the last visible
  // "附录A" reference in the document. This still avoids creating new sections.
  const tables = [...xml.matchAll(/<w:tbl>[\s\S]*?<\/w:tbl>/g)]
  for (const table of tables) {
    const colCount = (table[0].match(/<w:gridCol/g) || []).length
    const looksLikeAppendixTable = /<w:vMerge/.test(table[0]) || colCount >= 8
    if (!looksLikeAppendixTable) continue
    const beforeVisible = visibleText(xml.slice(0, table.index))
    if (keyIdx !== -1 && beforeVisible.lastIndexOf(normalizedKey) === -1) continue
    allAppendixCandidates.push({ start: table.index, end: table.index + table[0].length })
  }
  if (allAppendixCandidates.length === 1) {
    const [{ start, end }] = allAppendixCandidates
    return xml.slice(0, start) + tableXml + xml.slice(end)
  }

  const sectionSummary = sects
    .map((s, i) => `${i}:${/w:orient="landscape"/.test(s[0]) ? '横' : '纵'}`)
    .join(',')
  throw new Error(
    `未找到横向${key}客户名单表，请检查模板结构（分节=${sectionSummary || '无'}，横向候选表=${fallbackCandidates.length}，全文候选表=${allAppendixCandidates.length}）`,
  )
}

function removeMarkerParagraph(xml: string, key: string, requireHash = true): string {
  const re = new RegExp(keyPattern(key, requireHash))
  const m = re.exec(xml)
  if (!m) return xml
  const pStart = paragraphStart(xml, m.index)
  if (pStart === -1) return xml
  const closeIdx = xml.indexOf('</w:p>', m.index)
  if (closeIdx === -1) return xml
  const pEnd = closeIdx + '</w:p>'.length
  const para = xml.slice(pStart, pEnd)
  const sect = para.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)
  const replacement = sect ? `<w:p><w:pPr>${sect[0]}</w:pPr></w:p>` : ''
  return xml.slice(0, pStart) + replacement + xml.slice(pEnd)
}

// Replace the paragraph containing {{#key}} with the table (tables can't live
// inside a paragraph), followed by an empty paragraph so Word keeps them apart.
// When `landscape` is set, the table is isolated in its own landscape section
// (a portrait break before it, a landscape break after it) so only the appendix
// page is rotated while every other page keeps its original orientation.
export function insertTableAtMarker(
  xml: string,
  key: string,
  tableXml: string,
  requireHash = true,
  landscape = false,
): string {
  const re = new RegExp(keyPattern(key, requireHash))
  const m = re.exec(xml)
  if (!m) return xml
  const pStart = paragraphStart(xml, m.index)
  if (pStart === -1) return xml
  const closeIdx = xml.indexOf('</w:p>', m.index)
  if (closeIdx === -1) return xml
  const pEnd = closeIdx + '</w:p>'.length

  // Only fabricate a landscape section when the marker is NOT already inside a
  // landscape section. If the template already set the appendix region to
  // landscape, fall through and insert in place so the page layout matches the
  // template exactly (no extra section breaks, no extra page).
  if (landscape && !enclosingSectIsLandscape(xml, m.index)) {
    const base = primarySectPr(xml)
    const before = `<w:p><w:pPr>${buildSectPr(base, false)}</w:pPr></w:p>`
    const after = `<w:p><w:pPr>${buildSectPr(base, true)}</w:pPr></w:p>`
    const prevStart = previousParagraphStart(xml, pStart)
    if (prevStart !== -1) {
      const prevPara = xml.slice(prevStart, pStart)
      const prevText = visibleText(prevPara)
      const isAppendixTitle = prevText.startsWith(key.replace(/\s+/g, ''))
      if (isAppendixTitle && !/<w:sectPr[\s\S]*?<\/w:sectPr>/.test(prevPara)) {
        const restStart = skipFollowingBlankParagraphs(xml, pEnd)
        return xml.slice(0, prevStart) + before + removePageBreaks(prevPara) + tableXml + after + xml.slice(restStart)
      }
    }
    const restStart = skipFollowingBlankParagraphs(xml, pEnd)
    return xml.slice(0, pStart) + before + tableXml + after + xml.slice(restStart)
  }

  // If the replaced paragraph carries a section break, keep its <w:sectPr> so
  // page orientation is preserved.
  const para = xml.slice(pStart, pEnd)
  const sect = para.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/)
  const spacer = sect ? `<w:p><w:pPr>${sect[0]}</w:pPr></w:p>` : '<w:p/>'
  return xml.slice(0, pStart) + tableXml + spacer + xml.slice(pEnd)
}

// Collect placeholders from every text-bearing part of the docx.
function scanParts(zip: any): { textKeys: string[]; tableKeys: string[] } {
  const textKeys = new Set<string>()
  const tableKeys = new Set<string>()
  const patterns = DELIMS.map(([o, c, inner]) => new RegExp(`${o}\\s*(#?)\\s*([^${inner}]+?)\\s*${c}`, 'g'))
  for (const name of Object.keys(zip.files)) {
    if (/^word\/(document|header\d*|footer\d*)\.xml$/.test(name)) {
      const xml = repairPlaceholders(zip.file(name).asText())
      for (const re of patterns) {
        re.lastIndex = 0
        let m: RegExpExecArray | null
        while ((m = re.exec(xml))) {
          const key = m[2].trim()
          if (!key) continue
          if (m[1] === '#') tableKeys.add(key)
          else textKeys.add(key)
        }
      }
    }
  }
  return { textKeys: [...textKeys], tableKeys: [...tableKeys] }
}

function fillPattern(pattern: string, row: Record<string, string>): string {
  let out = pattern
  for (const [o, c, inner] of DELIMS) {
    out = out.replace(new RegExp(`${o}\\s*([^${inner}]+?)\\s*${c}`, 'g'), (_, k: string) =>
      String(row[k.trim()] ?? ''),
    )
  }
  return out
}

function safeFileName(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, '_')
}

function isRowEmpty(row: Record<string, string>): boolean {
  return Object.values(row).every((v) => String(v ?? '').trim() === '')
}

function enableUpdateFields(settingsXml: string | undefined): string {
  if (!settingsXml) {
    return '<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:updateFields w:val="true"/></w:settings>'
  }
  if (/<w:updateFields\b[^>]*\/>/.test(settingsXml)) {
    return settingsXml.replace(/<w:updateFields\b[^>]*\/>/, '<w:updateFields w:val="true"/>')
  }
  return settingsXml.replace('</w:settings>', '<w:updateFields w:val="true"/></w:settings>')
}

function markFieldsDirty(xml: string): string {
  return xml
    .replace(/<w:fldSimple\b(?![^>]*\bw:dirty=)([^>]*)>/g, '<w:fldSimple$1 w:dirty="true">')
    .replace(/<w:fldSimple\b([^>]*?)\bw:dirty="[^"]*"([^>]*)>/g, '<w:fldSimple$1w:dirty="true"$2>')
    .replace(/<w:fldChar\b([^>]*\bw:fldCharType="begin"[^>]*?)(?<!\/)>/g, (m) =>
      /\bw:dirty=/.test(m) ? m.replace(/\bw:dirty="[^"]*"/, 'w:dirty="true"') : m.replace(/>$/, ' w:dirty="true">'),
    )
    .replace(/<w:fldChar\b([^>]*\bw:fldCharType="begin"[^>]*?)\/>/g, (m) =>
      /\bw:dirty=/.test(m) ? m.replace(/\bw:dirty="[^"]*"/, 'w:dirty="true"') : m.replace(/\/>$/, ' w:dirty="true"/>'),
    )
}

function normalizePageNumbering(xml: string): string {
  let seenFirstSection = false
  return xml.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/g, (sect) => {
    if (!seenFirstSection) {
      seenFirstSection = true
      return sect
    }
    return sect.replace(/<w:pgNumType\b[^>]*\/>/g, (tag) => {
      const withoutStart = tag
        .replace(/\s*w:start="[^"]*"/g, '')
        .replace(/\s+/g, ' ')
        .replace(' />', '/>')
      return withoutStart === '<w:pgNumType/>' ? '' : withoutStart
    })
  })
}

// Build one document buffer from the template for a single main-sheet row.
export function buildDocx(
  templateBuffer: Buffer,
  row: Record<string, string>,
  textMap: Record<string, string>,
  tables: TableMapping[],
): Buffer {
  const PizZip = require('pizzip')
  const zip = new PizZip(templateBuffer)

  const values: Record<string, string> = {}
  for (const [key, col] of Object.entries(textMap)) {
    if (col) values[key] = String(row[col] ?? '')
  }

  // 1) Body: text placeholders, then inject tables at their markers.
  let docXml: string = zip.file('word/document.xml').asText()
  docXml = replaceText(docXml, values)
  for (const t of tables) {
    let tableRows = t.rows
    if (t.filterMainColumn && t.filterSheetColumn) {
      const want = String(row[t.filterMainColumn] ?? '').trim()
      tableRows = t.rows.filter((r) => String(r[t.filterSheetColumn!] ?? '').trim() === want)
    }
    const tableXml = buildTableXml(t.headers, tableRows, t.mergeColumns, t.parentHeaders)
    if (t.autoLocate === true) {
      const beforeAutoLocate = docXml
      try {
        docXml = replaceAutoLocatedAppendixTable(docXml, t.key, tableXml)
        docXml = removeMarkerParagraph(docXml, t.key, t.requireHash !== false)
      } catch (err) {
        docXml = insertTableAtMarker(
          beforeAutoLocate,
          t.key,
          tableXml,
          t.requireHash !== false,
          t.landscape === true,
        )
        if (docXml === beforeAutoLocate) throw err
      }
      continue
    }

    const beforeInsert = docXml
    docXml = insertTableAtMarker(
      docXml,
      t.key,
      tableXml,
      t.requireHash !== false,
      t.landscape === true,
    )
    if (docXml === beforeInsert) continue
  }
  zip.file('word/document.xml', docXml)

  // 2) Headers / footers: text placeholders only.
  for (const name of Object.keys(zip.files)) {
    if (/^word\/(header|footer)\d*\.xml$/.test(name)) {
      const xml = zip.file(name).asText()
      const replaced = HAS_PLACEHOLDER.test(xml) ? replaceText(xml, values) : xml
      zip.file(name, markFieldsDirty(replaced))
    }
  }

  const settings = zip.file('word/settings.xml')?.asText()
  zip.file('word/settings.xml', enableUpdateFields(settings))
  zip.file(
    'word/document.xml',
    normalizePageNumbering(normalizeLandscapeMargins(markFieldsDirty(zip.file('word/document.xml').asText()))),
  )

  return zip.generate({ type: 'nodebuffer', compression: 'DEFLATE' })
}

export interface PreviewPayload {
  templateBuffer: number[]
  row: Record<string, string>
  textMap: Record<string, string>
  tables: TableMapping[]
  fileNamePattern: string
}

export function setupContractHandlers() {
  // Render a single row and open the file directly for previewing.
  ipcMain.handle('preview-contract', async (_, payload: PreviewPayload) => {
    try {
      const buf = buildDocx(Buffer.from(payload.templateBuffer), payload.row, payload.textMap, payload.tables)
      const dir = path.join(app.getPath('temp'), 'vh-contract-preview')
      mkdirSync(dir, { recursive: true })
      const base = safeFileName(fillPattern(payload.fileNamePattern, payload.row).trim() || '预览')
      const file = path.join(dir, `${base}.docx`)
      writeFileSync(file, buf)
      shell.openPath(file)
      return { success: true, file }
    } catch (err) {
      return { success: false, error: err instanceof Error ? err.message : String(err) }
    }
  })

  ipcMain.handle('scan-placeholders', async (_, templateBuffer: number[]) => {
    try {
      const PizZip = require('pizzip')
      const zip = new PizZip(Buffer.from(templateBuffer))
      const { textKeys, tableKeys } = scanParts(zip)
      return { success: true, textKeys, tableKeys }
    } catch (err) {
      return { success: false, error: err instanceof Error ? err.message : String(err) }
    }
  })

  ipcMain.handle('generate-contracts', async (_, payload: GeneratePayload) => {
    try {
      const dirName = safeFileName(payload.outputDirName?.trim() || '生成文档')
      const dir = path.join(app.getPath('downloads'), dirName)
      mkdirSync(dir, { recursive: true })

      const templateBuffer = Buffer.from(payload.templateBuffer)
      const rows = payload.rows.filter((r) => !isRowEmpty(r))
      const files: string[] = []
      const used = new Set<string>()
      rows.forEach((row, i) => {
        const buf = buildDocx(templateBuffer, row, payload.textMap, payload.tables)
        let base = safeFileName(fillPattern(payload.fileNamePattern, row).trim() || `文档-${i + 1}`)
        // Avoid collisions when the naming column repeats or is blank.
        let fileName = `${base}.docx`
        let n = 1
        while (used.has(fileName)) fileName = `${base}(${++n}).docx`
        used.add(fileName)
        writeFileSync(path.join(dir, fileName), buf)
        files.push(fileName)
      })
      shell.openPath(dir)
      return { success: true, dir, files }
    } catch (err) {
      return { success: false, error: err instanceof Error ? err.message : String(err) }
    }
  })
}
