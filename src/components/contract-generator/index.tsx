import { useState, useCallback, useRef, useMemo, useEffect } from 'react'
import * as XLSX from 'xlsx'

interface SheetTable {
  headers: string[]
  rows: Record<string, string>[]
}

export interface ParentHeader {
  text: string
  span: number
}

interface TableConfig {
  sheet: string             // which attachment sheet to insert for this {{#key}}
  headerRow: number         // 0-based row index that holds the column headers
  columns: string[]         // which sheet columns to render, in order
  mergeColumns: string[]    // columns whose consecutive equal values merge vertically
  landscape: boolean        // put this table's page in landscape orientation
  filterMainColumn: string  // optional: data-row column used to filter
  filterSheetColumn: string // optional: attachment column matched against it
}

const NONE = ''
const AUTO_APPENDIX_KEY = '附录A'
const EMPTY_TABLE: SheetTable = { headers: [], rows: [] }

// Normalise Excel number-format strings that XLSX.SSF doesn't recognise
// (e.g. "###,###,##0.00" → "#,##0.00") so we can format them ourselves.
function normalizeNumFmt(z: string): string {
  return z.replace(/#{1,3},#{1,3},#{1,3}0/g, '#,##0')
}

// Format a single cell value to its display string, honouring the cell's
// number format.  `sheet_to_json` with `raw: false` skips cells whose format
// string is unsupported (w = undefined), so we read cells individually and
// fall back to XLSX.SSF.format with a normalised format string.
function cellText(cell: XLSX.CellObject): string {
  if (cell.w !== undefined) return cell.w           // library-provided formatted text
  if (cell.t === 'n' && cell.z) {                    // number with custom format
    try { return XLSX.SSF.format(normalizeNumFmt(cell.z), cell.v as number) }
    catch { /* fall through */ }
  }
  return String(cell.v ?? '')
}

// Read every sheet as a raw matrix; the header row is chosen later, because
// real-world sheets often have title/merged rows above the actual headers.
function readAllSheets(wb: XLSX.WorkBook): WorkbookState {
  const aoa: Record<string, string[][]> = {}
  const merges: Record<string, XLSX.Range[]> = {}
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name]
    if (!ws) continue
    const ref = ws['!ref']
    const rows: string[][] = []
    if (ref) {
      const rng = XLSX.utils.decode_range(ref)
      for (let r = rng.s.r; r <= rng.e.r; r++) {
        const row: string[] = []
        for (let c = rng.s.c; c <= rng.e.c; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })]
          row.push(cell ? cellText(cell) : '')
        }
        rows.push(row)
      }
    }
    aoa[name] = rows
    // !merges keeps absolute sheet coordinates. Subtract the !ref origin so
    // merge coordinates line up with aoa indices (e.g. a sheet starting at A2
    // shifts every row up by one).
    const rng = ref ? XLSX.utils.decode_range(ref) : { s: { r: 0, c: 0 } }
    merges[name] = (ws['!merges'] ?? []).map((m) => ({
      s: { r: m.s.r - rng.s.r, c: m.s.c - rng.s.c },
      e: { r: m.e.r - rng.s.r, c: m.e.c - rng.s.c },
    }))
  }
  const names = wb.SheetNames.filter((n) => aoa[n])
  return { aoa, merges, names, fileName: '' }
}

// Guess the header row: the one (within the first rows) with the most non-empty cells.
function autoHeaderRow(aoa: string[][]): number {
  let best = 0
  let bestCount = -1
  const lim = Math.min(aoa.length, 15)
  for (let i = 0; i < lim; i++) {
    const count = (aoa[i] || []).filter((v) => v.trim() !== '').length
    if (count > bestCount) { bestCount = count; best = i }
  }
  return best
}

function deriveHeaders(aoa: string[][], headerRow: number): string[] {
  const seen: Record<string, number> = {}
  return (aoa[headerRow] || []).map((h, i) => {
    let name = h.trim() || `列${i + 1}`
    if (seen[name] != null) { seen[name] += 1; name = `${name}_${seen[name]}` }
    else seen[name] = 0
    return name
  })
}

// Derive parent header row (row above headerRow) accounting for merged cells.
// Returns null when there is no row above, or when that row is all empty.
function deriveParentHeaders(
  aoa: string[][],
  sheetMerges: XLSX.Range[],
  headerRow: number,
  selectedCols: number[],  // column indices that are actually used
): ParentHeader[] | null {
  const parentRow = headerRow - 1
  if (parentRow < 0) return null
  const raw = aoa[parentRow] ?? []
  // Expand merge info into a colIndex -> { text, span } map for this row.
  // A horizontal merge on row `parentRow` means multiple consecutive columns share one label.
  const spanByCol: Record<number, { text: string; span: number }> = {}
  for (const merge of sheetMerges) {
    if (merge.s.r !== parentRow || merge.e.r !== parentRow) continue
    const text = raw[merge.s.c]?.trim() ?? ''
    for (let c = merge.s.c; c <= merge.e.c; c++) {
      spanByCol[c] = { text, span: merge.e.c - merge.s.c + 1 }
    }
  }
  // For unmerged cells, span = 1
  for (let c = 0; c < raw.length; c++) {
    if (!spanByCol[c]) spanByCol[c] = { text: raw[c]?.trim() ?? '', span: 1 }
  }

  // Build the parent headers only for the selected columns, collapsing adjacent
  // columns that belong to the same merge group into a single entry.
  const result: ParentHeader[] = []
  let i = 0
  while (i < selectedCols.length) {
    const colIdx = selectedCols[i]
    const entry = spanByCol[colIdx] ?? { text: '', span: 1 }
    // Count how many of the selected columns fall within this merge group
    const mergeEnd = colIdx + entry.span - 1
    let count = 1
    while (i + count < selectedCols.length && selectedCols[i + count] <= mergeEnd) count++
    result.push({ text: entry.text, span: count })
    i += count
  }

  if (result.every((p) => p.text === '')) return null
  return result
}
function deriveTable(aoa: string[][] | undefined, headerRow: number): SheetTable {
  if (!aoa) return EMPTY_TABLE
  const headers = deriveHeaders(aoa, headerRow)
  const rows = aoa.slice(headerRow + 1).map((r) => {
    const o: Record<string, string> = {}
    headers.forEach((h, i) => { o[h] = r[i] ?? '' })
    return o
  })
  return { headers, rows }
}

// Score how well two column names line up, for auto-mapping / auto-filter.
function similarity(a: string, b: string): number {
  const x = a.trim()
  const y = b.trim()
  if (!x || !y) return 0
  if (x === y) return 3
  if (x.includes(y) || y.includes(x)) return 2
  return 0
}

interface WorkbookState {
  aoa: Record<string, string[][]>
  merges: Record<string, XLSX.Range[]>
  names: string[]
  fileName: string
}

export default function ContractGenerator() {
  const [data, setData] = useState<WorkbookState | null>(null)
  const [attach, setAttach] = useState<WorkbookState | null>(null)

  const [templateBuffer, setTemplateBuffer] = useState<number[] | null>(null)
  const [templateName, setTemplateName] = useState('')
  const [textKeys, setTextKeys] = useState<string[]>([])
  const [tableKeys, setTableKeys] = useState<string[]>([])
  // Text placeholders the user manually marked as appendix tables (no # in template).
  const [promoted, setPromoted] = useState<string[]>([])
  const [scanning, setScanning] = useState(false)

  const [mainSheet, setMainSheet] = useState('')
  const [dataHeaderRow, setDataHeaderRow] = useState(0)
  const [textMap, setTextMap] = useState<Record<string, string>>({})
  const [tableConfig, setTableConfig] = useState<Record<string, TableConfig>>({})
  const [fileNamePattern, setFileNamePattern] = useState('')
  const [outputDirName, setOutputDirName] = useState('生成文档')

  const [previewIndex, setPreviewIndex] = useState(0)
  const [isPreviewing, setIsPreviewing] = useState(false)
  const [isProcessing, setIsProcessing] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [result, setResult] = useState<string | null>(null)

  const dataRef = useRef<HTMLInputElement>(null)
  const attachRef = useRef<HTMLInputElement>(null)
  const wordRef = useRef<HTMLInputElement>(null)

  const readExcel = useCallback(
    (file: File, set: (w: WorkbookState) => void) => {
      const reader = new FileReader()
      reader.onload = (e) => {
        try {
          const wb = XLSX.read(e.target?.result as ArrayBuffer, { type: 'array', cellNF: true })
          const { aoa, merges, names } = readAllSheets(wb)
          if (!names.length) { setError(`「${file.name}」中没有可用的工作表`); return }
          set({ aoa, merges, names, fileName: file.name })
          setError(null); setResult(null)
        } catch { setError(`无法读取「${file.name}」`) }
      }
      reader.readAsArrayBuffer(file)
    },
    [],
  )

  const handleWord = useCallback(async (file: File) => {
    try {
      const buf = await file.arrayBuffer()
      const arr = Array.from(new Uint8Array(buf))
      setTemplateBuffer(arr)
      setTemplateName(file.name)
      setError(null); setResult(null)
      setScanning(true)
      const res = await window.ipcRenderer.invoke('scan-placeholders', arr)
      if (res.success) {
        setTextKeys(res.textKeys as string[])
        setTableKeys(res.tableKeys as string[])
        setPromoted([])
        if (!res.textKeys.length && !res.tableKeys.length) {
          setError('未在该 Word 模板中检测到任何占位符，请确认变量已用 {{ }} 或 【 】 包裹')
        }
      } else {
        setError(res.error || '扫描占位符失败')
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : '无法读取 Word 模板')
    } finally {
      setScanning(false)
    }
  }, [])

  // Data workbook: pick the main sheet (each of its rows -> one document).
  useEffect(() => {
    if (!data) return
    setMainSheet((prev) => (prev && data.names.includes(prev) ? prev : data.names[0]))
  }, [data])

  // Auto-detect the header row whenever the chosen data sheet changes.
  useEffect(() => {
    if (!data || !mainSheet || !data.aoa[mainSheet]) return
    setDataHeaderRow(autoHeaderRow(data.aoa[mainSheet]))
  }, [data, mainSheet])

  const mainTable = useMemo(
    () => deriveTable(data?.aoa[mainSheet], dataHeaderRow),
    [data, mainSheet, dataHeaderRow],
  )
  const mainHeaders = mainTable.headers

  const autoAppendixKeys = useMemo(
    () => attach && !tableKeys.includes(AUTO_APPENDIX_KEY) ? [AUTO_APPENDIX_KEY] : [],
    [attach, tableKeys],
  )
  // Effective keys after promotion: a promoted text key becomes a table key.
  const effTextKeys = useMemo(
    () => textKeys.filter((k) => !promoted.includes(k) && !autoAppendixKeys.includes(k)),
    [textKeys, promoted, autoAppendixKeys],
  )
  const effTableKeys = useMemo(
    () => [
      ...tableKeys,
      ...promoted.filter((k) => !tableKeys.includes(k)),
      ...autoAppendixKeys.filter((k) => !tableKeys.includes(k) && !promoted.includes(k)),
    ],
    [tableKeys, promoted, autoAppendixKeys],
  )

  // Auto-map text placeholders to same-named (or closest) data columns.
  useEffect(() => {
    if (!effTextKeys.length) { setTextMap({}); return }
    setTextMap(() => {
      const next: Record<string, string> = {}
      for (const key of effTextKeys) {
        let best = NONE
        let bestScore = 0
        for (const h of mainHeaders) {
          const s = similarity(key, h)
          if (s > bestScore) { bestScore = s; best = h }
        }
        next[key] = best
      }
      return next
    })
  }, [effTextKeys, mainHeaders])

  // Auto-configure table placeholders against the attachment workbook. Existing
  // (manually edited) configs are preserved; only new keys get defaults.
  useEffect(() => {
    const attachNames = attach?.names ?? []
    setTableConfig((prev) => {
      const next: Record<string, TableConfig> = {}
      for (const key of effTableKeys) {
        if (prev[key]) { next[key] = prev[key]; continue }
        const sheet = attachNames.includes(key) ? key : attachNames[0] ?? NONE
        const headerRow = sheet && attach ? autoHeaderRow(attach.aoa[sheet]) : 0
        let filterMainColumn = NONE
        let filterSheetColumn = NONE
        const sheetHeaders = sheet && attach ? deriveHeaders(attach.aoa[sheet], headerRow) : []
        let bestScore = 0
        for (const mh of mainHeaders) {
          for (const sh of sheetHeaders) {
            const s = similarity(mh, sh)
            if (s > bestScore) {
              bestScore = s
              filterMainColumn = mh
              filterSheetColumn = sh
            }
          }
        }
        next[key] = { sheet, headerRow, columns: sheetHeaders, mergeColumns: [], landscape: true, filterMainColumn, filterSheetColumn }
      }
      return next
    })
  }, [effTableKeys, attach, mainHeaders])

  // Default file-name rule once the data columns are known.
  useEffect(() => {
    if (!mainHeaders.length) return
    setFileNamePattern((prev) => {
      if (prev) return prev
      const nameCol =
        mainHeaders.find((h) => /名称|名字|乙方|客户|name/i.test(h)) || mainHeaders[0]
      return `文档-{{${nameCol}}}`
    })
  }, [mainHeaders])

  const mainRows = useMemo(
    () => mainTable.rows.filter((r) => !Object.values(r).every((v) => v.trim() === '')),
    [mainTable],
  )

  const unmappedText = useMemo(
    () => effTextKeys.filter((k) => !textMap[k]),
    [effTextKeys, textMap],
  )

  const needAttach = effTableKeys.length > 0
  const ready =
    !!templateBuffer &&
    !!data &&
    !!mainSheet &&
    (effTextKeys.length > 0 || effTableKeys.length > 0) &&
    mainRows.length > 0 &&
    (!needAttach || !!attach)

  const buildTablesPayload = useCallback(() => {
    return effTableKeys
      .map((key) => {
        const cfg = tableConfig[key]
        const sheet = cfg?.sheet
        if (!sheet || !attach?.aoa[sheet]) return null
        const table = deriveTable(attach.aoa[sheet], cfg.headerRow)
        const columns = cfg.columns?.filter((c) => table.headers.includes(c))
        const effCols = columns && columns.length ? columns : table.headers
        // Compute selected column indices (positions in the header row) for parent header span math
        const allHeaders = deriveHeaders(attach.aoa[sheet], cfg.headerRow)
        const selectedColIdxs = effCols.map((c) => allHeaders.indexOf(c)).filter((i) => i >= 0)
        const sheetMerges = attach.merges?.[sheet] ?? []
        const parentHeaders = deriveParentHeaders(attach.aoa[sheet], sheetMerges, cfg.headerRow, selectedColIdxs)
        return {
          key,
          headers: effCols,
          parentHeaders: parentHeaders ?? undefined,
          rows: table.rows,
          requireHash: tableKeys.includes(key), // promoted text placeholders have no #
          mergeColumns: (cfg.mergeColumns ?? []).filter((c) => effCols.includes(c)),
          landscape: cfg.landscape !== false,
          autoLocate: key === AUTO_APPENDIX_KEY && !tableKeys.includes(key),
          filterMainColumn: cfg.filterMainColumn || undefined,
          filterSheetColumn: cfg.filterSheetColumn || undefined,
        }
      })
      .filter(Boolean)
  }, [effTableKeys, tableKeys, tableConfig, attach])

  const handlePreview = useCallback(async () => {
    if (!templateBuffer || !mainRows.length) return
    const row = mainRows[Math.min(previewIndex, mainRows.length - 1)]
    setIsPreviewing(true); setError(null); setResult(null)
    try {
      const res = await window.ipcRenderer.invoke('preview-contract', {
        templateBuffer,
        row,
        textMap,
        tables: buildTablesPayload(),
        fileNamePattern,
      })
      if (res.success) setResult(`已生成预览文件并打开：${res.file}`)
      else if (res.error) setError(res.error)
    } catch (err) {
      setError(err instanceof Error ? err.message : '预览失败')
    } finally { setIsPreviewing(false) }
  }, [templateBuffer, mainRows, previewIndex, textMap, buildTablesPayload, fileNamePattern])

  const handleGenerate = useCallback(async () => {
    if (!templateBuffer || !data || !mainSheet) return
    setIsProcessing(true); setError(null); setResult(null)
    try {
      const tables = buildTablesPayload()

      const res = await window.ipcRenderer.invoke('generate-contracts', {
        templateBuffer,
        rows: mainRows,
        textMap,
        tables,
        fileNamePattern,
        outputDirName,
      })
      if (res.success) setResult(`已生成 ${res.files.length} 份文档（已打开文件夹）：${res.dir}`)
      else if (res.error) setError(res.error)
    } catch (err) {
      setError(err instanceof Error ? err.message : '生成失败')
    } finally { setIsProcessing(false) }
  }, [templateBuffer, data, mainSheet, mainRows, textMap, buildTablesPayload, fileNamePattern, outputDirName])

  return (
    <div className="flex flex-col h-full">
      {/* Toolbar */}
      <div className="flex items-center gap-2.5 px-4 py-2.5 border-b border-gray-200 dark:border-gray-800 bg-white dark:bg-gray-900 flex-shrink-0">
        <h1 className="text-sm font-semibold text-gray-900 dark:text-gray-100">协议批量生成</h1>
        <div className="flex-1" />
        <UploadButton active={!!templateBuffer} onClick={() => wordRef.current?.click()}>
          {scanning ? '扫描中…' : templateBuffer ? `✓ ${templateName}` : '① Word 模板'}
        </UploadButton>
        <input ref={wordRef} type="file" accept=".docx" className="hidden"
          onChange={(e) => e.target.files?.[0] && handleWord(e.target.files[0])} />
        <UploadButton active={!!data} onClick={() => dataRef.current?.click()}>
          {data ? `✓ ${data.fileName}` : '② 数据表 Excel'}
        </UploadButton>
        <input ref={dataRef} type="file" accept=".xlsx,.xls" className="hidden"
          onChange={(e) => e.target.files?.[0] && readExcel(e.target.files[0], setData)} />
        <UploadButton active={!!attach} onClick={() => attachRef.current?.click()}>
          {attach ? `✓ ${attach.fileName}` : '③ 附件表 Excel'}
        </UploadButton>
        <input ref={attachRef} type="file" accept=".xlsx,.xls" className="hidden"
          onChange={(e) => e.target.files?.[0] && readExcel(e.target.files[0], setAttach)} />
        <button
          onClick={handleGenerate}
          disabled={!ready || isProcessing}
          className="px-3.5 py-1.5 text-sm bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg font-medium transition-colors disabled:bg-gray-200 dark:disabled:bg-gray-800 disabled:text-gray-400 dark:disabled:text-gray-600"
        >
          {isProcessing ? '生成中…' : `④ 批量生成 ${mainRows.length || ''} 份`}
        </button>
      </div>

      {(error || result) && (
        <div className="px-4 py-2 flex-shrink-0">
          {error && <Banner kind="error" message={error} />}
          {result && <Banner kind="success" message={result} />}
        </div>
      )}

      {!templateBuffer || !data ? (
        <EmptyState hasWord={!!templateBuffer} hasData={!!data} hasAttach={!!attach} />
      ) : (
        <div className="flex-1 overflow-auto bg-gray-50 dark:bg-gray-950">
          <div className="max-w-4xl mx-auto p-6 space-y-6">
            {/* Output settings */}
            <Section title="生成设置">
              <div className="grid grid-cols-2 gap-4">
                <Field label="数据表工作表（每行生成一份文档）">
                  <div className="flex items-center gap-2">
                    <select value={mainSheet} onChange={(e) => setMainSheet(e.target.value)} className={`${selectCls} flex-1`}>
                      {data.names.map((n) => (
                        <option key={n} value={n}>{n}（{data.aoa[n].length} 行）</option>
                      ))}
                    </select>
                    <HeaderRowInput value={dataHeaderRow} onChange={setDataHeaderRow} />
                  </div>
                </Field>
                <Field label="输出文件夹（位于「下载」目录下）">
                  <input value={outputDirName} onChange={(e) => setOutputDirName(e.target.value)} className={inputCls} />
                </Field>
              </div>
              <Field label="文件名规则（可用 {{列名}} 取值）">
                <input
                  value={fileNamePattern}
                  onChange={(e) => setFileNamePattern(e.target.value)}
                  className={inputCls}
                  placeholder="例如：终止协议-{{乙方}}"
                />
              </Field>
              <Field label="预览（先生成单份并用 Word 打开核对）">
                <div className="flex items-center gap-2">
                  <select
                    value={Math.min(previewIndex, Math.max(mainRows.length - 1, 0))}
                    onChange={(e) => setPreviewIndex(Number(e.target.value))}
                    className={`${selectCls} flex-1`}
                  >
                    {mainRows.map((r, i) => (
                      <option key={i} value={i}>
                        第 {i + 1} 行{mainHeaders[0] ? ` · ${r[mainHeaders[0]]}` : ''}
                      </option>
                    ))}
                  </select>
                  <button
                    onClick={handlePreview}
                    disabled={!templateBuffer || !mainRows.length || isPreviewing}
                    className="px-3 py-1.5 text-xs rounded-lg font-medium border border-indigo-200 dark:border-indigo-500/30 bg-indigo-50 dark:bg-indigo-500/15 text-indigo-600 dark:text-indigo-400 hover:bg-indigo-100 dark:hover:bg-indigo-500/25 transition-colors disabled:opacity-50 flex-shrink-0"
                  >
                    {isPreviewing ? '生成中…' : '预览此行'}
                  </button>
                </div>
              </Field>
            </Section>

            {/* Text placeholders */}
            <Section
              title={`文本占位符（检测到 ${effTextKeys.length} 个${
                unmappedText.length ? `，${unmappedText.length} 个待映射` : '，已全部自动映射'
              }）`}
            >
              {effTextKeys.length === 0 ? (
                <Empty>未检测到文本占位符</Empty>
              ) : (
                <div className="space-y-2">
                  {effTextKeys.map((key) => (
                    <div key={key} className="flex items-center gap-3">
                      <code className="w-52 flex-shrink-0 text-xs px-2 py-1.5 rounded bg-indigo-50 dark:bg-indigo-500/10 text-indigo-600 dark:text-indigo-300 truncate">
                        {`{{${key}}}`}
                      </code>
                      <span className="text-gray-400 text-xs">→</span>
                      <select
                        value={textMap[key] ?? NONE}
                        onChange={(e) => setTextMap((m) => ({ ...m, [key]: e.target.value }))}
                        className={`${selectCls} flex-1 ${textMap[key] ? '' : 'border-amber-300 dark:border-amber-500/50'}`}
                      >
                        <option value={NONE}>— 未映射（留空）—</option>
                        {mainHeaders.map((h) => (
                          <option key={h} value={h}>{h}</option>
                        ))}
                      </select>
                      <button
                        onClick={() => setPromoted((p) => [...p, key])}
                        title="把这个占位符改成插入整张附录表格"
                        className="text-[11px] px-2 py-1.5 rounded-md border border-purple-200 dark:border-purple-500/30 bg-purple-50 dark:bg-purple-500/10 text-purple-600 dark:text-purple-300 hover:bg-purple-100 dark:hover:bg-purple-500/20 transition-colors flex-shrink-0 whitespace-nowrap"
                      >
                        设为附录表格
                      </button>
                    </div>
                  ))}
                </div>
              )}
            </Section>

            {/* Table placeholders (from the attachment workbook) */}
            {effTableKeys.length > 0 && (
              <Section
                title={`表格占位符（${effTableKeys.length} 个）${attach ? '' : ' · 请上传附件表 Excel'}`}
              >
                {!attach ? (
                  <Empty>有表格占位符，请上传附件表 Excel 作为表格数据来源</Empty>
                ) : (
                  <div className="space-y-3">
                    {effTableKeys.map((key) => {
                      const cfg = tableConfig[key] ?? { sheet: NONE, headerRow: 0, columns: [], mergeColumns: [], landscape: true, filterMainColumn: NONE, filterSheetColumn: NONE }
                      const sheetHeaders = cfg.sheet ? deriveHeaders(attach.aoa[cfg.sheet], cfg.headerRow) : []
                      const isAutoLocated = key === AUTO_APPENDIX_KEY && !tableKeys.includes(key)
                      const isPromoted = promoted.includes(key) && !isAutoLocated
                      return (
                        <div key={key} className="border border-gray-200 dark:border-gray-800 rounded-lg p-3 space-y-2">
                          <div className="flex items-center gap-3">
                            <code className="w-52 flex-shrink-0 text-xs px-2 py-1.5 rounded bg-purple-50 dark:bg-purple-500/10 text-purple-600 dark:text-purple-300 truncate">
                              {isAutoLocated ? `自动定位：${key}` : isPromoted ? `{{${key}}}` : `{{#${key}}}`}
                            </code>
                            {isAutoLocated && (
                              <span className="text-[11px] text-amber-600 dark:text-amber-400 flex-shrink-0">
                                横向节内替换静态表
                              </span>
                            )}
                            {isPromoted && (
                              <button
                                onClick={() => setPromoted((p) => p.filter((x) => x !== key))}
                                title="改回普通文本占位符"
                                className="text-[11px] px-1.5 py-1 rounded border border-gray-200 dark:border-gray-700 text-gray-500 hover:bg-gray-100 dark:hover:bg-gray-800 flex-shrink-0"
                              >
                                ↩ 改回文本
                              </button>
                            )}
                            <span className="text-gray-400 text-xs">插入附件表</span>
                            <select
                              value={cfg.sheet}
                              onChange={(e) => {
                                const sheet = e.target.value
                                const headerRow = sheet ? autoHeaderRow(attach.aoa[sheet]) : 0
                                const columns = sheet ? deriveHeaders(attach.aoa[sheet], headerRow) : []
                                setTableConfig((c) => ({ ...c, [key]: { ...cfg, sheet, headerRow, columns, mergeColumns: [], filterSheetColumn: NONE } }))
                              }}
                              className={`${selectCls} flex-1`}
                            >
                              <option value={NONE}>— 不插入 —</option>
                              {attach.names.map((n) => (
                                <option key={n} value={n}>{n}</option>
                              ))}
                            </select>
                            {cfg.sheet && (
                              <HeaderRowInput
                                value={cfg.headerRow}
                                onChange={(v) => setTableConfig((c) => ({ ...c, [key]: { ...cfg, headerRow: v, columns: deriveHeaders(attach.aoa[cfg.sheet], v), mergeColumns: [], filterSheetColumn: NONE } }))}
                              />
                            )}
                            {cfg.sheet && (
                              <button
                                onClick={() => setTableConfig((c) => ({ ...c, [key]: { ...cfg, landscape: !(cfg.landscape !== false) } }))}
                                title="把该表格所在的页面设为横向（其它页面不变）"
                                className={`text-[11px] px-2 py-1.5 rounded-md border transition-colors flex-shrink-0 whitespace-nowrap ${
                                  cfg.landscape !== false
                                    ? 'bg-amber-50 dark:bg-amber-500/15 text-amber-600 dark:text-amber-400 border-amber-200 dark:border-amber-500/30'
                                    : 'bg-gray-50 dark:bg-gray-800 text-gray-500 border-gray-200 dark:border-gray-700 hover:bg-gray-100 dark:hover:bg-gray-700'
                                }`}
                              >
                                {cfg.landscape !== false ? '横向页面 ✓' : '横向页面'}
                              </button>
                            )}
                          </div>
                          {cfg.sheet && sheetHeaders.length > 0 && (
                            <div className="pl-2">
                              <div className="flex items-center justify-between mb-1.5">
                                <span className="text-[11px] text-gray-500">
                                  插入列（已选 {(cfg.columns ?? []).filter((c) => sheetHeaders.includes(c)).length}/{sheetHeaders.length}）
                                </span>
                                <div className="flex gap-2">
                                  <button
                                    onClick={() => setTableConfig((c) => ({ ...c, [key]: { ...cfg, columns: sheetHeaders } }))}
                                    className="text-[11px] text-indigo-600 dark:text-indigo-400 hover:underline"
                                  >全选</button>
                                  <button
                                    onClick={() => setTableConfig((c) => ({ ...c, [key]: { ...cfg, columns: [] } }))}
                                    className="text-[11px] text-gray-500 hover:underline"
                                  >清空</button>
                                </div>
                              </div>
                              <div className="flex flex-wrap gap-1.5">
                                {sheetHeaders.map((h) => {
                                  const on = (cfg.columns ?? []).includes(h)
                                  return (
                                    <button
                                      key={h}
                                      onClick={() =>
                                        setTableConfig((c) => {
                                          const cur = cfg.columns ?? []
                                          const columns = on
                                            ? cur.filter((x) => x !== h)
                                            : sheetHeaders.filter((x) => cur.includes(x) || x === h)
                                          const mergeColumns = (cfg.mergeColumns ?? []).filter((x) => columns.includes(x))
                                          return { ...c, [key]: { ...cfg, columns, mergeColumns } }
                                        })
                                      }
                                      className={`text-[11px] px-2 py-1 rounded-md border transition-colors ${
                                        on
                                          ? 'bg-indigo-50 dark:bg-indigo-500/15 text-indigo-600 dark:text-indigo-400 border-indigo-200 dark:border-indigo-500/30'
                                          : 'bg-gray-50 dark:bg-gray-800 text-gray-500 border-gray-200 dark:border-gray-700 hover:bg-gray-100 dark:hover:bg-gray-700'
                                      }`}
                                    >
                                      {on ? '✓ ' : ''}{h}
                                    </button>
                                  )
                                })}
                              </div>
                            </div>
                          )}
                          {cfg.sheet && (() => {
                            const selectedCols = (cfg.columns ?? []).filter((c) => sheetHeaders.includes(c))
                            return selectedCols.length > 0 ? (
                              <div className="pl-2">
                                <div className="text-[11px] text-gray-500 mb-1.5">合并相同值的列（同列上下相邻值相同则纵向合并单元格）</div>
                                <div className="flex flex-wrap gap-1.5">
                                  {selectedCols.map((h) => {
                                    const on = (cfg.mergeColumns ?? []).includes(h)
                                    return (
                                      <button
                                        key={h}
                                        onClick={() =>
                                          setTableConfig((c) => {
                                            const cur = cfg.mergeColumns ?? []
                                            const mergeColumns = on ? cur.filter((x) => x !== h) : [...cur, h]
                                            return { ...c, [key]: { ...cfg, mergeColumns } }
                                          })
                                        }
                                        className={`text-[11px] px-2 py-1 rounded-md border transition-colors ${
                                          on
                                            ? 'bg-emerald-50 dark:bg-emerald-500/15 text-emerald-600 dark:text-emerald-400 border-emerald-200 dark:border-emerald-500/30'
                                            : 'bg-gray-50 dark:bg-gray-800 text-gray-500 border-gray-200 dark:border-gray-700 hover:bg-gray-100 dark:hover:bg-gray-700'
                                        }`}
                                      >
                                        {on ? '⇕ ' : ''}{h}
                                      </button>
                                    )
                                  })}
                                </div>
                              </div>
                            ) : null
                          })()}
                          {cfg.sheet && (
                            <div className="flex items-center gap-2 pl-2 text-xs text-gray-500 flex-wrap">
                              <span>按</span>
                              <select
                                value={cfg.filterMainColumn}
                                onChange={(e) =>
                                  setTableConfig((c) => ({ ...c, [key]: { ...cfg, filterMainColumn: e.target.value } }))
                                }
                                className={selectCls}
                              >
                                <option value={NONE}>数据表（不过滤）</option>
                                {mainHeaders.map((h) => (
                                  <option key={h} value={h}>数据表.{h}</option>
                                ))}
                              </select>
                              <span>=</span>
                              <select
                                value={cfg.filterSheetColumn}
                                onChange={(e) =>
                                  setTableConfig((c) => ({ ...c, [key]: { ...cfg, filterSheetColumn: e.target.value } }))
                                }
                                className={selectCls}
                              >
                                <option value={NONE}>附件表（全部行）</option>
                                {sheetHeaders.map((h) => (
                                  <option key={h} value={h}>附件表.{h}</option>
                                ))}
                              </select>
                              <span className="text-gray-400">筛选出当前行对应的附件内容</span>
                            </div>
                          )}
                        </div>
                      )
                    })}
                  </div>
                )}
              </Section>
            )}
          </div>
        </div>
      )}
    </div>
  )
}

const selectCls =
  'text-xs px-2 py-1.5 rounded-lg border border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-900 text-gray-700 dark:text-gray-200 focus:outline-none focus:ring-2 focus:ring-indigo-400/40'
const inputCls =
  'w-full text-xs px-2.5 py-1.5 rounded-lg border border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-900 text-gray-700 dark:text-gray-200 focus:outline-none focus:ring-2 focus:ring-indigo-400/40'

function UploadButton({ active, onClick, children }: { active: boolean; onClick: () => void; children: React.ReactNode }) {
  return (
    <button
      onClick={onClick}
      className={`px-3 py-1.5 text-sm rounded-lg font-medium transition-colors border max-w-48 truncate ${
        active
          ? 'bg-indigo-50 dark:bg-indigo-500/15 text-indigo-600 dark:text-indigo-400 border-indigo-200 dark:border-indigo-500/30'
          : 'bg-gray-100 dark:bg-gray-800 text-gray-600 dark:text-gray-300 border-gray-200 dark:border-gray-700 hover:bg-gray-200 dark:hover:bg-gray-700'
      }`}
    >
      {children}
    </button>
  )
}

function Section({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="bg-white dark:bg-gray-900 rounded-xl border border-gray-200 dark:border-gray-800 p-4">
      <div className="text-xs font-semibold text-gray-700 dark:text-gray-300 mb-3">{title}</div>
      {children}
    </div>
  )
}

// 1-based "表头第 N 行" control backed by a 0-based index.
function HeaderRowInput({ value, onChange }: { value: number; onChange: (v: number) => void }) {
  return (
    <label className="flex items-center gap-1 text-[11px] text-gray-500 flex-shrink-0 whitespace-nowrap" title="选择 Excel 中作为列标题的那一行">
      表头第
      <input
        type="number"
        min={1}
        value={value + 1}
        onChange={(e) => onChange(Math.max(0, (Number(e.target.value) || 1) - 1))}
        className="w-12 text-xs px-1.5 py-1.5 rounded-lg border border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-900 text-gray-700 dark:text-gray-200 focus:outline-none focus:ring-2 focus:ring-indigo-400/40"
      />
      行
    </label>
  )
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <label className="block">
      <div className="text-[11px] text-gray-500 mb-1">{label}</div>
      {children}
    </label>
  )
}

function Empty({ children }: { children: React.ReactNode }) {
  return <div className="text-xs text-gray-400 dark:text-gray-600 py-2">{children}</div>
}

function EmptyState({ hasWord, hasData, hasAttach }: { hasWord: boolean; hasData: boolean; hasAttach: boolean }) {
  return (
    <div className="flex-1 flex items-center justify-center text-center p-10">
      <div className="text-gray-400 dark:text-gray-600">
        <div className="text-4xl mb-4">📄</div>
        <div className="text-sm mb-2">请按顺序上传以下文件</div>
        <div className="text-xs space-y-1 text-left inline-block">
          <div>{hasWord ? '✓' : '○'} Word 模板（.docx，变量写成 {'{{变量名}}'} 或 {'【变量名】'}，整张附件表写成 {'{{#名称}}'} 或 {'【#名称】'}）</div>
          <div>{hasData ? '✓' : '○'} 数据表 Excel（每行生成一份，列头作为可填字段）</div>
          <div>{hasAttach ? '✓' : '○'} 附件表 Excel（可选，仅当模板含 {'{{#名称}}'} 时需要）</div>
        </div>
      </div>
    </div>
  )
}

function Banner({ kind, message }: { kind: 'error' | 'success'; message: string }) {
  const styles = kind === 'error'
    ? 'bg-red-50 dark:bg-red-950/40 border-red-200 dark:border-red-500/30 text-red-600 dark:text-red-300'
    : 'bg-green-50 dark:bg-green-950/40 border-green-200 dark:border-green-500/30 text-green-600 dark:text-green-300'
  return (
    <div className={`flex items-start gap-2 border rounded-lg px-4 py-3 text-sm ${styles}`}>
      <span className="mt-0.5 flex-shrink-0">{kind === 'error' ? '⚠' : '✓'}</span>
      <span className="break-all">{message}</span>
    </div>
  )
}
