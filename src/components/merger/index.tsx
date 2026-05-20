import { useState, useCallback, useRef } from 'react'
import * as XLSX from 'xlsx'
import PizZip from 'pizzip'
import Docxtemplater from 'docxtemplater'
import mammoth from 'mammoth'

interface ExcelData {
  headers: string[]
  rows: Record<string, string>[]
}

export default function Merger() {
  const [excelData, setExcelData] = useState<ExcelData | null>(null)
  const [excelFileName, setExcelFileName] = useState('')
  const [wordBuffer, setWordBuffer] = useState<ArrayBuffer | null>(null)
  const [wordFileName, setWordFileName] = useState('')
  const [selectedRow, setSelectedRow] = useState(0)
  const [previewHtml, setPreviewHtml] = useState<string | null>(null)
  const [isProcessing, setIsProcessing] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [saveSuccess, setSaveSuccess] = useState<string | null>(null)

  const excelInputRef = useRef<HTMLInputElement>(null)
  const wordInputRef = useRef<HTMLInputElement>(null)

  const handleExcelFile = useCallback((file: File) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = e.target?.result as ArrayBuffer
        const workbook = XLSX.read(data, { type: 'array' })
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
        const jsonRows = XLSX.utils.sheet_to_json<Record<string, string>>(firstSheet, { defval: '' })
        if (jsonRows.length === 0) {
          setError('Excel 文件中没有数据，请检查文件内容')
          return
        }
        const headers = Object.keys(jsonRows[0])
        setExcelData({ headers, rows: jsonRows })
        setExcelFileName(file.name)
        setSelectedRow(0)
        setPreviewHtml(null)
        setSaveSuccess(null)
        setError(null)
      } catch {
        setError('无法读取 Excel 文件，请确保文件格式正确（.xlsx 或 .xls）')
      }
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const handleWordFile = useCallback((file: File) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      setWordBuffer(e.target?.result as ArrayBuffer)
      setWordFileName(file.name)
      setPreviewHtml(null)
      setSaveSuccess(null)
      setError(null)
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const handleFileDrop = useCallback(
    (type: 'excel' | 'word', e: React.DragEvent) => {
      e.preventDefault()
      const file = e.dataTransfer.files[0]
      if (!file) return
      if (type === 'excel') {
        if (!file.name.match(/\.(xlsx|xls)$/i)) {
          setError('请拖入 Excel 文件（.xlsx 或 .xls）')
          return
        }
        handleExcelFile(file)
      } else {
        if (!file.name.match(/\.docx$/i)) {
          setError('请拖入 Word 文档（.docx）')
          return
        }
        handleWordFile(file)
      }
    },
    [handleExcelFile, handleWordFile],
  )

  const buildMergedBuffer = useCallback(
    async (rowData: Record<string, string>): Promise<ArrayBuffer> => {
      const zip = new PizZip(wordBuffer!)
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        nullGetter: () => '',
      })
      doc.render(rowData)
      return doc.getZip().generate({ type: 'arraybuffer' }) as ArrayBuffer
    },
    [wordBuffer],
  )

  const handlePreview = async () => {
    if (!excelData || !wordBuffer) return
    setIsProcessing(true)
    setError(null)
    setSaveSuccess(null)
    try {
      const merged = await buildMergedBuffer(excelData.rows[selectedRow])
      const result = await mammoth.convertToHtml({ arrayBuffer: merged })
      setPreviewHtml(result.value || '<p style="color:#999">（文档内容为空）</p>')
    } catch (err) {
      setError(`合并失败：${err instanceof Error ? err.message : '请检查 Word 模板格式是否正确'}`)
    } finally {
      setIsProcessing(false)
    }
  }

  const handleSave = async () => {
    if (!excelData || !wordBuffer) return
    setIsProcessing(true)
    setError(null)
    setSaveSuccess(null)
    try {
      const merged = await buildMergedBuffer(excelData.rows[selectedRow])
      const buffer = Array.from(new Uint8Array(merged))
      const baseName = wordFileName.replace(/\.docx$/i, '')
      const result = await window.ipcRenderer.invoke('save-docx', {
        defaultName: `${baseName}_filled.docx`,
        buffer,
      })
      if (result.success) {
        setSaveSuccess(`文件已保存至：${result.filePath}`)
      }
    } catch (err) {
      setError(`保存失败：${err instanceof Error ? err.message : '未知错误'}`)
    } finally {
      setIsProcessing(false)
    }
  }

  const canMerge = excelData !== null && wordBuffer !== null

  return (
    <div className="p-6 min-h-full">
      {/* Tool Header */}
      <div className="mb-6">
        <h1 className="text-lg font-semibold text-gray-900 dark:text-gray-100">
          Excel → Word 数据填充
        </h1>
        <p className="text-gray-500 text-sm mt-1">
          在 Word 模板中使用{' '}
          <code className="bg-gray-100 dark:bg-gray-800 text-indigo-600 dark:text-indigo-400 px-1.5 py-0.5 rounded text-xs font-mono">
            {'{列名}'}
          </code>{' '}
          作为占位符，程序将自动替换为 Excel 对应列的数据
        </p>
      </div>

      {/* Upload Cards */}
      <div className="grid grid-cols-2 gap-4 mb-5">
        <UploadCard
          step={1}
          title="Excel 数据文件"
          hint=".xlsx  /  .xls"
          accept=".xlsx,.xls"
          done={excelData !== null}
          fileName={excelFileName}
          subText={excelData ? `${excelData.rows.length} 行数据，${excelData.headers.length} 列` : undefined}
          inputRef={excelInputRef}
          onFileSelect={handleExcelFile}
          onDrop={(e) => handleFileDrop('excel', e)}
        />
        <UploadCard
          step={2}
          title="Word 模板文件"
          hint=".docx"
          accept=".docx"
          done={wordBuffer !== null}
          fileName={wordFileName}
          subText={wordBuffer ? '模板已加载，可开始合并' : undefined}
          inputRef={wordInputRef}
          onFileSelect={handleWordFile}
          onDrop={(e) => handleFileDrop('word', e)}
        />
      </div>

      {/* Excel Data Table */}
      {excelData && (
        <div className="bg-white dark:bg-gray-900 border border-gray-200 dark:border-gray-800 rounded-xl mb-5 overflow-hidden">
          <div className="flex items-center justify-between px-4 py-3 border-b border-gray-200 dark:border-gray-800">
            <div className="flex items-center gap-2">
              <StepBadge n={3} done={false} />
              <span className="text-sm font-medium text-gray-800 dark:text-gray-200">
                选择要填入的数据行
              </span>
            </div>
            <span className="text-xs text-gray-500 bg-gray-100 dark:bg-gray-800 px-2.5 py-1 rounded-full">
              已选第 {selectedRow + 1} 行
            </span>
          </div>
          <div className="overflow-auto max-h-52">
            <table className="w-full text-sm">
              <thead>
                <tr className="text-gray-400 dark:text-gray-600 text-xs">
                  <th className="py-2 pl-4 pr-2 text-left font-normal w-10">#</th>
                  {excelData.headers.map((h) => (
                    <th key={h} className="py-2 px-3 text-left font-normal whitespace-nowrap">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {excelData.rows.map((row, i) => (
                  <tr
                    key={i}
                    onClick={() => {
                      setSelectedRow(i)
                      setPreviewHtml(null)
                    }}
                    className={`border-t border-gray-100 dark:border-gray-800/60 cursor-pointer transition-colors ${
                      i === selectedRow
                        ? 'bg-indigo-50 dark:bg-indigo-950/50 text-gray-900 dark:text-gray-100'
                        : 'text-gray-600 dark:text-gray-400 hover:bg-gray-50 dark:hover:bg-gray-800/50'
                    }`}
                  >
                    <td className="py-2.5 pl-4 pr-2">
                      <span
                        className={`inline-flex w-5 h-5 rounded-full items-center justify-center text-xs border ${
                          i === selectedRow
                            ? 'bg-indigo-500 border-indigo-500 text-white'
                            : 'border-gray-300 dark:border-gray-700 text-gray-400 dark:text-gray-600'
                        }`}
                      >
                        {i === selectedRow ? '✓' : i + 1}
                      </span>
                    </td>
                    {excelData.headers.map((h) => (
                      <td
                        key={h}
                        className="py-2.5 px-3 max-w-[180px] overflow-hidden text-ellipsis whitespace-nowrap"
                      >
                        {String(row[h] ?? '')}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* Error / Success */}
      {error && (
        <div className="flex items-start gap-2 bg-red-50 dark:bg-red-950/40 border border-red-200 dark:border-red-500/30 text-red-600 dark:text-red-300 rounded-lg px-4 py-3 mb-4 text-sm">
          <span className="mt-0.5 flex-shrink-0">⚠</span>
          <span>{error}</span>
        </div>
      )}
      {saveSuccess && (
        <div className="flex items-center gap-2 bg-green-50 dark:bg-green-950/40 border border-green-200 dark:border-green-500/30 text-green-600 dark:text-green-300 rounded-lg px-4 py-3 mb-4 text-sm">
          <span className="flex-shrink-0">✓</span>
          <span>{saveSuccess}</span>
        </div>
      )}

      {/* Actions */}
      <div className="flex items-center gap-3 mb-5">
        <button
          onClick={handlePreview}
          disabled={!canMerge || isProcessing}
          className="flex items-center gap-2 px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 disabled:bg-gray-100 dark:disabled:bg-gray-800 disabled:text-gray-400 dark:disabled:text-gray-600 text-white rounded-lg font-medium text-sm transition-colors"
        >
          {isProcessing ? '处理中…' : '预览效果'}
        </button>
        <button
          onClick={handleSave}
          disabled={!canMerge || isProcessing}
          className="flex items-center gap-2 px-5 py-2.5 bg-gray-100 dark:bg-gray-800 hover:bg-gray-200 dark:hover:bg-gray-700 disabled:bg-gray-50 dark:disabled:bg-gray-900 disabled:text-gray-400 dark:disabled:text-gray-700 text-gray-700 dark:text-gray-200 border border-gray-200 dark:border-gray-700 disabled:border-gray-100 dark:disabled:border-gray-800 rounded-lg font-medium text-sm transition-colors"
        >
          保存文档
        </button>
        {!canMerge && (
          <span className="text-gray-400 dark:text-gray-600 text-xs">
            {!excelData && !wordBuffer
              ? '请先上传两个文件'
              : !excelData
                ? '请上传 Excel 文件'
                : '请上传 Word 模板'}
          </span>
        )}
      </div>

      {/* Preview */}
      {previewHtml !== null && (
        <div className="rounded-xl overflow-hidden border border-gray-200 dark:border-gray-700">
          <div className="bg-gray-100 dark:bg-gray-800 px-4 py-2.5 flex items-center gap-1.5 border-b border-gray-200 dark:border-gray-700">
            <span className="w-3 h-3 rounded-full bg-red-400/70" />
            <span className="w-3 h-3 rounded-full bg-yellow-400/70" />
            <span className="w-3 h-3 rounded-full bg-green-400/70" />
            <span className="ml-3 text-gray-500 dark:text-gray-400 text-xs">文档预览</span>
          </div>
          <div className="bg-white overflow-auto max-h-[520px]">
            <div
              className="p-10 text-gray-900 text-sm leading-relaxed"
              dangerouslySetInnerHTML={{ __html: previewHtml }}
            />
          </div>
        </div>
      )}
    </div>
  )
}

interface UploadCardProps {
  step: number
  title: string
  hint: string
  accept: string
  done: boolean
  fileName: string
  subText?: string
  inputRef: React.RefObject<HTMLInputElement>
  onFileSelect: (file: File) => void
  onDrop: (e: React.DragEvent) => void
}

function UploadCard({
  step,
  title,
  hint,
  accept,
  done,
  fileName,
  subText,
  inputRef,
  onFileSelect,
  onDrop,
}: UploadCardProps) {
  return (
    <div
      className={`border rounded-xl p-5 cursor-pointer transition-all group ${
        done
          ? 'border-indigo-300 dark:border-indigo-500/40 bg-indigo-50 dark:bg-indigo-950/20'
          : 'border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-900 hover:border-indigo-300 dark:hover:border-indigo-500/40 hover:bg-gray-50 dark:hover:bg-gray-800/50'
      }`}
      onDrop={onDrop}
      onDragOver={(e) => e.preventDefault()}
      onClick={() => inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accept}
        className="hidden"
        onChange={(e) => e.target.files?.[0] && onFileSelect(e.target.files[0])}
      />
      <div className="flex items-start gap-3">
        <StepBadge n={step} done={done} />
        <div className="flex-1 min-w-0">
          <div className="font-medium text-sm text-gray-800 dark:text-gray-200 mb-1">{title}</div>
          {done ? (
            <>
              <div className="text-indigo-600 dark:text-indigo-400 text-sm font-medium truncate">
                {fileName}
              </div>
              {subText && (
                <div className="text-gray-500 text-xs mt-0.5">{subText}</div>
              )}
            </>
          ) : (
            <>
              <div className="text-gray-400 dark:text-gray-600 text-xs mb-2">{hint}</div>
              <div className="border border-dashed border-gray-300 dark:border-gray-700 group-hover:border-indigo-300 dark:group-hover:border-indigo-500/40 rounded-lg py-4 text-center text-gray-400 dark:text-gray-600 text-xs transition-colors">
                点击选择文件，或拖放到此处
              </div>
            </>
          )}
        </div>
        {done && <span className="text-green-500 dark:text-green-400 text-sm flex-shrink-0 mt-0.5">✓</span>}
      </div>
    </div>
  )
}

function StepBadge({ n, done }: { n: number; done: boolean }) {
  return (
    <span
      className={`flex-shrink-0 w-6 h-6 rounded-full flex items-center justify-center text-xs font-semibold mt-0.5 ${
        done
          ? 'bg-indigo-500 text-white'
          : 'bg-gray-100 dark:bg-gray-800 text-gray-500 dark:text-gray-400 border border-gray-300 dark:border-gray-700'
      }`}
    >
      {done ? '✓' : n}
    </span>
  )
}
