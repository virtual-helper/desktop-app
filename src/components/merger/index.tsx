import { useState, useCallback, useRef } from 'react'
import { useEditor, EditorContent } from '@tiptap/react'
import StarterKit from '@tiptap/starter-kit'
import Underline from '@tiptap/extension-underline'
import { Mark, mergeAttributes } from '@tiptap/core'
import * as XLSX from 'xlsx'
import mammoth from 'mammoth'

// Custom mark: renders {field} placeholders as styled chips
const FieldMark = Mark.create({
  name: 'field',
  inclusive: false,
  parseHTML() {
    return [{ tag: 'span.ph' }]
  },
  renderHTML({ HTMLAttributes }) {
    return ['span', mergeAttributes(HTMLAttributes, { class: 'ph' }), 0]
  },
})

interface ExcelData {
  headers: string[]
  rows: Record<string, string>[]
}

type Mode = 'upload' | 'edit' | 'preview'

export default function Merger() {
  const [mode, setMode] = useState<Mode>('upload')
  const [excelData, setExcelData] = useState<ExcelData | null>(null)
  const [excelFileName, setExcelFileName] = useState('')
  const [wordFileName, setWordFileName] = useState('')
  const [selectedRow, setSelectedRow] = useState(0)
  const [previewHtml, setPreviewHtml] = useState('')
  const [isProcessing, setIsProcessing] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const [saveSuccess, setSaveSuccess] = useState<string | null>(null)

  const excelInputRef = useRef<HTMLInputElement>(null)
  const wordInputRef = useRef<HTMLInputElement>(null)

  const editor = useEditor({
    extensions: [StarterKit, Underline, FieldMark],
    content: '',
    editorProps: {
      attributes: { class: 'outline-none' },
    },
  })

  const handleExcelFile = useCallback((file: File) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const workbook = XLSX.read(e.target?.result as ArrayBuffer, { type: 'array' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json<Record<string, string>>(sheet, { defval: '' })
        if (!rows.length) { setError('Excel 文件中没有数据'); return }
        setExcelData({ headers: Object.keys(rows[0]), rows })
        setExcelFileName(file.name)
        setSelectedRow(0)
        setError(null)
      } catch { setError('无法读取 Excel 文件') }
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const handleWordFile = useCallback(async (file: File) => {
    try {
      const arrayBuffer = await file.arrayBuffer()
      const result = await mammoth.convertToHtml(
        { arrayBuffer },
        { styleMap: ['r[underline] => u'] },
      )
      const html = result.value.replace(
        /\{([^}]+)\}/g,
        '<span class="ph">{$1}</span>',
      )
      editor?.commands.setContent(html || '<p></p>')
      setWordFileName(file.name)
      setMode('edit')
      setError(null)
      setSaveSuccess(null)
    } catch { setError('无法读取 Word 文件') }
  }, [editor])

  const insertField = useCallback((name: string) => {
    editor?.chain().focus().insertContent(`<span class="ph">{${name}}</span> `).run()
  }, [editor])

  const buildPreviewHtml = useCallback((rowData: Record<string, string>) => {
    const html = editor?.getHTML() ?? ''
    return html
      .replace(/<span class="ph">\{([^}]+)\}<\/span>/g, (_, k) => String(rowData[k] ?? `{${k}}`))
      .replace(/\{([^}]+)\}/g, (_, k) => String(rowData[k] ?? `{${k}}`))
  }, [editor])

  const handlePreview = useCallback(() => {
    if (!excelData) return
    setPreviewHtml(buildPreviewHtml(excelData.rows[selectedRow]))
    setMode('preview')
  }, [excelData, selectedRow, buildPreviewHtml])

  const handleSaveTemplate = useCallback(async () => {
    setIsProcessing(true)
    setError(null)
    setSaveSuccess(null)
    try {
      const raw = editor?.getHTML() ?? ''
      const html = raw.replace(/<span class="ph">(\{[^}]+\})<\/span>/g, '$1')
      const res = await window.ipcRenderer.invoke('generate-docx', {
        html,
        defaultName: wordFileName.replace(/\.docx$/i, '_template.docx'),
      })
      if (res.success) setSaveSuccess(`模板已保存：${res.filePath}`)
      else if (res.error) setError(res.error)
    } catch (err) {
      setError(err instanceof Error ? err.message : '保存失败')
    } finally { setIsProcessing(false) }
  }, [editor, wordFileName])

  const handleExportPreview = useCallback(async () => {
    setIsProcessing(true)
    setError(null)
    setSaveSuccess(null)
    try {
      const res = await window.ipcRenderer.invoke('generate-docx', {
        html: previewHtml,
        defaultName: wordFileName.replace(/\.docx$/i, '_export.docx'),
      })
      if (res.success) setSaveSuccess(`已导出：${res.filePath}`)
      else if (res.error) setError(res.error)
    } catch (err) {
      setError(err instanceof Error ? err.message : '导出失败')
    } finally { setIsProcessing(false) }
  }, [previewHtml, wordFileName])

  // ─── Upload Mode ────────────────────────────────────────────────────────────
  if (mode === 'upload') {
    return (
      <div className="p-6">
        <div className="mb-6">
          <h1 className="text-lg font-semibold text-gray-900 dark:text-gray-100">
            Excel → Word 数据填充
          </h1>
          <p className="text-gray-500 text-sm mt-1">
            上传 Excel 获取字段列表，上传 Word 模板后自动进入编辑器
          </p>
        </div>

        <div className="grid grid-cols-2 gap-4 mb-4">
          {/* Excel upload */}
          <div
            className={`border rounded-xl p-5 cursor-pointer transition-all group ${
              excelData
                ? 'border-indigo-300 dark:border-indigo-500/40 bg-indigo-50 dark:bg-indigo-950/20'
                : 'border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-900 hover:border-indigo-300 dark:hover:border-indigo-500/40'
            }`}
            onClick={() => excelInputRef.current?.click()}
            onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f) handleExcelFile(f) }}
            onDragOver={(e) => e.preventDefault()}
          >
            <input ref={excelInputRef} type="file" accept=".xlsx,.xls" className="hidden"
              onChange={(e) => e.target.files?.[0] && handleExcelFile(e.target.files[0])} />
            <div className="flex items-start gap-3">
              <StepBadge n={1} done={!!excelData} />
              <div className="flex-1 min-w-0">
                <div className="font-medium text-sm text-gray-800 dark:text-gray-200 mb-1">Excel 数据文件</div>
                {excelData ? (
                  <>
                    <div className="text-indigo-600 dark:text-indigo-400 text-sm font-medium truncate">{excelFileName}</div>
                    <div className="text-gray-500 text-xs mt-0.5">{excelData.rows.length} 行 × {excelData.headers.length} 列</div>
                  </>
                ) : (
                  <>
                    <div className="text-gray-400 dark:text-gray-600 text-xs mb-2">.xlsx / .xls</div>
                    <div className="border border-dashed border-gray-300 dark:border-gray-700 rounded-lg py-4 text-center text-gray-400 dark:text-gray-600 text-xs group-hover:border-indigo-300 dark:group-hover:border-indigo-500/40 transition-colors">
                      点击选择文件，或拖放到此处
                    </div>
                  </>
                )}
              </div>
              {excelData && <span className="text-green-500 dark:text-green-400 text-sm mt-0.5">✓</span>}
            </div>
          </div>

          {/* Word upload */}
          <div
            className="border rounded-xl p-5 cursor-pointer transition-all group border-gray-200 dark:border-gray-700 bg-white dark:bg-gray-900 hover:border-indigo-300 dark:hover:border-indigo-500/40"
            onClick={() => wordInputRef.current?.click()}
            onDrop={(e) => { e.preventDefault(); const f = e.dataTransfer.files[0]; if (f?.name.match(/\.docx$/i)) handleWordFile(f) }}
            onDragOver={(e) => e.preventDefault()}
          >
            <input ref={wordInputRef} type="file" accept=".docx" className="hidden"
              onChange={(e) => e.target.files?.[0] && handleWordFile(e.target.files[0])} />
            <div className="flex items-start gap-3">
              <StepBadge n={2} done={false} />
              <div className="flex-1 min-w-0">
                <div className="font-medium text-sm text-gray-800 dark:text-gray-200 mb-1">Word 模板文件</div>
                <div className="text-gray-400 dark:text-gray-600 text-xs mb-2">.docx</div>
                <div className="border border-dashed border-gray-300 dark:border-gray-700 rounded-lg py-4 text-center text-gray-400 dark:text-gray-600 text-xs group-hover:border-indigo-300 dark:group-hover:border-indigo-500/40 transition-colors">
                  点击选择文件，或拖放到此处
                </div>
              </div>
            </div>
          </div>
        </div>

        {error && <ErrorBanner message={error} />}
      </div>
    )
  }

  // ─── Edit Mode ──────────────────────────────────────────────────────────────
  if (mode === 'edit') {
    return (
      <div className="flex flex-col h-full">
        {/* Toolbar */}
        <div className="flex items-center gap-2 px-4 py-2.5 border-b border-gray-200 dark:border-gray-800 bg-white dark:bg-gray-900 flex-shrink-0">
          <button
            onClick={() => setMode('upload')}
            className="flex items-center gap-1.5 text-gray-500 hover:text-gray-700 dark:hover:text-gray-300 text-sm transition-colors mr-1"
          >
            <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M19 12H5"/><path d="M12 5l-7 7 7 7"/></svg>
            返回
          </button>
          <span className="text-sm text-gray-400 dark:text-gray-600 truncate max-w-[140px]">{wordFileName}</span>
          <div className="w-px h-4 bg-gray-200 dark:bg-gray-700 mx-1" />
          {/* Formatting */}
          <button
            onClick={() => editor?.chain().focus().toggleBold().run()}
            className={`w-7 h-7 rounded flex items-center justify-center text-sm font-bold transition-colors ${editor?.isActive('bold') ? 'bg-indigo-100 dark:bg-indigo-500/20 text-indigo-600 dark:text-indigo-400' : 'text-gray-500 hover:bg-gray-100 dark:hover:bg-gray-800'}`}
          >B</button>
          <button
            onClick={() => editor?.chain().focus().toggleItalic().run()}
            className={`w-7 h-7 rounded flex items-center justify-center text-sm italic transition-colors ${editor?.isActive('italic') ? 'bg-indigo-100 dark:bg-indigo-500/20 text-indigo-600 dark:text-indigo-400' : 'text-gray-500 hover:bg-gray-100 dark:hover:bg-gray-800'}`}
          >I</button>
          <button
            onClick={() => editor?.chain().focus().toggleUnderline().run()}
            className={`w-7 h-7 rounded flex items-center justify-center text-sm underline transition-colors ${editor?.isActive('underline') ? 'bg-indigo-100 dark:bg-indigo-500/20 text-indigo-600 dark:text-indigo-400' : 'text-gray-500 hover:bg-gray-100 dark:hover:bg-gray-800'}`}
          >U</button>
          <div className="flex-1" />
          <button
            onClick={handleSaveTemplate}
            disabled={isProcessing}
            className="px-3.5 py-1.5 text-sm bg-gray-100 dark:bg-gray-800 hover:bg-gray-200 dark:hover:bg-gray-700 text-gray-700 dark:text-gray-200 border border-gray-200 dark:border-gray-700 rounded-lg font-medium transition-colors disabled:opacity-50"
          >
            保存模板
          </button>
          <button
            onClick={handlePreview}
            disabled={!excelData || isProcessing}
            className="px-3.5 py-1.5 text-sm bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg font-medium transition-colors disabled:bg-gray-200 dark:disabled:bg-gray-800 disabled:text-gray-400 dark:disabled:text-gray-600"
          >
            预览效果
          </button>
        </div>

        {/* Status messages */}
        {(error || saveSuccess) && (
          <div className="px-4 py-2 flex-shrink-0">
            {error && <ErrorBanner message={error} />}
            {saveSuccess && <SuccessBanner message={saveSuccess} />}
          </div>
        )}

        {/* Editor + Panel */}
        <div className="flex flex-1 min-h-0">
          {/* TipTap editor */}
          <div className="flex-1 overflow-auto bg-gray-50 dark:bg-gray-950">
            <div className="max-w-3xl mx-auto my-6 bg-white dark:bg-gray-900 shadow-sm rounded-lg border border-gray-200 dark:border-gray-800 min-h-[600px] p-10">
              <EditorContent editor={editor} />
            </div>
          </div>

          {/* Right panel */}
          <div className="w-60 border-l border-gray-200 dark:border-gray-800 bg-white dark:bg-gray-900 flex flex-col overflow-y-auto flex-shrink-0">
            {/* Fields */}
            <div className="p-4 border-b border-gray-200 dark:border-gray-800">
              <div className="text-xs font-medium text-gray-500 dark:text-gray-500 uppercase tracking-wider mb-3">
                可插入字段
              </div>
              {excelData ? (
                <div className="flex flex-wrap gap-1.5">
                  {excelData.headers.map((h) => (
                    <button
                      key={h}
                      onClick={() => insertField(h)}
                      title={`插入 {${h}}`}
                      className="px-2.5 py-1 text-xs bg-indigo-50 dark:bg-indigo-500/15 text-indigo-600 dark:text-indigo-400 border border-indigo-200 dark:border-indigo-500/30 rounded-full hover:bg-indigo-100 dark:hover:bg-indigo-500/25 transition-colors font-medium"
                    >
                      {h}
                    </button>
                  ))}
                </div>
              ) : (
                <div className="text-xs text-gray-400 dark:text-gray-600 flex flex-col items-center gap-2 py-3">
                  <span>请先上传 Excel 文件</span>
                  <button
                    onClick={() => excelInputRef.current?.click()}
                    className="px-3 py-1.5 text-xs bg-gray-100 dark:bg-gray-800 hover:bg-gray-200 dark:hover:bg-gray-700 text-gray-600 dark:text-gray-400 rounded-lg transition-colors"
                  >
                    上传 Excel
                  </button>
                  <input ref={excelInputRef} type="file" accept=".xlsx,.xls" className="hidden"
                    onChange={(e) => e.target.files?.[0] && handleExcelFile(e.target.files[0])} />
                </div>
              )}
            </div>

            {/* Row selector */}
            {excelData && (
              <div className="p-4 flex-1">
                <div className="text-xs font-medium text-gray-500 dark:text-gray-500 uppercase tracking-wider mb-3">
                  预览数据行
                </div>
                <div className="space-y-1">
                  {excelData.rows.map((row, i) => {
                    const preview = excelData.headers.slice(0, 2).map((h) => row[h]).filter(Boolean).join('，')
                    return (
                      <button
                        key={i}
                        onClick={() => setSelectedRow(i)}
                        className={`w-full text-left px-3 py-2 rounded-lg text-xs transition-colors ${
                          i === selectedRow
                            ? 'bg-indigo-50 dark:bg-indigo-500/15 text-indigo-600 dark:text-indigo-400 font-medium'
                            : 'text-gray-600 dark:text-gray-400 hover:bg-gray-50 dark:hover:bg-gray-800'
                        }`}
                      >
                        <span className="text-gray-400 dark:text-gray-600 mr-1.5">第{i + 1}行</span>
                        <span className="truncate">{preview}</span>
                      </button>
                    )
                  })}
                </div>
              </div>
            )}
          </div>
        </div>
      </div>
    )
  }

  // ─── Preview Mode ────────────────────────────────────────────────────────────
  return (
    <div className="flex flex-col h-full">
      {/* Toolbar */}
      <div className="flex items-center gap-3 px-4 py-2.5 border-b border-gray-200 dark:border-gray-800 bg-white dark:bg-gray-900 flex-shrink-0">
        <button
          onClick={() => setMode('edit')}
          className="flex items-center gap-1.5 text-gray-500 hover:text-gray-700 dark:hover:text-gray-300 text-sm transition-colors"
        >
          <svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round"><path d="M19 12H5"/><path d="M12 5l-7 7 7 7"/></svg>
          返回编辑
        </button>
        <span className="text-sm text-gray-400 dark:text-gray-600">
          预览：第 {selectedRow + 1} 行数据
        </span>
        <div className="flex-1" />
        <button
          onClick={handleExportPreview}
          disabled={isProcessing}
          className="px-3.5 py-1.5 text-sm bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg font-medium transition-colors disabled:opacity-50"
        >
          {isProcessing ? '导出中…' : '导出预览'}
        </button>
      </div>

      {(error || saveSuccess) && (
        <div className="px-4 py-2 flex-shrink-0">
          {error && <ErrorBanner message={error} />}
          {saveSuccess && <SuccessBanner message={saveSuccess} />}
        </div>
      )}

      <div className="flex-1 overflow-auto bg-gray-50 dark:bg-gray-950">
        <div className="max-w-3xl mx-auto my-6 bg-white shadow-sm rounded-lg border border-gray-200 min-h-[600px] p-10">
          <div
            className="text-gray-900 text-sm leading-relaxed"
            dangerouslySetInnerHTML={{ __html: previewHtml }}
          />
        </div>
      </div>
    </div>
  )
}

function StepBadge({ n, done }: { n: number; done: boolean }) {
  return (
    <span className={`flex-shrink-0 w-6 h-6 rounded-full flex items-center justify-center text-xs font-semibold mt-0.5 ${
      done ? 'bg-indigo-500 text-white' : 'bg-gray-100 dark:bg-gray-800 text-gray-500 dark:text-gray-400 border border-gray-300 dark:border-gray-700'
    }`}>
      {done ? '✓' : n}
    </span>
  )
}

function ErrorBanner({ message }: { message: string }) {
  return (
    <div className="flex items-start gap-2 bg-red-50 dark:bg-red-950/40 border border-red-200 dark:border-red-500/30 text-red-600 dark:text-red-300 rounded-lg px-4 py-3 text-sm">
      <span className="mt-0.5 flex-shrink-0">⚠</span>
      <span>{message}</span>
    </div>
  )
}

function SuccessBanner({ message }: { message: string }) {
  return (
    <div className="flex items-center gap-2 bg-green-50 dark:bg-green-950/40 border border-green-200 dark:border-green-500/30 text-green-600 dark:text-green-300 rounded-lg px-4 py-3 text-sm">
      <span className="flex-shrink-0">✓</span>
      <span>{message}</span>
    </div>
  )
}
