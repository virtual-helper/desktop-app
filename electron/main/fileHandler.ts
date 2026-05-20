import { ipcMain, dialog } from 'electron'
import { writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'

const require = createRequire(import.meta.url)

export function setupFileHandlers() {
  ipcMain.handle('save-docx', async (_, { defaultName, buffer }: { defaultName: string; buffer: number[] }) => {
    const { canceled, filePath } = await dialog.showSaveDialog({
      defaultPath: defaultName,
      filters: [{ name: 'Word 文档', extensions: ['docx'] }],
    })
    if (canceled || !filePath) return { success: false }
    writeFileSync(filePath, Buffer.from(buffer))
    return { success: true, filePath }
  })

  ipcMain.handle('generate-docx', async (_, { html, defaultName }: { html: string; defaultName: string }) => {
    try {
      const HTMLtoDOCX = require('html-to-docx')
      const { canceled, filePath } = await dialog.showSaveDialog({
        defaultPath: defaultName,
        filters: [{ name: 'Word 文档', extensions: ['docx'] }],
      })
      if (canceled || !filePath) return { success: false }
      const buffer = await HTMLtoDOCX(
        `<!DOCTYPE html><html><body>${html}</body></html>`,
        null,
        { table: { row: { cantSplit: true } }, footer: false, pageNumber: false },
      )
      writeFileSync(filePath, buffer)
      return { success: true, filePath }
    } catch (err) {
      return { success: false, error: err instanceof Error ? err.message : String(err) }
    }
  })
}
