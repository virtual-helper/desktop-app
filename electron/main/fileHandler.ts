import { ipcMain, dialog } from 'electron'
import { writeFileSync } from 'node:fs'

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
}
