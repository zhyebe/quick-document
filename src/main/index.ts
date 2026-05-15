import { app, BrowserWindow, dialog, ipcMain, Menu, nativeImage, shell, Tray } from 'electron'
import { join } from 'node:path'
import type { ChatRequest, ChatResponse, WorkspaceSnapshot } from '@shared/types'
import { planDocumentWork } from './services/aiPlanner'
import { executeDocumentAction } from './services/documentService'
import { loadExternalAiConfig } from './services/externalAiConfig'
import { SettingsStore } from './services/settingsStore'
import { getEmbeddedOfficeSkillBrief } from './services/skillRegistry'
import { scanWorkspace } from './services/workspaceFiles'

let mainWindow: BrowserWindow | null = null
let tray: Tray | null = null
let isQuitting = false
let settingsStore: SettingsStore

const trayIcon = nativeImage.createFromDataURL(
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAKUlEQVR42mNgQANnz551YGRkZIABJgYGRgYGBiYGBpYGBgZGJgYGAAAKYgIrZbrZ2QAAAABJRU5ErkJggg=='
)

function createWindow(): void {
  mainWindow = new BrowserWindow({
    width: 1180,
    height: 780,
    minWidth: 940,
    minHeight: 620,
    title: 'Quick Document',
    backgroundColor: '#f7f7f5',
    titleBarStyle: process.platform === 'darwin' ? 'hiddenInset' : 'default',
    webPreferences: {
      preload: join(__dirname, '../preload/index.mjs'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false
    }
  })

  if (!app.isPackaged && process.env.ELECTRON_RENDERER_URL) {
    mainWindow.loadURL(process.env.ELECTRON_RENDERER_URL)
  } else {
    mainWindow.loadFile(join(__dirname, '../renderer/index.html'))
  }

  mainWindow.on('close', (event) => {
    if (!isQuitting && settingsStore.getPublicSettings().residentMode) {
      event.preventDefault()
      mainWindow?.hide()
    }
  })
}

function createTray(): void {
  if (tray) return
  trayIcon.setTemplateImage(process.platform === 'darwin')
  tray = new Tray(trayIcon)
  tray.setToolTip('Quick Document')
  tray.setContextMenu(
    Menu.buildFromTemplate([
      {
        label: '显示 Quick Document',
        click: () => showMainWindow()
      },
      {
        label: '退出',
        click: () => {
          isQuitting = true
          app.quit()
        }
      }
    ])
  )
  tray.on('click', () => showMainWindow())
}

function showMainWindow(): void {
  if (!mainWindow) createWindow()
  mainWindow?.show()
  mainWindow?.focus()
}

function registerIpc(): void {
  ipcMain.handle('settings:get', () => settingsStore.getPublicSettings())
  ipcMain.handle('settings:save', (_event, patch) => settingsStore.save(patch))
  ipcMain.handle('settings:import-external', () => {
    const config = loadExternalAiConfig()
    if (!config) return null
    const settings = settingsStore.importExternalConfig(config)
    return {
      settings,
      source: config.source
    }
  })
  ipcMain.handle('files:recent', () => settingsStore.getRecentFiles())

  ipcMain.handle('workspace:choose', async () => {
    const result = await dialog.showOpenDialog({
      title: '选择文档工作区',
      properties: ['openDirectory', 'createDirectory']
    })
    return result.canceled ? null : result.filePaths[0]
  })

  ipcMain.handle('workspace:scan', async (_event, rootPath?: string): Promise<WorkspaceSnapshot> => {
    const publicSettings = settingsStore.getPublicSettings()
    const root = rootPath?.trim() || publicSettings.workspacePath
    return scanWorkspace(root)
  })

  ipcMain.handle('files:open', async (_event, filePath: string) => shell.openPath(filePath))
  ipcMain.handle('files:reveal', (_event, filePath: string) => shell.showItemInFolder(filePath))

  ipcMain.handle('chat:send', async (_event, request: ChatRequest): Promise<ChatResponse> => {
    const publicSettings = settingsStore.getPublicSettings()
    const workspaceSnapshot = request.workspaceSnapshot || scanWorkspace(publicSettings.workspacePath)
    const plan = await planDocumentWork({
      messages: request.messages,
      targetFiles: request.targetFiles || [],
      workspaceSnapshot,
      skillBrief: getEmbeddedOfficeSkillBrief(),
      settings: {
        provider: publicSettings.provider,
        wireApi: publicSettings.wireApi,
        baseUrl: publicSettings.baseUrl,
        model: publicSettings.model,
        apiKey: settingsStore.getApiKey()
      }
    })

    const results = []
    for (const action of plan.actions) {
      const result = await executeDocumentAction(action, publicSettings.workspacePath)
      if (result.file && result.file.kind !== 'unknown') settingsStore.addRecentFile(result.file)
      results.push(result)
    }

    const generatedFiles = results.flatMap((result) => (result.file ? [result.file] : []))
    return {
      message: {
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        role: 'assistant',
        content: buildAssistantReply(plan.reply, results),
        createdAt: new Date().toISOString(),
        actions: results
      },
      generatedFiles
    }
  })
}

function buildAssistantReply(reply: string, results: Awaited<ReturnType<typeof executeDocumentAction>>[]): string {
  if (results.length === 0) return reply
  const lines = results.map((result) => {
    if (result.workflow) return `- ${result.summary}`
    if (result.ok && result.file) return `- ${result.summary}`
    return `- ${result.summary}: ${result.error || 'unknown error'}`
  })
  return `${reply}\n\n处理结果：\n${lines.join('\n')}`
}

app.whenReady().then(() => {
  app.setAppUserModelId('com.quickdocument.desktop')
  settingsStore = new SettingsStore()
  registerIpc()
  createWindow()
  createTray()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
    showMainWindow()
  })
})

app.on('before-quit', () => {
  isQuitting = true
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin' && !settingsStore?.getPublicSettings().residentMode) {
    app.quit()
  }
})
