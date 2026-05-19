import { app, BrowserWindow, dialog, ipcMain, Menu, nativeImage, screen, shell, Tray } from 'electron'
import { appendFileSync, existsSync, mkdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import type {
  ChatGuidanceRequest,
  ChatMessage,
  ChatRequest,
  ChatResponse,
  ChatStreamEvent,
  UpdateStatus,
  VoiceTranscriptionRequest,
  WorkspaceSnapshot
} from '@shared/types'
import { runDocumentAgent } from './services/documentAgent'
import { buildDocumentPreviewContext } from './services/documentTextPreview'
import { buildDoclingPreviewContext, getDoclingStatus, installDocling } from './services/doclingService'
import { loadExternalAiConfig } from './services/externalAiConfig'
import { SettingsStore } from './services/settingsStore'
import { getEmbeddedOfficeSkillBrief, getRelevantOfficeSkillContext } from './services/skillRegistry'
import { checkForUpdates, downloadAndOpenUpdate } from './services/updateService'
import { transcribeVoiceInput } from './services/voiceTranscription'
import { scanWorkspace } from './services/workspaceFiles'

let mainWindow: BrowserWindow | null = null
let tray: Tray | null = null
let isQuitting = false
let settingsStore: SettingsStore
const activeChatControllers = new Map<string, AbortController>()
const activeGuidanceSessions = new Map<string, { pending: ChatMessage[]; all: ChatMessage[] }>()
const mainDir = dirname(fileURLToPath(import.meta.url))

app.commandLine.appendSwitch('disable-features', 'MacWebContentsOcclusion,CalculateNativeWinOcclusion')

const gotSingleInstanceLock = app.requestSingleInstanceLock()
if (!gotSingleInstanceLock) {
  app.quit()
} else {
  app.on('second-instance', () => {
    logMain('second-instance: showing existing window')
    showMainWindow()
  })
}

const fallbackTrayIcon = nativeImage.createFromDataURL(
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAQAAAC1+jfqAAAAKUlEQVR42mNgQANnz551YGRkZIABJgYGRgYGBiYGBpYGBgZGJgYGAAAKYgIrZbrZ2QAAAABJRU5ErkJggg=='
)

function getResourcePath(...segments: string[]): string {
  return app.isPackaged
    ? join(process.resourcesPath, 'resources', ...segments)
    : join(mainDir, '../../resources', ...segments)
}

function getAppIconPath(): string {
  return getResourcePath('icons', process.platform === 'win32' ? 'icon.ico' : 'icon.png')
}

function createAppIcon(): Electron.NativeImage {
  const iconPath = getAppIconPath()
  const icon = existsSync(iconPath) ? nativeImage.createFromPath(iconPath) : fallbackTrayIcon
  return icon.isEmpty() ? fallbackTrayIcon : icon
}

function logMain(message: string, error?: unknown): void {
  try {
    const logDir = app.getPath('userData')
    mkdirSync(logDir, { recursive: true })
    const details =
      error instanceof Error
        ? `${error.message}\n${error.stack || ''}`
        : error
          ? String(error)
          : ''
    appendFileSync(join(logDir, 'main.log'), `[${new Date().toISOString()}] ${message}${details ? `\n${details}` : ''}\n`, 'utf8')
  } catch {
    // Logging must never prevent the app from opening.
  }
}

function createWindow(): BrowserWindow {
  if (mainWindow && !mainWindow.isDestroyed()) return mainWindow

  logMain('createWindow:start')
  mainWindow = new BrowserWindow({
    width: 1180,
    height: 780,
    minWidth: 940,
    minHeight: 620,
    title: 'Quick Document',
    icon: createAppIcon(),
    show: true,
    backgroundColor: '#f7f7f5',
    webPreferences: {
      preload: join(mainDir, '../preload/index.mjs'),
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false
    }
  })

  logMain('createWindow:browserWindow-created')
  ensureWindowVisible(mainWindow)
  logMain('createWindow:window-visible-checked')
  wirePermissionHandlers(mainWindow)
  wireWindowDiagnostics(mainWindow)
  logMain('createWindow:diagnostics-wired')
  showMainWindow()
  logMain('createWindow:initial-show-called')

  const loadPromise =
    !app.isPackaged && process.env.ELECTRON_RENDERER_URL
      ? mainWindow.loadURL(process.env.ELECTRON_RENDERER_URL)
      : mainWindow.loadFile(join(mainDir, '../renderer/index.html'))

  loadPromise.catch((error) => {
    logMain('window:load failed', error)
    mainWindow?.loadURL(
      `data:text/html;charset=utf-8,${encodeURIComponent(
        '<!doctype html><meta charset="utf-8"><body style="font:14px -apple-system,BlinkMacSystemFont,Segoe UI,sans-serif;padding:32px;background:#f7f7f5;color:#26231f"><h2>Quick Document 启动失败</h2><p>渲染页面加载失败，请把 main.log 发给开发者。</p></body>'
      )}`
    )
  })

  if (!app.isPackaged && process.env.ELECTRON_RENDERER_URL) {
    logMain(`window:load url ${process.env.ELECTRON_RENDERER_URL}`)
  } else {
    logMain(`window:load file ${join(mainDir, '../renderer/index.html')}`)
  }

  mainWindow.on('close', (event) => {
    if (!isQuitting && settingsStore.getPublicSettings().residentMode) {
      event.preventDefault()
      mainWindow?.hide()
    }
  })

  mainWindow.on('closed', () => {
    logMain('window:closed')
    mainWindow = null
  })

  setTimeout(() => showMainWindow(), 500)
  setTimeout(() => showMainWindow(), 1800)
  return mainWindow
}

function wireWindowDiagnostics(window: BrowserWindow): void {
  window.once('ready-to-show', () => {
    logMain('window:ready-to-show')
    showMainWindow()
  })
  window.webContents.on('did-finish-load', () => logMain('renderer:did-finish-load'))
  window.webContents.on('did-fail-load', (_event, errorCode, errorDescription, validatedURL) => {
    logMain(`renderer:did-fail-load ${errorCode} ${errorDescription} ${validatedURL}`)
  })
  window.webContents.on('render-process-gone', (_event, details) => {
    logMain(`renderer:render-process-gone ${details.reason} ${details.exitCode}`)
  })
  window.webContents.on('unresponsive', () => logMain('window:unresponsive'))
  window.webContents.on('console-message', (_event, level, message, line, sourceId) => {
    if (level >= 2) logMain(`renderer:console ${level} ${message} (${sourceId}:${line})`)
  })
}

function wirePermissionHandlers(window: BrowserWindow): void {
  window.webContents.session.setPermissionRequestHandler((_webContents, permission, callback) => {
    callback(permission === 'media')
  })
}

function createTray(): void {
  if (tray) return
  const trayIcon = createAppIcon()
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
  const window = mainWindow && !mainWindow.isDestroyed() ? mainWindow : createWindow()
  ensureWindowVisible(window)
  if (window.isMinimized()) window.restore()
  window.setSkipTaskbar(false)
  window.show()
  window.moveTop()
  window.focus()
  app.focus({ steal: true })
  if (process.platform === 'darwin') {
    window.setAlwaysOnTop(true, 'floating')
    setTimeout(() => {
      if (!window.isDestroyed()) window.setAlwaysOnTop(false)
    }, 250)
  }
  logMain(`showMainWindow visible=${window.isVisible()} minimized=${window.isMinimized()}`)
}

function ensureWindowVisible(window: BrowserWindow): void {
  const bounds = window.getBounds()
  const display = screen.getDisplayMatching(bounds) || screen.getPrimaryDisplay()
  const area = display.workArea
  const isOutside =
    bounds.x + Math.min(bounds.width, 120) < area.x ||
    bounds.y + Math.min(bounds.height, 80) < area.y ||
    bounds.x > area.x + area.width - 120 ||
    bounds.y > area.y + area.height - 80
  if (isOutside) window.center()
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
  ipcMain.handle('chat:history:get', () => settingsStore.getChatHistory())
  ipcMain.handle('chat:history:save', (_event, messages) => settingsStore.saveChatHistory(messages))
  ipcMain.handle('chat:history:clear', () => settingsStore.clearChatHistory())
  ipcMain.handle('chat:cancel', (_event, requestId: string) => {
    const controller = activeChatControllers.get(requestId)
    if (!controller) return false
    controller.abort()
    return true
  })
  ipcMain.handle('chat:guide', (_event, request: ChatGuidanceRequest) => {
    const session = activeGuidanceSessions.get(request.requestId)
    if (!session) return false
    session.pending.push(request.message)
    session.all.push(request.message)
    return true
  })
  ipcMain.handle('docling:status', () => getDoclingStatus())
  ipcMain.handle('docling:install', async () => {
    const status = await getDoclingStatus()
    if (status.installed) return { ...status, ok: true, log: 'Docling already installed.' }
    const result = await dialog.showMessageBox({
      type: 'question',
      buttons: ['允许安装', '取消'],
      defaultId: 0,
      cancelId: 1,
      title: '安装 Docling',
      message: 'Quick Document 需要安装 Docling 来增强文档解析。',
      detail: `将执行：${status.installCommand || 'python -m pip install --user docling'}\n\n该操作需要访问 Python/PyPI，Windows 和 macOS 都会使用当前系统 Python 环境。`
    })
    if (result.response !== 0) {
      return { ...status, ok: false, message: '已取消安装 Docling。' }
    }
    return installDocling()
  })

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
  ipcMain.handle('updates:check', () => checkForUpdates())
  ipcMain.handle('updates:download', (_event, status?: UpdateStatus) => downloadAndOpenUpdate(status))
  ipcMain.handle('voice:transcribe', (_event, request: VoiceTranscriptionRequest) => {
    return transcribeVoiceInput(request, {
      xfyunVoiceConfig: settingsStore.getXfyunVoiceConfig()
    })
  })

  ipcMain.handle('chat:send', async (event, request: ChatRequest): Promise<ChatResponse> => {
    const requestId = request.requestId || `${Date.now()}-${Math.random().toString(16).slice(2)}`
    const abortController = new AbortController()
    const guidanceSession = { pending: [] as ChatMessage[], all: [] as ChatMessage[] }
    activeChatControllers.set(requestId, abortController)
    activeGuidanceSessions.set(requestId, guidanceSession)
    const emit = (payload: Omit<ChatStreamEvent, 'requestId'>): void => {
      event.sender.send('chat:stream', { requestId, ...payload })
    }

    try {
      const publicSettings = settingsStore.getPublicSettings()
      emit({ type: 'status', message: '正在读取当前文档目录...' })
      const workspaceSnapshot = scanWorkspace(publicSettings.workspacePath)
      const latestPrompt = request.messages[request.messages.length - 1]?.content || ''
      emit({ type: 'status', message: '正在读取相关文档预览...' })
      const documentPreviewContext = await buildDocumentPreviewContext(
        latestPrompt,
        request.targetFiles || [],
        workspaceSnapshot
      )
      const doclingPreviewContext = await buildDoclingPreviewContext(
        request.targetFiles.length > 0 ? request.targetFiles : workspaceSnapshot.files
      )
      emit({ type: 'status', message: '正在发送给 AI，等待模型返回...' })
      const agentInput = {
        messages: request.messages,
        targetFiles: request.targetFiles || [],
        workspaceSnapshot,
        documentPreviewContext: [documentPreviewContext, doclingPreviewContext ? `Docling 解析结果：\n${doclingPreviewContext}` : '']
          .filter(Boolean)
          .join('\n\n---\n\n'),
        skillBrief: [
          getEmbeddedOfficeSkillBrief(),
          getRelevantOfficeSkillContext(latestPrompt, request.targetFiles || [], workspaceSnapshot)
        ]
          .filter(Boolean)
          .join('\n\n---\n\n'),
        settings: {
          provider: publicSettings.provider,
          wireApi: publicSettings.wireApi,
          baseUrl: publicSettings.baseUrl,
          model: publicSettings.model,
          apiKey: settingsStore.getApiKey()
        },
        signal: abortController.signal,
        onProgress: (message: string) => emit({ type: 'status', message }),
        onAssistantDelta: (delta: string) => emit({ type: 'assistant-delta', delta }),
        consumeGuidance: () => guidanceSession.pending.splice(0)
      }
      const agentResult = await runDocumentAgent(agentInput)

      const results = agentResult.actionResults
      for (const result of results) {
        if (result.file && result.file.kind !== 'unknown') settingsStore.addRecentFile(result.file)
        emit({
          type: result.ok ? 'step-done' : 'error',
          message: result.ok ? result.summary : `${result.summary}${result.error ? `：${result.error}` : ''}`
        })
      }
      const generatedFiles = results.flatMap((result) => (result.file ? [result.file] : []))
      const assistantMessage = {
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        role: 'assistant' as const,
        content: buildAssistantReply(agentResult.reply, results),
        createdAt: new Date().toISOString(),
        actions: results
      }
      settingsStore.saveChatHistory([...request.messages, ...guidanceSession.all, assistantMessage])
      emit({ type: 'done', message: generatedFiles.length > 0 ? '处理完成，已刷新文档目录和产物列表。' : '处理完成。' })

      return {
        message: assistantMessage,
        generatedFiles
      }
    } catch (error) {
      const message = chatFailureMessage(error)
      emit({ type: 'error', message })
      emit({ type: 'done', message: '处理结束，未修改文档。' })
      const assistantMessage = {
        id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        role: 'assistant' as const,
        content: message,
        createdAt: new Date().toISOString(),
        actions: []
      }
      settingsStore.saveChatHistory([...request.messages, ...guidanceSession.all, assistantMessage])
      return {
        message: assistantMessage,
        generatedFiles: []
      }
    } finally {
      activeChatControllers.delete(requestId)
      activeGuidanceSessions.delete(requestId)
    }
  })
}

function buildAssistantReply(reply: string, results: Array<{ ok: boolean; summary: string; error?: string; file?: unknown }>): string {
  const text = reply.trim()
  if (text) return text
  if (results.length === 0) return ''
  return results
    .map((result) => (result.ok ? result.summary : `${result.summary}: ${result.error || 'unknown error'}`))
    .join('\n')
}

function chatFailureMessage(error: unknown): string {
  const raw = error instanceof Error ? error.message : String(error)
  const message = raw
    .replace(/Error invoking remote method ['"][^'"]+['"]:\s*/g, '')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 500)

  if (/\b504\b|gateway|timeout|timed out/i.test(message)) {
    return 'AI 接口或代理暂时超时，当前没有修改任何文档。请稍后重试，或切换到可用的 cc-switch / OpenAI 代理配置。'
  }
  if (/手动停止|cancelled|aborted|abort/i.test(message)) {
    return '已手动停止当前处理，后续步骤没有继续执行。'
  }
  if (/\b401\b|unauthorized|api key|authentication/i.test(message)) {
    return 'AI Key 不可用或认证失败，当前没有修改任何文档。请检查 cc-switch / API Key 配置。'
  }
  return `处理没有完成，当前没有修改任何文档。原因：${message || '未知错误'}`
}

process.on('uncaughtException', (error) => logMain('process:uncaughtException', error))
process.on('unhandledRejection', (reason) => logMain('process:unhandledRejection', reason))

app.whenReady().then(() => {
  app.setAppUserModelId('com.quickdocument.desktop')
  logMain('app:ready')
  settingsStore = new SettingsStore()
  registerIpc()
  createWindow()
  createTray()

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
    showMainWindow()
  })
}).catch((error) => {
  logMain('app:ready failed', error)
})

app.on('before-quit', () => {
  logMain('app:before-quit')
  isQuitting = true
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin' && !settingsStore?.getPublicSettings().residentMode) {
    app.quit()
  }
})
