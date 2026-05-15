import { contextBridge, ipcRenderer } from 'electron'
import type {
  AppSettings,
  ChatRequest,
  ChatResponse,
  ChatHistorySnapshot,
  ChatStreamEvent,
  DoclingInstallResult,
  DoclingStatus,
  GeneratedFile,
  SettingsPatch,
  UpdateDownloadResult,
  UpdateStatus,
  WorkspaceSnapshot
} from '@shared/types'

const api = {
  getSettings: (): Promise<AppSettings> => ipcRenderer.invoke('settings:get'),
  saveSettings: (patch: SettingsPatch): Promise<AppSettings> =>
    ipcRenderer.invoke('settings:save', patch),
  importExternalSettings: (): Promise<{ settings: AppSettings; source: string } | null> =>
    ipcRenderer.invoke('settings:import-external'),
  chooseWorkspace: (): Promise<string | null> => ipcRenderer.invoke('workspace:choose'),
  scanWorkspace: (rootPath?: string): Promise<WorkspaceSnapshot> =>
    ipcRenderer.invoke('workspace:scan', rootPath),
  sendMessage: (request: ChatRequest): Promise<ChatResponse> =>
    ipcRenderer.invoke('chat:send', request),
  cancelMessage: (requestId: string): Promise<boolean> =>
    ipcRenderer.invoke('chat:cancel', requestId),
  onChatStream: (callback: (event: ChatStreamEvent) => void): (() => void) => {
    const listener = (_event: Electron.IpcRendererEvent, payload: ChatStreamEvent): void => callback(payload)
    ipcRenderer.on('chat:stream', listener)
    return () => ipcRenderer.removeListener('chat:stream', listener)
  },
  getDoclingStatus: (): Promise<DoclingStatus> => ipcRenderer.invoke('docling:status'),
  installDocling: (): Promise<DoclingInstallResult> => ipcRenderer.invoke('docling:install'),
  getChatHistory: (): Promise<ChatHistorySnapshot> => ipcRenderer.invoke('chat:history:get'),
  saveChatHistory: (messages: ChatHistorySnapshot['messages']): Promise<ChatHistorySnapshot> =>
    ipcRenderer.invoke('chat:history:save', messages),
  clearChatHistory: (): Promise<ChatHistorySnapshot> => ipcRenderer.invoke('chat:history:clear'),
  getRecentFiles: (): Promise<GeneratedFile[]> => ipcRenderer.invoke('files:recent'),
  openFile: (filePath: string): Promise<string> => ipcRenderer.invoke('files:open', filePath),
  revealFile: (filePath: string): Promise<void> => ipcRenderer.invoke('files:reveal', filePath),
  checkForUpdates: (): Promise<UpdateStatus> => ipcRenderer.invoke('updates:check'),
  downloadUpdate: (): Promise<UpdateDownloadResult> => ipcRenderer.invoke('updates:download')
}

contextBridge.exposeInMainWorld('quickDocument', api)

export type QuickDocumentApi = typeof api
