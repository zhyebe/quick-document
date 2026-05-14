import { contextBridge, ipcRenderer } from 'electron'
import type {
  AppSettings,
  ChatRequest,
  ChatResponse,
  GeneratedFile,
  SettingsPatch,
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
  getRecentFiles: (): Promise<GeneratedFile[]> => ipcRenderer.invoke('files:recent'),
  openFile: (filePath: string): Promise<string> => ipcRenderer.invoke('files:open', filePath),
  revealFile: (filePath: string): Promise<void> => ipcRenderer.invoke('files:reveal', filePath)
}

contextBridge.exposeInMainWorld('quickDocument', api)

export type QuickDocumentApi = typeof api
