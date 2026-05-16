import { app, safeStorage } from 'electron'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import type {
  AiProvider,
  AiWireApi,
  AppSettings,
  ChatHistorySnapshot,
  ChatMessage,
  ExternalAiConfig,
  GeneratedFile,
  SettingsPatch
} from '@shared/types'
import { loadExternalAiConfig } from './externalAiConfig'

interface StoredSettings {
  provider: AiProvider
  wireApi: AiWireApi
  baseUrl: string
  model: string
  workspacePath: string
  residentMode: boolean
  apiKeyCipher?: string
  apiKeyPlain?: string
  xfyunVoiceAppId?: string
  xfyunVoiceApiKeyCipher?: string
  xfyunVoiceApiKeyPlain?: string
  xfyunVoiceApiSecretCipher?: string
  xfyunVoiceApiSecretPlain?: string
  recentFiles: GeneratedFile[]
  chatHistory: ChatHistorySnapshot
}

export interface XfyunVoiceConfig {
  appId: string
  apiKey: string
  apiSecret: string
}

const DEFAULT_BASE_URL = 'https://api.openai.com/v1'
const DEFAULT_MODEL = 'gpt-4.1-mini'

export class SettingsStore {
  private readonly settingsPath: string

  constructor() {
    this.settingsPath = join(app.getPath('userData'), 'settings.json')
  }

  public getPublicSettings(): AppSettings {
    const stored = this.read()
    const storedApiKey = this.getStoredApiKey(stored)
    const xfyunVoiceConfig = this.getStoredXfyunVoiceConfig(stored)
    const external = loadExternalAiConfig()
    this.ensureWorkspace(stored.workspacePath)

    return {
      provider: external?.provider || stored.provider,
      wireApi: external?.wireApi || stored.wireApi,
      baseUrl: external?.baseUrl || stored.baseUrl,
      model: external?.model || stored.model,
      workspacePath: stored.workspacePath,
      residentMode: stored.residentMode,
      hasApiKey: Boolean(external?.apiKey || storedApiKey),
      apiConfigSource: external?.source || (storedApiKey ? 'Quick Document 设置' : undefined),
      usesExternalApiConfig: Boolean(external?.apiKey),
      cachedMessageCount: stored.chatHistory.messages.length,
      hasXfyunVoiceConfig: Boolean(xfyunVoiceConfig),
      xfyunVoiceAppId: stored.xfyunVoiceAppId
    }
  }

  public save(patch: SettingsPatch): AppSettings {
    const stored = this.read()
    const next: StoredSettings = {
      ...stored,
      provider: patch.provider || stored.provider,
      wireApi: patch.wireApi || stored.wireApi,
      baseUrl: patch.baseUrl?.trim() || stored.baseUrl,
      model: patch.model?.trim() || stored.model,
      workspacePath: patch.workspacePath?.trim() || stored.workspacePath,
      residentMode:
        typeof patch.residentMode === 'boolean' ? patch.residentMode : stored.residentMode
    }

    if (patch.clearApiKey) {
      delete next.apiKeyCipher
      delete next.apiKeyPlain
    }

    if (patch.apiKey?.trim()) {
      const apiKey = patch.apiKey.trim()
      this.setSecret(next, 'apiKey', apiKey)
    }

    if (patch.clearXfyunVoiceConfig) {
      delete next.xfyunVoiceAppId
      delete next.xfyunVoiceApiKeyCipher
      delete next.xfyunVoiceApiKeyPlain
      delete next.xfyunVoiceApiSecretCipher
      delete next.xfyunVoiceApiSecretPlain
    }

    if (typeof patch.xfyunVoiceAppId === 'string') {
      next.xfyunVoiceAppId = patch.xfyunVoiceAppId.trim()
    }
    if (patch.xfyunVoiceApiKey?.trim()) {
      this.setSecret(next, 'xfyunVoiceApiKey', patch.xfyunVoiceApiKey.trim())
    }
    if (patch.xfyunVoiceApiSecret?.trim()) {
      this.setSecret(next, 'xfyunVoiceApiSecret', patch.xfyunVoiceApiSecret.trim())
    }

    this.ensureWorkspace(next.workspacePath)
    this.write(next)
    return this.getPublicSettings()
  }

  public importExternalConfig(config: ExternalAiConfig): AppSettings {
    return this.save({
      provider: config.provider,
      wireApi: config.wireApi,
      baseUrl: config.baseUrl,
      model: config.model,
      apiKey: config.apiKey
    })
  }

  public getApiKey(): string {
    const stored = this.read()
    const external = loadExternalAiConfig()
    if (external?.apiKey) return external.apiKey

    const storedApiKey = this.getStoredApiKey(stored)
    if (storedApiKey) return storedApiKey

    return ''
  }

  public getXfyunVoiceConfig(): XfyunVoiceConfig | null {
    return this.getStoredXfyunVoiceConfig(this.read())
  }

  private getStoredApiKey(stored: StoredSettings): string {
    return this.getSecret(stored.apiKeyCipher, stored.apiKeyPlain)
  }

  private getStoredXfyunVoiceConfig(stored: StoredSettings): XfyunVoiceConfig | null {
    const appId = stored.xfyunVoiceAppId?.trim() || ''
    const apiKey = this.getSecret(stored.xfyunVoiceApiKeyCipher, stored.xfyunVoiceApiKeyPlain)
    const apiSecret = this.getSecret(stored.xfyunVoiceApiSecretCipher, stored.xfyunVoiceApiSecretPlain)
    if (!appId || !apiKey || !apiSecret) return null
    return { appId, apiKey, apiSecret }
  }

  private getSecret(cipher: string | undefined, plain: string | undefined): string {
    if (cipher && safeStorage.isEncryptionAvailable()) {
      try {
        return safeStorage.decryptString(Buffer.from(cipher, 'base64'))
      } catch {
        return ''
      }
    }

    return plain || ''
  }

  private setSecret(
    settings: StoredSettings,
    field: 'apiKey' | 'xfyunVoiceApiKey' | 'xfyunVoiceApiSecret',
    value: string
  ): void {
    const cipherKey = `${field}Cipher` as keyof StoredSettings
    const plainKey = `${field}Plain` as keyof StoredSettings
    const mutable = settings as unknown as Record<string, unknown>
    if (safeStorage.isEncryptionAvailable()) {
      mutable[cipherKey] = safeStorage.encryptString(value).toString('base64')
      delete mutable[plainKey]
    } else {
      mutable[plainKey] = value
      delete mutable[cipherKey]
    }
  }

  public getRecentFiles(): GeneratedFile[] {
    return this.read().recentFiles.slice(0, 20)
  }

  public addRecentFile(file: GeneratedFile): void {
    const stored = this.read()
    const deduped = stored.recentFiles.filter((item) => item.path !== file.path)
    this.write({
      ...stored,
      recentFiles: [file, ...deduped].slice(0, 20)
    })
  }

  public getChatHistory(): ChatHistorySnapshot {
    return this.read().chatHistory
  }

  public saveChatHistory(messages: ChatMessage[]): ChatHistorySnapshot {
    const stored = this.read()
    const sanitized = messages
      .filter((message) => message.role === 'user' || message.role === 'assistant')
      .slice(-80)
      .map((message) => ({
        ...message,
        attachments: message.attachments?.map((attachment) => ({
          ...attachment,
          dataUrl:
            attachment.dataUrl.length > 2_000_000
              ? ''
              : attachment.dataUrl
        }))
      }))
    const chatHistory = {
      messages: sanitized,
      updatedAt: new Date().toISOString()
    }
    this.write({
      ...stored,
      chatHistory
    })
    return chatHistory
  }

  public clearChatHistory(): ChatHistorySnapshot {
    const stored = this.read()
    const chatHistory = {
      messages: [],
      updatedAt: new Date().toISOString()
    }
    this.write({
      ...stored,
      chatHistory
    })
    return chatHistory
  }

  private read(): StoredSettings {
    if (!existsSync(this.settingsPath)) {
      const defaults = this.defaults()
      this.write(defaults)
      return defaults
    }

    try {
      return {
        ...this.defaults(),
        ...JSON.parse(readFileSync(this.settingsPath, 'utf8'))
      }
    } catch {
      const defaults = this.defaults()
      this.write(defaults)
      return defaults
    }
  }

  private write(settings: StoredSettings): void {
    mkdirSync(dirname(this.settingsPath), { recursive: true })
    writeFileSync(this.settingsPath, JSON.stringify(settings, null, 2), 'utf8')
  }

  private defaults(): StoredSettings {
    return {
      provider: 'openai',
      wireApi: 'chat_completions',
      baseUrl: DEFAULT_BASE_URL,
      model: DEFAULT_MODEL,
      workspacePath: join(app.getPath('documents'), 'Quick Document'),
      residentMode: true,
      recentFiles: [],
      chatHistory: {
        messages: []
      }
    }
  }

  private ensureWorkspace(workspacePath: string): void {
    mkdirSync(workspacePath, { recursive: true })
  }
}
