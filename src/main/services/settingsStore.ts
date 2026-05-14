import { app, safeStorage } from 'electron'
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import type { AiProvider, AppSettings, ExternalAiConfig, GeneratedFile, SettingsPatch } from '@shared/types'
import { loadExternalAiConfig } from './externalAiConfig'

interface StoredSettings {
  provider: AiProvider
  baseUrl: string
  model: string
  workspacePath: string
  residentMode: boolean
  apiKeyCipher?: string
  apiKeyPlain?: string
  recentFiles: GeneratedFile[]
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
    const external = storedApiKey ? null : loadExternalAiConfig()
    this.ensureWorkspace(stored.workspacePath)

    return {
      provider: external?.provider || stored.provider,
      baseUrl: external?.baseUrl || stored.baseUrl,
      model: external?.model || stored.model,
      workspacePath: stored.workspacePath,
      residentMode: stored.residentMode,
      hasApiKey: Boolean(storedApiKey || external?.apiKey),
      apiConfigSource: storedApiKey ? 'Quick Document 设置' : external?.source,
      usesExternalApiConfig: Boolean(!storedApiKey && external?.apiKey)
    }
  }

  public save(patch: SettingsPatch): AppSettings {
    const stored = this.read()
    const next: StoredSettings = {
      ...stored,
      provider: patch.provider || stored.provider,
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
      if (safeStorage.isEncryptionAvailable()) {
        next.apiKeyCipher = safeStorage.encryptString(apiKey).toString('base64')
        delete next.apiKeyPlain
      } else {
        next.apiKeyPlain = apiKey
        delete next.apiKeyCipher
      }
    }

    this.ensureWorkspace(next.workspacePath)
    this.write(next)
    return this.getPublicSettings()
  }

  public importExternalConfig(config: ExternalAiConfig): AppSettings {
    return this.save({
      provider: config.provider,
      baseUrl: config.baseUrl,
      model: config.model,
      apiKey: config.apiKey
    })
  }

  public getApiKey(): string {
    const stored = this.read()
    const storedApiKey = this.getStoredApiKey(stored)
    if (storedApiKey) return storedApiKey

    return loadExternalAiConfig()?.apiKey || ''
  }

  private getStoredApiKey(stored: StoredSettings): string {
    if (stored.apiKeyCipher && safeStorage.isEncryptionAvailable()) {
      try {
        return safeStorage.decryptString(Buffer.from(stored.apiKeyCipher, 'base64'))
      } catch {
        return ''
      }
    }

    return stored.apiKeyPlain || ''
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
      baseUrl: DEFAULT_BASE_URL,
      model: DEFAULT_MODEL,
      workspacePath: join(app.getPath('documents'), 'Quick Document'),
      residentMode: true,
      recentFiles: []
    }
  }

  private ensureWorkspace(workspacePath: string): void {
    mkdirSync(workspacePath, { recursive: true })
  }
}
