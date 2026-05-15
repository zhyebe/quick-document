import { execFileSync } from 'node:child_process'
import { existsSync, readFileSync, readdirSync, statSync } from 'node:fs'
import { homedir } from 'node:os'
import { extname } from 'node:path'
import { join } from 'node:path'
import type { AiProvider, AiWireApi, ExternalAiConfig } from '@shared/types'

interface CandidateConfig {
  path: string
  source: string
  providerHint?: AiProvider
  reader?: () => ExternalAiConfig | null
}

export function loadExternalAiConfig(): ExternalAiConfig | null {
  const ccSwitchAppPaths = join(
    homedir(),
    'Library',
    'Application Support',
    'com.ccswitch.desktop',
    'app_paths.json'
  )
  const currentApp = readCcSwitchLastApp()
  const currentCandidates =
    currentApp === 'codex'
      ? [codexCandidate('cc-switch current Codex config')]
      : currentApp === 'claude'
        ? [claudeCandidate('cc-switch current Claude config')]
        : []
  const codexFallbackSource =
    currentApp && currentApp !== 'codex'
      ? `Codex config (preferred OpenAI fallback; cc-switch current app: ${currentApp})`
      : 'Codex config'

  const candidates: CandidateConfig[] = [
    ...currentCandidates,
    codexCandidate(codexFallbackSource),
    ...readCandidatePaths(ccSwitchAppPaths).map((path) => ({
      path,
      source: 'cc-switch linked config'
    })),
    claudeCandidate('Claude settings')
  ]

  for (const candidate of candidates) {
    const config = readCandidate(candidate)
    if (config) return config
  }

  return null
}

function codexCandidate(source: string): CandidateConfig {
  return {
    path: join(homedir(), '.codex', 'config.toml'),
    source,
    providerHint: 'openai',
    reader: () => readCodexConfig(source)
  }
}

function claudeCandidate(source: string): CandidateConfig {
  return {
    path: join(homedir(), '.claude', 'settings.json'),
    source,
    providerHint: 'anthropic'
  }
}

function readCandidatePaths(appPathsFile: string): string[] {
  if (!existsSync(appPathsFile)) return []

  try {
    const parsed = JSON.parse(readFileSync(appPathsFile, 'utf8')) as unknown
    const values = Object.values(flattenObject(parsed))
    return values.filter((value): value is string => {
      if (typeof value !== 'string') return false
      const trimmed = value.trim()
      return Boolean(trimmed) && existsSync(trimmed) && /\.(json|toml|yaml|yml|env)$/i.test(trimmed)
    })
  } catch {
    return []
  }
}

function readCandidate(candidate: CandidateConfig): ExternalAiConfig | null {
  if (candidate.reader) return candidate.reader()
  if (!existsSync(candidate.path)) return null

  try {
    const raw = readStructuredConfig(candidate.path)
    const flattened = flattenObject(raw)
    const provider = inferProvider(flattened, candidate.providerHint)
    const baseUrl = pickString(flattened, [
      'ANTHROPIC_BASE_URL',
      'OPENAI_BASE_URL',
      'baseUrl',
      'base_url',
      'apiBaseUrl',
      'api_base_url'
    ])
    const model = pickString(flattened, [
      'ANTHROPIC_MODEL',
      'OPENAI_MODEL',
      'model',
      'defaultModel',
      'default_model'
    ])
    const wireApi = normalizeWireApi(
      pickString(flattened, ['wireApi', 'wire_api', 'apiType', 'api_type']),
      provider
    )
    const apiKey = pickString(flattened, [
      'ANTHROPIC_AUTH_TOKEN',
      'ANTHROPIC_API_KEY',
      'OPENAI_API_KEY',
      'apiKey',
      'api_key',
      'authToken',
      'auth_token',
      'token'
    ])

    if (!apiKey || !model) return null

    return {
      provider,
      wireApi,
      baseUrl: baseUrl || defaultBaseUrl(provider),
      model,
      apiKey,
      source: candidate.source
    }
  } catch {
    return null
  }
}

function readCodexConfig(source: string): ExternalAiConfig | null {
  const configPath = join(homedir(), '.codex', 'config.toml')
  const authPath = join(homedir(), '.codex', 'auth.json')
  if (!existsSync(configPath)) return null

  try {
    const flattened = flattenTomlLike(readFileSync(configPath, 'utf8'))
    const providerName = pickString(flattened, ['model_provider']) || 'openai'
    const providerPrefix = `model_providers.${providerName}`
    const model = pickString(flattened, ['model'])
    const baseUrl =
      pickString(flattened, [`${providerPrefix}.base_url`, `${providerPrefix}.baseUrl`, 'base_url']) ||
      defaultBaseUrl('openai')
    const wireApi = normalizeWireApi(
      pickString(flattened, [`${providerPrefix}.wire_api`, `${providerPrefix}.wireApi`, 'wire_api']),
      'openai'
    )
    const apiKey = readCodexApiKey(authPath)

    if (!apiKey || !model) return null

    return {
      provider: 'openai',
      wireApi,
      baseUrl,
      model,
      apiKey,
      source
    }
  } catch {
    return null
  }
}

function readCodexApiKey(authPath: string): string {
  if (!existsSync(authPath)) return ''
  try {
    const raw = JSON.parse(readFileSync(authPath, 'utf8')) as unknown
    const flattened = flattenObject(raw)
    return pickString(flattened, ['OPENAI_API_KEY', 'openai_api_key', 'apiKey', 'api_key'])
  } catch {
    return ''
  }
}

function inferProvider(flattened: Record<string, unknown>, hint?: AiProvider): AiProvider {
  if (hint) return hint
  if (hasAnyKey(flattened, ['ANTHROPIC_AUTH_TOKEN', 'ANTHROPIC_API_KEY', 'ANTHROPIC_MODEL'])) {
    return 'anthropic'
  }
  return 'openai'
}

function defaultBaseUrl(provider: AiProvider): string {
  return provider === 'anthropic' ? 'https://api.anthropic.com' : 'https://api.openai.com/v1'
}

function normalizeWireApi(value: string, provider: AiProvider): AiWireApi {
  if (provider === 'anthropic') return 'anthropic_messages'
  const normalized = value.trim().toLowerCase().replace(/-/g, '_')
  if (normalized === 'responses' || normalized === 'response') return 'responses'
  return 'chat_completions'
}

function hasAnyKey(record: Record<string, unknown>, keys: string[]): boolean {
  return keys.some((key) => typeof record[key] === 'string' && Boolean(String(record[key]).trim()))
}

function pickString(record: Record<string, unknown>, keys: string[]): string {
  for (const key of keys) {
    const value = record[key]
    if (typeof value === 'string' && value.trim()) return value.trim()
  }
  return ''
}

function readStructuredConfig(path: string): unknown {
  const raw = readFileSync(path, 'utf8')
  const extension = extname(path).toLowerCase()
  if (extension === '.toml') return flattenTomlLike(raw)
  if (extension === '.env') return parseEnvLike(raw)
  return JSON.parse(raw) as unknown
}

function parseEnvLike(raw: string): Record<string, unknown> {
  const output: Record<string, unknown> = {}
  raw.split(/\r?\n/).forEach((line) => {
    const trimmed = line.trim()
    if (!trimmed || trimmed.startsWith('#')) return
    const match = trimmed.match(/^([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/)
    if (!match) return
    output[match[1]] = unquote(match[2].trim())
  })
  return output
}

function flattenTomlLike(raw: string): Record<string, unknown> {
  const output: Record<string, unknown> = {}
  let section = ''

  raw.split(/\r?\n/).forEach((line) => {
    const trimmed = stripTomlComment(line).trim()
    if (!trimmed) return

    const sectionMatch = trimmed.match(/^\[([^\]]+)]$/)
    if (sectionMatch) {
      section = sectionMatch[1].trim()
      return
    }

    const match = trimmed.match(/^([A-Za-z0-9_.-]+)\s*=\s*(.*)$/)
    if (!match) return
    const key = section ? `${section}.${match[1]}` : match[1]
    output[key] = parseTomlValue(match[2].trim())
  })

  return output
}

function stripTomlComment(line: string): string {
  let quote: '"' | "'" | '' = ''
  for (let index = 0; index < line.length; index += 1) {
    const char = line[index]
    if ((char === '"' || char === "'") && line[index - 1] !== '\\') {
      quote = quote === char ? '' : quote || char
    }
    if (char === '#' && !quote) return line.slice(0, index)
  }
  return line
}

function parseTomlValue(value: string): unknown {
  if (value.startsWith('[') && value.endsWith(']')) {
    return value
      .slice(1, -1)
      .split(',')
      .map((item) => unquote(item.trim()))
      .filter(Boolean)
  }
  if (value === 'true') return true
  if (value === 'false') return false
  return unquote(value)
}

function unquote(value: string): string {
  const trimmed = value.trim()
  if (
    (trimmed.startsWith('"') && trimmed.endsWith('"')) ||
    (trimmed.startsWith("'") && trimmed.endsWith("'"))
  ) {
    return trimmed.slice(1, -1)
  }
  return trimmed
}

function readCcSwitchLastApp(): string {
  if (process.platform !== 'darwin') return ''

  const root = join(
    homedir(),
    'Library',
    'WebKit',
    'com.ccswitch.desktop',
    'WebsiteData'
  )
  const databases = findFiles(root, 'localstorage.sqlite3')

  for (const database of databases) {
    const values = readLocalStorage(database)
    const lastApp = values['cc-switch-last-app']
    if (lastApp) return lastApp.trim().toLowerCase()
  }

  return ''
}

function findFiles(root: string, filename: string, depth = 0): string[] {
  if (!existsSync(root) || depth > 8) return []

  try {
    return readdirSync(root).flatMap((entry) => {
      const path = join(root, entry)
      const stats = statSync(path)
      if (stats.isDirectory()) return findFiles(path, filename, depth + 1)
      return entry === filename ? [path] : []
    })
  } catch {
    return []
  }
}

function readLocalStorage(databasePath: string): Record<string, string> {
  try {
    const output = execFileSync('/usr/bin/sqlite3', [
      databasePath,
      'select key, quote(value) from ItemTable;'
    ]).toString('utf8')

    return Object.fromEntries(
      output
        .split(/\r?\n/)
        .filter(Boolean)
        .map((line) => {
          const separator = line.indexOf('|')
          const key = line.slice(0, separator)
          const value = decodeSqliteQuotedBlob(line.slice(separator + 1))
          return [key, value]
        })
    )
  } catch {
    return {}
  }
}

function decodeSqliteQuotedBlob(value: string): string {
  if (!value.startsWith("X'")) return unquote(value)
  const buffer = Buffer.from(value.slice(2, -1), 'hex')
  const utf16 = buffer.toString('utf16le').replace(/\0+$/g, '')
  if (utf16.trim()) return utf16
  return buffer.toString('utf8').replace(/\0+$/g, '')
}

function flattenObject(value: unknown, prefix = '', output: Record<string, unknown> = {}): Record<string, unknown> {
  if (!value || typeof value !== 'object') return output

  Object.entries(value as Record<string, unknown>).forEach(([key, nested]) => {
    output[key] = nested
    if (prefix) output[`${prefix}.${key}`] = nested
    if (nested && typeof nested === 'object' && !Array.isArray(nested)) {
      flattenObject(nested, prefix ? `${prefix}.${key}` : key, output)
    }
  })

  return output
}
