import { existsSync, readFileSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import type { AiProvider, ExternalAiConfig } from '@shared/types'

interface CandidateConfig {
  path: string
  source: string
  providerHint?: AiProvider
}

export function loadExternalAiConfig(): ExternalAiConfig | null {
  const ccSwitchAppPaths = join(
    homedir(),
    'Library',
    'Application Support',
    'com.ccswitch.desktop',
    'app_paths.json'
  )
  const candidates: CandidateConfig[] = [
    {
      path: ccSwitchAppPaths,
      source: 'cc-switch desktop'
    },
    ...readCandidatePaths(ccSwitchAppPaths).map((path) => ({
      path,
      source: 'cc-switch linked config'
    })),
    {
      path: join(homedir(), '.claude', 'settings.json'),
      source: 'cc-switch / Claude settings',
      providerHint: 'anthropic'
    }
  ]

  for (const candidate of candidates) {
    const config = readCandidate(candidate)
    if (config) return config
  }

  return null
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
  if (!existsSync(candidate.path)) return null

  try {
    const raw = JSON.parse(readFileSync(candidate.path, 'utf8')) as unknown
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
      baseUrl: baseUrl || defaultBaseUrl(provider),
      model,
      apiKey,
      source: candidate.source
    }
  } catch {
    return null
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
