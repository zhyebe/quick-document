import type { AiProvider, VoiceTranscriptionRequest, VoiceTranscriptionResult } from '@shared/types'

interface TranscriptionSettings {
  provider: AiProvider
  baseUrl: string
  apiKey: string
}

const TRANSCRIPTION_TIMEOUT_MS = 60_000
const MAX_AUDIO_BYTES = 20 * 1024 * 1024
const TRANSCRIPTION_MODELS = ['whisper-1', 'gpt-4o-mini-transcribe', 'gpt-4o-transcribe']

export async function transcribeVoiceInput(
  request: VoiceTranscriptionRequest,
  settings: TranscriptionSettings
): Promise<VoiceTranscriptionResult> {
  if (!settings.apiKey) {
    return {
      ok: false,
      message: '还没有可用的 AI Key，不能进行语音转文字。请先配置 API Key 或启用 cc-switch/Codex 配置。'
    }
  }
  if (settings.provider !== 'openai') {
    return {
      ok: false,
      message: '当前语音转文字需要 OpenAI 兼容接口。请切换到 OpenAI-compatible 配置。'
    }
  }

  try {
    const audio = audioFromDataUrl(request.dataUrl, request.mimeType)
    if (audio.bytes.byteLength > MAX_AUDIO_BYTES) {
      return {
        ok: false,
        message: '录音太长，暂时无法转文字。请缩短后再试。'
      }
    }

    let lastError = ''
    for (const model of TRANSCRIPTION_MODELS) {
      try {
        const text = await callTranscriptionApi(audio, request.language || 'zh', model, settings)
        if (text.trim()) {
          return {
            ok: true,
            text: text.trim(),
            message: '语音已转成文字。'
          }
        }
        lastError = '接口没有返回文字。'
      } catch (error) {
        lastError = error instanceof Error ? error.message : String(error)
      }
    }

    return {
      ok: false,
      message: `语音转文字失败：${lastError || '未知错误'}`
    }
  } catch (error) {
    return {
      ok: false,
      message: `语音转文字失败：${error instanceof Error ? error.message : String(error)}`
    }
  }
}

async function callTranscriptionApi(
  audio: { bytes: Uint8Array; mimeType: string; filename: string },
  language: string,
  model: string,
  settings: TranscriptionSettings
): Promise<string> {
  const form = new FormData()
  form.append('model', model)
  form.append('language', language)
  form.append('response_format', 'json')
  form.append('file', new Blob([Buffer.from(audio.bytes)], { type: audio.mimeType }), audio.filename)

  const response = await fetchWithTimeout(`${settings.baseUrl.replace(/\/$/, '')}/audio/transcriptions`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${settings.apiKey}`
    },
    body: form
  })

  if (!response.ok) {
    const detail = await response.text().catch(() => '')
    throw new Error(`${model} 返回 ${response.status}${detail ? `：${sanitizeRemoteError(detail)}` : ''}`)
  }

  const contentType = response.headers.get('content-type') || ''
  if (contentType.includes('application/json')) {
    const data = (await response.json()) as { text?: unknown }
    return typeof data.text === 'string' ? data.text : ''
  }
  return response.text()
}

async function fetchWithTimeout(url: string, init: RequestInit): Promise<Response> {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), TRANSCRIPTION_TIMEOUT_MS)
  try {
    return await fetch(url, {
      ...init,
      signal: controller.signal
    })
  } catch (error) {
    if (error instanceof Error && /abort|aborted/i.test(error.message)) {
      throw new Error(`请求超时（${Math.round(TRANSCRIPTION_TIMEOUT_MS / 1000)}s）`)
    }
    throw error
  } finally {
    clearTimeout(timer)
  }
}

function audioFromDataUrl(dataUrl: string, mimeType: string): { bytes: Uint8Array; mimeType: string; filename: string } {
  const parsed = parseAudioDataUrl(dataUrl, mimeType)
  const bytes = Uint8Array.from(Buffer.from(parsed.base64, 'base64'))
  if (bytes.byteLength === 0) throw new Error('录音数据为空。')
  return {
    bytes,
    mimeType: parsed.mimeType,
    filename: `voice-input.${audioExtension(parsed.mimeType)}`
  }
}

function parseAudioDataUrl(dataUrl: string, fallbackMimeType: string): { mimeType: string; base64: string } {
  const trimmed = dataUrl.trim()
  const commaIndex = trimmed.indexOf(',')
  if (!trimmed.startsWith('data:') || commaIndex < 0) {
    throw new Error('录音数据格式不正确。')
  }

  const header = trimmed.slice(5, commaIndex)
  const body = trimmed.slice(commaIndex + 1).replace(/\s/g, '')
  const parts = header.split(';').filter(Boolean)
  const detectedMimeType = parts.find((part) => part.includes('/')) || fallbackMimeType || 'audio/webm'
  const isBase64 = parts.some((part) => part.toLowerCase() === 'base64')
  if (!isBase64 || !body) throw new Error('录音数据格式不正确。')
  return {
    mimeType: detectedMimeType,
    base64: body
  }
}

function audioExtension(mimeType: string): string {
  if (/mp4|m4a/i.test(mimeType)) return 'm4a'
  if (/mpeg|mp3/i.test(mimeType)) return 'mp3'
  if (/ogg/i.test(mimeType)) return 'ogg'
  if (/wav/i.test(mimeType)) return 'wav'
  return 'webm'
}

function sanitizeRemoteError(text: string): string {
  return text
    .trim()
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .slice(0, 400)
}
