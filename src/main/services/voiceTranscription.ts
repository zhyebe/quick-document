import type { AiProvider, VoiceTranscriptionRequest, VoiceTranscriptionResult } from '@shared/types'

interface TranscriptionSettings {
  provider: AiProvider
  baseUrl: string
  model: string
  apiKey: string
}

const TRANSCRIPTION_TIMEOUT_MS = 60_000
const MAX_AUDIO_BYTES = 20 * 1024 * 1024
const TRANSCRIPTION_MODELS = ['whisper-1', 'gpt-4o-mini-transcribe', 'gpt-4o-transcribe']
const AUDIO_CHAT_MODELS = ['gpt-audio', 'gpt-4o-audio-preview']
const DIRECT_AUDIO_PROMPT = '请把这段语音完整转写成文字。只返回转写文本，不要解释，不要添加标点以外的说明。如果听不清，请尽量按原话转写。'

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

    const errors: string[] = []
    for (const model of audioChatModelCandidates(settings.model)) {
      try {
        const text = await callDirectAudioChat(audio, { ...settings, model })
        const transcription = normalizeTranscriptionText(text)
        if (transcription) {
          return {
            ok: true,
            text: transcription,
            message: '语音已由当前 AI 模型转成文字。'
          }
        }
        errors.push(`${model} 没有收到可识别的语音内容。`)
      } catch (error) {
        errors.push(`${model} 音频输入失败：${error instanceof Error ? error.message : String(error)}`)
      }
    }

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
        errors.push(`${model} 接口没有返回文字。`)
      } catch (error) {
        errors.push(`${model}：${error instanceof Error ? error.message : String(error)}`)
      }
    }

    return {
      ok: false,
      message: `语音转文字失败：${errors.slice(-3).join('；') || '未知错误'}`
    }
  } catch (error) {
    return {
      ok: false,
      message: `语音转文字失败：${error instanceof Error ? error.message : String(error)}`
    }
  }
}

async function callTranscriptionApi(
  audio: AudioPayload,
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

async function callDirectAudioChat(audio: AudioPayload, settings: TranscriptionSettings): Promise<string> {
  const audioErrorMessages: string[] = []
  try {
    return await callDirectAudioChatWithContent(settings, [
      { type: 'text', text: DIRECT_AUDIO_PROMPT },
      {
        type: 'input_audio',
        input_audio: {
          data: audio.base64,
          format: audioInputFormat(audio.mimeType)
        }
      }
    ])
  } catch (error) {
    audioErrorMessages.push(error instanceof Error ? error.message : String(error))
  }

  try {
    return await callDirectAudioChatWithContent(settings, [
      { type: 'text', text: DIRECT_AUDIO_PROMPT },
      {
        type: 'file',
        file: {
          filename: audio.filename,
          file_data: audio.dataUrl
        }
      }
    ])
  } catch (error) {
    audioErrorMessages.push(error instanceof Error ? error.message : String(error))
  }

  throw new Error(audioErrorMessages.join('；') || '当前模型不支持音频输入。')
}

async function callDirectAudioChatWithContent(
  settings: TranscriptionSettings,
  content: Array<Record<string, unknown>>
): Promise<string> {
  const response = await fetchWithTimeout(`${settings.baseUrl.replace(/\/$/, '')}/chat/completions`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${settings.apiKey}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: settings.model,
      temperature: 0,
      store: false,
      messages: [
        {
          role: 'system',
          content: '你是一个语音转文字助手，只输出转写文本。'
        },
        {
          role: 'user',
          content
        }
      ]
    })
  })

  if (!response.ok) {
    const detail = await response.text().catch(() => '')
    throw new Error(`当前模型返回 ${response.status}${detail ? `：${sanitizeRemoteError(detail)}` : ''}`)
  }

  const data = (await response.json()) as { choices?: Array<{ message?: { content?: unknown } }> }
  return textFromMessageContent(data.choices?.[0]?.message?.content)
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

interface AudioPayload {
  bytes: Uint8Array
  mimeType: string
  filename: string
  base64: string
  dataUrl: string
}

function audioFromDataUrl(dataUrl: string, mimeType: string): AudioPayload {
  const parsed = parseAudioDataUrl(dataUrl, mimeType)
  const bytes = Uint8Array.from(Buffer.from(parsed.base64, 'base64'))
  if (bytes.byteLength === 0) throw new Error('录音数据为空。')
  return {
    bytes,
    mimeType: parsed.mimeType,
    filename: `voice-input.${audioExtension(parsed.mimeType)}`,
    base64: parsed.base64,
    dataUrl
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

function audioInputFormat(mimeType: string): string {
  if (/wav|x-wav/i.test(mimeType)) return 'wav'
  if (/mpeg|mp3/i.test(mimeType)) return 'mp3'
  if (/flac/i.test(mimeType)) return 'flac'
  if (/ogg|opus/i.test(mimeType)) return 'mp3'
  return 'wav'
}

function audioChatModelCandidates(currentModel: string): string[] {
  const candidates = [currentModel.trim(), ...AUDIO_CHAT_MODELS]
    .filter(Boolean)
    .filter((model, index, array) => array.indexOf(model) === index)
  return candidates
}

function normalizeTranscriptionText(value: string): string {
  const text = value
    .trim()
    .replace(/^["'“”‘’]+|["'“”‘’]+$/g, '')
    .trim()
  if (!text || isNoAudioReply(text)) return ''
  return text
}

function isNoAudioReply(text: string): boolean {
  return /未收到.*语音|没有收到.*语音|没.*收到.*音频|没有.*音频|无法.*(听到|访问|读取).*音频|no audio|did(?: not|n't) receive.*audio|cannot access.*audio|unable to access.*audio/i.test(text)
}

function textFromMessageContent(content: unknown): string {
  if (typeof content === 'string') return content
  if (!Array.isArray(content)) return ''
  return content
    .map((part) => {
      if (!part || typeof part !== 'object') return ''
      const record = part as Record<string, unknown>
      return typeof record.text === 'string' ? record.text : ''
    })
    .join('')
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
