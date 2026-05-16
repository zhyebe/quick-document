import { createHmac } from 'node:crypto'
import type { VoiceTranscriptionRequest, VoiceTranscriptionResult } from '@shared/types'
import type { XfyunVoiceConfig } from './settingsStore'

interface TranscriptionSettings {
  xfyunVoiceConfig: XfyunVoiceConfig | null
}

const XFYUN_IAT_HOST = 'iat-api.xfyun.cn'
const XFYUN_IAT_PATH = '/v2/iat'
const XFYUN_IAT_URL = `wss://${XFYUN_IAT_HOST}${XFYUN_IAT_PATH}`
const TRANSCRIPTION_TIMEOUT_MS = 60_000
const MAX_AUDIO_BYTES = 20 * 1024 * 1024
const XFYUN_SAMPLE_RATE = 16_000
const XFYUN_FRAME_BYTES = 1280

export async function transcribeVoiceInput(
  request: VoiceTranscriptionRequest,
  settings: TranscriptionSettings
): Promise<VoiceTranscriptionResult> {
  const config = settings.xfyunVoiceConfig
  if (!config) {
    return {
      ok: false,
      message: '还没有配置讯飞语音听写。请在设置中填写 APPID、APIKey 和 APISecret。'
    }
  }

  try {
    const audio = audioFromDataUrl(request.dataUrl)
    if (audio.byteLength > MAX_AUDIO_BYTES) {
      return {
        ok: false,
        message: '录音太长，暂时无法转文字。请缩短后再试。'
      }
    }

    const pcm = wavToPcm16k(audio)
    if (pcm.byteLength < XFYUN_SAMPLE_RATE * 2 * 0.15) {
      return {
        ok: false,
        message: '没有录到有效语音。'
      }
    }

    const text = await callXfyunIat(pcm, config)
    return text.trim()
      ? { ok: true, text: text.trim(), message: '语音已转成文字。' }
      : { ok: false, message: '讯飞没有返回可识别文字。' }
  } catch (error) {
    return {
      ok: false,
      message: `语音转文字失败：${error instanceof Error ? error.message : String(error)}`
    }
  }
}

function callXfyunIat(pcm: Buffer, config: XfyunVoiceConfig): Promise<string> {
  return new Promise((resolve, reject) => {
    const socket = new WebSocket(buildXfyunAuthUrl(config))
    const chunks: string[] = []
    let opened = false
    let settled = false

    const timer = setTimeout(() => {
      finish(new Error(`请求超时（${Math.round(TRANSCRIPTION_TIMEOUT_MS / 1000)}s）`))
    }, TRANSCRIPTION_TIMEOUT_MS)

    const finish = (error?: Error): void => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      try {
        socket.close()
      } catch {
        // no-op
      }
      if (error) reject(error)
      else resolve(chunks.join(''))
    }

    socket.onopen = () => {
      opened = true
      void sendXfyunAudioFrames(socket, pcm, config).catch((error) => finish(error))
    }

    socket.onerror = () => {
      finish(new Error(opened ? '讯飞 WebSocket 连接异常。' : '无法连接讯飞语音听写服务。'))
    }

    socket.onmessage = (event) => {
      try {
        const payload = JSON.parse(String(event.data)) as XfyunResponse
        if (payload.code !== 0) {
          finish(new Error(payload.message || `讯飞返回错误码 ${payload.code}`))
          return
        }
        const text = extractXfyunText(payload)
        if (text) chunks.push(text)
        if (payload.data?.status === 2) finish()
      } catch (error) {
        finish(error instanceof Error ? error : new Error(String(error)))
      }
    }

    socket.onclose = () => {
      if (!settled && opened) finish()
    }
  })
}

async function sendXfyunAudioFrames(socket: WebSocket, pcm: Buffer, config: XfyunVoiceConfig): Promise<void> {
  let offset = 0
  let status = 0
  while (offset < pcm.byteLength) {
    const chunk = pcm.subarray(offset, offset + XFYUN_FRAME_BYTES)
    socket.send(
      JSON.stringify({
        ...(status === 0
          ? {
              common: { app_id: config.appId },
              business: {
                language: 'zh_cn',
                domain: 'iat',
                accent: 'mandarin',
                vad_eos: 5000
              }
            }
          : {}),
        data: {
          status,
          format: 'audio/L16;rate=16000',
          encoding: 'raw',
          audio: chunk.toString('base64')
        }
      })
    )
    status = 1
    offset += XFYUN_FRAME_BYTES
    await sleep(40)
  }

  socket.send(
    JSON.stringify({
      data: {
        status: 2,
        format: 'audio/L16;rate=16000',
        encoding: 'raw',
        audio: ''
      }
    })
  )
}

function buildXfyunAuthUrl(config: XfyunVoiceConfig): string {
  const date = new Date().toUTCString()
  const signatureOrigin = `host: ${XFYUN_IAT_HOST}\ndate: ${date}\nGET ${XFYUN_IAT_PATH} HTTP/1.1`
  const signature = createHmac('sha256', config.apiSecret)
    .update(signatureOrigin)
    .digest('base64')
  const authorizationOrigin = `api_key="${config.apiKey}", algorithm="hmac-sha256", headers="host date request-line", signature="${signature}"`
  const params = new URLSearchParams({
    authorization: Buffer.from(authorizationOrigin).toString('base64'),
    date,
    host: XFYUN_IAT_HOST
  })
  return `${XFYUN_IAT_URL}?${params.toString()}`
}

function audioFromDataUrl(dataUrl: string): Buffer {
  const trimmed = dataUrl.trim()
  const commaIndex = trimmed.indexOf(',')
  if (!trimmed.startsWith('data:') || commaIndex < 0) {
    throw new Error('录音数据格式不正确。')
  }
  const header = trimmed.slice(5, commaIndex)
  const body = trimmed.slice(commaIndex + 1).replace(/\s/g, '')
  if (!header.toLowerCase().includes('base64') || !body) {
    throw new Error('录音数据格式不正确。')
  }
  return Buffer.from(body, 'base64')
}

function wavToPcm16k(wav: Buffer): Buffer {
  if (wav.toString('ascii', 0, 4) !== 'RIFF' || wav.toString('ascii', 8, 12) !== 'WAVE') {
    throw new Error('讯飞听写当前需要 WAV 录音数据。')
  }

  let offset = 12
  let channels = 1
  let sampleRate = XFYUN_SAMPLE_RATE
  let bitsPerSample = 16
  let data: Buffer | null = null

  while (offset + 8 <= wav.byteLength) {
    const chunkId = wav.toString('ascii', offset, offset + 4)
    const chunkSize = wav.readUInt32LE(offset + 4)
    const chunkStart = offset + 8
    const chunkEnd = Math.min(chunkStart + chunkSize, wav.byteLength)

    if (chunkId === 'fmt ') {
      const audioFormat = wav.readUInt16LE(chunkStart)
      channels = wav.readUInt16LE(chunkStart + 2)
      sampleRate = wav.readUInt32LE(chunkStart + 4)
      bitsPerSample = wav.readUInt16LE(chunkStart + 14)
      if (audioFormat !== 1 || bitsPerSample !== 16) {
        throw new Error('讯飞听写当前仅支持 16-bit PCM WAV。')
      }
    } else if (chunkId === 'data') {
      data = wav.subarray(chunkStart, chunkEnd)
    }

    offset = chunkEnd + (chunkSize % 2)
  }

  if (!data) throw new Error('WAV 录音缺少音频数据。')
  const mono = pcm16ToMonoSamples(data, channels)
  const resampled = resamplePcm16(mono, sampleRate, XFYUN_SAMPLE_RATE)
  return samplesToPcm16Buffer(resampled)
}

function pcm16ToMonoSamples(data: Buffer, channels: number): Int16Array {
  const safeChannels = Math.max(1, channels)
  const frameCount = Math.floor(data.byteLength / 2 / safeChannels)
  const output = new Int16Array(frameCount)
  for (let frame = 0; frame < frameCount; frame += 1) {
    let sum = 0
    for (let channel = 0; channel < safeChannels; channel += 1) {
      sum += data.readInt16LE((frame * safeChannels + channel) * 2)
    }
    output[frame] = Math.max(-32768, Math.min(32767, Math.round(sum / safeChannels)))
  }
  return output
}

function resamplePcm16(input: Int16Array, fromRate: number, toRate: number): Int16Array {
  if (fromRate === toRate) return input
  const outputLength = Math.max(1, Math.round((input.length * toRate) / fromRate))
  const output = new Int16Array(outputLength)
  const ratio = fromRate / toRate
  for (let index = 0; index < outputLength; index += 1) {
    const sourceIndex = index * ratio
    const before = Math.floor(sourceIndex)
    const after = Math.min(input.length - 1, before + 1)
    const weight = sourceIndex - before
    output[index] = Math.round(input[before] * (1 - weight) + input[after] * weight)
  }
  return output
}

function samplesToPcm16Buffer(samples: Int16Array): Buffer {
  const buffer = Buffer.alloc(samples.length * 2)
  for (let index = 0; index < samples.length; index += 1) {
    buffer.writeInt16LE(samples[index], index * 2)
  }
  return buffer
}

interface XfyunResponse {
  code: number
  message?: string
  data?: {
    status?: number
    result?: {
      ws?: Array<{
        cw?: Array<{
          w?: string
        }>
      }>
    }
  }
}

function extractXfyunText(payload: XfyunResponse): string {
  return (payload.data?.result?.ws || [])
    .map((item) => item.cw?.[0]?.w || '')
    .join('')
}

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms))
}
