import { app } from 'electron'
import { mkdirSync, readdirSync, writeFileSync } from 'node:fs'
import { basename, dirname, join, resolve, sep } from 'node:path'
import type {
  ActionResult,
  AiProvider,
  ChatAttachment,
  ChatMessage,
  OfficeAction,
  WorkspaceFile,
  WorkspaceSnapshot
} from '@shared/types'
import { executeDocumentAction } from './documentService'
import { scanWorkspace } from './workspaceFiles'

interface AgentSettings {
  provider: AiProvider
  wireApi: 'chat_completions' | 'responses' | 'anthropic_messages'
  baseUrl: string
  model: string
  apiKey: string
  requestMaxRetries?: number
  requestTimeoutMs?: number
}

export interface DocumentAgentInput {
  messages: ChatMessage[]
  targetFiles: WorkspaceFile[]
  workspaceSnapshot: WorkspaceSnapshot
  documentPreviewContext?: string
  skillBrief?: string
  settings: AgentSettings
  signal?: AbortSignal
  onProgress?: (message: string) => void
  onAssistantDelta?: (delta: string) => void
  consumeGuidance?: () => ChatMessage[]
}

export interface DocumentAgentOutput {
  reply: string
  actionResults: ActionResult[]
}

interface AgentToolCall {
  id: string
  name: string
  argumentsText: string
}

interface AgentToolResult {
  ok: boolean
  message: string
  files?: WorkspaceFile[]
  actionResult?: ActionResult
  skillPath?: string
  error?: string
}

const DEFAULT_AGENT_TURN_LIMIT = Number.POSITIVE_INFINITY
const DEFAULT_AI_REQUEST_TIMEOUT_MS = 120_000
const UNLIMITED_RETRY_LABEL = '持续'

const DOCUMENT_AGENT_PROMPT = `
You are Quick Document's AI document agent. Behave like a focused Codex-style agent for local Word, Excel, and PowerPoint work.

The desktop app does not plan workflow for you. You decide what to inspect, copy, edit, create, and verify by calling the provided tools.

Core rules:
- The user's latest instruction is the source of truth.
- Work only on document tasks: Word .docx, Excel .xlsx/.xls/.csv, PowerPoint .ppt/.pptx.
- Use Documents skill guidance for .docx, Spreadsheets skill guidance for .xlsx/.xls/.csv, and Presentations skill guidance for .ppt/.pptx.
- If the user says modify a specific file, modify that file directly. Do not create a copy unless the user asks for a copy/version/backup.
- If the user says copy, preserve the file structure by copying the actual file first, then continue later edits on the copied file only when that matches the request.
- Before relying on an inferred filename, call list_workspace_files and use the actual current path.
- If a tool fails, read the error and continue in the same conversation. Do not repeat the same wrong path.
- Prefer run_document_script for all real document changes. The script can use quickDocument.ExcelJS, quickDocument.docx, quickDocument.pptxgen, quickDocument.JSZip, quickDocument.fs, and quickDocument.path.
- In packaged desktop mode, prefer quickDocument.ExcelJS/JSZip/docx/pptxgen instead of bare package imports. If you import those packages anyway, the tool will try to map common imports for compatibility.
- Scripts must write files themselves and call quickDocument.writeResult({ filePath }) for the primary changed file.
- If the user asks to create or update a SKILL, use write_skill_file and author the SKILL content yourself. The app only stores it.
- Keep final Chinese replies concise and result-focused. Do not output JSON workflow to the user.
`.trim()

export async function runDocumentAgent(input: DocumentAgentInput): Promise<DocumentAgentOutput> {
  if (!input.settings.apiKey) {
    return {
      reply: '还没有可用的 AI Key，不能执行 AI 文档修改。请先配置 API Key 或启用 cc-switch/Codex 配置。',
      actionResults: []
    }
  }

  try {
    ensureNotCancelled(input.signal)
    if (input.settings.provider === 'anthropic') return await runAnthropicAgent(input)
    if (input.settings.wireApi === 'responses') {
      try {
        return await runOpenAiResponsesAgent(input)
      } catch (error) {
        if (!shouldFallbackToChatCompletions(error)) throw error
        input.onProgress?.('Responses 工具调用不可用，正在切换到 OpenAI 兼容聊天工具调用...')
        return await runOpenAiChatAgent(input)
      }
    }
    return await runOpenAiChatAgent(input)
  } catch (error) {
    const reply = formatAgentFailureMessage(error)
    input.onProgress?.(reply)
    return {
      reply,
      actionResults: []
    }
  }
}

async function runOpenAiChatAgent(input: DocumentAgentInput): Promise<DocumentAgentOutput> {
  const actionResults: ActionResult[] = []
  const messages: Array<Record<string, unknown>> = buildOpenAiChatMessages(input)
  const maxTurns = agentTurnLimit()

  for (let turn = 0; turn < maxTurns; turn += 1) {
    ensureNotCancelled(input.signal)
    appendOpenAiChatGuidance(input, messages)
    const assistantMessage = await callOpenAiChatTurn(input, messages)
    const toolCalls = extractOpenAiChatToolCalls(assistantMessage)
    const content = stringValue(assistantMessage.content)

    if (toolCalls.length === 0) {
      if (appendOpenAiChatGuidance(input, messages, content)) continue
      if (shouldAskAiToRepairAfterToolFailure(actionResults)) {
        input.onProgress?.('文档工具上一步失败，AI 没有继续调用工具，正在要求它换方法重试...')
        messages.push({
          role: 'assistant',
          content: content || ''
        })
        messages.push({
          role: 'user',
          content: buildToolFailureRepairPrompt(actionResults)
        })
        continue
      }
      const reply = content || buildDefaultReply(actionResults)
      if (!content) input.onAssistantDelta?.(reply)
      return {
        reply,
        actionResults: visibleActionResults(actionResults)
      }
    }

    messages.push({
      role: 'assistant',
      content: content || '',
      tool_calls: assistantMessage.tool_calls
    })

    for (const toolCall of toolCalls) {
      ensureNotCancelled(input.signal)
      const result = await executeAgentTool(input, toolCall, actionResults)
      messages.push({
        role: 'tool',
        tool_call_id: toolCall.id,
        content: JSON.stringify(toolResultForModel(result))
      })
    }
  }

  return {
    reply: buildTurnLimitReply(maxTurns, actionResults),
    actionResults: visibleActionResults(actionResults)
  }
}

async function callOpenAiChatTurn(
  input: DocumentAgentInput,
  messages: Array<Record<string, unknown>>
): Promise<Record<string, unknown>> {
  return withRetries(input, async () => {
    const response = await fetchWithTimeout(
      `${input.settings.baseUrl.replace(/\/$/, '')}/chat/completions`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${input.settings.apiKey}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          model: input.settings.model,
          temperature: 0.2,
          messages,
          tools: openAiTools(),
          tool_choice: 'auto',
          parallel_tool_calls: false,
          stream: true
        })
      },
      input.settings.requestTimeoutMs,
      input.signal
    )
    if (!response.ok) {
      const text = await response.text()
      throw createRemoteRequestError('OpenAI-compatible request', response.status, text)
    }
    if (!isEventStreamResponse(response)) {
      const data = (await response.json()) as { choices?: Array<{ message?: Record<string, unknown> }> }
      const message = data.choices?.[0]?.message || {}
      const content = stringValue(message.content)
      if (content) input.onAssistantDelta?.(content)
      return message
    }

    return readOpenAiChatStream(response, input.onAssistantDelta)
  })
}

interface ChatCompletionToolCallDraft {
  id?: string
  type?: string
  name: string
  argumentsText: string
}

function isEventStreamResponse(response: Response): boolean {
  return /text\/event-stream|stream/i.test(response.headers.get('content-type') || '')
}

async function readOpenAiChatStream(
  response: Response,
  onDelta?: (delta: string) => void
): Promise<Record<string, unknown>> {
  if (!response.body) return {}
  const reader = response.body.getReader()
  const decoder = new TextDecoder()
  const toolCalls = new Map<number, ChatCompletionToolCallDraft>()
  let buffer = ''
  let raw = ''
  let content = ''
  let sawSseData = false

  const processEvent = (event: string): void => {
    const dataLines = event
      .split(/\r?\n/)
      .filter((line) => line.startsWith('data:'))
      .map((line) => line.slice(5).trim())
    if (dataLines.length === 0) return
    sawSseData = true

    for (const dataLine of dataLines) {
      if (!dataLine || dataLine === '[DONE]') continue
      let data: unknown
      try {
        data = JSON.parse(dataLine)
      } catch {
        continue
      }
      if (!isRecord(data)) continue
      const choices = Array.isArray(data.choices) ? data.choices : []
      for (const choice of choices) {
        if (!isRecord(choice) || !isRecord(choice.delta)) continue
        const delta = choice.delta
        const textDelta = stringValue(delta.content) || ''
        if (textDelta) {
          content += textDelta
          onDelta?.(textDelta)
        }

        const calls = Array.isArray(delta.tool_calls) ? delta.tool_calls : []
        for (const call of calls) {
          if (!isRecord(call)) continue
          const rawIndex = typeof call.index === 'number' ? call.index : Number(call.index)
          const index = Number.isInteger(rawIndex)
            ? rawIndex
            : toolCalls.size > 0
              ? toolCalls.size - 1
              : 0
          const current = toolCalls.get(index) || { name: '', argumentsText: '' }
          const fn = isRecord(call.function) ? call.function : {}
          toolCalls.set(index, {
            id: stringValue(call.id) || current.id,
            type: stringValue(call.type) || current.type || 'function',
            name: stringValue(fn.name) || current.name,
            argumentsText: current.argumentsText + (stringValue(fn.arguments) || '')
          })
        }
      }
    }
  }

  while (true) {
    const { done, value } = await reader.read()
    if (done) break
    const chunk = decoder.decode(value, { stream: true })
    raw += chunk
    buffer += chunk
    const events = buffer.split(/\n\n/)
    buffer = events.pop() || ''
    events.forEach(processEvent)
  }

  if (buffer.trim()) processEvent(buffer)
  if (!sawSseData) {
    const message = parseBufferedChatCompletion(raw)
    const fallbackContent = stringValue(message.content)
    if (fallbackContent) onDelta?.(fallbackContent)
    return message
  }

  const normalizedToolCalls = Array.from(toolCalls.entries())
    .sort(([left], [right]) => left - right)
    .map(([, call]) => ({
      id: call.id || `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      type: call.type || 'function',
      function: {
        name: call.name,
        arguments: call.argumentsText || '{}'
      }
    }))
    .filter((call) => call.function.name)

  return {
    role: 'assistant',
    content,
    ...(normalizedToolCalls.length > 0 ? { tool_calls: normalizedToolCalls } : {})
  }
}

function parseBufferedChatCompletion(raw: string): Record<string, unknown> {
  try {
    const data = JSON.parse(raw) as { choices?: Array<{ message?: Record<string, unknown> }> }
    return data.choices?.[0]?.message || {}
  } catch {
    return {}
  }
}

async function runOpenAiResponsesAgent(input: DocumentAgentInput): Promise<DocumentAgentOutput> {
  const actionResults: ActionResult[] = []
  let previousResponseId = ''
  let nextInput: unknown = buildOpenAiResponsesInput(input)
  const maxTurns = agentTurnLimit()

  for (let turn = 0; turn < maxTurns; turn += 1) {
    ensureNotCancelled(input.signal)
    const guidanceInput = consumeOpenAiResponsesGuidance(input)
    if (guidanceInput.length > 0) nextInput = mergeResponsesInput(nextInput, guidanceInput)
    const responseData = await callOpenAiResponsesTurn(input, nextInput, previousResponseId)
    previousResponseId = stringValue(responseData.id) || previousResponseId
    const toolCalls = extractOpenAiResponsesToolCalls(responseData)
    if (toolCalls.length === 0) {
      const reply = extractResponseOutputText(responseData) || buildDefaultReply(actionResults)
      const guidedInput = consumeOpenAiResponsesGuidance(input)
      if (guidedInput.length > 0) {
        input.onAssistantDelta?.(reply ? `${reply}\n` : '')
        nextInput = guidedInput
        continue
      }
      if (shouldAskAiToRepairAfterToolFailure(actionResults)) {
        input.onProgress?.('文档工具上一步失败，AI 没有继续调用工具，正在要求它换方法重试...')
        nextInput = buildOpenAiResponsesRepairInput(actionResults)
        continue
      }
      input.onAssistantDelta?.(reply)
      return {
        reply,
        actionResults: visibleActionResults(actionResults)
      }
    }

    const toolOutputs: Array<Record<string, unknown>> = []
    for (const toolCall of toolCalls) {
      ensureNotCancelled(input.signal)
      const result = await executeAgentTool(input, toolCall, actionResults)
      toolOutputs.push({
        type: 'function_call_output',
        call_id: toolCall.id,
        output: JSON.stringify(toolResultForModel(result))
      })
    }
    nextInput = toolOutputs
  }

  return {
    reply: buildTurnLimitReply(maxTurns, actionResults),
    actionResults: visibleActionResults(actionResults)
  }
}

async function callOpenAiResponsesTurn(
  input: DocumentAgentInput,
  nextInput: unknown,
  previousResponseId: string
): Promise<Record<string, unknown>> {
  return withRetries(input, async () => {
    const body: Record<string, unknown> = {
      model: input.settings.model,
      temperature: 0.2,
      instructions: buildSystemText(input),
      tools: responsesTools(),
      tool_choice: 'auto',
      parallel_tool_calls: false,
      store: false,
      input: nextInput
    }
    if (previousResponseId) body.previous_response_id = previousResponseId

    const response = await fetchWithTimeout(
      `${input.settings.baseUrl.replace(/\/$/, '')}/responses`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${input.settings.apiKey}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(body)
      },
      input.settings.requestTimeoutMs,
      input.signal
    )
    if (!response.ok) {
      const text = await response.text()
      throw createRemoteRequestError('OpenAI Responses request', response.status, text)
    }
    return (await response.json()) as Record<string, unknown>
  })
}

async function runAnthropicAgent(input: DocumentAgentInput): Promise<DocumentAgentOutput> {
  const actionResults: ActionResult[] = []
  const messages: Array<Record<string, unknown>> = buildAnthropicMessages(input)
  const maxTurns = agentTurnLimit()

  for (let turn = 0; turn < maxTurns; turn += 1) {
    ensureNotCancelled(input.signal)
    appendAnthropicGuidance(input, messages)
    const responseData = await callAnthropicTurn(input, messages)
    const content = Array.isArray(responseData.content) ? responseData.content : []
    const toolCalls = content
      .filter((item): item is Record<string, unknown> => isRecord(item) && item.type === 'tool_use')
      .map((item) => ({
        id: stringValue(item.id) || `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        name: stringValue(item.name) || '',
        argumentsText: JSON.stringify(isRecord(item.input) ? item.input : {})
      }))
      .filter((toolCall) => toolCall.name)

    if (toolCalls.length === 0) {
      const reply = content
          .filter((item): item is Record<string, unknown> => isRecord(item) && item.type === 'text')
          .map((item) => stringValue(item.text))
          .filter(Boolean)
          .join('\n')
          .trim() || buildDefaultReply(actionResults)
      if (appendAnthropicGuidance(input, messages, content)) {
        input.onAssistantDelta?.(reply ? `${reply}\n` : '')
        continue
      }
      if (shouldAskAiToRepairAfterToolFailure(actionResults)) {
        input.onProgress?.('文档工具上一步失败，AI 没有继续调用工具，正在要求它换方法重试...')
        messages.push({ role: 'assistant', content })
        messages.push({ role: 'user', content: buildToolFailureRepairPrompt(actionResults) })
        continue
      }
      input.onAssistantDelta?.(reply)
      return {
        reply,
        actionResults: visibleActionResults(actionResults)
      }
    }

    messages.push({ role: 'assistant', content })
    messages.push({
      role: 'user',
      content: await Promise.all(
        toolCalls.map(async (toolCall) => {
          ensureNotCancelled(input.signal)
          const result = await executeAgentTool(input, toolCall, actionResults)
          return {
            type: 'tool_result',
            tool_use_id: toolCall.id,
            content: JSON.stringify(toolResultForModel(result))
          }
        })
      )
    })
  }

  return {
    reply: buildTurnLimitReply(maxTurns, actionResults),
    actionResults: visibleActionResults(actionResults)
  }
}

async function callAnthropicTurn(
  input: DocumentAgentInput,
  messages: Array<Record<string, unknown>>
): Promise<Record<string, unknown>> {
  return withRetries(input, async () => {
    const response = await fetchWithTimeout(
      buildAnthropicUrl(input.settings.baseUrl),
      {
        method: 'POST',
        headers: {
          'x-api-key': input.settings.apiKey,
          'anthropic-version': '2023-06-01',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          model: input.settings.model,
          max_tokens: 4096,
          temperature: 0.2,
          system: buildSystemText(input),
          tools: anthropicTools(),
          messages
        })
      },
      input.settings.requestTimeoutMs,
      input.signal
    )
    if (!response.ok) {
      const text = await response.text()
      throw createRemoteRequestError('Anthropic-compatible request', response.status, text)
    }
    return (await response.json()) as Record<string, unknown>
  })
}

async function executeAgentTool(
  input: DocumentAgentInput,
  toolCall: AgentToolCall,
  actionResults: ActionResult[]
): Promise<AgentToolResult> {
  ensureNotCancelled(input.signal)
  const args = parseToolArguments(toolCall.argumentsText)
  if (!args.ok) {
    return { ok: false, message: '工具参数不是有效 JSON。', error: args.error }
  }

  if (toolCall.name === 'list_workspace_files') {
    input.onProgress?.('AI 正在查看当前目录文件...')
    const query = stringValue(args.value.query)?.toLowerCase().trim()
    const snapshot = scanWorkspace(input.workspaceSnapshot.rootPath)
    const files = query
      ? snapshot.files.filter((file) => `${file.name} ${file.path}`.toLowerCase().includes(query))
      : snapshot.files
    return {
      ok: true,
      message: `当前目录有 ${snapshot.files.length} 个 Office 文件，返回 ${files.length} 个匹配文件。`,
      files: files.slice(0, 200)
    }
  }

  if (toolCall.name === 'run_document_script') {
    const script = stringValue(args.value.script)
    if (!script) return { ok: false, message: '缺少 script。', error: 'script is required' }
    const title = stringValue(args.value.title) || 'AI 文档脚本'
    const action: OfficeAction = {
      type: 'run_javascript',
      title,
      script,
      targetPaths: stringArrayValue(args.value.targetPaths),
      sourcePath: stringValue(args.value.sourcePath),
      destinationPath: stringValue(args.value.destinationPath)
    }
    input.onProgress?.(`AI 正在执行文档工具：${title}`)
    const latestSnapshot = scanWorkspace(input.workspaceSnapshot.rootPath)
    const result = await executeDocumentAction(action, input.workspaceSnapshot.rootPath, latestSnapshot, {
      signal: input.signal
    })
    ensureNotCancelled(input.signal)
    if (result.ok) {
      actionResults.push(result)
      input.onProgress?.(result.summary)
      return {
        ok: true,
        message: result.summary,
        actionResult: result,
        files: scanWorkspace(input.workspaceSnapshot.rootPath).files.slice(0, 200)
      }
    }

    actionResults.push(result)
    input.onProgress?.(`工具执行失败，错误已交回 AI：${result.error || result.summary}`)
    return {
      ok: false,
      message: result.summary,
      error: result.error || result.summary,
      actionResult: result,
      files: scanWorkspace(input.workspaceSnapshot.rootPath).files.slice(0, 200)
    }
  }

  if (toolCall.name === 'list_skills') {
    input.onProgress?.('AI 正在查看本地 SKILL...')
    const skills = listLocalSkills()
    return {
      ok: true,
      message: `当前有 ${skills.length} 个本地 SKILL。`,
      files: skills.map((skill) => ({
        id: skill.path,
        name: skill.name,
        path: skill.path,
        kind: 'unknown'
      }))
    }
  }

  if (toolCall.name === 'write_skill_file') {
    const skillName = stringValue(args.value.skillName)
    const content = typeof args.value.content === 'string' ? args.value.content : ''
    const relativePath = stringValue(args.value.relativePath) || 'SKILL.md'
    if (!skillName) return { ok: false, message: '缺少 skillName。', error: 'skillName is required' }
    if (!content.trim()) return { ok: false, message: '缺少 content。', error: 'content is required' }
    const skillPath = writeLocalSkillFile(skillName, relativePath, content)
    input.onProgress?.(`AI 已写入 SKILL：${skillPath}`)
    return {
      ok: true,
      message: `已写入 SKILL 文件：${skillPath}`,
      skillPath
    }
  }

  return { ok: false, message: `未知工具：${toolCall.name}`, error: `Unknown tool ${toolCall.name}` }
}

function buildSystemText(input: DocumentAgentInput): string {
  return [
    DOCUMENT_AGENT_PROMPT,
    input.skillBrief ? `Embedded Office SKILL context:\n${input.skillBrief}` : ''
  ]
    .filter(Boolean)
    .join('\n\n---\n\n')
}

function appendOpenAiChatGuidance(
  input: DocumentAgentInput,
  messages: Array<Record<string, unknown>>,
  assistantContent?: string
): boolean {
  const guidance = consumeGuidanceMessages(input)
  if (guidance.length === 0) return false
  if (typeof assistantContent === 'string') {
    messages.push({ role: 'assistant', content: assistantContent })
  }
  messages.push(...guidance.map(openAiChatGuidanceMessage))
  return true
}

function consumeOpenAiResponsesGuidance(input: DocumentAgentInput): Array<Record<string, unknown>> {
  return consumeGuidanceMessages(input).map(openAiResponsesGuidanceMessage)
}

function mergeResponsesInput(current: unknown, guidance: Array<Record<string, unknown>>): unknown {
  if (guidance.length === 0) return current
  if (Array.isArray(current)) return [...current, ...guidance]
  return [current, ...guidance].filter(Boolean)
}

function appendAnthropicGuidance(
  input: DocumentAgentInput,
  messages: Array<Record<string, unknown>>,
  assistantContent?: unknown
): boolean {
  const guidance = consumeGuidanceMessages(input)
  if (guidance.length === 0) return false
  if (typeof assistantContent !== 'undefined') {
    messages.push({ role: 'assistant', content: assistantContent })
  }
  messages.push(...guidance.map(anthropicGuidanceMessage))
  return true
}

function consumeGuidanceMessages(input: DocumentAgentInput): ChatMessage[] {
  return (input.consumeGuidance?.() || [])
    .filter((message) => message.role === 'user')
    .map((message) => ({
      ...message,
      content: appendAttachmentContext(message.content, message.attachments),
      attachments: message.attachments?.filter((attachment) =>
        attachment.kind === 'image' || attachment.kind === 'file' || attachment.kind === 'audio'
      )
    }))
}

function openAiChatGuidanceMessage(message: ChatMessage): Record<string, unknown> {
  return {
    role: 'user',
    content: message.attachments?.length
      ? [
          { type: 'text', text: message.content },
          ...message.attachments.flatMap((attachment) => openAiChatAttachmentParts(attachment))
        ]
      : message.content
  }
}

function openAiResponsesGuidanceMessage(message: ChatMessage): Record<string, unknown> {
  return {
    role: 'user',
    content: [
      { type: 'input_text', text: message.content },
      ...(message.attachments || []).flatMap((attachment) => openAiResponsesAttachmentParts(attachment))
    ]
  }
}

function anthropicGuidanceMessage(message: ChatMessage): Record<string, unknown> {
  return {
    role: 'user',
    content: message.attachments?.length
      ? [
          { type: 'text', text: message.content },
          ...message.attachments.flatMap((attachment) => anthropicAttachmentParts(attachment))
        ]
      : message.content
  }
}

function buildOpenAiChatMessages(input: DocumentAgentInput): Array<Record<string, unknown>> {
  return [
    { role: 'system', content: buildSystemText(input) },
    ...buildAgentContextMessages(input).map((message) => ({
      role: message.role,
      content: message.attachments?.length
        ? [
            { type: 'text', text: message.content },
            ...message.attachments.flatMap((attachment) => openAiChatAttachmentParts(attachment))
          ]
        : message.content
    }))
  ]
}

function buildOpenAiResponsesInput(input: DocumentAgentInput): Array<Record<string, unknown>> {
  return buildAgentContextMessages(input).map((message) => ({
    role: message.role === 'assistant' ? 'assistant' : 'user',
    content: [
      {
        type: message.role === 'assistant' ? 'output_text' : 'input_text',
        text: message.content
      },
      ...(message.attachments || []).flatMap((attachment) => openAiResponsesAttachmentParts(attachment))
    ]
  }))
}

function buildOpenAiResponsesRepairInput(actionResults: ActionResult[]): Array<Record<string, unknown>> {
  return [
    {
      role: 'user',
      content: [
        {
          type: 'input_text',
          text: buildToolFailureRepairPrompt(actionResults)
        }
      ]
    }
  ]
}

function buildAnthropicMessages(input: DocumentAgentInput): Array<Record<string, unknown>> {
  return buildAgentContextMessages(input).map((message) => ({
    role: message.role === 'assistant' ? 'assistant' : 'user',
    content: message.attachments?.length
      ? [
          { type: 'text', text: message.content },
          ...message.attachments.flatMap((attachment) => anthropicAttachmentParts(attachment))
        ]
      : message.content
  }))
}

function buildAgentContextMessages(
  input: DocumentAgentInput
): Array<{ role: 'user' | 'assistant'; content: string; attachments?: ChatAttachment[] }> {
  const workspaceContext = `当前允许处理的文档目录：${input.workspaceSnapshot.rootPath}
目录文件索引：
${input.workspaceSnapshot.files
  .slice(0, 300)
  .map((file) => {
    const modified = file.modifiedAt ? ` | modified ${file.modifiedAt}` : ''
    return `- ${file.name} | ${file.kind} | ${file.path}${modified}`
  })
  .join('\n')}${input.workspaceSnapshot.truncated ? '\n- [index truncated]' : ''}`

  const targetContext = input.targetFiles.length
    ? `用户当前选中的目标文件：
${input.targetFiles.map((file) => `- ${file.name} | ${file.kind} | ${file.path}`).join('\n')}`
    : ''

  const previewContext = input.documentPreviewContext
    ? `可能相关的文档预览，仅供理解，不得覆盖用户原话：
${input.documentPreviewContext}`
    : ''

  const recent = input.messages.slice(-10).map((message) => ({
    role: message.role === 'assistant' ? ('assistant' as const) : ('user' as const),
    content: appendAttachmentContext(message.content, message.attachments),
    attachments: message.attachments?.filter((attachment) =>
      attachment.kind === 'image' || attachment.kind === 'file' || attachment.kind === 'audio'
    )
  }))

  return [
    { role: 'user' as const, content: workspaceContext },
    ...(targetContext ? [{ role: 'user' as const, content: targetContext }] : []),
    ...(previewContext ? [{ role: 'user' as const, content: previewContext }] : []),
    ...recent
  ]
}

function openAiTools(): Array<Record<string, unknown>> {
  return [
    {
      type: 'function',
      function: {
        name: 'list_workspace_files',
        description: 'Refresh and list actual Office files in the selected workspace. Use this before relying on an inferred path.',
        parameters: {
          type: 'object',
          additionalProperties: false,
          properties: {
            query: { type: 'string', description: 'Optional filename/path keyword filter.' }
          }
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'run_document_script',
        description:
          'Run an AI-authored JavaScript ES module in the selected document workspace to read, copy, modify, or create Word/Excel/PPT files.',
        parameters: documentScriptParameters()
      }
    },
    {
      type: 'function',
      function: {
        name: 'list_skills',
        description: 'List app-managed local Quick Document skills that the AI can use or update.',
        parameters: {
          type: 'object',
          additionalProperties: false,
          properties: {}
        }
      }
    },
    {
      type: 'function',
      function: {
        name: 'write_skill_file',
        description:
          'Create or update an app-managed Quick Document skill file. Use this when the user asks AI to create a SKILL.',
        parameters: skillFileParameters()
      }
    }
  ]
}

function responsesTools(): Array<Record<string, unknown>> {
  return openAiTools().map((tool) => {
    const fn = isRecord(tool.function) ? tool.function : {}
    return {
      type: 'function',
      name: fn.name,
      description: fn.description,
      parameters: fn.parameters
    }
  })
}

function anthropicTools(): Array<Record<string, unknown>> {
  return responsesTools().map((tool) => ({
    name: tool.name,
    description: tool.description,
    input_schema:
      tool.name === 'run_document_script'
        ? documentScriptParameters()
        : tool.name === 'write_skill_file'
          ? skillFileParameters()
          : {
              type: 'object',
              additionalProperties: false,
              properties: tool.name === 'list_workspace_files' ? { query: { type: 'string' } } : {}
            }
  }))
}

function documentScriptParameters(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['script'],
    properties: {
      title: { type: 'string', description: 'Short Chinese description of this document operation.' },
      sourcePath: { type: 'string', description: 'Optional primary source path.' },
      destinationPath: { type: 'string', description: 'Optional output path or filename.' },
      targetPaths: {
        type: 'array',
        items: { type: 'string' },
        description: 'Optional paths that the script intends to read or modify.'
      },
      script: {
        type: 'string',
        description:
          'JavaScript ES module code. Use quickDocument.ExcelJS/fs/path/JSZip/docx/pptxgen and call quickDocument.writeResult({ filePath }).'
      }
    }
  }
}

function skillFileParameters(): Record<string, unknown> {
  return {
    type: 'object',
    additionalProperties: false,
    required: ['skillName', 'content'],
    properties: {
      skillName: {
        type: 'string',
        description: 'Folder-safe skill name, for example invoice-cleanup or monthly-report.'
      },
      relativePath: {
        type: 'string',
        description: 'File path inside the skill directory. Defaults to SKILL.md.'
      },
      content: {
        type: 'string',
        description: 'Full file content to write.'
      }
    }
  }
}

function extractOpenAiChatToolCalls(message: Record<string, unknown>): AgentToolCall[] {
  const calls = Array.isArray(message.tool_calls) ? message.tool_calls : []
  return calls
    .map((call) => {
      if (!isRecord(call) || !isRecord(call.function)) return null
      return {
        id: stringValue(call.id) || `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        name: stringValue(call.function.name) || '',
        argumentsText: stringValue(call.function.arguments) || '{}'
      }
    })
    .filter((call): call is AgentToolCall => Boolean(call?.name))
}

function extractOpenAiResponsesToolCalls(response: Record<string, unknown>): AgentToolCall[] {
  const output = Array.isArray(response.output) ? response.output : []
  return output
    .map((item) => {
      if (!isRecord(item) || item.type !== 'function_call') return null
      return {
        id: stringValue(item.call_id) || stringValue(item.id) || `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        name: stringValue(item.name) || '',
        argumentsText: stringValue(item.arguments) || '{}'
      }
    })
    .filter((call): call is AgentToolCall => Boolean(call?.name))
}

function toolResultForModel(result: AgentToolResult): Record<string, unknown> {
  return {
    ok: result.ok,
    message: result.message,
    error: result.error,
    skillPath: result.skillPath,
    file: result.actionResult?.file,
    files: result.files?.map((file) => ({
      name: file.name,
      path: file.path,
      kind: file.kind,
      size: file.size,
      modifiedAt: file.modifiedAt
    }))
  }
}

function localSkillsDir(): string {
  return join(app.getPath('userData'), 'skills')
}

function listLocalSkills(): Array<{ name: string; path: string }> {
  try {
    return readdirSync(localSkillsDir(), { withFileTypes: true })
      .filter((entry) => entry.isDirectory())
      .map((entry) => ({
        name: entry.name,
        path: join(localSkillsDir(), entry.name, 'SKILL.md')
      }))
  } catch {
    return []
  }
}

function writeLocalSkillFile(skillName: string, relativePath: string, content: string): string {
  const safeName = sanitizeSkillPathPart(skillName) || 'quick-document-skill'
  const safeRelativePath =
    relativePath
      .split(/[\\/]+/)
      .filter((part) => part && part !== '.' && part !== '..')
      .map((part) => sanitizeSkillPathPart(part) || 'file')
      .join(sep) || 'SKILL.md'
  const root = resolve(localSkillsDir(), safeName)
  const target = resolve(root, safeRelativePath)
  if (target !== root && !target.startsWith(root + sep)) {
    throw new Error('SKILL 文件路径超出允许目录。')
  }
  mkdirSync(dirname(target), { recursive: true })
  writeFileSync(target, content, 'utf8')
  return target
}

function sanitizeSkillPathPart(value: string): string {
  return basename(value)
    .replace(/[<>:"|?*\u0000-\u001f]/g, '-')
    .replace(/^\.+$/, '')
    .trim()
    .slice(0, 120)
}

function visibleActionResults(results: ActionResult[]): ActionResult[] {
  const successes = results.filter((result) => result.ok)
  return successes.length > 0 ? successes : results.slice(-3)
}

function shouldAskAiToRepairAfterToolFailure(results: ActionResult[]): boolean {
  const last = results[results.length - 1]
  return Boolean(last && !last.ok)
}

function buildToolFailureRepairPrompt(results: ActionResult[]): string {
  const recentFailures = results
    .filter((result) => !result.ok)
    .slice(-3)
    .map((result, index) => {
      const error = result.error ? `\n错误：${result.error}` : ''
      return `${index + 1}. ${result.summary}${error}`
    })
    .join('\n\n')

  return `上一轮文档工具执行失败，但用户任务还没有完成。请像 Codex 一样继续处理，不要只解释原因，也不要输出工作流。

请基于下面的失败信息换一种可执行方法继续调用工具：
${recentFailures || '工具返回失败，但没有提供详细错误。'}

要求：
- 如果是路径、文件名或月份判断错误，先调用 list_workspace_files 读取当前目录真实文件。
- 如果要复制文件，先用 quickDocument.fs.copyFileSync 复制真实文件，再只在新文件上继续修改。
- 如果要修改 Word/Excel/PPT，继续调用 run_document_script 直接读写目标文件。
- 在脚本里优先使用 quickDocument.fs/path/ExcelJS/docx/pptxgen/JSZip，不要依赖用户手动处理。
- 完成后必须调用 quickDocument.writeResult({ filePath }) 报告实际改好的文件。`
}

function buildDefaultReply(results: ActionResult[]): string {
  const successes = results.filter((result) => result.ok)
  if (successes.length === 0) return '我没有完成任何文档修改。'
  return successes.map((result) => result.summary).join('\n')
}

function buildTurnLimitReply(maxTurns: number, results: ActionResult[]): string {
  const limit = Number.isFinite(maxTurns) ? `${maxTurns}` : UNLIMITED_RETRY_LABEL
  return `AI 已达到本轮最大工具调用次数（${limit}）。${buildDefaultReply(results)}`
}

function parseToolArguments(value: string): { ok: true; value: Record<string, unknown> } | { ok: false; error: string } {
  try {
    const parsed = JSON.parse(value || '{}')
    return isRecord(parsed) ? { ok: true, value: parsed } : { ok: false, error: 'arguments must be object' }
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  }
}

function extractResponseOutputText(response: Record<string, unknown>): string {
  if (typeof response.output_text === 'string') return response.output_text
  const output = Array.isArray(response.output) ? response.output : []
  return output
    .flatMap((item) => (isRecord(item) && Array.isArray(item.content) ? item.content : []))
    .map((content) => (isRecord(content) && typeof content.text === 'string' ? content.text : ''))
    .join('')
    .trim()
}

function appendAttachmentContext(content: string, attachments: ChatAttachment[] | undefined): string {
  if (!attachments?.length) return content
  const lines = attachments.map((attachment) => {
    const label =
      attachment.kind === 'image'
        ? '图片/截图'
        : attachment.kind === 'audio'
          ? '音频'
          : attachment.kind === 'video'
            ? '视频'
            : '文件'
    return `- ${label}: ${attachment.name || attachment.mimeType} (${attachment.mimeType})`
  })
  return `${content}\n\n用户同时附加了以下材料：\n${lines.join('\n')}`
}

function openAiChatAttachmentParts(attachment: ChatAttachment): Array<Record<string, unknown>> {
  if (attachment.kind === 'image') return [{ type: 'image_url', image_url: { url: attachment.dataUrl } }]
  if (attachment.kind === 'audio') {
    const audio = audioAttachmentPayload(attachment)
    return audio ? [{ type: 'input_audio', input_audio: audio }] : []
  }
  return []
}

function openAiResponsesAttachmentParts(attachment: ChatAttachment): Array<Record<string, unknown>> {
  if (attachment.kind === 'audio') {
    const audio = audioAttachmentPayload(attachment)
    return audio ? [{ type: 'input_audio', input_audio: audio }] : []
  }
  if (attachment.kind === 'file') {
    return [
      {
        type: 'input_file',
        filename: attachment.name || `attachment.${attachment.mimeType.split('/')[1] || 'bin'}`,
        file_data: attachment.dataUrl
      }
    ]
  }
  if (attachment.kind !== 'image') return []
  return [{ type: 'input_image', image_url: attachment.dataUrl, detail: 'auto' }]
}

function audioAttachmentPayload(attachment: ChatAttachment): { data: string; format: string } | null {
  const [, data = ''] = attachment.dataUrl.split(',')
  if (!data) return null
  return {
    data,
    format: audioInputFormat(attachment.mimeType, attachment.name)
  }
}

function audioInputFormat(mimeType: string, name?: string): string {
  const value = `${mimeType} ${name || ''}`
  if (/wav|x-wav/i.test(value)) return 'wav'
  if (/mpeg|mp3/i.test(value)) return 'mp3'
  return 'wav'
}

function anthropicAttachmentParts(attachment: ChatAttachment): Array<Record<string, unknown>> {
  if (attachment.kind !== 'image') return []
  const [header, data] = attachment.dataUrl.split(',')
  const mediaType = header.match(/^data:([^;]+);base64$/)?.[1] || attachment.mimeType
  if (!data) return []
  return [{ type: 'image', source: { type: 'base64', media_type: mediaType, data } }]
}

function buildAnthropicUrl(baseUrl: string): string {
  const normalized = baseUrl.replace(/\/$/, '')
  return normalized.endsWith('/v1') ? `${normalized}/messages` : `${normalized}/v1/messages`
}

function shouldFallbackToChatCompletions(error: unknown): boolean {
  if (!(error instanceof Error)) return false
  return /Responses request failed|\/responses|tools?|function_call|unsupported|not found|Cannot POST|404|405|400/i.test(error.message)
}

function createRemoteRequestError(label: string, status: number, text: string): Error {
  const error = new Error(`${label} failed: ${status} ${sanitizeRemoteError(text)}`)
  error.name = `RemoteRequest${status}`
  return error
}

function sanitizeRemoteError(text: string): string {
  const compact = text
    .trim()
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/\s+/g, ' ')
    .slice(0, 500)
  return compact || 'empty response'
}

async function fetchWithTimeout(
  url: string,
  init: RequestInit,
  timeoutMs = DEFAULT_AI_REQUEST_TIMEOUT_MS,
  signal?: AbortSignal
): Promise<Response> {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), Math.max(10_000, timeoutMs))
  const abort = (): void => controller.abort()
  if (signal?.aborted) abort()
  signal?.addEventListener('abort', abort, { once: true })
  try {
    return await fetch(url, {
      ...init,
      signal: controller.signal
    })
  } finally {
    clearTimeout(timer)
    signal?.removeEventListener('abort', abort)
  }
}

async function withRetries<T>(input: DocumentAgentInput, fn: () => Promise<T>): Promise<T> {
  let lastError: unknown
  const total = retryLimit(input.settings.requestMaxRetries)
  for (let index = 0; index < total; index += 1) {
    try {
      ensureNotCancelled(input.signal)
      return await fn()
    } catch (error) {
      lastError = error
      ensureNotCancelled(input.signal)
      if (!isRetryableError(error) || index === total - 1) break
      const delay = retryDelayMs(index)
      input.onProgress?.(retryProgressMessage(error, index + 2, total, delay))
      await waitForRetry(delay, input.signal)
    }
  }
  throw lastError instanceof Error ? lastError : new Error('Agent request failed.')
}

function agentTurnLimit(): number {
  const configured = Number(process.env.QUICK_DOCUMENT_AGENT_MAX_TURNS || '')
  if (Number.isFinite(configured) && configured > 0) return Math.floor(configured)
  return DEFAULT_AGENT_TURN_LIMIT
}

function retryLimit(value: number | undefined): number {
  if (typeof value === 'number' && Number.isFinite(value) && value > 0) return Math.floor(value)
  return Number.POSITIVE_INFINITY
}

function isRetryableError(error: unknown): boolean {
  if (!(error instanceof Error)) return false
  if (/\b401\b|\b403\b|unauthorized|forbidden|invalid[_ -]?api[_ -]?key|authentication|insufficient[_ -]?quota|billing/i.test(error.message)) {
    return false
  }
  return /(?:\b429\b|\b5\d\d\b|\bnetwork\b|\btimeout\b|\babort\b|\bfetch failed\b|ECONNRESET|ETIMEDOUT|ENOTFOUND|EAI_AGAIN)/i.test(
    error.message
  )
}

function retryDelayMs(index: number): number {
  return Math.min(12_000, 800 * 2 ** index)
}

function retryProgressMessage(error: unknown, attempt: number, total: number, delayMs: number): string {
  const detail = retryErrorLabel(error)
  const waitSeconds = Math.max(1, Math.round(delayMs / 1000))
  const totalText = Number.isFinite(total) ? `${total}` : UNLIMITED_RETRY_LABEL
  return `AI 接口暂时不可用（${detail}），${waitSeconds}s 后第 ${attempt}/${totalText} 次自动重试。可点击停止，或继续输入补充引导。`
}

function retryErrorLabel(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error)
  if (/\b429\b|rate limit/i.test(message)) return '限流'
  if (/\b504\b|gateway|timeout|timed out/i.test(message)) return '代理超时'
  if (/\b5\d\d\b/i.test(message)) return '服务端错误'
  if (/network|fetch failed|ECONNRESET|ETIMEDOUT|ENOTFOUND|EAI_AGAIN/i.test(message)) return '网络异常'
  if (/abort/i.test(message)) return '请求超时'
  return '可恢复错误'
}

function waitForRetry(delayMs: number, signal?: AbortSignal): Promise<void> {
  return new Promise((resolve, reject) => {
    if (signal?.aborted) {
      reject(new Error('已手动停止当前处理。'))
      return
    }
    const done = (): void => {
      signal?.removeEventListener('abort', abort)
      resolve()
    }
    const timer = setTimeout(done, delayMs)
    const abort = (): void => {
      clearTimeout(timer)
      signal?.removeEventListener('abort', abort)
      reject(new Error('已手动停止当前处理。'))
    }
    signal?.addEventListener('abort', abort, { once: true })
  })
}

function formatAgentFailureMessage(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error)
  if (/手动停止|cancelled|aborted/i.test(message)) {
    return '已手动停止当前处理，后续步骤没有继续执行。'
  }
  if (/\b504\b|gateway|timeout|timed out|abort/i.test(message)) {
    return 'AI 接口或代理暂时超时，当前没有执行任何文档修改。你可以稍后重试，或者切换到可用的 cc-switch / OpenAI 代理配置。'
  }
  if (/\b401\b|unauthorized|api key|authentication/i.test(message)) {
    return 'AI Key 不可用或认证失败，当前没有执行任何文档修改。请检查 cc-switch / API Key 配置。'
  }
  if (/\b429\b|rate limit|quota/i.test(message)) {
    return 'AI 接口限流或额度不足，当前没有执行任何文档修改。请稍后重试或切换模型。'
  }
  return `AI 调用没有成功，所以没有修改任何文档。原因：${sanitizeRemoteError(message)}`
}

function ensureNotCancelled(signal?: AbortSignal): void {
  if (!signal?.aborted) return
  throw new Error('已手动停止当前处理。')
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value))
}

function stringValue(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() ? value.trim() : undefined
}

function stringArrayValue(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) return undefined
  const strings = value.filter((item): item is string => typeof item === 'string' && item.trim().length > 0)
  return strings.length ? strings : undefined
}
