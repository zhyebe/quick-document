import type {
  AiProvider,
  ChatAttachment,
  ChatMessage,
  OfficeAction,
  OfficeActionType,
  PlannedResponse,
  WorkspaceFile,
  WorkspaceSnapshot
} from '@shared/types'

interface PlannerSettings {
  provider: AiProvider
  wireApi: 'chat_completions' | 'responses' | 'anthropic_messages'
  baseUrl: string
  model: string
  apiKey: string
}

interface PlannerInput {
  messages: ChatMessage[]
  targetFiles: WorkspaceFile[]
  workspaceSnapshot?: WorkspaceSnapshot
  settings: PlannerSettings
  skillBrief?: string
}

const SYSTEM_PROMPT = `
You are the workflow planner for Quick Document, a desktop AI document-processing assistant.
The product is not a general chatbot. It only handles Word, Excel, and PowerPoint document workflows.
Users select local folders or file paths. The AI plans a skill-driven workflow that operates directly on those paths.
User instructions can be short. Infer the target file from selected paths, workspace filenames, and filename mentions.
Users may attach screenshots, images, audio, video, or files, but attachments are only context for Office document work.

Return strict JSON only. Do not wrap it in markdown.

Schema:
{
  "reply": "Short Chinese response for the chat UI.",
  "actions": [
    {
      "type": "skill_task | create_docx | create_xlsx | create_pptx | revise_docx | revise_xlsx | revise_pptx",
      "title": "Human title",
      "filename": "safe filename with extension",
      "targetPaths": ["absolute local paths to files or directories"],
      "skillName": "documents | spreadsheets | presentations",
      "instructions": "Clear workflow instructions for the skill executor.",
      "expectedOutput": "Expected output files or in-place changes.",
      "sourcePath": "optional source path for legacy create/revise actions",
      "sections": [{"heading": "string", "paragraphs": ["string"], "bullets": ["string"]}],
      "sheets": [{"name": "string", "columns": ["string"], "rows": [{"Column": "value"}]}],
      "slides": [{"title": "string", "bullets": ["string"], "notes": "string"}]
    }
  ]
}

Rules:
- Prefer skill_task for real document workflows. This lets the local executor call embedded skills directly on local paths.
- Use documents for .docx, spreadsheets for .xlsx, and presentations for .ppt/.pptx.
- Filter every request through the Office skills scope before answering. If the user asks for unrelated general work, politely redirect them to a Word, Excel, or PowerPoint document workflow and return an empty actions array.
- Use screenshots/images/files only to understand document content, target locations, formatting, or edit requirements. Do not become a general image/audio/video assistant.
- If one workflow touches multiple file types, return multiple skill_task actions in execution order.
- Keep instructions concise and executable. Do not ask for long briefs when the user's edit intent is clear.
- Use create_docx/create_xlsx/create_pptx only for simple new-file fallback drafts.
- Use revise_docx/revise_xlsx/revise_pptx only when a legacy executor is enough.
- If the request asks for Word, Excel, and PPT together, return multiple actions.
- Keep filenames concise, cross-platform safe, and include the correct extension.
- Prefer Chinese content unless the user asks for another language.
- If no target path is selected, ask the user to select a document folder or target files and return an empty actions array.
`

export async function planDocumentWork(input: PlannerInput): Promise<PlannedResponse> {
  if (!input.settings.apiKey) {
    return localFallbackPlan(input.messages, input.targetFiles, input.workspaceSnapshot)
  }

  try {
    const content = await callPlannerModel(input)
    return normalizePlan(parseJson(content))
  } catch (error) {
    return {
      reply: `AI 调用没有成功，我先切换到本地基础规划。原因：${
        error instanceof Error ? error.message : String(error)
      }`,
      actions: localFallbackPlan(input.messages, input.targetFiles, input.workspaceSnapshot).actions
    }
  }
}

async function callPlannerModel(input: PlannerInput): Promise<string> {
  if (input.settings.provider === 'anthropic') {
    return callAnthropicPlanner(input)
  }

  if (input.settings.wireApi === 'responses') {
    return callOpenAiResponsesPlanner(input)
  }

  return callOpenAiPlanner(input)
}

async function callOpenAiPlanner(input: PlannerInput): Promise<string> {
  const response = await fetch(`${input.settings.baseUrl.replace(/\/$/, '')}/chat/completions`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${input.settings.apiKey}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify({
        model: input.settings.model,
        temperature: 0.2,
        response_format: { type: 'json_object' },
        messages: buildOpenAiChatMessages(input)
      })
    })

  if (!response.ok) {
    const text = await response.text()
    throw new Error(`OpenAI-compatible request failed: ${response.status} ${text.slice(0, 300)}`)
  }

  const data = (await response.json()) as {
    choices?: Array<{ message?: { content?: string } }>
  }
  return data.choices?.[0]?.message?.content || ''
}

async function callOpenAiResponsesPlanner(input: PlannerInput): Promise<string> {
  const response = await fetch(`${input.settings.baseUrl.replace(/\/$/, '')}/responses`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${input.settings.apiKey}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: input.settings.model,
      temperature: 0.2,
      instructions: buildSystemText(input),
      input: buildOpenAiResponsesInput(input)
    })
  })

  if (!response.ok) {
    const text = await response.text()
    throw new Error(`OpenAI Responses request failed: ${response.status} ${text.slice(0, 300)}`)
  }

  const data = (await response.json()) as {
    output_text?: string
    output?: Array<{
      content?: Array<{ type?: string; text?: string }>
    }>
  }
  return (
    data.output_text ||
    data.output?.flatMap((item) => item.content || []).find((item) => item.text)?.text ||
    ''
  )
}

async function callAnthropicPlanner(input: PlannerInput): Promise<string> {
  const response = await fetch(buildAnthropicUrl(input.settings.baseUrl), {
    method: 'POST',
    headers: {
      'x-api-key': input.settings.apiKey,
      'anthropic-version': '2023-06-01',
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({
      model: input.settings.model,
      max_tokens: 2000,
      temperature: 0.2,
      system: buildSystemText(input),
      messages: buildAnthropicMessages(input)
    })
  })

  if (!response.ok) {
    const text = await response.text()
    throw new Error(`Anthropic-compatible request failed: ${response.status} ${text.slice(0, 300)}`)
  }

  const data = (await response.json()) as {
    content?: Array<{ type?: string; text?: string }>
  }
  return data.content?.find((item) => item.type === 'text' && item.text)?.text || ''
}

function buildAnthropicUrl(baseUrl: string): string {
  const normalized = baseUrl.replace(/\/$/, '')
  return normalized.endsWith('/v1') ? `${normalized}/messages` : `${normalized}/v1/messages`
}

function buildSystemText(input: PlannerInput): string {
  return input.skillBrief
    ? `${SYSTEM_PROMPT}\n\nEmbedded Office skill brief:\n${input.skillBrief}`
    : SYSTEM_PROMPT
}

function buildAnthropicMessages(
  input: PlannerInput
): Array<{ role: 'user' | 'assistant'; content: string | Array<Record<string, unknown>> }> {
  return buildPlannerMessageParts(input)
    .filter((message) => message.role !== 'system')
    .map((message) => ({
      role: message.role === 'assistant' ? 'assistant' : 'user',
      content: message.attachments?.length
        ? [
            { type: 'text', text: message.content },
            ...message.attachments.flatMap((attachment) => anthropicAttachmentParts(attachment))
          ]
        : message.content
    }))
}

function buildOpenAiChatMessages(input: PlannerInput): Array<{ role: string; content: unknown }> {
  return buildPlannerMessageParts(input).map((message) => ({
    role: message.role,
    content: message.attachments?.length
      ? [
          { type: 'text', text: message.content },
          ...message.attachments.flatMap((attachment) => openAiChatAttachmentParts(attachment))
        ]
      : message.content
  }))
}

function buildOpenAiResponsesInput(input: PlannerInput): Array<{ role: string; content: Array<Record<string, unknown>> }> {
  return buildPlannerMessageParts(input)
    .filter((message) => message.role !== 'system')
    .map((message) => ({
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

function buildPlannerMessageParts(
  input: PlannerInput
): Array<{ role: string; content: string; attachments?: ChatAttachment[] }> {
  const recent = input.messages.slice(-8).map((message) => ({
    role: message.role === 'assistant' ? 'assistant' : 'user',
    content: appendAttachmentContext(message.content, message.attachments),
    attachments: message.attachments?.filter((attachment) => attachment.kind === 'image')
  }))

  const targetContext =
    input.targetFiles.length > 0
      ? [
          {
            role: 'user',
            content: `当前选中的目标路径：\n${input.targetFiles
              .map((file) => `- ${file.name} | ${file.kind} | ${file.path}`)
              .join('\n\n---\n\n')}`
          }
        ]
      : []

  const workspaceContext = input.workspaceSnapshot
    ? [
        {
          role: 'user',
          content: `当前工作目录：${input.workspaceSnapshot.rootPath}\n可处理 Office 文件索引：\n${input.workspaceSnapshot.files
            .slice(0, 80)
            .map((file) => `- ${file.name} | ${file.kind} | ${file.path}`)
            .join('\n')}${input.workspaceSnapshot.truncated ? '\n- [index truncated]' : ''}`
        }
      ]
    : []

  const skillBrief = input.skillBrief
    ? [
        {
          role: 'system',
          content: `Embedded Office skill brief:\n${input.skillBrief}`
        }
      ]
    : []

  return [
    { role: 'system', content: SYSTEM_PROMPT },
    ...skillBrief,
    ...workspaceContext,
    ...targetContext,
    ...recent
  ]
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
    return `- ${label}: ${attachment.name || attachment.mimeType} (${attachment.mimeType}, ${formatAttachmentSize(
      attachment.size
    )})`
  })
  return `${content}\n\n用户同时附加了以下材料：\n${lines.join('\n')}`
}

function formatAttachmentSize(size: number | undefined): string {
  if (!size) return 'unknown size'
  if (size < 1024) return `${size} B`
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`
  return `${(size / 1024 / 1024).toFixed(1)} MB`
}

function openAiChatAttachmentParts(attachment: ChatAttachment): Array<Record<string, unknown>> {
  if (attachment.kind !== 'image') return []
  return [
    {
      type: 'image_url',
      image_url: {
        url: attachment.dataUrl
      }
    }
  ]
}

function openAiResponsesAttachmentParts(attachment: ChatAttachment): Array<Record<string, unknown>> {
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
  return [
    {
      type: 'input_image',
      image_url: attachment.dataUrl,
      detail: 'auto'
    }
  ]
}

function anthropicAttachmentParts(attachment: ChatAttachment): Array<Record<string, unknown>> {
  if (attachment.kind !== 'image') return []
  const [header, data] = attachment.dataUrl.split(',')
  const mediaType = header.match(/^data:([^;]+);base64$/)?.[1] || attachment.mimeType
  if (!data) return []
  return [
    {
      type: 'image',
      source: {
        type: 'base64',
        media_type: mediaType,
        data
      }
    }
  ]
}

function parseJson(content: string): unknown {
  const trimmed = content.trim().replace(/^```json\s*/i, '').replace(/```$/i, '').trim()
  try {
    return JSON.parse(trimmed)
  } catch {
    const match = trimmed.match(/\{[\s\S]*\}/)
    if (!match) throw new Error('AI response did not contain JSON.')
    return JSON.parse(match[0])
  }
}

function normalizePlan(value: unknown): PlannedResponse {
  const record = isRecord(value) ? value : {}
  const actions = Array.isArray(record.actions)
    ? record.actions.map(normalizeAction).filter((action): action is OfficeAction => Boolean(action))
    : []

  return {
    reply: typeof record.reply === 'string' ? record.reply : '我已经根据你的请求整理好了处理动作。',
    actions
  }
}

function normalizeAction(value: unknown): OfficeAction | null {
  if (!isRecord(value)) return null
  const type = normalizeActionType(value.type)
  if (!type) return null

  return {
    type,
    title: stringValue(value.title),
    filename: stringValue(value.filename),
    sourcePath: stringValue(value.sourcePath),
    targetPaths: stringArrayValue(value.targetPaths),
    skillName: normalizeSkillName(value.skillName),
    instructions: stringValue(value.instructions),
    expectedOutput: stringValue(value.expectedOutput),
    sections: Array.isArray(value.sections) ? value.sections : undefined,
    sheets: Array.isArray(value.sheets) ? value.sheets : undefined,
    slides: Array.isArray(value.slides) ? value.slides : undefined
  }
}

function normalizeActionType(value: unknown): OfficeActionType | null {
  const raw = typeof value === 'string' ? value : ''
  const aliases: Record<string, OfficeActionType> = {
    create_word: 'create_docx',
    skill_task: 'skill_task',
    task: 'skill_task',
    workflow: 'skill_task',
    create_docx: 'create_docx',
    revise_word: 'revise_docx',
    revise_docx: 'revise_docx',
    update_word: 'revise_docx',
    create_excel: 'create_xlsx',
    create_xlsx: 'create_xlsx',
    revise_excel: 'revise_xlsx',
    revise_xlsx: 'revise_xlsx',
    update_excel: 'revise_xlsx',
    create_ppt: 'create_pptx',
    create_pptx: 'create_pptx',
    create_powerpoint: 'create_pptx',
    revise_ppt: 'revise_pptx',
    revise_pptx: 'revise_pptx',
    revise_powerpoint: 'revise_pptx',
    update_ppt: 'revise_pptx'
  }
  return aliases[raw] || null
}

function localFallbackPlan(
  messages: ChatMessage[],
  targetFiles: WorkspaceFile[],
  workspaceSnapshot?: WorkspaceSnapshot
): PlannedResponse {
  const prompt = messages[messages.length - 1]?.content || ''
  const latestAttachments = messages[messages.length - 1]?.attachments || []
  const normalized = prompt.toLowerCase()
  const wantsWord = /word|docx|文档|报告|合同|方案/.test(normalized)
  const wantsExcel = /excel|xlsx|表格|台账|清单|预算|数据/.test(normalized)
  const wantsPpt = /ppt|pptx|powerpoint|演示|幻灯|汇报/.test(normalized)
  const wantsAll = /word.*excel.*ppt|三种|全部|全套|office/.test(normalized)
  const title = inferTitle(prompt)
  const actions: OfficeAction[] = []
  const targets = resolveTargets(prompt, targetFiles, workspaceSnapshot)

  if (targets.length === 0 && !wantsWord && !wantsExcel && !wantsPpt && !wantsAll) {
    return {
      reply:
        latestAttachments.length > 0
          ? '我已收到附件。请再选择要处理的 Word、Excel 或 PPT 文件，或说明要把附件内容整理到哪个文档里。'
          : '先选择一个文档目录或目标文件，然后直接告诉我要怎么改 Word、Excel 或 PPT。',
      actions: []
    }
  }

  const groupedTargets = groupTargetsBySkill(targets)

  if (targets.length > 0) {
    groupedTargets.forEach((paths, skillName) => {
      actions.push({
        type: 'skill_task',
        title,
        targetPaths: paths,
        skillName,
        instructions: prompt,
        expectedOutput: '按用户指令直接修改目标文档；如需保留原件，由执行器创建修订副本或备份。'
      })
    })

    return {
      reply: '已根据目录和文件路径生成文档处理工作流。',
      actions
    }
  }

  if (wantsAll || wantsWord || (!wantsExcel && !wantsPpt)) {
    actions.push({
      type: 'create_docx',
      title,
      filename: `${title}.docx`,
      sections: [
        {
          heading: '目标',
          paragraphs: [prompt || '请描述要创建或修改的 Word 文档内容。']
        },
        {
          heading: '建议结构',
          bullets: ['背景与目标', '关键内容', '执行步骤', '风险与下一步']
        }
      ]
    })
  }

  if (wantsAll || wantsExcel) {
    actions.push({
      type: 'create_xlsx',
      title,
      filename: `${title}.xlsx`,
      sheets: [
        {
          name: '任务清单',
          columns: ['序号', '事项', '说明', '状态'],
          rows: [
            { 序号: 1, 事项: '需求整理', 说明: prompt || '根据对话补充需求', 状态: '草稿' },
            { 序号: 2, 事项: '内容生成', 说明: '配置 AI Key 后可自动扩展', 状态: '待处理' },
            { 序号: 3, 事项: '人工确认', 说明: '打开文件检查并继续对话修改', 状态: '待确认' }
          ]
        }
      ]
    })
  }

  if (wantsAll || wantsPpt) {
    actions.push({
      type: 'create_pptx',
      title,
      filename: `${title}.pptx`,
      slides: [
        { title, bullets: ['目标与背景', prompt || '请描述演示主题'] },
        { title: '核心内容', bullets: ['关键观点一', '关键观点二', '关键观点三'] },
        { title: '下一步', bullets: ['补充材料', '完善页面', '导出并分享'] }
      ]
    })
  }

  return {
    reply:
      actions.length > 0
        ? '当前未配置 AI Key，我用本地规则先生成一版基础文档或文档处理任务。'
        : '选择目录后直接说要修改哪个文档、哪个位置、怎么改即可。',
    actions
  }
}

function resolveTargets(
  prompt: string,
  targetFiles: WorkspaceFile[],
  workspaceSnapshot?: WorkspaceSnapshot
): WorkspaceFile[] {
  if (targetFiles.length > 0) return targetFiles
  if (!workspaceSnapshot) return []

  const normalizedPrompt = prompt.toLowerCase()
  const mentioned = workspaceSnapshot.files.filter((file) => {
    const name = file.name.toLowerCase()
    const stem = name.replace(/\.(docx|xlsx|ppt|pptx)$/i, '')
    return normalizedPrompt.includes(name) || normalizedPrompt.includes(stem)
  })

  if (mentioned.length > 0) return mentioned.slice(0, 12)

  const extensionMatched = workspaceSnapshot.files.filter((file) => {
    if (file.kind === 'word') return /word|docx|文档|报告|合同|方案/.test(normalizedPrompt)
    if (file.kind === 'excel') return /excel|xlsx|表格|台账|清单|预算|数据/.test(normalizedPrompt)
    if (file.kind === 'powerpoint') return /ppt|pptx|powerpoint|演示|幻灯|汇报/.test(normalizedPrompt)
    return false
  })

  return extensionMatched.length === 1 ? extensionMatched : []
}

function groupTargetsBySkill(targets: WorkspaceFile[]): Map<'documents' | 'spreadsheets' | 'presentations', string[]> {
  const groups = new Map<'documents' | 'spreadsheets' | 'presentations', string[]>()

  targets.forEach((target) => {
    const skillName = skillNameForKind(target.kind)
    if (!skillName) return
    groups.set(skillName, [...(groups.get(skillName) || []), target.path])
  })

  return groups
}

function skillNameForKind(kind: WorkspaceFile['kind']): 'documents' | 'spreadsheets' | 'presentations' | null {
  if (kind === 'word') return 'documents'
  if (kind === 'excel') return 'spreadsheets'
  if (kind === 'powerpoint') return 'presentations'
  return null
}

function inferTitle(prompt: string): string {
  const compact = prompt
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
  if (!compact) return 'Quick Document Draft'
  return compact.slice(0, 34)
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value))
}

function stringValue(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() ? value.trim() : undefined
}

function stringArrayValue(value: unknown): string[] | undefined {
  if (!Array.isArray(value)) return undefined
  const strings = value.filter(
    (item): item is string => typeof item === 'string' && item.trim().length > 0
  )
  return strings.length > 0 ? strings : undefined
}

function normalizeSkillName(value: unknown): OfficeAction['skillName'] {
  if (value === 'documents' || value === 'spreadsheets' || value === 'presentations') return value
  return undefined
}
