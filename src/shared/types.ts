export type MessageRole = 'user' | 'assistant' | 'system'

export type OfficeKind = 'word' | 'excel' | 'powerpoint' | 'unknown'
export type WorkflowStageId = 'select' | 'plan' | 'execute' | 'verify' | 'done'
export type WorkflowStageStatus = 'idle' | 'active' | 'complete' | 'blocked'
export type AiProvider = 'openai' | 'anthropic'
export type AiWireApi = 'chat_completions' | 'responses' | 'anthropic_messages'

export type OfficeActionType =
  | 'create_docx'
  | 'create_xlsx'
  | 'create_pptx'
  | 'revise_docx'
  | 'revise_xlsx'
  | 'revise_pptx'
  | 'copy_file'
  | 'find_files'
  | 'replace_text'
  | 'update_excel_cells'
  | 'run_javascript'
  | 'skill_task'

export interface ChatMessage {
  id: string
  role: MessageRole
  content: string
  createdAt: string
  targetFiles?: WorkspaceFile[]
  attachments?: ChatAttachment[]
  actions?: ActionResult[]
  events?: ChatProcessEvent[]
}

export interface ChatProcessEvent {
  id: string
  createdAt: string
  message: string
  status?: 'running' | 'done' | 'error'
}

export interface ChatAttachment {
  id: string
  kind: 'image' | 'file' | 'audio' | 'video'
  name?: string
  mimeType: string
  dataUrl: string
  size?: number
}

export interface WorkspaceFile {
  id: string
  name: string
  path: string
  kind: OfficeKind
  size?: number
  modifiedAt?: string
  isDirectory?: boolean
}

export interface WorkspaceSnapshot {
  rootPath: string
  files: WorkspaceFile[]
  truncated: boolean
}

export interface WorkflowStage {
  id: WorkflowStageId
  label: string
  status: WorkflowStageStatus
  detail?: string
}

export interface DocumentWorkflowRun {
  id: string
  title: string
  targetPaths: string[]
  stages: WorkflowStage[]
  createdAt: string
  taskFilePath?: string
}

export interface AppSettings {
  provider: AiProvider
  wireApi: AiWireApi
  baseUrl: string
  model: string
  workspacePath: string
  residentMode: boolean
  hasApiKey: boolean
  apiConfigSource?: string
  usesExternalApiConfig: boolean
  cachedMessageCount: number
}

export interface DoclingStatus {
  installed: boolean
  engine?: string
  installCommand?: string
  message: string
}

export interface DoclingInstallResult extends DoclingStatus {
  ok: boolean
  log?: string
}

export interface SettingsPatch {
  provider?: AiProvider
  wireApi?: AiWireApi
  baseUrl?: string
  model?: string
  workspacePath?: string
  residentMode?: boolean
  apiKey?: string
  clearApiKey?: boolean
}

export interface ExternalAiConfig {
  provider: AiProvider
  wireApi?: AiWireApi
  baseUrl: string
  model: string
  apiKey: string
  source: string
}

export interface WordSection {
  heading?: string
  paragraphs?: string[]
  bullets?: string[]
}

export interface SheetPlan {
  name?: string
  columns?: string[]
  rows?: Array<Record<string, unknown> | unknown[]>
  summary?: string
}

export interface SlidePlan {
  title?: string
  bullets?: string[]
  notes?: string
}

export interface OfficeAction {
  type: OfficeActionType
  title?: string
  filename?: string
  sourcePath?: string
  destinationPath?: string
  targetPaths?: string[]
  query?: string
  skillName?: 'documents' | 'spreadsheets' | 'presentations'
  instructions?: string
  expectedOutput?: string
  script?: string
  replacements?: Array<{ find: string; replace: string }>
  cellUpdates?: Array<{ sheet?: string; cell: string; value: unknown }>
  sections?: WordSection[]
  sheets?: SheetPlan[]
  slides?: SlidePlan[]
}

export interface PlannedResponse {
  reply: string
  actions: OfficeAction[]
  executedResults?: ActionResult[]
}

export interface GeneratedFile {
  id: string
  name: string
  path: string
  kind: OfficeKind
  size: number
  createdAt: string
  summary?: string
}

export interface ActionResult {
  ok: boolean
  actionType: OfficeActionType
  kind: OfficeKind
  summary: string
  file?: GeneratedFile
  workflow?: DocumentWorkflowRun
  error?: string
}

export interface ChatRequest {
  requestId?: string
  messages: ChatMessage[]
  targetFiles: WorkspaceFile[]
  workspaceSnapshot?: WorkspaceSnapshot
}

export interface ChatResponse {
  message: ChatMessage
  generatedFiles: GeneratedFile[]
}

export interface ChatHistorySnapshot {
  messages: ChatMessage[]
  updatedAt?: string
}

export interface ChatStreamEvent {
  requestId: string
  type: 'status' | 'done' | 'error' | 'step-done'
  message?: string
}

export interface UpdateAsset {
  name: string
  url: string
  size?: number
}

export interface UpdateStatus {
  currentVersion: string
  latestVersion?: string
  available: boolean
  releaseUrl?: string
  asset?: UpdateAsset
  message: string
}

export interface UpdateDownloadResult {
  ok: boolean
  message: string
  filePath?: string
  releaseUrl?: string
}
