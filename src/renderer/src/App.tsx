import {
  Bot,
  Check,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  KeyRound,
  Loader2,
  MessageSquareText,
  MonitorUp,
  Presentation,
  RefreshCw,
  Send,
  Settings,
  Sparkles,
  X
} from 'lucide-react'
import { FormEvent, KeyboardEvent, useEffect, useMemo, useRef, useState } from 'react'
import type {
  ActionResult,
  AiProvider,
  AppSettings,
  ChatMessage,
  GeneratedFile,
  OfficeKind,
  SettingsPatch,
  WorkspaceFile,
  WorkspaceSnapshot
} from '@shared/types'

const welcomeMessage: ChatMessage = {
  id: 'welcome',
  role: 'assistant',
  createdAt: new Date().toISOString(),
  content: '选择一个文档目录，然后直接告诉我要改哪个 Word、Excel 或 PPT。'
}

export function App(): JSX.Element {
  const [settings, setSettings] = useState<AppSettings | null>(null)
  const [workspaceSnapshot, setWorkspaceSnapshot] = useState<WorkspaceSnapshot | null>(null)
  const [messages, setMessages] = useState<ChatMessage[]>([welcomeMessage])
  const [input, setInput] = useState('')
  const [targetFiles, setTargetFiles] = useState<WorkspaceFile[]>([])
  const [recentFiles, setRecentFiles] = useState<GeneratedFile[]>([])
  const [busy, setBusy] = useState(false)
  const [settingsOpen, setSettingsOpen] = useState(false)
  const [error, setError] = useState('')
  const scrollRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    void bootstrap()
  }, [])

  useEffect(() => {
    scrollRef.current?.scrollTo({ top: scrollRef.current.scrollHeight, behavior: 'smooth' })
  }, [messages, busy])

  const statusText = useMemo(() => {
    if (!settings) return '正在加载'
    if (settings.hasApiKey) {
      const source = settings.usesExternalApiConfig ? `自动读取：${settings.apiConfigSource}` : '手动配置'
      return `AI 已连接：${settings.model}（${source}）`
    }
    return '未配置 AI Key，本地规划模式'
  }, [settings])

  async function bootstrap(): Promise<void> {
    if (typeof window.quickDocument === 'undefined') {
      setError('当前页面没有连接到桌面端桥接，请通过 Electron 启动应用。')
      return
    }

    const nextSettings = await window.quickDocument.getSettings()
    setSettings(nextSettings)
    setRecentFiles(await window.quickDocument.getRecentFiles())
    setWorkspaceSnapshot(await window.quickDocument.scanWorkspace(nextSettings.workspacePath))
  }

  async function chooseWorkspaceRoot(): Promise<void> {
    const selected = await window.quickDocument.chooseWorkspace()
    if (!selected) return
    const nextSettings = await window.quickDocument.saveSettings({ workspacePath: selected })
    setSettings(nextSettings)
    setTargetFiles([])
    setWorkspaceSnapshot(await window.quickDocument.scanWorkspace(selected))
  }

  async function refreshWorkspace(): Promise<void> {
    if (!settings) return
    setWorkspaceSnapshot(await window.quickDocument.scanWorkspace(settings.workspacePath))
  }

  function toggleWorkspaceFile(file: WorkspaceFile): void {
    setTargetFiles((current) => {
      if (current.some((item) => item.path === file.path)) {
        return current.filter((item) => item.path !== file.path)
      }
      return [...current, file]
    })
  }

  async function submit(event?: FormEvent): Promise<void> {
    event?.preventDefault()
    const content = input.trim()
    if (!content || busy) return

    const userMessage: ChatMessage = {
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      role: 'user',
      content,
      targetFiles,
      createdAt: new Date().toISOString()
    }
    const nextMessages = [...messages, userMessage]
    setMessages(nextMessages)
    setInput('')
    setBusy(true)
    setError('')

    try {
      const response = await window.quickDocument.sendMessage({
        messages: nextMessages,
        targetFiles,
        workspaceSnapshot: workspaceSnapshot || undefined
      })
      setMessages((current) => [...current, response.message])
      if (response.generatedFiles.length > 0) {
        setRecentFiles(await window.quickDocument.getRecentFiles())
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err))
    } finally {
      setBusy(false)
    }
  }

  function onComposerKeyDown(event: KeyboardEvent<HTMLTextAreaElement>): void {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault()
      void submit()
    }
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand-block">
          <div className="brand-mark">
            <Sparkles size={18} />
          </div>
          <div>
            <strong>Quick Document</strong>
            <span>Document Workflow</span>
          </div>
        </div>

        <button className="primary-nav active" type="button">
          <MessageSquareText size={18} />
          文档工作流
        </button>

        <section className="side-section">
          <div className="section-title">状态</div>
          <div className="status-card">
            <MonitorUp size={17} />
            <span>{statusText}</span>
          </div>
        </section>

        <section className="side-section recent-list">
          <div className="section-title">产物</div>
          {recentFiles.length === 0 ? (
            <p className="muted">完成后的文件会显示在这里。</p>
          ) : (
            recentFiles.slice(0, 6).map((file) => (
              <button
                key={file.id}
                className="recent-item"
                type="button"
                onClick={() => void window.quickDocument.openFile(file.path)}
                title={file.path}
              >
                {kindIcon(file.kind)}
                <span>{file.name}</span>
              </button>
            ))
          )}
        </section>

        <div className="sidebar-spacer" />

        <button className="secondary-nav" type="button" onClick={() => setSettingsOpen(true)}>
          <Settings size={18} />
          设置
        </button>
      </aside>

      <main className="workspace">
        <header className="topbar">
          <div>
            <h1>本地文档处理</h1>
            <p>{settings?.workspacePath || '未选择目录'}</p>
          </div>
          <div className="topbar-actions">
            <button className="icon-button" type="button" onClick={chooseWorkspaceRoot} title="选择目录">
              <FolderOpen size={18} />
            </button>
            <button className="icon-button" type="button" onClick={refreshWorkspace} title="刷新目录">
              <RefreshCw size={18} />
            </button>
            <button className="icon-button" type="button" onClick={() => setSettingsOpen(true)} title="AI 设置">
              <KeyRound size={18} />
            </button>
          </div>
        </header>

        <div className="workflow-layout">
          <section className="file-panel">
            <div className="panel-heading">
              <strong>目录文件</strong>
              <button type="button" onClick={chooseWorkspaceRoot}>
                <FolderOpen size={15} />
                选择文件夹
              </button>
            </div>
            <div className="file-list">
              {(workspaceSnapshot?.files || []).length === 0 ? (
                <p className="muted">当前目录没有可处理的 Office 文件。</p>
              ) : (
                workspaceSnapshot!.files.map((file) => {
                  const selected = targetFiles.some((item) => item.path === file.path)
                  return (
                    <button
                      key={file.path}
                      className={`file-row ${selected ? 'selected' : ''}`}
                      type="button"
                      onClick={() => toggleWorkspaceFile(file)}
                      title={file.path}
                    >
                      {kindIcon(file.kind)}
                      <span>{file.name}</span>
                    </button>
                  )
                })
              )}
            </div>
          </section>

          <section className="chat-panel">
            <div className="messages" ref={scrollRef}>
              {messages.map((message) => (
                <MessageBubble key={message.id} message={message} />
              ))}
              {busy && (
                <div className="message assistant">
                  <div className="avatar">
                    <Bot size={16} />
                  </div>
                  <div className="bubble compact">
                    <Loader2 className="spin" size={16} />
                    正在生成文档处理任务...
                  </div>
                </div>
              )}
            </div>

            {error && <div className="error-bar">{error}</div>}

            {targetFiles.length > 0 && (
              <div className="target-strip">
                {targetFiles.map((file) => (
                  <span className="target-pill" key={file.path}>
                    {kindIcon(file.kind)}
                    {file.name}
                    <button
                      type="button"
                      onClick={() =>
                        setTargetFiles((current) => current.filter((item) => item.path !== file.path))
                      }
                      title="移除"
                    >
                      <X size={14} />
                    </button>
                  </span>
                ))}
              </div>
            )}

            <form className="composer" onSubmit={submit}>
              <button className="tool-button" type="button" onClick={chooseWorkspaceRoot} title="选择文件夹">
                <FolderOpen size={18} />
              </button>
              <textarea
                value={input}
                onChange={(event) => setInput(event.target.value)}
                onKeyDown={onComposerKeyDown}
                rows={1}
                placeholder="例如：把 xxx.docx 中第二段润色一下"
              />
              <button className="send-button" type="submit" disabled={!input.trim() || busy}>
                {busy ? <Loader2 className="spin" size={18} /> : <Send size={18} />}
              </button>
            </form>
          </section>
        </div>
      </main>

      {settingsOpen && settings && (
        <SettingsPanel
          settings={settings}
          onClose={() => setSettingsOpen(false)}
          onSaved={async (next) => {
            setSettings(next)
            setWorkspaceSnapshot(await window.quickDocument.scanWorkspace(next.workspacePath))
            setSettingsOpen(false)
          }}
        />
      )}
    </div>
  )
}

function MessageBubble({ message }: { message: ChatMessage }): JSX.Element {
  const isAssistant = message.role === 'assistant'

  return (
    <div className={`message ${isAssistant ? 'assistant' : 'user'}`}>
      <div className="avatar">{isAssistant ? <Bot size={16} /> : <span>你</span>}</div>
      <div className="message-stack">
        <div className="bubble">{message.content}</div>
        {message.targetFiles && message.targetFiles.length > 0 && (
          <div className="bubble-files">
            {message.targetFiles.map((file) => (
              <span key={file.path}>
                {kindIcon(file.kind)}
                {file.name}
              </span>
            ))}
          </div>
        )}
        {message.actions && message.actions.length > 0 && (
          <div className="action-results">
            {message.actions.map((action, index) => (
              <ActionCard action={action} key={`${action.summary}-${index}`} />
            ))}
          </div>
        )}
      </div>
    </div>
  )
}

function ActionCard({ action }: { action: ActionResult }): JSX.Element {
  return (
    <div className={`action-card ${action.ok ? 'ok' : 'fail'}`}>
      <div className="action-main">
        {action.ok ? <Check size={16} /> : <X size={16} />}
        <div>
          <strong>{action.summary}</strong>
          {action.workflow?.taskFilePath && <span>{action.workflow.taskFilePath}</span>}
          {action.file && <span>{formatBytes(action.file.size)}</span>}
          {action.error && <span>{action.error}</span>}
        </div>
      </div>
      <div className="action-buttons">
        {action.workflow?.taskFilePath && (
          <button type="button" onClick={() => void window.quickDocument.revealFile(action.workflow!.taskFilePath!)}>
            定位
          </button>
        )}
        {action.file && (
          <>
            <button type="button" onClick={() => void window.quickDocument.openFile(action.file!.path)}>
              打开
            </button>
            <button type="button" onClick={() => void window.quickDocument.revealFile(action.file!.path)}>
              定位
            </button>
          </>
        )}
      </div>
    </div>
  )
}

function SettingsPanel({
  settings,
  onClose,
  onSaved
}: {
  settings: AppSettings
  onClose: () => void
  onSaved: (settings: AppSettings) => void | Promise<void>
}): JSX.Element {
  const [form, setForm] = useState<SettingsPatch>({
    provider: settings.provider,
    baseUrl: settings.baseUrl,
    model: settings.model,
    workspacePath: settings.workspacePath,
    residentMode: settings.residentMode,
    apiKey: ''
  })
  const [saving, setSaving] = useState(false)
  const [importMessage, setImportMessage] = useState('')

  async function chooseWorkspace(): Promise<void> {
    const selected = await window.quickDocument.chooseWorkspace()
    if (selected) setForm((current) => ({ ...current, workspacePath: selected }))
  }

  async function save(): Promise<void> {
    setSaving(true)
    try {
      const next = await window.quickDocument.saveSettings(form)
      await onSaved(next)
    } finally {
      setSaving(false)
    }
  }

  async function importExternalSettings(): Promise<void> {
    setSaving(true)
    setImportMessage('')
    try {
      const imported = await window.quickDocument.importExternalSettings()
      if (!imported) {
        setImportMessage('未找到可导入的 cc-switch / Claude 配置。')
        return
      }

      await onSaved(imported.settings)
      setImportMessage(`已导入：${imported.source}`)
    } finally {
      setSaving(false)
    }
  }

  return (
    <div className="modal-layer" role="presentation" onMouseDown={onClose}>
      <section className="settings-panel" role="dialog" aria-modal="true" onMouseDown={(event) => event.stopPropagation()}>
        <header>
          <div>
            <h2>设置</h2>
            <p>AI 接口和本地文档目录。</p>
          </div>
          <button className="icon-button" type="button" onClick={onClose} title="关闭">
            <X size={18} />
          </button>
        </header>

        <div className="import-row">
          <button type="button" onClick={importExternalSettings} disabled={saving}>
            <KeyRound size={16} />
            导入 cc-switch / Claude 配置
          </button>
          <span>
            {importMessage ||
              (settings.usesExternalApiConfig
                ? `已自动读取：${settings.apiConfigSource}`
                : settings.hasApiKey
                  ? '当前使用手动配置'
                  : '未检测到外部配置')}
          </span>
        </div>

        <label>
          <span>Provider</span>
          <select
            value={form.provider || settings.provider}
            onChange={(event) =>
              setForm((current) => ({
                ...current,
                provider: event.target.value as AiProvider
              }))
            }
          >
            <option value="openai">OpenAI-compatible</option>
            <option value="anthropic">Anthropic-compatible</option>
          </select>
        </label>

        <label>
          <span>API Base URL</span>
          <input
            value={form.baseUrl || ''}
            onChange={(event) => setForm((current) => ({ ...current, baseUrl: event.target.value }))}
            placeholder="https://api.openai.com/v1"
          />
        </label>

        <label>
          <span>Model</span>
          <input
            value={form.model || ''}
            onChange={(event) => setForm((current) => ({ ...current, model: event.target.value }))}
            placeholder="gpt-4.1-mini"
          />
        </label>

        <label>
          <span>API Key {settings.hasApiKey ? '（已保存，留空则不变）' : ''}</span>
          <input
            type="password"
            value={form.apiKey || ''}
            onChange={(event) => setForm((current) => ({ ...current, apiKey: event.target.value }))}
            placeholder={settings.hasApiKey ? '输入新 Key 可替换' : 'sk-...'}
          />
        </label>

        <label>
          <span>文档目录</span>
          <div className="path-row">
            <input
              value={form.workspacePath || ''}
              onChange={(event) =>
                setForm((current) => ({ ...current, workspacePath: event.target.value }))
              }
            />
            <button type="button" onClick={chooseWorkspace}>
              <FolderOpen size={16} />
            </button>
          </div>
        </label>

        <label className="toggle-row">
          <input
            type="checkbox"
            checked={Boolean(form.residentMode)}
            onChange={(event) =>
              setForm((current) => ({ ...current, residentMode: event.target.checked }))
            }
          />
          <span>关闭窗口后驻留在托盘</span>
        </label>

        <footer>
          <button className="plain-button" type="button" onClick={onClose}>
            取消
          </button>
          <button className="save-button" type="button" onClick={save} disabled={saving}>
            {saving ? <Loader2 className="spin" size={16} /> : <Check size={16} />}
            保存
          </button>
        </footer>
      </section>
    </div>
  )
}

function kindIcon(kind: OfficeKind): JSX.Element {
  if (kind === 'excel') return <FileSpreadsheet size={16} />
  if (kind === 'powerpoint') return <Presentation size={16} />
  return <FileText size={16} />
}

function formatBytes(value: number): string {
  if (value < 1024) return `${value} B`
  if (value < 1024 * 1024) return `${(value / 1024).toFixed(1)} KB`
  return `${(value / 1024 / 1024).toFixed(1)} MB`
}
