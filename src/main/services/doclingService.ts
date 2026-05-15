import { app } from 'electron'
import { execFile, execFileSync } from 'node:child_process'
import { existsSync, mkdtempSync, mkdirSync, readdirSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, extname, join } from 'node:path'
import { promisify } from 'node:util'
import type { DoclingInstallResult, DoclingStatus, WorkspaceFile } from '@shared/types'

const execFileAsync = promisify(execFile)

const MAX_DOCLING_FILES = 4
const MAX_DOCLING_CHARS_PER_FILE = 5000
const DOCLING_TIMEOUT_MS = 45_000
const INSTALL_TIMEOUT_MS = 15 * 60_000
const MIN_DOCLING_PYTHON = { major: 3, minor: 10 }
const PYTHON_PATHS_DARWIN = [
  '/Library/Frameworks/Python.framework/Versions/3.13/bin/python3',
  '/Library/Frameworks/Python.framework/Versions/3.12/bin/python3',
  '/Library/Frameworks/Python.framework/Versions/3.11/bin/python3',
  '/opt/homebrew/bin/python3',
  '/usr/local/bin/python3',
  '/opt/anaconda3/bin/python3',
  '/usr/bin/python3',
  'python3',
  'python'
]
const PIP_INDEX_URLS = [
  '',
  'https://pypi.org/simple',
  'https://pypi.tuna.tsinghua.edu.cn/simple',
  'https://pypi.mirrors.ustc.edu.cn/simple'
]
const DOCLING_EXTENSIONS = new Set([
  '.pdf',
  '.docx',
  '.xlsx',
  '.pptx',
  '.html',
  '.htm',
  '.md',
  '.csv',
  '.txt',
  '.png',
  '.jpg',
  '.jpeg',
  '.webp',
  '.tif',
  '.tiff'
])

interface DoclingParseResult {
  ok: boolean
  engine?: string
  markdown?: string
  error?: string
}

export async function buildDoclingPreviewContext(files: WorkspaceFile[]): Promise<string> {
  const candidates = files.filter((file) => DOCLING_EXTENSIONS.has(extname(file.path).toLowerCase()))
  if (candidates.length === 0) return ''
  const status = await getDoclingStatus()
  if (!status.installed) return ''

  const previews: string[] = []
  for (const file of candidates.slice(0, MAX_DOCLING_FILES)) {
    const result = await parseWithDocling(file.path)
    if (!result.ok || !result.markdown?.trim()) continue
    previews.push(
      `文件：${file.name}\n路径：${file.path}\n解析器：${result.engine || 'docling'}\nDocling Markdown：\n${clip(
        result.markdown,
        MAX_DOCLING_CHARS_PER_FILE
      )}`
    )
  }

  return previews.length > 0 ? previews.join('\n\n---\n\n') : ''
}

export async function getDoclingStatus(): Promise<DoclingStatus> {
  const managedProbe = await probeManagedVenvDocling()
  if (managedProbe.installed) return managedProbe

  const pythonProbe = await probePythonDocling()
  if (pythonProbe.installed) return pythonProbe

  const cliProbe = await probeCommand('docling', ['--version'])
  if (cliProbe.installed) return cliProbe

  return {
    installed: false,
    installCommand: recommendedInstallCommand().join(' '),
    message: '未检测到 Docling。点击安装后会创建 Quick Document 专用 Python 环境安装 docling。'
  }
}

export async function installDocling(): Promise<DoclingInstallResult> {
  const before = await getDoclingStatus()
  if (before.installed) return { ...before, ok: true, log: 'Docling already installed.' }

  const logs: string[] = []
  const managedInstall = await installManagedVenvDocling(logs)
  if (managedInstall.ok) return managedInstall

  const attempts = installCommandAttempts()

  if (attempts.length === 0 && !managedInstall.ok) {
    return {
      ...before,
      ok: false,
      message: `Docling 安装失败。需要 Python ${MIN_DOCLING_PYTHON.major}.${MIN_DOCLING_PYTHON.minor}+；未检测到可用的 Python。`,
      log: logs.join('\n\n')
    }
  }

  for (const attempt of attempts) {
    try {
      const { stdout, stderr } = await execFileAsync(attempt.command, attempt.args, {
        timeout: 10 * 60_000,
        maxBuffer: 16 * 1024 * 1024
      })
      logs.push(`$ ${[attempt.command, ...attempt.args].join(' ')}\n${stdout}${stderr}`)
      const after = await getDoclingStatus()
      if (after.installed) return { ...after, ok: true, log: logs.join('\n\n') }
    } catch (error) {
      logs.push(`$ ${[attempt.command, ...attempt.args].join(' ')}\n${formatExecError(error)}`)
    }
  }

  const after = await getDoclingStatus()
  return {
    ...after,
    ok: false,
    message: `Docling 安装失败。${lastLogSummary(logs)}`,
    log: logs.join('\n\n')
  }
}

export async function parseWithDocling(filePath: string): Promise<DoclingParseResult> {
  if (!DOCLING_EXTENSIONS.has(extname(filePath).toLowerCase())) {
    return { ok: false, error: `Docling 不支持该文件类型：${basename(filePath)}` }
  }

  const pythonResult = await parseWithPythonModule(filePath)
  if (pythonResult.ok) return pythonResult

  const cliResult = await parseWithDoclingCli(filePath, 'docling', ['--to', 'md'])
  if (cliResult.ok) return cliResult

  if (process.env.QUICK_DOCUMENT_DOCLING_AUTO_UVX === '1') {
    const uvxResult = await parseWithDoclingCli(filePath, 'uvx', ['--from', 'docling', 'docling', '--to', 'md'])
    if (uvxResult.ok) return uvxResult
  }

  return {
    ok: false,
    error: pythonResult.error || cliResult.error || 'Docling 未安装或解析失败。'
  }
}

async function parseWithPythonModule(filePath: string): Promise<DoclingParseResult> {
  const script = [
    'import json, sys',
    'from docling.document_converter import DocumentConverter',
    'doc = DocumentConverter().convert(sys.argv[1]).document',
    'print(json.dumps({"markdown": doc.export_to_markdown()}, ensure_ascii=False))'
  ].join('\n')

  for (const command of [managedVenvPython(), ...pythonCommands()].filter(Boolean) as string[]) {
    try {
      const { stdout } = await execFileAsync(command, [...pythonVersionArgs(command), '-c', script, filePath], {
        timeout: DOCLING_TIMEOUT_MS,
        maxBuffer: 8 * 1024 * 1024
      })
      const data = JSON.parse(stdout) as { markdown?: string }
      if (data.markdown?.trim()) {
        return { ok: true, engine: `${command} docling module`, markdown: data.markdown }
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      if (!/ENOENT|No module named|ModuleNotFoundError/.test(message)) {
        return { ok: false, error: message }
      }
    }
  }

  return { ok: false, error: 'Python docling module not found.' }
}

async function probePythonDocling(): Promise<DoclingStatus> {
  for (const command of pythonCommands()) {
    try {
      const { stdout } = await execFileAsync(command, [
        ...pythonVersionArgs(command),
        '-c',
        'import docling,sys; print(getattr(docling, "__version__", "installed"))'
      ])
      return {
        installed: true,
        engine: `${command} docling module`,
        message: `已检测到 Docling：${stdout.trim() || command}`
      }
    } catch {
      // Try the next Python launcher.
    }
  }
  return { installed: false, message: 'Python docling module not found.' }
}

async function probeManagedVenvDocling(): Promise<DoclingStatus> {
  const python = managedVenvPython()
  if (!python || !existsSync(python)) return { installed: false, message: 'Managed Docling venv not found.' }
  try {
    const { stdout } = await execFileAsync(python, [
      '-c',
      'import docling,sys; print(getattr(docling, "__version__", "installed"))'
    ])
    return {
      installed: true,
      engine: `${python} docling module`,
      message: `已检测到 Quick Document 专用 Docling：${stdout.trim() || 'installed'}`
    }
  } catch {
    return { installed: false, message: 'Managed Docling venv exists but docling is not installed.' }
  }
}

async function probeCommand(command: string, args: string[]): Promise<DoclingStatus> {
  try {
    const { stdout, stderr } = await execFileAsync(command, args, {
      timeout: 10_000,
      maxBuffer: 1024 * 1024
    })
    return {
      installed: true,
      engine: `${command} CLI`,
      message: `已检测到 Docling CLI：${(stdout || stderr).trim() || command}`
    }
  } catch {
    return { installed: false, message: `${command} not found.` }
  }
}

function pythonCommands(): string[] {
  const commands =
    process.platform === 'win32'
      ? ['py', 'python']
      : process.platform === 'darwin'
        ? PYTHON_PATHS_DARWIN
        : ['/usr/bin/python3', '/usr/local/bin/python3', 'python3', 'python']
  return uniqueExistingCommands(commands)
}

function installCommandAttempts(): Array<{ command: string; args: string[] }> {
  const commands: Array<{ command: string; args: string[] }> = [
    ...pipxCommands().map((command) => ({ command, args: ['install', 'docling'] }))
  ]

  for (const python of installPythonCommands()) {
    commands.push({
      command: python,
      args: [
        ...pythonVersionArgs(python),
        '-m',
        'pip',
        'install',
        '--user',
        '--prefer-binary',
        'docling'
      ]
    })
  }

  return commands
}

function recommendedInstallCommand(): string[] {
  const [python] = installPythonCommands()
  return python
    ? [python, ...pythonVersionArgs(python), '-m', 'venv', managedVenvDir()]
    : [`Python ${MIN_DOCLING_PYTHON.major}.${MIN_DOCLING_PYTHON.minor}+`]
}

function pythonVersionArgs(command: string): string[] {
  return process.platform === 'win32' && command === 'py' ? ['-3'] : []
}

function formatExecError(error: unknown): string {
  if (!error || typeof error !== 'object') return String(error)
  const record = error as { message?: string; stdout?: string; stderr?: string }
  return [record.message, record.stdout, record.stderr].filter(Boolean).join('\n')
}

async function installManagedVenvDocling(logs: string[]): Promise<DoclingInstallResult> {
  const venvDir = managedVenvDir()
  mkdirSync(app.getPath('userData'), { recursive: true })

  const existingVenvPython = managedVenvPython()
  if (existsSync(existingVenvPython) && !isSupportedPythonCommand(existingVenvPython)) {
    logs.push(
      `Managed venv uses unsupported Python (${pythonVersionLabel(existingVenvPython)}). Recreating ${venvDir}.`
    )
    rmSync(venvDir, { recursive: true, force: true })
  }

  const candidates = installPythonCommands()
  if (candidates.length === 0) {
    logs.push(`No Python ${MIN_DOCLING_PYTHON.major}.${MIN_DOCLING_PYTHON.minor}+ command was found.`)
    return {
      installed: false,
      ok: false,
      installCommand: recommendedInstallCommand().join(' '),
      message: `需要 Python ${MIN_DOCLING_PYTHON.major}.${MIN_DOCLING_PYTHON.minor}+ 才能安装 Docling。`,
      log: logs.join('\n\n')
    }
  }

  for (const python of candidates) {
    try {
      if (!existsSync(managedVenvPython())) {
        const args = [...pythonVersionArgs(python), '-m', 'venv', venvDir]
        const { stdout, stderr } = await execFileAsync(python, args, {
          timeout: INSTALL_TIMEOUT_MS,
          maxBuffer: 16 * 1024 * 1024
        })
        logs.push(`$ ${[python, ...args].join(' ')}\n${stdout}${stderr}`)
      }

      const venvPython = managedVenvPython()
      if (!venvPython || !existsSync(venvPython)) {
        logs.push(`Managed venv Python not found at ${venvPython || '(unknown path)'}`)
        continue
      }

      await ensureVenvPipReady(venvPython, logs)

      for (const indexUrl of PIP_INDEX_URLS) {
        const installArgs = [
          '-m',
          'pip',
          'install',
          '--timeout',
          '60',
          '--retries',
          '2',
          '--prefer-binary',
          ...(indexUrl ? ['--index-url', indexUrl] : []),
          'docling'
        ]
        try {
          const install = await execFileAsync(venvPython, installArgs, {
            timeout: INSTALL_TIMEOUT_MS,
            maxBuffer: 32 * 1024 * 1024,
            env: installEnvironment()
          })
          logs.push(`$ ${[venvPython, ...installArgs].join(' ')}\n${install.stdout}${install.stderr}`)

          const after = await probeManagedVenvDocling()
          if (after.installed) return { ...after, ok: true, log: logs.join('\n\n') }
        } catch (error) {
          logs.push(`$ ${[venvPython, ...installArgs].join(' ')}\n${formatExecError(error)}`)
        }
      }

    } catch (error) {
      logs.push(`$ ${python} -m venv/pip docling\n${formatExecError(error)}`)
    }
  }

  const after = await probeManagedVenvDocling()
  return { ...after, ok: false, log: logs.join('\n\n') }
}

async function ensureVenvPipReady(venvPython: string, logs: string[]): Promise<void> {
  try {
    const ensure = await execFileAsync(venvPython, ['-m', 'ensurepip', '--upgrade'], {
      timeout: 120_000,
      maxBuffer: 4 * 1024 * 1024
    })
    logs.push(`$ ${[venvPython, '-m', 'ensurepip', '--upgrade'].join(' ')}\n${ensure.stdout}${ensure.stderr}`)
  } catch (error) {
    logs.push(`$ ${[venvPython, '-m', 'ensurepip', '--upgrade'].join(' ')}\n${formatExecError(error)}`)
  }

  try {
    const upgrade = await execFileAsync(
      venvPython,
      ['-m', 'pip', 'install', '--upgrade', '--prefer-binary', 'pip', 'setuptools', 'wheel'],
      {
        timeout: INSTALL_TIMEOUT_MS,
        maxBuffer: 16 * 1024 * 1024,
        env: installEnvironment()
      }
    )
    logs.push(
      `$ ${[venvPython, '-m', 'pip', 'install', '--upgrade', '--prefer-binary', 'pip', 'setuptools', 'wheel'].join(
        ' '
      )}\n${upgrade.stdout}${upgrade.stderr}`
    )
  } catch (error) {
    logs.push(
      `$ ${[
        venvPython,
        '-m',
        'pip',
        'install',
        '--upgrade',
        '--prefer-binary',
        'pip',
        'setuptools',
        'wheel'
      ].join(' ')}\n${formatExecError(error)}`
    )
  }
}

function pipxCommands(): string[] {
  return uniqueExistingCommands(
    process.platform === 'win32'
      ? ['pipx']
      : ['/opt/homebrew/bin/pipx', '/usr/local/bin/pipx', 'pipx']
  )
}

function uniqueExistingCommands(commands: string[]): string[] {
  const seen = new Set<string>()
  return commands.filter((command) => {
    if (seen.has(command)) return false
    seen.add(command)
    return !command.includes('/') || existsSync(command)
  })
}

function installPythonCommands(): string[] {
  return pythonCommands().filter((command) => isSupportedPythonCommand(command))
}

function isSupportedPythonCommand(command: string): boolean {
  const version = pythonVersion(command)
  if (!version) return false
  if (version.major > MIN_DOCLING_PYTHON.major) return true
  return version.major === MIN_DOCLING_PYTHON.major && version.minor >= MIN_DOCLING_PYTHON.minor
}

function pythonVersion(command: string): { major: number; minor: number; patch: number } | null {
  try {
    const stdout = execFileSync(command, [
      ...pythonVersionArgs(command),
      '-c',
      'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}")'
    ], {
      timeout: 10_000,
      encoding: 'utf8'
    }).trim()
    const match = stdout.match(/^(\d+)\.(\d+)\.(\d+)/)
    if (!match) return null
    return {
      major: Number(match[1]),
      minor: Number(match[2]),
      patch: Number(match[3])
    }
  } catch {
    return null
  }
}

function pythonVersionLabel(command: string): string {
  const version = pythonVersion(command)
  return version ? `${version.major}.${version.minor}.${version.patch}` : 'unknown'
}

function installEnvironment(): NodeJS.ProcessEnv {
  return {
    ...process.env,
    PATH: augmentedPath(),
    PIP_DISABLE_PIP_VERSION_CHECK: '1'
  }
}

function augmentedPath(): string {
  const extras =
    process.platform === 'win32'
      ? []
      : ['/Library/Frameworks/Python.framework/Versions/3.13/bin', '/Library/Frameworks/Python.framework/Versions/3.12/bin', '/Library/Frameworks/Python.framework/Versions/3.11/bin', '/opt/homebrew/bin', '/usr/local/bin', '/usr/bin', '/bin', '/usr/sbin', '/sbin', '/opt/anaconda3/bin']
  return [...extras, process.env.PATH || ''].filter(Boolean).join(process.platform === 'win32' ? ';' : ':')
}

function managedVenvDir(): string {
  return join(app.getPath('userData'), 'docling-venv')
}

function managedVenvPython(): string {
  const venvDir = managedVenvDir()
  return process.platform === 'win32'
    ? join(venvDir, 'Scripts', 'python.exe')
    : join(venvDir, 'bin', 'python')
}

function lastLogSummary(logs: string[]): string {
  const last = logs[logs.length - 1] || ''
  const lines = last
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(-4)
  return lines.length > 0 ? lines.join(' ') : '请确认本机已安装 Python，并且网络可访问 PyPI。'
}

async function parseWithDoclingCli(
  filePath: string,
  command: string,
  baseArgs: string[]
): Promise<DoclingParseResult> {
  const outputDir = mkdtempSync(join(tmpdir(), 'quick-document-docling-'))
  try {
    await execFileAsync(command, [...baseArgs, '--output', outputDir, filePath], {
      timeout: DOCLING_TIMEOUT_MS,
      maxBuffer: 8 * 1024 * 1024
    })
    const markdown = readFirstMarkdown(outputDir)
    if (!markdown.trim()) return { ok: false, error: `${command} did not produce Markdown.` }
    return { ok: true, engine: command === 'uvx' ? 'uvx docling' : 'docling CLI', markdown }
  } catch (error) {
    return { ok: false, error: error instanceof Error ? error.message : String(error) }
  } finally {
    rmSync(outputDir, { recursive: true, force: true })
  }
}

function readFirstMarkdown(outputDir: string): string {
  const markdownPath = findFirstMarkdownPath(outputDir)
  return markdownPath ? readFileSync(markdownPath, 'utf8') : ''
}

function findFirstMarkdownPath(directory: string): string | null {
  for (const entry of readdirSync(directory, { withFileTypes: true })) {
    const fullPath = join(directory, entry.name)
    if (entry.isFile() && entry.name.toLowerCase().endsWith('.md')) return fullPath
    if (entry.isDirectory()) {
      const nested = findFirstMarkdownPath(fullPath)
      if (nested) return nested
    }
  }
  return null
}

function clip(value: string, maxLength: number): string {
  return value.length > maxLength ? `${value.slice(0, maxLength)}\n...[已截断]` : value
}
