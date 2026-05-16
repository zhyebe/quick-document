import { app, net, shell } from 'electron'
import { spawn } from 'node:child_process'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { UpdateAsset, UpdateDownloadResult, UpdateStatus } from '@shared/types'

const RELEASE_API_URL = 'https://api.github.com/repos/zhyebe/quick-document/releases/latest'
const RELEASE_PAGE_URL = 'https://github.com/zhyebe/quick-document/releases/latest'
const UPDATE_REQUEST_TIMEOUT_MS = 15_000
const INSTALLER_DOWNLOAD_TIMEOUT_MS = 10 * 60_000

interface GitHubRelease {
  tag_name?: string
  html_url?: string
  assets?: Array<{
    name?: string
    size?: number
    browser_download_url?: string
  }>
}

export async function checkForUpdates(): Promise<UpdateStatus> {
  const currentVersion = app.getVersion()

  try {
    const release = await fetchLatestReleaseInfo(currentVersion)
    const available = compareVersions(release.latestVersion, currentVersion) > 0
    return {
      currentVersion,
      latestVersion: release.latestVersion,
      available,
      releaseUrl: release.releaseUrl,
      asset: release.asset,
      message: available
        ? release.asset
          ? `发现新版本 v${release.latestVersion}`
          : `发现新版本 v${release.latestVersion}，但没有匹配当前系统的安装包。`
        : release.asset
          ? `当前已是最新版本 v${currentVersion}`
          : '当前已是最新版本，且发布页上暂时没有匹配当前系统的安装包。'
    }
  } catch (error) {
    return {
      currentVersion,
      available: false,
      message: `检查更新失败：${error instanceof Error ? error.message : String(error)}`
    }
  }
}

export async function downloadAndOpenUpdate(cachedStatus?: UpdateStatus): Promise<UpdateDownloadResult> {
  const status = cachedStatus?.asset ? cachedStatus : await checkForUpdates()
  if (!status.asset) {
    return {
      ok: false,
      message: status.available
        ? '发现新版本，但没有匹配当前系统的安装包，无法自动更新。'
        : status.message,
      releaseUrl: status.releaseUrl
    }
  }

  try {
    const filePath = await downloadInstaller(status.asset)
    await openInstaller(filePath)
    return {
      ok: true,
      filePath,
      releaseUrl: status.releaseUrl,
      message: buildInstallerOpenedMessage(status.asset, filePath)
    }
  } catch (error) {
    return {
      ok: false,
      releaseUrl: status.releaseUrl,
      message: `下载安装包失败：${formatUpdateError(error)}`
    }
  }
}

async function fetchForUpdate(url: string, init: RequestInit): Promise<Response> {
  const requestInit: RequestInit = {
    ...init,
    redirect: 'follow'
  }
  try {
    return await fetchWithTimeout((signal) => net.fetch(url, { ...requestInit, signal }))
  } catch (error) {
    if (isAbortError(error)) throw error
    return fetchWithTimeout((signal) => fetch(url, { ...requestInit, signal }))
  }
}

async function fetchLatestReleaseInfo(currentVersion: string): Promise<{
  latestVersion: string
  releaseUrl: string
  asset?: UpdateAsset
}> {
  try {
    const response = await fetchForUpdate(RELEASE_API_URL, {
      headers: {
        Accept: 'application/vnd.github+json',
        'User-Agent': `QuickDocument/${currentVersion}`
      }
    })

    if (!response.ok) {
      throw new Error(`GitHub 返回 ${response.status}`)
    }

    const release = (await response.json()) as GitHubRelease
    const latestVersion = normalizeVersion(release.tag_name || '')
    if (!latestVersion) {
      throw new Error('没有找到可用的 release 版本。')
    }

    return {
      latestVersion,
      releaseUrl: release.html_url || RELEASE_PAGE_URL,
      asset: selectInstallerAsset(release.assets || [])
    }
  } catch {
    const fallback = await fetchReleasePageInfo(currentVersion)
    if (fallback) return fallback
    throw new Error('无法连接到 GitHub release 页面。')
  }
}

async function fetchReleasePageInfo(currentVersion: string): Promise<{
  latestVersion: string
  releaseUrl: string
  asset?: UpdateAsset
} | null> {
  const response = await fetchForUpdate(RELEASE_PAGE_URL, {
    headers: {
      Accept: 'text/html,application/xhtml+xml',
      'User-Agent': `QuickDocument/${currentVersion}`
    }
  })

  if (!response.ok) return null
  const releaseUrl = response.url || RELEASE_PAGE_URL
  const html = await response.text()
  const latestVersion = extractVersionFromReleaseUrl(releaseUrl) || extractVersionFromReleaseHtml(html)
  if (!latestVersion) return null
  return {
    latestVersion,
    releaseUrl,
    asset: selectInstallerAsset(extractAssetsFromReleaseHtml(html))
  }
}

async function downloadInstaller(asset: UpdateAsset): Promise<string> {
  const targetDir = join(app.getPath('downloads'), 'Quick Document Updates')
  mkdirSync(targetDir, { recursive: true })
  const filePath = join(targetDir, sanitizeFileName(asset.name))
  const response = await fetchInstaller(asset.url)
  if (!response.ok) {
    const detail = await response.text().catch(() => '')
    throw new Error(`GitHub 返回 ${response.status}${detail ? `：${detail.slice(0, 180)}` : ''}`)
  }

  const bytes = Buffer.from(await response.arrayBuffer())
  if (asset.size && bytes.length < asset.size) {
    throw new Error(`安装包下载不完整：${bytes.length}/${asset.size} bytes`)
  }
  writeFileSync(filePath, bytes)
  return filePath
}

async function fetchInstaller(url: string): Promise<Response> {
  const requestInit: RequestInit = {
    redirect: 'follow',
    headers: {
      Accept: 'application/octet-stream',
      'User-Agent': `QuickDocument/${app.getVersion()}`
    }
  }
  try {
    return await fetchWithTimeout(
      (signal) => net.fetch(url, { ...requestInit, signal }),
      INSTALLER_DOWNLOAD_TIMEOUT_MS
    )
  } catch (error) {
    if (isAbortError(error)) throw error
    return fetchWithTimeout(
      (signal) => fetch(url, { ...requestInit, signal }),
      INSTALLER_DOWNLOAD_TIMEOUT_MS
    )
  }
}

async function openInstaller(filePath: string): Promise<void> {
  if (process.platform === 'win32') {
    await new Promise<void>((resolve, reject) => {
      const child = spawn(filePath, [], {
        detached: true,
        stdio: 'ignore'
      })
      child.once('error', reject)
      child.unref()
      resolve()
    })
    return
  }

  const error = await shell.openPath(filePath)
  if (error) throw new Error(error)
}

function buildInstallerOpenedMessage(asset: UpdateAsset, filePath: string): string {
  if (process.platform === 'win32') {
    return `已下载并启动安装器：${asset.name}。如果安装器提示文件占用，请退出 Quick Document 后继续安装。`
  }
  if (process.platform === 'darwin') {
    return `已下载并打开安装包：${asset.name}。下载位置：${filePath}`
  }
  return `已下载安装包：${filePath}`
}

function formatUpdateError(error: unknown): string {
  if (error instanceof Error) return error.message
  return String(error)
}

async function fetchWithTimeout(
  fn: (signal: AbortSignal) => Promise<Response>,
  timeoutMs = UPDATE_REQUEST_TIMEOUT_MS
): Promise<Response> {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), timeoutMs)
  try {
    return await fn(controller.signal)
  } catch (error) {
    if (isAbortError(error)) {
      throw new Error(`请求超时（${Math.round(timeoutMs / 1000)}s）`)
    }
    throw error
  } finally {
    clearTimeout(timer)
  }
}

function isAbortError(error: unknown): boolean {
  return error instanceof Error && /abort|aborted|timeout/i.test(error.message)
}

function selectInstallerAsset(assets: GitHubRelease['assets']): UpdateAsset | undefined {
  const candidates = (assets || [])
    .filter((asset) => asset.name && asset.browser_download_url)
    .map((asset) => ({
      name: asset.name!,
      url: asset.browser_download_url!,
      size: asset.size
    }))

  if (process.platform === 'win32') {
    return candidates.find((asset) => /\.exe$/i.test(asset.name))
  }

  if (process.platform === 'darwin') {
    const dmgAssets = candidates.filter((asset) => /\.dmg$/i.test(asset.name))
    if (process.arch === 'arm64') {
      return dmgAssets.find((asset) => /arm64/i.test(asset.name)) || dmgAssets[0]
    }
    return dmgAssets.find((asset) => /x64|universal/i.test(asset.name)) || dmgAssets[0]
  }

  return undefined
}

function normalizeVersion(value: string): string {
  return value.trim().replace(/^v/i, '')
}

function extractVersionFromReleaseUrl(url: string): string {
  const match = url.match(/\/releases\/tag\/v?([^/?#]+)/i)
  return match ? normalizeVersion(match[1] || '') : ''
}

function extractVersionFromReleaseHtml(html: string): string {
  const match = html.match(/\/zhyebe\/quick-document\/releases\/tag\/v?([0-9][^"'<>/]*)/i)
  return match ? normalizeVersion(match[1] || '') : ''
}

function extractAssetsFromReleaseHtml(html: string): GitHubRelease['assets'] {
  const assets = new Map<string, { name: string; browser_download_url: string }>()
  const pattern = /href=["']([^"']*\/zhyebe\/quick-document\/releases\/download\/v[^"']+\.(?:dmg|exe|zip))["']/gi
  let match: RegExpExecArray | null
  while ((match = pattern.exec(html))) {
    const url = absoluteGithubUrl(decodeHtml(match[1] || ''))
    const name = decodeURIComponent(url.split('/').pop() || '')
    if (name) assets.set(url, { name, browser_download_url: url })
  }
  return Array.from(assets.values())
}

function absoluteGithubUrl(value: string): string {
  if (/^https?:\/\//i.test(value)) return value
  return `https://github.com${value.startsWith('/') ? '' : '/'}${value}`
}

function decodeHtml(value: string): string {
  return value.replace(/&amp;/g, '&')
}

function sanitizeFileName(value: string): string {
  return value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '-').trim() || 'Quick.Document.Update'
}

function compareVersions(first: string, second: string): number {
  const firstParts = normalizeVersion(first).split('.').map(numberPart)
  const secondParts = normalizeVersion(second).split('.').map(numberPart)
  const length = Math.max(firstParts.length, secondParts.length)
  for (let index = 0; index < length; index += 1) {
    const diff = (firstParts[index] || 0) - (secondParts[index] || 0)
    if (diff !== 0) return diff
  }
  return 0
}

function numberPart(value: string): number {
  const parsed = Number(value.replace(/[^0-9].*$/, ''))
  return Number.isFinite(parsed) ? parsed : 0
}
