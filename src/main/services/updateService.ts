import { app, net, shell } from 'electron'
import type { UpdateAsset, UpdateDownloadResult, UpdateStatus } from '@shared/types'

const RELEASE_API_URL = 'https://api.github.com/repos/zhyebe/quick-document/releases/latest'
const RELEASE_PAGE_URL = 'https://github.com/zhyebe/quick-document/releases/latest'
const UPDATE_REQUEST_TIMEOUT_MS = 15_000

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
          : `发现新版本 v${release.latestVersion}，请打开发布页下载对应安装包。`
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
  return openUpdateInBrowser(status, buildOpenUpdateMessage(status))
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

async function openUpdateInBrowser(status: UpdateStatus, message: string): Promise<UpdateDownloadResult> {
  const target = status.asset?.url || status.releaseUrl || RELEASE_PAGE_URL
  if (!target) {
    return {
      ok: false,
      message,
      releaseUrl: status.releaseUrl
    }
  }

  try {
    await shell.openExternal(target)
    return {
      ok: true,
      message,
      releaseUrl: status.releaseUrl
    }
  } catch (error) {
    return {
      ok: false,
      message: `${message} 但浏览器打开失败：${formatUpdateError(error)}`,
      releaseUrl: status.releaseUrl
    }
  }
}

function buildOpenUpdateMessage(status: UpdateStatus): string {
  if (status.asset) {
    return `已在浏览器打开安装包下载页面：${status.asset.name}。下载完成后运行安装包完成更新。`
  }
  return '已在浏览器打开发布页，请下载对应平台的安装包并手动安装。'
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
  const latestVersion = extractVersionFromReleaseUrl(releaseUrl)
  if (!latestVersion) return null
  return {
    latestVersion,
    releaseUrl,
    asset: undefined
  }
}

function formatUpdateError(error: unknown): string {
  if (error instanceof Error) return error.message
  return String(error)
}

async function fetchWithTimeout(fn: (signal: AbortSignal) => Promise<Response>): Promise<Response> {
  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), UPDATE_REQUEST_TIMEOUT_MS)
  try {
    return await fn(controller.signal)
  } catch (error) {
    if (isAbortError(error)) {
      throw new Error(`请求超时（${Math.round(UPDATE_REQUEST_TIMEOUT_MS / 1000)}s）`)
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
