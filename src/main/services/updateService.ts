import { app, shell } from 'electron'
import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import type { UpdateAsset, UpdateDownloadResult, UpdateStatus } from '@shared/types'

const RELEASE_API_URL = 'https://api.github.com/repos/zhyebe/quick-document/releases/latest'

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
    const response = await fetch(RELEASE_API_URL, {
      headers: {
        Accept: 'application/vnd.github+json',
        'User-Agent': `QuickDocument/${currentVersion}`
      }
    })

    if (!response.ok) {
      return {
        currentVersion,
        available: false,
        message: `检查更新失败：GitHub 返回 ${response.status}`
      }
    }

    const release = (await response.json()) as GitHubRelease
    const latestVersion = normalizeVersion(release.tag_name || '')
    if (!latestVersion) {
      return {
        currentVersion,
        available: false,
        releaseUrl: release.html_url,
        message: '没有找到可用的 release 版本。'
      }
    }

    const asset = selectInstallerAsset(release.assets || [])
    const available = compareVersions(latestVersion, currentVersion) > 0 && Boolean(asset)
    return {
      currentVersion,
      latestVersion,
      available,
      releaseUrl: release.html_url,
      asset,
      message: available
        ? `发现新版本 v${latestVersion}`
        : asset
          ? `当前已是最新版本 v${currentVersion}`
          : '发现新版，但没有匹配当前系统的安装包。'
    }
  } catch (error) {
    return {
      currentVersion,
      available: false,
      message: `检查更新失败：${error instanceof Error ? error.message : String(error)}`
    }
  }
}

export async function downloadAndOpenUpdate(): Promise<UpdateDownloadResult> {
  const status = await checkForUpdates()
  if (!status.available || !status.asset) {
    return {
      ok: false,
      message: status.message,
      releaseUrl: status.releaseUrl
    }
  }

  try {
    const downloadDir = join(app.getPath('userData'), 'updates')
    mkdirSync(downloadDir, { recursive: true })
    const filePath = join(downloadDir, sanitizeFilename(status.asset.name))
    const response = await fetch(status.asset.url, {
      headers: {
        'User-Agent': `QuickDocument/${status.currentVersion}`
      }
    })

    if (!response.ok) {
      return {
        ok: false,
        message: `下载安装包失败：GitHub 返回 ${response.status}`,
        releaseUrl: status.releaseUrl
      }
    }

    writeFileSync(filePath, Buffer.from(await response.arrayBuffer()))
    const openError = await shell.openPath(filePath)
    return {
      ok: !openError,
      message: openError ? `安装包已下载，但无法打开：${openError}` : `已下载并打开安装包：${status.asset.name}`,
      filePath,
      releaseUrl: status.releaseUrl
    }
  } catch (error) {
    return {
      ok: false,
      message: `下载安装包失败：${error instanceof Error ? error.message : String(error)}`,
      releaseUrl: status.releaseUrl
    }
  }
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

function sanitizeFilename(value: string): string {
  return value.replace(/[<>:"/\\|?*\u0000-\u001f]/g, '-').slice(0, 160)
}
