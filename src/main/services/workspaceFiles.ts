import { readdirSync, statSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import type { OfficeKind, WorkspaceFile, WorkspaceSnapshot } from '@shared/types'

const INDEXED_DOCUMENT_EXTENSIONS = new Set([
  '.docx',
  '.xlsx',
  '.ppt',
  '.pptx',
  '.pdf',
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
const SKIPPED_DIRECTORIES = new Set([
  'node_modules',
  '.git',
  '.idea',
  '.vscode',
  'out',
  'dist',
  'release',
  '.quick-document'
])

export function inferOfficeKind(filePath: string): OfficeKind {
  const ext = extname(filePath).toLowerCase()
  if (ext === '.docx') return 'word'
  if (ext === '.xlsx') return 'excel'
  if (ext === '.ppt' || ext === '.pptx') return 'powerpoint'
  return 'unknown'
}

export function toWorkspaceFile(filePath: string): WorkspaceFile {
  const stats = statSync(filePath)
  return {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    name: basename(filePath),
    path: filePath,
    kind: stats.isDirectory() ? 'unknown' : inferOfficeKind(filePath),
    size: stats.isFile() ? stats.size : undefined,
    modifiedAt: stats.mtime.toISOString(),
    isDirectory: stats.isDirectory()
  }
}

export function scanWorkspace(rootPath: string, limit = 2000): WorkspaceSnapshot {
  const files: WorkspaceFile[] = []
  walk(rootPath, files, limit)
  files.sort((first, second) => {
    const firstTime = first.modifiedAt ? Date.parse(first.modifiedAt) : 0
    const secondTime = second.modifiedAt ? Date.parse(second.modifiedAt) : 0
    return secondTime - firstTime
  })

  return {
    rootPath,
    files,
    truncated: files.length >= limit
  }
}

function walk(currentPath: string, files: WorkspaceFile[], limit: number): void {
  if (files.length >= limit) return

  let entries = []
  try {
    entries = readdirSync(currentPath, { withFileTypes: true })
  } catch {
    return
  }

  for (const entry of entries) {
    if (files.length >= limit) return
    const fullPath = join(currentPath, entry.name)

    if (entry.isDirectory()) {
      if (!SKIPPED_DIRECTORIES.has(entry.name)) walk(fullPath, files, limit)
      continue
    }

    if (!entry.isFile() || !INDEXED_DOCUMENT_EXTENSIONS.has(extname(entry.name).toLowerCase())) continue

    try {
      files.push(toWorkspaceFile(fullPath))
    } catch {
      // Ignore files that disappeared or cannot be read while scanning.
    }
  }
}
