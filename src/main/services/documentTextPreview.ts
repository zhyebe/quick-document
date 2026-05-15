import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import { readFileSync } from 'node:fs'
import { extname } from 'node:path'
import type { WorkspaceFile, WorkspaceSnapshot } from '@shared/types'

const MAX_PREVIEW_FILES = 8
const MAX_PREVIEW_CHARS = 18000

export async function buildDocumentPreviewContext(
  prompt: string,
  targetFiles: WorkspaceFile[],
  workspaceSnapshot?: WorkspaceSnapshot
): Promise<string> {
  const files = selectPreviewFiles(prompt, targetFiles, workspaceSnapshot)
  if (files.length === 0) return ''

  const previews: string[] = []
  for (const file of files) {
    try {
      const text = await extractOfficeText(file)
      if (!text.trim()) continue
      previews.push(`文件：${file.name}\n路径：${file.path}\n内容预览：\n${clip(text, 2600)}`)
    } catch {
      previews.push(`文件：${file.name}\n路径：${file.path}\n内容预览：读取失败，请基于文件名和用户指令谨慎规划。`)
    }
  }

  return clip(previews.join('\n\n---\n\n'), MAX_PREVIEW_CHARS)
}

function selectPreviewFiles(
  prompt: string,
  targetFiles: WorkspaceFile[],
  workspaceSnapshot?: WorkspaceSnapshot
): WorkspaceFile[] {
  if (targetFiles.length > 0) return targetFiles.slice(0, MAX_PREVIEW_FILES)
  if (!workspaceSnapshot) return []

  const normalizedPrompt = normalizeText(prompt)
  const scoredFiles = workspaceSnapshot.files
    .filter((file) => {
      const wantedKind = kindHintFromPrompt(normalizedPrompt)
      return wantedKind === 'unknown' || file.kind === wantedKind
    })
    .map((file) => ({
      file,
      score: scoreFileForPrompt(file, normalizedPrompt)
    }))
    .filter((entry) => entry.score > 0)
    .sort((first, second) => second.score - first.score)

  return scoredFiles
    .map((entry) => entry.file)
    .slice(0, MAX_PREVIEW_FILES)
}

function scoreFileForPrompt(file: WorkspaceFile, normalizedPrompt: string): number {
  const name = normalizeText(file.name)
  const stem = normalizeText(file.name.replace(/\.(docx|xlsx|ppt|pptx)$/i, ''))
  let score = 0
  if (normalizedPrompt.includes(name)) score += 20
  if (stem && normalizedPrompt.includes(stem)) score += 16
  for (const token of promptTokens(normalizedPrompt)) {
    if (name.includes(token)) score += token.length >= 3 ? 4 : 1
  }
  const modifiedAt = file.modifiedAt ? Date.parse(file.modifiedAt) : 0
  if (modifiedAt) score += Math.max(0, Math.min(2, modifiedAt / Date.now()))
  return score
}

function kindHintFromPrompt(normalizedPrompt: string): WorkspaceFile['kind'] {
  if (/excel|xlsx|表格|工作簿|台账|清单/.test(normalizedPrompt)) return 'excel'
  if (/word|docx|文档|报告|合同|方案/.test(normalizedPrompt)) return 'word'
  if (/ppt|pptx|powerpoint|演示|幻灯|汇报/.test(normalizedPrompt)) return 'powerpoint'
  return 'unknown'
}

function promptTokens(value: string): string[] {
  return Array.from(new Set(value.split(/[^a-z0-9\u4e00-\u9fa5]+/i).filter((token) => token.length >= 2)))
}

function normalizeText(value: string): string {
  return value
    .toLowerCase()
    .replace(/五月/g, '5月')
    .replace(/五\s*月/g, '5月')
    .replace(/\s+/g, '')
}

async function extractOfficeText(file: WorkspaceFile): Promise<string> {
  if (file.kind === 'word') return extractDocxText(file.path)
  if (file.kind === 'powerpoint') return extractPresentationText(file.path)
  if (file.kind === 'excel') return extractWorkbookText(file.path)
  return ''
}

export async function extractOfficeTextByPath(filePath: string): Promise<string> {
  const extension = extname(filePath).toLowerCase()
  if (extension === '.docx') return extractDocxText(filePath)
  if (extension === '.xlsx') return extractWorkbookText(filePath)
  if (extension === '.ppt' || extension === '.pptx') return extractPresentationText(filePath)
  return ''
}

async function extractDocxText(filePath: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(filePath))
  const xml = await zip.file('word/document.xml')?.async('string')
  if (!xml) return ''
  return extractTextTags(xml, 'w:t').join('')
}

async function extractPresentationText(filePath: string): Promise<string> {
  const zip = await JSZip.loadAsync(readFileSync(filePath))
  const slideFiles = zip.file(/^ppt\/slides\/slide\d+\.xml$/)
  const slides: string[] = []
  for (const slideFile of slideFiles.slice(0, 20)) {
    const xml = await slideFile.async('string')
    const text = extractTextTags(xml, 'a:t').join(' ').replace(/\s+/g, ' ').trim()
    if (text) slides.push(`${slideFile.name}: ${text}`)
  }
  return slides.join('\n')
}

async function extractWorkbookText(filePath: string): Promise<string> {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.readFile(filePath)
  const sheets: string[] = []

  workbook.worksheets.slice(0, 6).forEach((worksheet) => {
    const rows: string[] = []
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      if (rowNumber > 60) return
      const values = row.values
      if (!Array.isArray(values)) return
      const cells = values
        .slice(1, 14)
        .map((value, index) => {
          const formatted = formatCellValue(value)
          return formatted ? `${columnLetters(index + 1)}${rowNumber}=${formatted}` : ''
        })
        .filter((value) => value.length > 0)
      if (cells.length > 0) rows.push(`${rowNumber}: ${cells.join(' | ')}`)
    })
    if (rows.length > 0) sheets.push(`工作表 ${worksheet.name}\n${rows.join('\n')}`)
  })

  return sheets.join('\n\n')
}

function extractTextTags(xml: string, tagName: string): string[] {
  const escapedTag = tagName.replace(':', '\\:')
  const regex = new RegExp(`<${escapedTag}[^>]*>([\\s\\S]*?)<\\/${escapedTag}>`, 'g')
  return Array.from(xml.matchAll(regex), (match) => unescapeXmlText(match[1]))
}

function unescapeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

function formatCellValue(value: unknown): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'object') {
    if ('text' in value && typeof value.text === 'string') return value.text
    if ('result' in value) return String(value.result ?? '')
    if ('formula' in value && typeof value.formula === 'string') return `=${value.formula}`
  }
  return String(value)
}

function columnLetters(index: number): string {
  let value = index
  let letters = ''
  while (value > 0) {
    const remainder = (value - 1) % 26
    letters = String.fromCharCode(65 + remainder) + letters
    value = Math.floor((value - 1) / 26)
  }
  return letters
}

function clip(value: string, maxLength: number): string {
  return value.length > maxLength ? `${value.slice(0, maxLength)}\n...[已截断]` : value
}
