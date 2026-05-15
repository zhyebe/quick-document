import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType
} from 'docx'
import ExcelJS from 'exceljs'
import JSZip from 'jszip'
import pptxgen from 'pptxgenjs'
import { execFile } from 'node:child_process'
import { copyFileSync, mkdirSync, mkdtempSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { basename, dirname, extname, isAbsolute, join, posix, resolve } from 'node:path'
import { promisify } from 'node:util'
import type {
  ActionResult,
  GeneratedFile,
  OfficeAction,
  OfficeActionType,
  OfficeKind,
  SheetPlan,
  SlidePlan,
  WorkspaceSnapshot,
  WordSection
} from '@shared/types'

const execFileAsync = promisify(execFile)

export async function executeDocumentAction(
  action: OfficeAction,
  workspacePath: string,
  workspaceSnapshot?: WorkspaceSnapshot
): Promise<ActionResult> {
  mkdirSync(workspacePath, { recursive: true })

  try {
    if (action.type === 'skill_task') {
      return unsupportedSkillTask(action)
    }

    if (action.type === 'find_files') {
      return findOfficeFiles(action, workspacePath, workspaceSnapshot)
    }

    if (action.type === 'copy_file') {
      const file = copyOfficeFile(action, workspacePath)
      return {
        ok: true,
        actionType: action.type,
        kind: file.kind,
        file,
        summary: `已复制文件：${file.name}`
      }
    }

    if (action.type === 'replace_text') {
      const file = await replaceTextInOfficeFile(action, workspacePath)
      return {
        ok: true,
        actionType: action.type,
        kind: file.kind,
        file,
        summary: `已修改文本：${file.name}`
      }
    }

    if (action.type === 'update_excel_cells') {
      const file = await updateExcelCells(action, workspacePath)
      return {
        ok: true,
        actionType: action.type,
        kind: 'excel',
        file,
        summary: `已修改表格：${file.name}`
      }
    }

    if (action.type === 'run_javascript') {
      const file = await runJavascriptDocumentAction(action, workspacePath, workspaceSnapshot)
      return {
        ok: true,
        actionType: action.type,
        kind: file.kind,
        file,
        summary: `AI 已写入文档：${file.name}`
      }
    }

    if (action.type === 'create_docx' || action.type === 'revise_docx') {
      const file = await createWordDocument(action, workspacePath)
      return success(action.type, 'word', file, action.type === 'revise_docx')
    }

    if (action.type === 'create_xlsx' || action.type === 'revise_xlsx') {
      const file = await createExcelWorkbook(action, workspacePath)
      return success(action.type, 'excel', file, action.type === 'revise_xlsx')
    }

    if (action.type === 'create_pptx' || action.type === 'revise_pptx') {
      const file = await createPowerPointDeck(action, workspacePath)
      return success(action.type, 'powerpoint', file, action.type === 'revise_pptx')
    }

    return {
      ok: false,
      actionType: action.type,
      kind: 'unknown',
      summary: 'Unsupported document action.',
      error: `Unsupported action type: ${action.type}`
    }
  } catch (error) {
    return {
      ok: false,
      actionType: action.type,
      kind: kindFromAction(action.type),
      summary: 'Document action failed.',
      error: error instanceof Error ? error.message : String(error)
    }
  }
}

function unsupportedSkillTask(action: OfficeAction): ActionResult {
  const skillName = action.skillName || '对应 Office'
  return {
    ok: false,
    actionType: action.type,
    kind: kindFromSkillName(action.skillName),
    summary: `已调用 ${skillName} skill 思考，但缺少可落盘的工具动作。`,
    error:
      'AI 没有返回可执行的复制、创建、文本替换或单元格修改动作，因此没有修改任何文件。请更明确说明要改哪段文字、哪个单元格，或要创建/复制什么文档。'
  }
}

function copyOfficeFile(action: OfficeAction, workspacePath: string): GeneratedFile {
  const sourcePath = resolveSourcePath(action)
  const sourceKind = kindFromExtension(sourcePath)
  ensureSupportedOfficeFile(sourcePath, ['word', 'excel', 'powerpoint'])

  const destinationPath = resolveOutputPath(
    workspacePath,
    action.destinationPath || action.filename || `${basename(sourcePath, extname(sourcePath))}-副本${extname(sourcePath)}`,
    extname(sourcePath)
  )
  mkdirSync(dirname(destinationPath), { recursive: true })
  copyFileSync(sourcePath, destinationPath)
  return toGeneratedFile(destinationPath, sourceKind, 'Copied Office file')
}

function findOfficeFiles(
  action: OfficeAction,
  workspacePath: string,
  workspaceSnapshot?: WorkspaceSnapshot
): ActionResult {
  const query = (action.query || action.title || action.instructions || '').trim()
  const normalizedQuery = normalizeSearchText(query)
  const files = workspaceSnapshot?.files || []
  const kind = kindFromSkillName(action.skillName)
  const matched = files.filter((file) => {
    if (kind !== 'unknown' && file.kind !== kind) return false
    if (!normalizedQuery) return true
    const haystack = normalizeSearchText(`${file.name} ${file.path}`)
    return queryTokens(normalizedQuery).every((token) => haystack.includes(token))
  })

  if (matched.length === 0) {
    return {
      ok: false,
      actionType: action.type,
      kind,
      summary: query ? `没有找到匹配文件：${query}` : '没有找到匹配文件',
      error: `当前目录没有匹配的 Office 文件：${workspacePath}`
    }
  }

  const file = matched[0]
  return {
    ok: true,
    actionType: action.type,
    kind: file.kind,
    summary:
      matched.length === 1
        ? `已找到文件：${file.name}`
        : `已找到 ${matched.length} 个匹配文件，请选择要处理的文件。`,
    file: matched.length === 1 ? toGeneratedFile(file.path, file.kind, 'Matched Office file') : undefined,
    error: matched.length === 1 ? undefined : matched.map((item) => item.name).join('、')
  }
}

async function replaceTextInOfficeFile(action: OfficeAction, workspacePath: string): Promise<GeneratedFile> {
  const targetPath = resolveTargetPath(action)
  const kind = kindFromExtension(targetPath)
  ensureSupportedOfficeFile(targetPath, ['word', 'powerpoint'])
  if (!action.replacements?.length) {
    throw new Error('缺少可执行的文本替换规则。')
  }

  const outputPath = outputPathForEdit(action, workspacePath, targetPath)
  if (outputPath !== targetPath) {
    mkdirSync(dirname(outputPath), { recursive: true })
    copyFileSync(targetPath, outputPath)
  }

  const changed = await replaceTextInZipXml(outputPath, action.replacements, kind)
  if (changed === 0) {
    throw new Error('没有在目标文档中找到要替换的文本。')
  }

  return toGeneratedFile(outputPath, kind, `${changed} text replacement(s)`)
}

async function updateExcelCells(action: OfficeAction, workspacePath: string): Promise<GeneratedFile> {
  const targetPath = resolveTargetPath(action)
  ensureSupportedOfficeFile(targetPath, ['excel'])
  if (!action.cellUpdates?.length) {
    throw new Error('缺少可执行的 Excel 单元格修改规则。')
  }

  const outputPath = outputPathForEdit(action, workspacePath, targetPath)
  if (outputPath !== targetPath) {
    mkdirSync(dirname(outputPath), { recursive: true })
    copyFileSync(targetPath, outputPath)
  }

  await updateExcelCellsInZip(outputPath, action.cellUpdates)
  return toGeneratedFile(outputPath, 'excel', `${action.cellUpdates.length} cell update(s)`)
}

async function updateExcelCellsInZip(
  filePath: string,
  cellUpdates: Array<{ sheet?: string; cell: string; value: unknown }>
): Promise<void> {
  const zip = await JSZip.loadAsync(readFileSync(filePath))
  const sheets = await readWorkbookSheets(zip)
  if (sheets.length === 0) throw new Error('目标 Excel 没有可编辑工作表。')

  const groupedUpdates = new Map<string, Array<{ cell: string; value: unknown }>>()
  for (const update of cellUpdates) {
    const sheet = resolveWorkbookSheet(sheets, update.sheet)
    if (!sheet) throw new Error(update.sheet ? `找不到工作表：${update.sheet}` : '目标 Excel 没有可编辑工作表。')
    const cellRef = normalizeCellRef(update.cell)
    const updates = groupedUpdates.get(sheet.path) || []
    updates.push({ cell: cellRef, value: update.value })
    groupedUpdates.set(sheet.path, updates)
  }

  for (const [sheetPath, updates] of groupedUpdates) {
    const entry = zip.file(sheetPath)
    if (!entry) throw new Error(`找不到工作表文件：${sheetPath}`)
    let xml = await entry.async('string')
    for (const update of updates) {
      xml = setSheetCellValue(xml, update.cell, update.value)
    }
    zip.file(sheetPath, xml)
  }

  const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
  writeFileSync(filePath, buffer)
}

async function runJavascriptDocumentAction(
  action: OfficeAction,
  workspacePath: string,
  workspaceSnapshot?: WorkspaceSnapshot
): Promise<GeneratedFile> {
  if (!action.script?.trim()) throw new Error('AI 未提供可执行脚本。')
  const workspaceRoot = resolve(workspacePath)
  const tempDir = mkdtempSync(join(tmpdir(), 'quick-document-ai-script-'))
  const scriptPath = join(tempDir, 'action.mjs')
  const resultPath = join(tempDir, 'result.json')
  const bootstrap = buildJavascriptActionBootstrap(action, workspaceRoot, workspaceSnapshot, resultPath)
  const script = rewriteKnownPackageImports(action.script)

  writeFileSync(scriptPath, `${bootstrap}\n\n${script}\n`, 'utf8')

  try {
    const { stdout, stderr } = await execFileAsync(process.execPath, [scriptPath], {
      cwd: workspaceRoot,
      timeout: 120_000,
      maxBuffer: 16 * 1024 * 1024,
      env: {
        ...process.env,
        ELECTRON_RUN_AS_NODE: '1',
        QUICK_DOCUMENT_WORKSPACE: workspaceRoot,
        QUICK_DOCUMENT_RESULT_PATH: resultPath
      }
    })
    const generatedPath = readJavascriptActionResult(resultPath, stdout, action)
    const safePath = assertPathInsideWorkspace(generatedPath, workspaceRoot)
    statSync(safePath)
    return toGeneratedFile(safePath, kindFromExtension(safePath), stderr.trim() || 'AI JavaScript document action')
  } catch (error) {
    throw new Error(formatScriptError(error))
  } finally {
    rmSync(tempDir, { recursive: true, force: true })
  }
}

function rewriteKnownPackageImports(script: string | undefined): string {
  if (!script) return ''
  return script
    .replace(/^\s*import\s+ExcelJS\s+from\s+['"]exceljs['"]\s*;?\s*$/gm, 'const ExcelJS = quickDocument.ExcelJS')
    .replace(/^\s*import\s+\*\s+as\s+ExcelJS\s+from\s+['"]exceljs['"]\s*;?\s*$/gm, 'const ExcelJS = quickDocument.ExcelJS')
    .replace(/^\s*import\s+JSZip\s+from\s+['"]jszip['"]\s*;?\s*$/gm, 'const JSZip = quickDocument.JSZip')
    .replace(/^\s*import\s+\*\s+as\s+JSZip\s+from\s+['"]jszip['"]\s*;?\s*$/gm, 'const JSZip = quickDocument.JSZip')
    .replace(/^\s*import\s+pptxgen\s+from\s+['"]pptxgenjs['"]\s*;?\s*$/gim, 'const pptxgen = quickDocument.pptxgen')
    .replace(/^\s*import\s+PptxGenJS\s+from\s+['"]pptxgenjs['"]\s*;?\s*$/gm, 'const PptxGenJS = quickDocument.pptxgen')
    .replace(/^\s*import\s+\*\s+as\s+docx\s+from\s+['"]docx['"]\s*;?\s*$/gm, 'const docx = quickDocument.docx')
    .replace(/^\s*import\s+\{([^}]+)\}\s+from\s+['"]docx['"]\s*;?\s*$/gm, (_match, names: string) => {
      return names
        .split(',')
        .map((name) => name.trim())
        .filter(Boolean)
        .map((name) => {
          const [source, alias] = name.split(/\s+as\s+/i).map((part) => part.trim())
          return `const ${alias || source} = quickDocument.docx.${source}`
        })
        .join('\n')
    })
}

function buildJavascriptActionBootstrap(
  action: OfficeAction,
  workspacePath: string,
  workspaceSnapshot: WorkspaceSnapshot | undefined,
  resultPath: string
): string {
  return `
{
  const { createRequire: __quickDocumentCreateRequire } = await import('node:module')
  const __quickDocumentFs = await import('node:fs')
  const __quickDocumentPath = await import('node:path')
  const __quickDocumentRequire = __quickDocumentCreateRequire(${JSON.stringify(import.meta.url)})
  const __quickDocumentWorkspacePath = ${JSON.stringify(workspacePath)}
  const __quickDocumentAction = ${JSON.stringify(action)}
  const __quickDocumentWorkspaceFiles = ${JSON.stringify(workspaceSnapshot?.files || [])}
  const __quickDocumentTargetPaths = ${JSON.stringify(action.targetPaths || [])}
  const __quickDocumentSourcePath = ${JSON.stringify(action.sourcePath || '')}
  const __quickDocumentDestinationPath = ${JSON.stringify(action.destinationPath || action.filename || '')}
  const __quickDocumentResultPath = ${JSON.stringify(resultPath)}
  const __quickDocumentExcelJS = __quickDocumentRequire('exceljs')
  const __quickDocumentJSZip = __quickDocumentRequire('jszip')
  const __quickDocumentDocx = __quickDocumentRequire('docx')
  const __quickDocumentPptxgen = __quickDocumentRequire('pptxgenjs')
  const __quickDocumentAssertInsideWorkspace = (filePath) => {
    const resolved = __quickDocumentPath.resolve(filePath)
    const root = __quickDocumentPath.resolve(__quickDocumentWorkspacePath)
    if (resolved !== root && !resolved.startsWith(root + __quickDocumentPath.sep)) {
      throw new Error('脚本尝试访问工作区外路径：' + filePath)
    }
    return resolved
  }
  const __quickDocumentWriteResult = (result) => {
    __quickDocumentFs.writeFileSync(__quickDocumentResultPath, JSON.stringify(result || {}, null, 2), 'utf8')
  }
  globalThis.quickDocument = {
    workspacePath: __quickDocumentWorkspacePath,
    action: __quickDocumentAction,
    workspaceFiles: __quickDocumentWorkspaceFiles,
    targetPaths: __quickDocumentTargetPaths,
    sourcePath: __quickDocumentSourcePath,
    destinationPath: __quickDocumentDestinationPath,
    ExcelJS: __quickDocumentExcelJS,
    JSZip: __quickDocumentJSZip,
    docx: __quickDocumentDocx,
    pptxgen: __quickDocumentPptxgen,
    fs: __quickDocumentFs,
    path: __quickDocumentPath,
    require: __quickDocumentRequire,
    createRequire: __quickDocumentCreateRequire,
    assertInsideWorkspace: __quickDocumentAssertInsideWorkspace,
    writeResult: __quickDocumentWriteResult
  }
  Object.assign(globalThis, {
    workspacePath: __quickDocumentWorkspacePath,
    action: __quickDocumentAction,
    workspaceFiles: __quickDocumentWorkspaceFiles,
    targetPaths: __quickDocumentTargetPaths,
    sourcePath: __quickDocumentSourcePath,
    destinationPath: __quickDocumentDestinationPath,
    ExcelJS: __quickDocumentExcelJS,
    JSZip: __quickDocumentJSZip,
    docx: __quickDocumentDocx,
    pptxgen: __quickDocumentPptxgen,
    fs: __quickDocumentFs,
    path: __quickDocumentPath,
    require: __quickDocumentRequire,
    createRequire: __quickDocumentCreateRequire,
    assertInsideWorkspace: __quickDocumentAssertInsideWorkspace,
    writeResult: __quickDocumentWriteResult
  })
}
`.trim()
}

function readJavascriptActionResult(resultPath: string, stdout: string, action: OfficeAction): string {
  const resultText = existsByStat(resultPath) ? readFileSync(resultPath, 'utf8') : extractLastJsonObject(stdout)
  const result = resultText.trim() ? JSON.parse(resultText) as Record<string, unknown> : {}
  const filePath = stringResultValue(result.filePath) || firstStringResultValue(result.filePaths)
  if (filePath) return filePath
  const targetPath = action.destinationPath || action.targetPaths?.[0] || action.sourcePath
  if (targetPath) return targetPath
  throw new Error('AI 脚本执行完成，但没有返回写入的文件路径。')
}

function extractLastJsonObject(value: string): string {
  const trimmed = value.trim()
  const start = trimmed.lastIndexOf('{')
  return start >= 0 ? trimmed.slice(start) : ''
}

function stringResultValue(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() ? value.trim() : undefined
}

function firstStringResultValue(value: unknown): string | undefined {
  return Array.isArray(value) ? value.find((item): item is string => typeof item === 'string' && item.trim().length > 0) : undefined
}

function assertPathInsideWorkspace(filePath: string, workspaceRoot: string): string {
  const resolvedPath = resolve(filePath)
  const resolvedRoot = resolve(workspaceRoot)
  if (resolvedPath !== resolvedRoot && !resolvedPath.startsWith(`${resolvedRoot}/`) && !resolvedPath.startsWith(`${resolvedRoot}\\`)) {
    throw new Error(`AI 脚本返回了工作区外路径：${filePath}`)
  }
  return resolvedPath
}

function formatScriptError(error: unknown): string {
  if (!error || typeof error !== 'object') return String(error)
  const record = error as { message?: string; stdout?: string; stderr?: string }
  return [record.message, record.stdout, record.stderr].filter(Boolean).join('\n')
}

interface WorkbookSheetInfo {
  name: string
  relationshipId: string
  path: string
}

async function readWorkbookSheets(zip: JSZip): Promise<WorkbookSheetInfo[]> {
  const workbookXml = await zip.file('xl/workbook.xml')?.async('string')
  const relationshipsXml = await zip.file('xl/_rels/workbook.xml.rels')?.async('string')
  if (!workbookXml || !relationshipsXml) return []

  const relationshipTargets = new Map<string, string>()
  for (const relationship of relationshipsXml.matchAll(/<Relationship\b([^>]*)\/?>/g)) {
    const attrs = parseXmlAttributes(relationship[1])
    const id = attrs.get('Id')
    const target = attrs.get('Target')
    if (!id || !target) continue
    relationshipTargets.set(id, workbookRelationshipTargetToPath(target))
  }

  const sheets: WorkbookSheetInfo[] = []
  for (const sheet of workbookXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
    const attrs = parseXmlAttributes(sheet[1])
    const name = attrs.get('name')
    const relationshipId = attrs.get('r:id') || attrs.get('id')
    if (!name || !relationshipId) continue
    const path = relationshipTargets.get(relationshipId)
    if (!path) continue
    sheets.push({
      name: unescapeXmlText(name),
      relationshipId,
      path
    })
  }
  return sheets
}

function workbookRelationshipTargetToPath(target: string): string {
  const cleanTarget = target.replace(/\\/g, '/')
  if (cleanTarget.startsWith('/')) return cleanTarget.replace(/^\/+/, '')
  return posix.normalize(posix.join('xl', cleanTarget))
}

function resolveWorkbookSheet(
  sheets: WorkbookSheetInfo[],
  wantedName: string | undefined
): WorkbookSheetInfo | undefined {
  if (!wantedName) return sheets[0]
  const wanted = normalizeSheetLookupText(wantedName)
  return (
    sheets.find((sheet) => normalizeSheetLookupText(sheet.name) === wanted) ||
    sheets.find((sheet) => normalizeSheetLookupText(sheet.name).includes(wanted)) ||
    sheets[0]
  )
}

function normalizeSheetLookupText(value: string): string {
  return value.toLowerCase().replace(/\s+/g, '')
}

function normalizeCellRef(value: string): string {
  const normalized = value.replace(/\$/g, '').toUpperCase().trim()
  if (!/^[A-Z]{1,3}[1-9][0-9]*$/.test(normalized)) {
    throw new Error(`Excel 单元格地址不正确：${value}`)
  }
  return normalized
}

function setSheetCellValue(xml: string, cellRef: string, value: unknown): string {
  const rowNumber = Number(cellRef.match(/[0-9]+$/)?.[0])
  const rowRegex = new RegExp(`(<row\\b[^>]*\\br=["']${rowNumber}["'][^>]*>)([\\s\\S]*?)(<\\/row>)`)
  const cellXml = buildCellXml(cellRef, value, findCellStyle(xml, cellRef))

  if (rowRegex.test(xml)) {
    return xml.replace(rowRegex, (_rowMatch, rowOpen: string, rowInner: string, rowClose: string) => {
      const cellRegex = new RegExp(`<c\\b[^>]*\\br=["']${escapeRegExp(cellRef)}["'][\\s\\S]*?<\\/c>|<c\\b[^>]*\\br=["']${escapeRegExp(cellRef)}["'][^>]*/>`)
      if (cellRegex.test(rowInner)) {
        return `${rowOpen}${rowInner.replace(cellRegex, cellXml)}${rowClose}`
      }
      return `${rowOpen}${insertCellInRow(rowInner, cellXml, cellRef)}${rowClose}`
    })
  }

  const sheetDataRegex = /(<sheetData[^>]*>)([\s\S]*?)(<\/sheetData>)/
  const rowXml = `<row r="${rowNumber}">${cellXml}</row>`
  if (sheetDataRegex.test(xml)) {
    return xml.replace(sheetDataRegex, (_match, open: string, inner: string, close: string) => {
      return `${open}${insertRowInSheetData(inner, rowXml, rowNumber)}${close}`
    })
  }

  return xml.replace(/(<worksheet\b[^>]*>)/, `$1<sheetData>${rowXml}</sheetData>`)
}

function buildCellXml(cellRef: string, value: unknown, styleAttribute: string): string {
  const prefix = `<c r="${cellRef}"${styleAttribute}`
  if (value === null || value === undefined || value === '') return `${prefix}/>`
  if (typeof value === 'number') return `${prefix}><v>${Number.isFinite(value) ? value : ''}</v></c>`
  if (typeof value === 'boolean') return `${prefix} t="b"><v>${value ? 1 : 0}</v></c>`
  const text = typeof value === 'string' ? value : JSON.stringify(value)
  return `${prefix} t="inlineStr"><is><t${textNeedsPreserveSpace(text) ? ' xml:space="preserve"' : ''}>${escapeXmlText(
    text
  )}</t></is></c>`
}

function findCellStyle(xml: string, cellRef: string): string {
  const cellMatch = xml.match(new RegExp(`<c\\b([^>]*)\\br=["']${escapeRegExp(cellRef)}["']([^>]*)[\\s\\S]*?(?:<\\/c>|\\/>)`))
  if (!cellMatch) return ''
  const attrs = parseXmlAttributes(`${cellMatch[1]} ${cellMatch[2]}`)
  const style = attrs.get('s')
  return style ? ` s="${escapeXmlAttribute(style)}"` : ''
}

function insertCellInRow(rowInner: string, cellXml: string, cellRef: string): string {
  const targetColumn = columnNumberFromCellRef(cellRef)
  const cellMatches = Array.from(rowInner.matchAll(/<c\b[^>]*\br=["']([A-Z]+)[0-9]+["'][\s\S]*?(?:<\/c>|\/>)/g))
  for (const match of cellMatches) {
    if (columnNumberFromLetters(match[1]) > targetColumn) {
      return `${rowInner.slice(0, match.index)}${cellXml}${rowInner.slice(match.index)}`
    }
  }
  return `${rowInner}${cellXml}`
}

function insertRowInSheetData(inner: string, rowXml: string, rowNumber: number): string {
  const rowMatches = Array.from(inner.matchAll(/<row\b[^>]*\br=["']([0-9]+)["'][\s\S]*?<\/row>/g))
  for (const match of rowMatches) {
    if (Number(match[1]) > rowNumber) {
      return `${inner.slice(0, match.index)}${rowXml}${inner.slice(match.index)}`
    }
  }
  return `${inner}${rowXml}`
}

function columnNumberFromCellRef(cellRef: string): number {
  return columnNumberFromLetters(cellRef.replace(/[0-9]/g, ''))
}

function columnNumberFromLetters(letters: string): number {
  return letters.split('').reduce((total, letter) => total * 26 + letter.charCodeAt(0) - 64, 0)
}

function parseXmlAttributes(value: string): Map<string, string> {
  const attrs = new Map<string, string>()
  for (const match of value.matchAll(/([\w:.-]+)\s*=\s*(["'])(.*?)\2/g)) {
    attrs.set(match[1], match[3])
  }
  return attrs
}

function textNeedsPreserveSpace(value: string): boolean {
  return /^\s|\s$/.test(value)
}

function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replace(/"/g, '&quot;')
}

async function replaceTextInZipXml(
  filePath: string,
  replacements: Array<{ find: string; replace: string }>,
  kind: OfficeKind
): Promise<number> {
  const zip = await JSZip.loadAsync(readFileSync(filePath))
  const xmlPaths =
    kind === 'word'
      ? ['word/document.xml', ...zip.file(/^word\/(header|footer)\d+\.xml$/).map((file) => file.name)]
      : zip.file(/^ppt\/slides\/slide\d+\.xml$/).map((file) => file.name)

  let changes = 0
  for (const xmlPath of xmlPaths) {
    const entry = zip.file(xmlPath)
    if (!entry) continue
    const original = await entry.async('string')
    const replacementResult =
      kind === 'word'
        ? replaceTextAcrossXmlRuns(original, 'w:p', 'w:t', replacements)
        : replaceTextAcrossXmlRuns(original, 'a:p', 'a:t', replacements)
    const updated = replacementResult.xml
    changes += replacementResult.changes
    if (updated !== original) zip.file(xmlPath, updated)
  }

  if (changes > 0) {
    const buffer = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    writeFileSync(filePath, buffer)
  }
  return changes
}

function replaceTextAcrossXmlRuns(
  xml: string,
  paragraphTag: string,
  textTag: string,
  replacements: Array<{ find: string; replace: string }>
): { xml: string; changes: number } {
  const paragraphRegex = new RegExp(`<${escapeRegExp(paragraphTag)}[\\s\\S]*?<\\/${escapeRegExp(paragraphTag)}>`, 'g')
  const textRegex = new RegExp(`(<${escapeRegExp(textTag)}[^>]*>)([\\s\\S]*?)(<\\/${escapeRegExp(textTag)}>)`, 'g')
  let changes = 0

  const updatedXml = xml.replace(paragraphRegex, (paragraph) => {
    const matches = Array.from(paragraph.matchAll(textRegex))
    if (matches.length === 0) return paragraph

    let visibleText = matches.map((match) => unescapeXmlText(match[2])).join('')
    let nextText = visibleText
    for (const replacement of replacements) {
      const count = countOccurrences(nextText, replacement.find)
      if (count > 0) {
        nextText = nextText.split(replacement.find).join(replacement.replace)
        changes += count
      }
    }

    if (nextText === visibleText) return paragraph

    let index = 0
    return paragraph.replace(textRegex, (match, open, _text, close) => {
      const nextValue = index === 0 ? escapeXmlText(nextText) : ''
      index += 1
      return `${open}${nextValue}${close}`
    })
  })

  return { xml: updatedXml, changes }
}

async function createWordDocument(action: OfficeAction, workspacePath: string): Promise<GeneratedFile> {
  const title = action.title || titleFromFilename(action.filename) || 'Quick Document'
  const sections = normalizeWordSections(action.sections, title)
  const filePath = uniqueFilePath(workspacePath, action.filename || title, '.docx')
  const children: Array<Paragraph | Table> = [
    new Paragraph({
      text: title,
      heading: HeadingLevel.TITLE,
      spacing: { after: 240 }
    }),
    new Paragraph({
      children: [
        new TextRun({
          text: action.type === 'revise_docx' ? 'AI-assisted revision' : 'AI-generated document',
          italics: true,
          color: '56616F'
        })
      ],
      spacing: { after: 320 }
    })
  ]

  sections.forEach((section) => {
    if (section.heading) {
      children.push(
        new Paragraph({
          text: section.heading,
          heading: HeadingLevel.HEADING_1,
          spacing: { before: 180, after: 120 }
        })
      )
    }

    section.paragraphs?.forEach((text) => {
      children.push(
        new Paragraph({
          children: [new TextRun({ text })],
          spacing: { after: 160 },
          alignment: AlignmentType.LEFT
        })
      )
    })

    section.bullets?.forEach((text) => {
      children.push(
        new Paragraph({
          text,
          bullet: { level: 0 },
          spacing: { after: 80 }
        })
      )
    })
  })

  const doc = new Document({
    creator: 'Quick Document',
    title,
    description: 'Created by Quick Document desktop assistant',
    styles: {
      paragraphStyles: [
        {
          id: 'Normal',
          name: 'Normal',
          run: {
            font: 'Arial',
            size: 22,
            color: '1F2937'
          },
          paragraph: {
            spacing: { line: 320, after: 120 }
          }
        },
        {
          id: 'Title',
          name: 'Title',
          basedOn: 'Normal',
          run: {
            font: 'Arial',
            size: 40,
            bold: true,
            color: '111827'
          },
          paragraph: {
            spacing: { after: 220 }
          }
        },
        {
          id: 'Heading1',
          name: 'Heading 1',
          basedOn: 'Normal',
          next: 'Normal',
          quickFormat: true,
          run: {
            font: 'Arial',
            size: 28,
            bold: true,
            color: '0F766E'
          },
          paragraph: {
            spacing: { before: 260, after: 100 }
          }
        }
      ]
    },
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: 1440,
              right: 1440,
              bottom: 1440,
              left: 1440
            }
          }
        },
        children
      }
    ]
  })

  const buffer = await Packer.toBuffer(doc)
  writeFileSync(filePath, buffer)
  return toGeneratedFile(filePath, 'word', `${sections.length} section Word document`)
}

async function createExcelWorkbook(
  action: OfficeAction,
  workspacePath: string
): Promise<GeneratedFile> {
  const title = action.title || titleFromFilename(action.filename) || 'Quick Workbook'
  const sheets = normalizeSheets(action.sheets)
  const filePath = uniqueFilePath(workspacePath, action.filename || title, '.xlsx')
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Quick Document'
  workbook.created = new Date()

  sheets.forEach((sheetPlan) => {
    const worksheet = workbook.addWorksheet(safeSheetName(sheetPlan.name || 'Sheet 1'), {
      views: [{ state: 'frozen', ySplit: 1 }]
    })
    const columns = deriveColumns(sheetPlan)
    worksheet.columns = columns.map((column) => ({
      header: column,
      key: column,
      width: Math.min(Math.max(column.length + 8, 14), 34)
    }))

    const rows = sheetPlan.rows || []
    rows.forEach((row) => {
      if (Array.isArray(row)) {
        const record = Object.fromEntries(columns.map((column, index) => [column, row[index] ?? '']))
        worksheet.addRow(record)
      } else {
        worksheet.addRow(row)
      }
    })

    const header = worksheet.getRow(1)
    header.font = { bold: true, color: { argb: 'FFFFFFFF' } }
    header.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F766E' }
    }
    header.alignment = { vertical: 'middle', horizontal: 'center' }
    header.height = 22

    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        cell.alignment = {
          vertical: 'middle',
          wrapText: true
        }
        cell.border = {
          top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
          right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
        }
      })
      if (rowNumber > 1) row.height = 24
    })
  })

  await workbook.xlsx.writeFile(filePath)
  return toGeneratedFile(filePath, 'excel', `${sheets.length} sheet Excel workbook`)
}

async function createPowerPointDeck(
  action: OfficeAction,
  workspacePath: string
): Promise<GeneratedFile> {
  const title = action.title || titleFromFilename(action.filename) || 'Quick Presentation'
  const filePath = uniqueFilePath(workspacePath, action.filename || title, '.pptx')
  const pptx = new pptxgen()
  pptx.author = 'Quick Document'
  pptx.subject = 'AI-generated presentation'
  pptx.title = title
  pptx.company = 'Quick Document'
  pptx.layout = 'LAYOUT_WIDE'
  pptx.theme = {
    headFontFace: 'Aptos Display',
    bodyFontFace: 'Aptos'
  }

  const slides = normalizeSlides(action.slides, title)
  slides.forEach((slidePlan, index) => {
    const slide = pptx.addSlide()
    slide.background = { color: index === 0 ? 'F7F7F5' : 'FFFFFF' }
    slide.addText(slidePlan.title || title, {
      x: 0.65,
      y: 0.45,
      w: 11.0,
      h: 0.7,
      fontFace: 'Aptos Display',
      fontSize: index === 0 ? 30 : 25,
      bold: true,
      color: '111827',
      margin: 0
    })

    const bullets = slidePlan.bullets || []
    if (bullets.length === 0 && !slidePlan.notes) {
      throw new Error(`AI 未提供幻灯片「${slidePlan.title || title}」的正文内容，未创建文件。`)
    }
    slide.addText(
      bullets.map((bullet) => ({ text: bullet, options: { bullet: { type: 'bullet' } } })),
      {
        x: 0.85,
        y: 1.45,
        w: 10.8,
        h: 4.6,
        fontFace: 'Aptos',
        fontSize: 18,
        color: '1F2937',
        breakLine: false,
        fit: 'shrink',
        paraSpaceAfter: 10
      }
    )

    if (slidePlan.notes) {
      slide.addNotes(slidePlan.notes)
    }

    slide.addShape(pptx.ShapeType.line, {
      x: 0.65,
      y: 6.75,
      w: 11.0,
      h: 0,
      line: { color: 'D1D5DB', width: 1 }
    })
    slide.addText('Quick Document', {
      x: 0.65,
      y: 6.85,
      w: 3.0,
      h: 0.25,
      fontSize: 8,
      color: '6B7280',
      margin: 0
    })
  })

  await pptx.writeFile({ fileName: filePath })
  return toGeneratedFile(filePath, 'powerpoint', `${slides.length} slide PowerPoint deck`)
}

function normalizeWordSections(sections: WordSection[] | undefined, title: string): WordSection[] {
  if (sections && sections.length > 0) return sections
  throw new Error('AI 未提供可写入 Word 的正文结构，未创建文件。')
}

function normalizeSheets(sheets: SheetPlan[] | undefined): SheetPlan[] {
  const validSheets = (sheets || []).filter((sheet) => {
    const columns = deriveColumns(sheet)
    return columns.length > 0 && (sheet.rows?.length || sheet.columns?.length)
  })
  if (validSheets.length > 0) return validSheets
  throw new Error('AI 未提供可写入 Excel 的表格内容，未创建文件。')
}

function normalizeSlides(slides: SlidePlan[] | undefined, title: string): SlidePlan[] {
  if (slides && slides.length > 0) return slides
  throw new Error('AI 未提供可写入 PowerPoint 的幻灯片内容，未创建文件。')
}

function deriveColumns(sheet: SheetPlan): string[] {
  if (sheet.columns && sheet.columns.length > 0) return sheet.columns
  const firstRow = sheet.rows?.[0]
  if (firstRow && !Array.isArray(firstRow) && typeof firstRow === 'object') {
    return Object.keys(firstRow)
  }
  if (Array.isArray(firstRow)) {
    return firstRow.map((_value, index) => `Column ${index + 1}`)
  }
  return []
}

function safeSheetName(name: string): string {
  return name.replace(/[\\/?*[\]:]/g, ' ').trim().slice(0, 31) || 'Sheet'
}

function success(
  actionType: OfficeActionType,
  kind: OfficeKind,
  file: GeneratedFile,
  revised: boolean
): ActionResult {
  return {
    ok: true,
    actionType,
    kind,
    file,
    summary: revised ? `已创建修订副本：${file.name}` : `已创建文件：${file.name}`
  }
}

function kindFromAction(actionType: OfficeActionType): OfficeKind {
  if (actionType.endsWith('docx')) return 'word'
  if (actionType.endsWith('xlsx')) return 'excel'
  if (actionType.endsWith('pptx')) return 'powerpoint'
  if (actionType === 'update_excel_cells') return 'excel'
  return 'unknown'
}

function kindFromSkillName(skillName: OfficeAction['skillName']): OfficeKind {
  if (skillName === 'documents') return 'word'
  if (skillName === 'spreadsheets') return 'excel'
  if (skillName === 'presentations') return 'powerpoint'
  return 'unknown'
}

function uniqueFilePath(workspacePath: string, wantedName: string, extension: string): string {
  let candidate = baseOutputPath(workspacePath, wantedName, extension)
  let index = 2

  while (existsByStat(candidate)) {
    const directory = dirname(candidate)
    const baseName = sanitizeFileName(basename(candidate, extname(candidate))) || 'document'
    candidate = join(directory, `${baseName}-${index}${extension}`)
    index += 1
  }

  return candidate
}

function resolveSourcePath(action: OfficeAction): string {
  const sourcePath = action.sourcePath || action.targetPaths?.[0]
  if (!sourcePath) throw new Error('没有找到源文件路径。')
  return sourcePath
}

function resolveTargetPath(action: OfficeAction): string {
  const targetPath = action.targetPaths?.[0] || action.sourcePath
  if (!targetPath) throw new Error('没有找到目标文件路径。')
  return targetPath
}

function outputPathForEdit(action: OfficeAction, workspacePath: string, targetPath: string): string {
  const wantedPath = action.destinationPath || action.filename
  if (!wantedPath) return targetPath
  return resolveOutputPath(workspacePath, wantedPath, extname(targetPath))
}

function resolveOutputPath(workspacePath: string, wantedNameOrPath: string, extension: string): string {
  return uniqueFilePath(workspacePath, wantedNameOrPath, extension)
}

function baseOutputPath(workspacePath: string, wantedNameOrPath: string, extension: string): string {
  const raw = wantedNameOrPath.trim()
  const rawExtension = extname(raw)
  const nameWithExtension = rawExtension ? raw : `${raw}${extension}`
  if (isAbsolute(nameWithExtension)) return nameWithExtension
  return join(workspacePath, sanitizeRelativePath(nameWithExtension))
}

function sanitizeRelativePath(pathValue: string): string {
  return pathValue
    .split(/[\\/]+/)
    .filter((part) => part && part !== '.' && part !== '..')
    .map((part, index, parts) => {
      const sanitized = sanitizeFileName(part)
      if (sanitized) return sanitized
      return index === parts.length - 1 ? 'document' : 'folder'
    })
    .join('/')
}

function ensureSupportedOfficeFile(filePath: string, allowedKinds: OfficeKind[]): void {
  const kind = kindFromExtension(filePath)
  if (!allowedKinds.includes(kind)) {
    throw new Error(`当前操作不支持这个文件类型：${basename(filePath)}`)
  }
  statSync(filePath)
}

function kindFromExtension(filePath: string): OfficeKind {
  const extension = extname(filePath).toLowerCase()
  if (extension === '.docx') return 'word'
  if (extension === '.xlsx') return 'excel'
  if (extension === '.ppt' || extension === '.pptx') return 'powerpoint'
  return 'unknown'
}

function normalizeSearchText(value: string): string {
  return value
    .toLowerCase()
    .replace(/五月/g, '5月')
    .replace(/五\s*月/g, '5月')
    .replace(/\bmay\b/g, '5月')
    .replace(/\s+/g, '')
}

function queryTokens(normalizedQuery: string): string[] {
  const tokens = [normalizedQuery]
  const monthMatch = normalizedQuery.match(/(?:^|[^0-9])([1-9]|1[0-2])月/)
  if (monthMatch) tokens.push(`${monthMatch[1]}月`)
  if (/excel|xlsx|表格/.test(normalizedQuery)) tokens.push('.xlsx')
  if (/word|docx|文档/.test(normalizedQuery)) tokens.push('.docx')
  if (/ppt|pptx|powerpoint|演示|幻灯/.test(normalizedQuery)) tokens.push('.ppt')
  return Array.from(new Set(tokens.filter(Boolean)))
}

function escapeXmlText(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
}

function unescapeXmlText(value: string): string {
  return value
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&')
}

function escapeRegExp(value: string): string {
  return value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
}

function countOccurrences(value: string, search: string): number {
  if (!search) return 0
  return value.split(search).length - 1
}

function existsByStat(filePath: string): boolean {
  try {
    statSync(filePath)
    return true
  } catch {
    return false
  }
}

function sanitizeFileName(name: string): string {
  return name
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 90)
}

function titleFromFilename(filename: string | undefined): string {
  if (!filename) return ''
  return basename(filename, extname(filename)).replace(/[-_]+/g, ' ').trim()
}

function toGeneratedFile(filePath: string, kind: OfficeKind, summary: string): GeneratedFile {
  const stats = statSync(filePath)
  return {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    name: basename(filePath),
    path: filePath,
    kind,
    size: stats.size,
    createdAt: new Date().toISOString(),
    summary
  }
}
