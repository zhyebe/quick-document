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
import pptxgen from 'pptxgenjs'
import { mkdirSync, statSync, writeFileSync } from 'node:fs'
import { basename, extname, join } from 'node:path'
import type {
  ActionResult,
  DocumentWorkflowRun,
  GeneratedFile,
  OfficeAction,
  OfficeActionType,
  OfficeKind,
  SheetPlan,
  SlidePlan,
  WordSection
} from '@shared/types'

export async function executeDocumentAction(
  action: OfficeAction,
  workspacePath: string
): Promise<ActionResult> {
  mkdirSync(workspacePath, { recursive: true })

  try {
    if (action.type === 'skill_task') {
      const workflow = createSkillTask(action, workspacePath)
      return {
        ok: true,
        actionType: action.type,
        kind: kindFromSkillName(action.skillName),
        summary: `已生成文档处理工作流：${workflow.title}`,
        workflow
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

function createSkillTask(action: OfficeAction, workspacePath: string): DocumentWorkflowRun {
  const title = action.title || 'Document workflow'
  const tasksPath = join(workspacePath, '.quick-document', 'tasks')
  mkdirSync(tasksPath, { recursive: true })

  const workflow: DocumentWorkflowRun = {
    id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
    title,
    targetPaths: action.targetPaths || (action.sourcePath ? [action.sourcePath] : []),
    createdAt: new Date().toISOString(),
    stages: [
      {
        id: 'select',
        label: '选择目录和目标文档',
        status: 'complete',
        detail: (action.targetPaths || []).join('\n')
      },
      {
        id: 'plan',
        label: `规划 ${action.skillName || 'documents'} skill 任务`,
        status: 'complete',
        detail: action.instructions
      },
      {
        id: 'execute',
        label: '执行文档修改',
        status: 'blocked',
        detail: '等待接入本地 skill runner 后自动执行。'
      },
      {
        id: 'verify',
        label: '校验输出',
        status: 'idle',
        detail: action.expectedOutput
      },
      {
        id: 'done',
        label: '完成',
        status: 'idle'
      }
    ]
  }

  const taskFilePath = join(tasksPath, `${sanitizeFileName(title) || 'workflow'}-${workflow.id}.json`)
  workflow.taskFilePath = taskFilePath
  writeFileSync(
    taskFilePath,
    JSON.stringify(
      {
        workflow,
        action
      },
      null,
      2
    ),
    'utf8'
  )

  return workflow
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
  const filePath = uniqueFilePath(workspacePath, action.filename || title, '.xlsx')
  const workbook = new ExcelJS.Workbook()
  workbook.creator = 'Quick Document'
  workbook.created = new Date()

  const sheets = normalizeSheets(action.sheets, title)
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

    const bullets = slidePlan.bullets && slidePlan.bullets.length > 0 ? slidePlan.bullets : ['Add detail here']
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
  return [
    {
      heading: '概述',
      paragraphs: [`${title} 已由 Quick Document 创建。请在对话中继续说明目标、受众、结构或需要补充的材料。`]
    },
    {
      heading: '下一步',
      bullets: ['补充更明确的文档目标', '选择本地目录或目标文件', '要求 AI 改写、扩展或转换格式']
    }
  ]
}

function normalizeSheets(sheets: SheetPlan[] | undefined, title: string): SheetPlan[] {
  if (sheets && sheets.length > 0) return sheets
  return [
    {
      name: 'Overview',
      columns: ['Item', 'Description', 'Status'],
      rows: [
        { Item: title, Description: 'Generated by Quick Document', Status: 'Draft' },
        { Item: 'Next step', Description: 'Continue editing through chat', Status: 'Open' }
      ]
    }
  ]
}

function normalizeSlides(slides: SlidePlan[] | undefined, title: string): SlidePlan[] {
  if (slides && slides.length > 0) return slides
  return [
    {
      title,
      bullets: ['Quick Document 已创建演示文稿草稿', '继续对话即可调整结构、内容和风格']
    },
    {
      title: '下一步',
      bullets: ['补充演讲对象和目标', '上传参考文档', '要求 AI 扩写、压缩或重排页面']
    }
  ]
}

function deriveColumns(sheet: SheetPlan): string[] {
  if (sheet.columns && sheet.columns.length > 0) return sheet.columns
  const firstRow = sheet.rows?.[0]
  if (firstRow && !Array.isArray(firstRow) && typeof firstRow === 'object') {
    return Object.keys(firstRow)
  }
  return ['Column 1', 'Column 2', 'Column 3']
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
    summary: revised ? `Created revised copy: ${file.name}` : `Created file: ${file.name}`
  }
}

function kindFromAction(actionType: OfficeActionType): OfficeKind {
  if (actionType.endsWith('docx')) return 'word'
  if (actionType.endsWith('xlsx')) return 'excel'
  if (actionType.endsWith('pptx')) return 'powerpoint'
  return 'unknown'
}

function kindFromSkillName(skillName: OfficeAction['skillName']): OfficeKind {
  if (skillName === 'documents') return 'word'
  if (skillName === 'spreadsheets') return 'excel'
  if (skillName === 'presentations') return 'powerpoint'
  return 'unknown'
}

function uniqueFilePath(workspacePath: string, wantedName: string, extension: string): string {
  const baseName = sanitizeFileName(wantedName.replace(extname(wantedName), '')) || 'document'
  let candidate = join(workspacePath, `${baseName}${extension}`)
  let index = 2

  while (existsByStat(candidate)) {
    candidate = join(workspacePath, `${baseName}-${index}${extension}`)
    index += 1
  }

  return candidate
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
