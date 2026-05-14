import { app } from 'electron'
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'

interface SkillEntry {
  name: string
  path: string
  summary: string
}

const OFFICE_SKILL_RULES = `
Quick Document ships with embedded Office skills. Use this compact routing brief when planning work:

Documents / Word:
- For new DOCX creation, choose an appropriate document archetype and stable design preset before drafting.
- Use real document structures: Word styles, headings, numbering, explicit table geometry, comments, tracked changes, and OOXML patches when needed.
- Prefer minimal local edits for existing DOCX files. Preserve structure unless a surgical restructure is required.
- Treat rendering and visual inspection as the quality gate for polished DOCX output when the local executor supports it.
- Avoid fake bullets, fake headings, default table geometry, cramped tables, and unverified layout.

Spreadsheets / Excel:
- Build workbooks as analyst-grade artifacts: clear sheet architecture, formulas for derived values, tables, filters, freeze panes, validation, and charts when the prompt implies summary analysis.
- For edits, preserve formulas and local formatting patterns; extend dependent ranges, tables, charts, and conditional formats when adding rows or columns.
- Verify key ranges, formula errors, and visual readability before final export when the local executor supports it.

Presentations / PowerPoint:
- Route both .ppt and .pptx files to the presentations skill.
- Build a claim-led deck, not a generic slide dump. Every slide needs a clear point, proof object, and readable hierarchy.
- Pick a deck profile before drafting: strategy, finance, product, GTM, engineering, consumer, appendix-heavy, or targeted edit.
- Prefer editable native PPTX content. Use rendered preview and contact-sheet critique as the quality gate when the local executor supports it.

Cross-artifact behavior:
- If the user asks for Word, Excel, and PPT together, plan separate actions with aligned titles and consistent source assumptions.
- If the user attaches a source document, summarize its role and use revise_* actions for same-format transformations.
- Keep filenames safe for macOS and Windows. Generate final Office files in the configured local workspace.
`.trim()

export function getSkillsRoot(): string {
  if (app.isPackaged) {
    return join(process.resourcesPath, 'resources', 'skills')
  }

  return join(app.getAppPath(), 'resources', 'skills')
}

export function getEmbeddedOfficeSkillBrief(): string {
  const skills = getEmbeddedSkills()
  const installed = skills
    .map((skill) => `- ${skill.name}: ${skill.path}\n  ${skill.summary}`)
    .join('\n')

  return `${OFFICE_SKILL_RULES}\n\nInstalled skill packs:\n${installed || '- none found'}`
}

export function getEmbeddedSkills(): SkillEntry[] {
  const root = getSkillsRoot()
  if (!existsSync(root)) return []

  return readdirSync(root, { withFileTypes: true })
    .filter((entry) => entry.isDirectory())
    .map((entry) => {
      const skillPath = join(root, entry.name)
      return {
        name: entry.name,
        path: skillPath,
        summary: readSkillSummary(join(skillPath, 'SKILL.md'))
      }
    })
}

function readSkillSummary(skillFile: string): string {
  if (!existsSync(skillFile)) return 'No SKILL.md found.'

  const content = readFileSync(skillFile, 'utf8')
  const frontmatterDescription = content.match(/description:\s*"?([^"\n]+)"?/i)?.[1]
  const firstParagraph = content
    .replace(/^---[\s\S]*?---/, '')
    .split('\n')
    .map((line) => line.trim())
    .find((line) => line && !line.startsWith('#'))

  return (frontmatterDescription || firstParagraph || 'Embedded Office skill.').slice(0, 360)
}
