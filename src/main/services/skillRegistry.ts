import { app } from 'electron'
import { existsSync, readdirSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import type { WorkspaceFile, WorkspaceSnapshot } from '@shared/types'

interface SkillEntry {
  name: string
  path: string
  summary: string
}

const OFFICE_SKILL_RULES = `
Quick Document ships with embedded Office skills. Treat these as capability notes, not intent rules. The user's wording decides what to do.

Documents / Word:
- Use the Documents skill for .docx understanding and editing decisions.
- Prefer preserving structure for edits unless the user asks for larger rewriting or redesign.

Spreadsheets / Excel:
- Use the Spreadsheets skill for .xlsx understanding and editing decisions.
- Preserve formulas, formatting, sheets, and layout unless the user asks to change them.

Presentations / PowerPoint:
- Use the Presentations skill for both .ppt and .pptx understanding and editing decisions.
- Preserve slide structure for targeted edits unless the user asks to redesign or rebuild.

Cross-artifact behavior:
- If the user asks for multiple document types, plan separate actions in the user's requested order.
- Keep filenames safe for macOS and Windows when creating or copying files.
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

export function getRelevantOfficeSkillContext(
  prompt: string,
  targetFiles: WorkspaceFile[],
  workspaceSnapshot?: WorkspaceSnapshot
): string {
  const skills = inferRelevantSkills(prompt, targetFiles, workspaceSnapshot)
  const sections = skills
    .map((skillName) => {
      const content = readSkillContext(skillName)
      return content ? `# ${skillName} skill\n${content}` : ''
    })
    .filter(Boolean)

  return sections.join('\n\n---\n\n')
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

function inferRelevantSkills(
  prompt: string,
  targetFiles: WorkspaceFile[],
  workspaceSnapshot?: WorkspaceSnapshot
): Array<'documents' | 'spreadsheets' | 'presentations'> {
  const skills = new Set<'documents' | 'spreadsheets' | 'presentations'>()
  const normalizedPrompt = prompt.toLowerCase()

  for (const file of targetFiles) {
    const skillName = skillNameForKind(file.kind)
    if (skillName) skills.add(skillName)
  }

  if (workspaceSnapshot) {
    for (const file of workspaceSnapshot.files) {
      const name = file.name.toLowerCase()
      const stem = name.replace(/\.(docx|xlsx|ppt|pptx)$/i, '')
      if (!normalizedPrompt.includes(name) && !normalizedPrompt.includes(stem)) continue
      const skillName = skillNameForKind(file.kind)
      if (skillName) skills.add(skillName)
    }
  }

  if (/word|docx|文档|报告|合同|方案/.test(normalizedPrompt)) skills.add('documents')
  if (/excel|xlsx|表格|台账|清单|预算|数据|单元格/.test(normalizedPrompt)) skills.add('spreadsheets')
  if (/ppt|pptx|powerpoint|演示|幻灯|汇报/.test(normalizedPrompt)) skills.add('presentations')

  return Array.from(skills).slice(0, 4)
}

function skillNameForKind(kind: WorkspaceFile['kind']): 'documents' | 'spreadsheets' | 'presentations' | null {
  if (kind === 'word') return 'documents'
  if (kind === 'excel') return 'spreadsheets'
  if (kind === 'powerpoint') return 'presentations'
  return null
}

function readSkillContext(skillName: 'documents' | 'spreadsheets' | 'presentations'): string {
  const skillPath = join(getSkillsRoot(), skillName, 'SKILL.md')
  if (!existsSync(skillPath)) return ''
  const content = readFileSync(skillPath, 'utf8')
  return compactSkillContext(content).slice(0, 14000)
}

function compactSkillContext(content: string): string {
  return content
    .replace(/^---[\s\S]*?---/, '')
    .split('\n')
    .filter((line) => {
      const trimmed = line.trim()
      if (!trimmed) return true
      if (/^!\[/.test(trimmed)) return false
      if (/^<img/i.test(trimmed)) return false
      return true
    })
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim()
}
