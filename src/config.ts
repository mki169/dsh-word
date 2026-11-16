import { homedir } from 'node:os'
import { resolve } from 'node:path'
import z from '@deepseek-ai/schemastery'

export const ENTRY_ID = 'office-word'

export const WORD_APP_PATH = '/Applications/Microsoft Word.app'
export const LIBREOFFICE_PATH = '/Applications/LibreOffice.app/Contents/MacOS/soffice'

export type PdfEngine = 'auto' | 'word' | 'libreoffice' | 'none'

export interface Config {
  python?: string
  workspace?: string
  allowOutsideWorkspace?: boolean
  backup?: boolean
  pdfEngine?: PdfEngine
  sofficePath?: string
  wordPath?: string
  timeoutMs?: number
  maxReadChars?: number
  renderWidth?: number
  maxRenderPages?: number
}

export const Config: z<Config> = z.object({
  python: z.string().default('python3').volatile().i18n({
    'zh-CN': { $description: '运行文档辅助脚本的 Python 解释器；需要 python-docx（导出/渲染还需要 PyMuPDF）' },
    'en-US': { $description: 'Python interpreter for the document helper; needs python-docx (PyMuPDF for export/render)' },
  }),
  workspace: z.string().default('').volatile().i18n({
    'zh-CN': { $description: '文档工作目录；留空则使用 Harness 会话的当前目录。相对路径按这个目录解析' },
    'en-US': { $description: 'Document workspace; empty uses the Harness session cwd. Relative paths resolve against it' },
  }),
  allowOutsideWorkspace: z.boolean().default(false).volatile().i18n({
    'zh-CN': { $description: '允许读写工作目录之外的文件；默认关闭' },
    'en-US': { $description: 'Allow reading and writing files outside the workspace; off by default' },
  }),
  backup: z.boolean().default(true).volatile().i18n({
    'zh-CN': { $description: '编辑前把原文件复制为 .bak；关闭后只做原子替换' },
    'en-US': { $description: 'Copy the original to .bak before editing; off keeps only the atomic replace' },
  }),
  pdfEngine: z.string().default('auto').volatile().i18n({
    'zh-CN': { $description: 'PDF 导出方式：auto（优先 LibreOffice，其次 Word）| libreoffice | word | none（禁用导出）' },
    'en-US': { $description: 'PDF export engine: auto (LibreOffice first, then Word) | libreoffice | word | none (disable export)' },
  }),
  sofficePath: z.string().default(LIBREOFFICE_PATH).volatile().i18n({
    'zh-CN': { $description: 'LibreOffice soffice 可执行文件路径' },
    'en-US': { $description: 'Path to the LibreOffice soffice executable' },
  }),
  wordPath: z.string().default(WORD_APP_PATH).volatile().i18n({
    'zh-CN': { $description: 'Microsoft Word.app 路径' },
    'en-US': { $description: 'Path to Microsoft Word.app' },
  }),
  timeoutMs: z.natural().min(5000).max(300000).default(90000).volatile().i18n({
    'zh-CN': { $description: '单次文档操作的超时时间（毫秒）' },
    'en-US': { $description: 'Timeout for one document operation in milliseconds' },
  }),
  maxReadChars: z.natural().min(2000).max(200000).default(40000).volatile().i18n({
    'zh-CN': { $description: 'word_read 返回的 Markdown 最大字符数' },
    'en-US': { $description: 'Maximum characters of Markdown returned by word_read' },
  }),
  renderWidth: z.natural().min(400).max(2400).default(1000).volatile().i18n({
    'zh-CN': { $description: 'word_render 给模型查看的页面图片宽度（像素）' },
    'en-US': { $description: 'Page image width in pixels for word_render' },
  }),
  maxRenderPages: z.natural().min(1).max(40).default(8).volatile().i18n({
    'zh-CN': { $description: 'word_render 一次最多渲染的页数' },
    'en-US': { $description: 'Maximum pages rendered by one word_render call' },
  }),
}) as unknown as z<Config>

export interface Settings {
  python: string
  workspace: string
  workspaceConfigured: boolean
  allowOutsideWorkspace: boolean
  backup: boolean
  pdfEngine: PdfEngine
  sofficePath: string
  wordPath: string
  timeoutMs: number
  maxReadChars: number
  renderWidth: number
  maxRenderPages: number
}

function valueOf(value: unknown): unknown {
  if (value !== null && typeof value === 'object' && typeof (value as { get?: unknown }).get === 'function') {
    return (value as { get: () => unknown }).get()
  }
  return value
}

function expandHome(value: string): string {
  if (value === '~') return homedir()
  if (value.startsWith('~/')) return resolve(homedir(), value.slice(2))
  return value
}

function pathValue(value: unknown, fallback: string): string {
  return typeof value === 'string' && value.trim() !== '' ? resolve(expandHome(value.trim())) : fallback
}

export function resolveConfig(raw: Config = {}): Settings {
  const values: Record<string, unknown> = {}
  for (const [key, value] of Object.entries(raw)) values[key] = valueOf(value)
  const engine = values.pdfEngine
  const workspaceConfigured = typeof values.workspace === 'string' && values.workspace.trim() !== ''
  const workspace = workspaceConfigured
    ? pathValue(values.workspace, process.cwd())
    : process.cwd()
  const timeout = values.timeoutMs
  const maxRead = values.maxReadChars
  const width = values.renderWidth
  const pages = values.maxRenderPages
  return {
    python: typeof values.python === 'string' && values.python.trim() !== '' ? values.python.trim() : 'python3',
    workspace,
    workspaceConfigured,
    allowOutsideWorkspace: typeof values.allowOutsideWorkspace === 'boolean' ? values.allowOutsideWorkspace : false,
    backup: typeof values.backup === 'boolean' ? values.backup : true,
    pdfEngine: engine === 'word' || engine === 'libreoffice' || engine === 'none' ? engine : 'auto',
    sofficePath: pathValue(values.sofficePath, LIBREOFFICE_PATH),
    wordPath: pathValue(values.wordPath, WORD_APP_PATH),
    timeoutMs: typeof timeout === 'number' && Number.isFinite(timeout) ? Math.min(300000, Math.max(5000, Math.round(timeout))) : 90000,
    maxReadChars: typeof maxRead === 'number' && Number.isFinite(maxRead) ? Math.min(200000, Math.max(2000, Math.round(maxRead))) : 40000,
    renderWidth: typeof width === 'number' && Number.isFinite(width) ? Math.min(2400, Math.max(400, Math.round(width))) : 1000,
    maxRenderPages: typeof pages === 'number' && Number.isFinite(pages) ? Math.min(40, Math.max(1, Math.round(pages))) : 8,
  }
}
