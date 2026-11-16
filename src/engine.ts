import { spawn } from 'node:child_process'
import { constants as fsConstants } from 'node:fs'
import { access, copyFile, mkdir, mkdtemp, realpath, rename, rm, stat } from 'node:fs/promises'
import { homedir, tmpdir } from 'node:os'
import { basename, dirname, extname, isAbsolute, join, relative, resolve, sep } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'
import { HarnessError } from '@deepseek-ai/dsh-llm'
import type { PdfEngine, Settings } from './config.js'
import type { AppResult, CreateResult, EditResult, ExportResult, HelperResponse, ProbeResult, ReadResult, RenderResult } from './types.js'

const MAX_OUTPUT = 32 * 1024 * 1024
const MAX_STDIN = 32 * 1024 * 1024

export class WordError extends HarnessError {
  constructor(code: string, message: string, options?: ErrorOptions) {
    super(message, code, options)
  }
}

export interface LoggerLike {
  warn(message: string): void
}

interface ProcessOptions {
  input?: string
  timeoutMs?: number
  signal?: AbortSignal
  env?: NodeJS.ProcessEnv
  cwd?: string
}

interface ProcessResult {
  stdout: string
  stderr: string
  code: number | null
}

function processError(command: string, error: NodeJS.ErrnoException): WordError {
  if (error.code === 'ENOENT') {
    return new WordError('COMMAND_NOT_FOUND', `Command not found: ${command}. Install it or point the plugin config at it.`, { cause: error })
  }
  return new WordError('PROCESS_FAILED', `Cannot run ${command}: ${error.message}`, { cause: error })
}

export function runProcess(command: string, args: string[], options: ProcessOptions = {}): Promise<ProcessResult> {
  return new Promise((resolvePromise, reject) => {
    let settled = false
    let timer: NodeJS.Timeout | undefined
    let abortListener: (() => void) | undefined
    const finish = (error?: Error, result?: ProcessResult): void => {
      if (settled) return
      settled = true
      if (timer !== undefined) clearTimeout(timer)
      if (abortListener !== undefined) options.signal?.removeEventListener('abort', abortListener)
      if (error) reject(error)
      else resolvePromise(result as ProcessResult)
    }
    if (options.signal?.aborted) {
      finish(new WordError('ABORTED', 'The operation was cancelled.'))
      return
    }
    const child = spawn(command, args, {
      cwd: options.cwd,
      env: options.env ?? process.env,
      stdio: ['pipe', 'pipe', 'pipe'],
    })
    const timeoutMs = options.timeoutMs
    if (timeoutMs !== undefined) {
      timer = setTimeout(() => {
        child.kill('SIGKILL')
        finish(new WordError('TIMEOUT', `${command} did not finish within ${Math.round(timeoutMs / 1000)} seconds.`))
      }, timeoutMs)
    }
    abortListener = () => {
      child.kill('SIGKILL')
      finish(new WordError('ABORTED', 'The operation was cancelled.'))
    }
    options.signal?.addEventListener('abort', abortListener, { once: true })
    let stdout = ''
    let stderr = ''
    child.on('error', (error: NodeJS.ErrnoException) => {
      finish(processError(command, error))
    })
    child.stdin?.on('error', () => undefined)
    child.stdout?.on('data', (chunk: Buffer) => {
      stdout += chunk.toString('utf8')
      if (stdout.length > MAX_OUTPUT) {
        child.kill('SIGKILL')
        finish(new WordError('OUTPUT_TOO_LARGE', `${command} produced more than ${MAX_OUTPUT / (1024 * 1024)} MB of output.`))
      }
    })
    child.stderr?.on('data', (chunk: Buffer) => {
      stderr += chunk.toString('utf8')
      if (stderr.length > MAX_OUTPUT) stderr = stderr.slice(-MAX_OUTPUT)
    })
    child.on('close', (code: number | null) => {
      finish(undefined, { stdout, stderr, code })
    })
    if (options.input !== undefined) {
      if (Buffer.byteLength(options.input, 'utf8') > MAX_STDIN) {
        child.kill('SIGKILL')
        finish(new WordError('INPUT_TOO_LARGE', `The input is larger than ${MAX_STDIN / (1024 * 1024)} MB.`))
        return
      }
      child.stdin?.end(options.input)
    } else {
      child.stdin?.end()
    }
  })
}

function expandHome(input: string): string {
  if (input === '~') return homedir()
  if (input.startsWith('~/')) return join(homedir(), input.slice(2))
  return input
}

function isWithin(root: string, target: string): boolean {
  const rel = relative(root, target)
  return rel === '' || (rel !== '..' && !rel.startsWith(`..${sep}`) && !isAbsolute(rel))
}

async function exists(target: string): Promise<boolean> {
  try {
    await access(target, fsConstants.F_OK)
    return true
  } catch {
    return false
  }
}

async function realpathSafe(target: string): Promise<string> {
  try {
    return await realpath(target)
  } catch {
    return resolve(target)
  }
}

async function resolvedRealTarget(target: string): Promise<string> {
  const missing: string[] = []
  let current = resolve(target)
  while (true) {
    try {
      const real = await realpath(current)
      return join(real, ...missing.reverse())
    } catch {
      const parent = dirname(current)
      if (parent === current) return resolve(target)
      missing.push(basename(current))
      current = parent
    }
  }
}

export type PathMode = 'read' | 'write' | 'create' | 'pdf'

export class WordEngine {
  readonly settings: Settings
  private readonly logger: LoggerLike
  private readonly helperPath: string
  private readonly wordScriptPath: string
  private probeCache?: Promise<ProbeResult>
  private readonly locks = new Map<string, Promise<unknown>>()

  constructor(settings: Settings, logger: LoggerLike) {
    this.settings = settings
    this.logger = logger
    this.helperPath = fileURLToPath(new URL('../helper/word_helper.py', import.meta.url))
    this.wordScriptPath = fileURLToPath(new URL('../helper/export_pdf.applescript', import.meta.url))
  }

  async resolvePath(input: string, mode: PathMode, workspace: string = this.settings.workspace): Promise<string> {
    const raw = String(input ?? '').trim()
    if (raw === '') throw new WordError('INVALID_ARGUMENT', 'path is required.')
    let absolute = expandHome(raw)
    if (!isAbsolute(absolute)) absolute = resolve(workspace, absolute)
    absolute = resolve(absolute)
    let suffix = extname(absolute).toLowerCase()
    if (mode === 'pdf') {
      if (suffix === '') {
        absolute += '.pdf'
        suffix = '.pdf'
      }
      if (suffix !== '.pdf') throw new WordError('UNSUPPORTED_FORMAT', `PDF output must end in .pdf, got ${suffix}.`)
    } else {
      if (suffix === '') {
        absolute += '.docx'
        suffix = '.docx'
      }
      if (suffix === '.doc') {
        throw new WordError('UNSUPPORTED_FORMAT', 'Legacy .doc files are not supported. Save the file as .docx in Word first.')
      }
      if (suffix !== '.docx') throw new WordError('UNSUPPORTED_FORMAT', `Only .docx files are supported, got ${suffix}.`)
    }
    await this.assertAllowed(absolute, workspace)
    if (mode === 'read' || mode === 'write') {
      let info
      try {
        info = await stat(absolute)
      } catch {
        throw new WordError('FILE_NOT_FOUND', `File not found: ${absolute}`)
      }
      if (!info.isFile()) throw new WordError('NOT_A_FILE', `Not a file: ${absolute}`)
    }
    return absolute
  }

  async resolveAssetPath(input: string, workspace: string = this.settings.workspace): Promise<string> {
    const raw = String(input ?? '').trim()
    if (raw === '') throw new WordError('INVALID_ARGUMENT', 'path is required.')
    let absolute = expandHome(raw)
    if (!isAbsolute(absolute)) absolute = resolve(workspace, absolute)
    absolute = resolve(absolute)
    await this.assertAllowed(absolute, workspace)
    let info
    try {
      info = await stat(absolute)
    } catch {
      throw new WordError('IMAGE_NOT_FOUND', `Image not found: ${absolute}`)
    }
    if (!info.isFile()) throw new WordError('NOT_A_FILE', `Not a file: ${absolute}`)
    return absolute
  }

  private async assertAllowed(absolute: string, workspace: string): Promise<void> {
    if (this.settings.allowOutsideWorkspace) return
    const workspaceReal = await realpathSafe(workspace)
    const targetReal = await resolvedRealTarget(absolute)
    if (!isWithin(workspaceReal, targetReal)) {
      throw new WordError(
        'PATH_OUTSIDE_WORKSPACE',
        `Path ${absolute} is outside the workspace ${workspace}. Set allowOutsideWorkspace: true to allow it.`,
      )
    }
  }

  async probe(signal?: AbortSignal): Promise<ProbeResult> {
    if (!this.probeCache) {
      this.probeCache = this.runHelper<ProbeResult>({ command: 'probe' }, signal).catch((error: unknown) => {
        this.probeCache = undefined
        throw error
      })
    }
    return this.probeCache
  }

  async ensureReady(signal?: AbortSignal): Promise<ProbeResult> {
    const probe = await this.probe(signal)
    if (!probe.python_docx) {
      throw new WordError(
        'PYTHON_DOCX_MISSING',
        `python-docx is not installed for ${this.settings.python}. Install it with: ${this.settings.python} -m pip install --user python-docx`,
      )
    }
    return probe
  }

  async ensureRenderReady(signal?: AbortSignal): Promise<ProbeResult> {
    const probe = await this.ensureReady(signal)
    if (!probe.pymupdf) {
      throw new WordError(
        'PYMUPDF_MISSING',
        `PyMuPDF is not installed for ${this.settings.python}. Install it with: ${this.settings.python} -m pip install --user pymupdf`,
      )
    }
    return probe
  }

  async read(path: string, options: { includeHeaderFooter?: boolean; maxChars?: number } = {}, signal?: AbortSignal): Promise<ReadResult> {
    await this.ensureReady(signal)
    return this.withLock(path, () => this.runHelper<ReadResult>({
      command: 'read',
      path,
      max_chars: options.maxChars ?? this.settings.maxReadChars,
      include_header_footer: options.includeHeaderFooter ?? true,
    }, signal))
  }

  async create(path: string, payload: Record<string, unknown>, signal?: AbortSignal): Promise<CreateResult> {
    await this.ensureReady(signal)
    return this.withLock(path, () => this.runHelper<CreateResult>({
      ...payload,
      command: 'create',
      path,
      backup: payload.backup ?? this.settings.backup,
    }, signal))
  }

  async edit(path: string, payload: Record<string, unknown>, signal?: AbortSignal): Promise<EditResult> {
    await this.ensureReady(signal)
    return this.withLock(path, () => this.runHelper<EditResult>({
      ...payload,
      command: 'edit',
      path,
      backup: payload.backup ?? this.settings.backup,
    }, signal))
  }

  private async stagingDirectory(engine: 'libreoffice' | 'word'): Promise<string> {
    if (engine === 'word') {
      const container = join(homedir(), 'Library', 'Containers', 'com.microsoft.Word', 'Data', 'tmp')
      try {
        await access(container, fsConstants.W_OK)
        return await mkdtemp(join(container, 'dsh-word-'))
      } catch {
        return mkdtemp(join(tmpdir(), 'dsh-word-pdf-'))
      }
    }
    return mkdtemp(join(tmpdir(), 'dsh-word-pdf-'))
  }

  async exportPdf(input: string, output: string, engine: PdfEngine, overwrite: boolean, signal?: AbortSignal): Promise<ExportResult> {
    await mkdir(dirname(output), { recursive: true })
    if (!overwrite && await exists(output)) {
      throw new WordError('FILE_EXISTS', `${output} already exists. Pass overwrite: true to replace it.`)
    }
    const resolved = await this.resolveEngine(engine)
    const temp = await this.stagingDirectory(resolved)
    try {
      const source = join(temp, 'input.docx')
      await copyFile(input, source)
      const pdf = resolved === 'libreoffice' ? join(temp, 'input.pdf') : join(temp, 'output.pdf')
      if (resolved === 'libreoffice') {
        const result = await runProcess(this.settings.sofficePath, [
          `-env:UserInstallation=${pathToFileURL(join(temp, 'profile')).href}`,
          '--headless',
          '--norestore',
          '--convert-to',
          'pdf:writer_pdf_Export',
          '--outdir',
          temp,
          source,
        ], {
          timeoutMs: Math.max(this.settings.timeoutMs, 60000),
          signal,
          env: { ...process.env, HOME: process.env.HOME ?? homedir() },
        })
        if (!await exists(pdf)) {
          const detail = (result.stderr || result.stdout).trim().slice(-2000)
          throw new WordError('PDF_EXPORT_FAILED', `LibreOffice did not produce a PDF.${detail ? ` ${detail}` : ''}`)
        }
      } else {
        const result = await runProcess('/usr/bin/osascript', [this.wordScriptPath, source, pdf], {
          timeoutMs: Math.max(this.settings.timeoutMs, 180000),
          signal,
        })
        if (!await exists(pdf)) {
          const detail = (result.stderr || result.stdout).trim().slice(-2000)
          throw new WordError(
            'PDF_EXPORT_FAILED',
            `Microsoft Word did not produce a PDF. Check that Word is allowed to be automated in System Settings > Privacy & Security > Automation.${detail ? ` ${detail}` : ''}`,
          )
        }
      }
      const temporaryOutput = `${output}.${process.pid}.${Date.now()}.tmp`
      await copyFile(pdf, temporaryOutput)
      try {
        await rename(temporaryOutput, output)
      } catch (error) {
        await rm(temporaryOutput, { force: true })
        throw error
      }
      const info = await stat(output)
      return { input, output, engine: resolved, bytes: info.size }
    } finally {
      await rm(temp, { recursive: true, force: true })
    }
  }

  async renderPages(
    document: string,
    options: { pages?: string; width?: number; engine?: PdfEngine; maxPages?: number } = {},
    signal?: AbortSignal,
  ): Promise<RenderResult & { cleanup: () => Promise<void> }> {
    await this.ensureRenderReady(signal)
    const temp = await mkdtemp(join(tmpdir(), 'dsh-word-render-'))
    const pdf = join(temp, 'document.pdf')
    const pagesDir = join(temp, 'pages')
    try {
      await this.exportPdf(document, pdf, options.engine ?? 'auto', true, signal)
      const result = await this.runHelper<RenderResult>({
        command: 'render_pdf',
        pdf,
        out_dir: pagesDir,
        pages: options.pages ?? '1',
        width: options.width ?? this.settings.renderWidth,
        max_pages: options.maxPages ?? this.settings.maxRenderPages,
      }, signal)
      return {
        ...result,
        cleanup: async () => {
          await rm(temp, { recursive: true, force: true })
        },
      }
    } catch (error) {
      await rm(temp, { recursive: true, force: true })
      throw error
    }
  }

  async appStatus(document?: string, signal?: AbortSignal): Promise<AppResult> {
    let wordVersion: string | null = null
    try {
      wordVersion = (await this.probe(signal)).word_version
    } catch (error) {
      if (error instanceof WordError && error.code === 'ABORTED') throw error
    }
    const installed = await exists(this.settings.wordPath)
    const open = document === undefined ? null : await this.isOpenInWord(document, signal)
    return {
      installed,
      version: installed ? wordVersion : null,
      appPath: this.settings.wordPath,
      document: document ?? null,
      openInWord: open,
      action: 'status',
      message: installed
        ? 'Use word_open to bring a document to the front. word_edit works on the file directly and does not require Word to be open.'
        : 'Microsoft Word is not installed at the configured path. File-level word_* tools still work.',
    }
  }

  async openInWord(document: string, signal?: AbortSignal): Promise<AppResult> {
    const installed = await exists(this.settings.wordPath)
    if (!installed) {
      throw new WordError('WORD_NOT_INSTALLED', `Microsoft Word.app not found at ${this.settings.wordPath}.`)
    }
    await runProcess('/usr/bin/open', ['-a', this.settings.wordPath, document], { timeoutMs: 15000, signal })
    let version: string | null = null
    try {
      version = (await this.probe(signal)).word_version
    } catch (error) {
      if (error instanceof WordError && error.code === 'ABORTED') throw error
    }
    return {
      installed: true,
      version,
      appPath: this.settings.wordPath,
      document,
      openInWord: true,
      action: 'open',
      message: `Opened ${document} in Microsoft Word.`,
    }
  }

  async isOpenInWord(document: string, signal?: AbortSignal): Promise<boolean | null> {
    try {
      const result = await runProcess('/usr/sbin/lsof', ['-Fpc', '--', document], { timeoutMs: 4000, signal })
      if (result.code === 1) return false
      if (result.code !== 0) return null
      return /cMicrosoft Word|Microso.*Word/i.test(result.stdout)
    } catch (error) {
      if (error instanceof WordError && error.code === 'COMMAND_NOT_FOUND') return null
      if (error instanceof WordError && error.code === 'ABORTED') throw error
      return null
    }
  }

  private async resolveEngine(requested: PdfEngine): Promise<'libreoffice' | 'word'> {
    if (requested === 'none') {
      throw new WordError('PDF_ENGINE_DISABLED', 'PDF export is disabled by the plugin configuration (pdfEngine: none).')
    }
    if (requested === 'libreoffice') {
      if (!await exists(this.settings.sofficePath)) {
        throw new WordError('PDF_ENGINE_UNAVAILABLE', `LibreOffice not found at ${this.settings.sofficePath}.`)
      }
      return 'libreoffice'
    }
    if (requested === 'word') {
      if (!await exists(this.settings.wordPath)) {
        throw new WordError('PDF_ENGINE_UNAVAILABLE', `Microsoft Word.app not found at ${this.settings.wordPath}.`)
      }
      return 'word'
    }
    if (await exists(this.settings.sofficePath)) return 'libreoffice'
    if (await exists(this.settings.wordPath)) return 'word'
    throw new WordError(
      'PDF_ENGINE_UNAVAILABLE',
      'No PDF engine is available. Install LibreOffice, install Microsoft Word, or set pdfEngine to libreoffice or word.',
    )
  }

  private async runHelper<T>(payload: Record<string, unknown>, signal?: AbortSignal): Promise<T> {
    const result = await runProcess(this.settings.python, [this.helperPath], {
      input: JSON.stringify(payload),
      timeoutMs: this.settings.timeoutMs,
      signal,
      env: {
        ...process.env,
        PYTHONIOENCODING: 'utf-8',
        PYTHONUNBUFFERED: '1',
      },
    })
    const output = result.stdout.trim()
    if (output === '') {
      const detail = result.stderr.trim().slice(-2000)
      throw new WordError('HELPER_CRASH', `The Word helper produced no output.${detail ? ` ${detail}` : ''}`)
    }
    let parsed: HelperResponse<T>
    try {
      parsed = JSON.parse(output) as HelperResponse<T>
    } catch {
      const detail = result.stderr.trim().slice(-2000)
      throw new WordError('HELPER_CRASH', `The Word helper returned invalid JSON.${detail ? ` ${detail}` : ''}`)
    }
    if (parsed.ok !== true) {
      throw new WordError(parsed.error?.code || 'HELPER_FAILED', parsed.error?.message || 'The Word helper failed.')
    }
    return parsed.result
  }

  private withLock<T>(key: string, work: () => Promise<T>): Promise<T> {
    const previous = this.locks.get(key) ?? Promise.resolve()
    const current = previous.catch(() => undefined).then(work)
    const tracked = current.finally(() => {
      if (this.locks.get(key) === tracked) this.locks.delete(key)
    }).catch(() => undefined)
    this.locks.set(key, tracked)
    return current
  }
}
