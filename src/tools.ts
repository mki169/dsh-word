import { readFile } from 'node:fs/promises'
import { basename } from 'node:path'
import type { ImageAttachmentRef } from '@deepseek-ai/dsh-attachment'
import { defineTool } from '@deepseek-ai/dsh-tools'
import type { PdfEngine } from './config.js'
import type { WordEngine } from './engine.js'
import { WordError } from './engine.js'
import { formatApp, formatCreate, formatEdit, formatExport, formatRead, formatRender } from './format.js'
import type { AttachmentJson, EditResult, RenderValue } from './types.js'

type ToolDefinition = ReturnType<typeof defineTool>

export interface ToolHost {
  engine: WordEngine
  saveImage(data: Buffer, name: string): Promise<ImageAttachmentRef>
}

const IMAGE = {
  type: 'object',
  additionalProperties: false,
  properties: {
    attachmentId: { type: 'string', required: true },
    mediaType: { type: 'string', required: true },
    bytes: { type: 'integer', required: true },
    width: { type: 'integer', required: true },
    height: { type: 'integer', required: true },
    name: { type: 'string' },
    originalDimensions: {
      type: 'object',
      additionalProperties: false,
      properties: {
        width: { type: 'integer', required: true },
        height: { type: 'integer', required: true },
      },
    },
  },
} as const

const TEXT_OUTPUT = {
  schema: {
    type: 'object',
    additionalProperties: false,
    properties: {
      text: { type: 'string', required: true },
    },
  },
  render: (_args: unknown, value: { text: string }) => [{ type: 'text' as const, text: value.text }],
} as const

const RENDER_OUTPUT = {
  schema: {
    type: 'object',
    additionalProperties: false,
    properties: {
      text: { type: 'string', required: true },
      image: IMAGE,
      more: { type: 'array', items: IMAGE },
    },
  },
  render: (_args: unknown, value: RenderValue) => [
    { type: 'text' as const, text: value.text },
    ...[value.image, ...(value.more ?? [])]
      .filter((image): image is AttachmentJson => image !== undefined)
      .map(image => ({ type: 'image' as const, attachment: image as unknown as ImageAttachmentRef })),
  ],
} as const

async function resolveBlockPaths(engine: WordEngine, blocks: Array<Record<string, unknown>>, workspace: string): Promise<Array<Record<string, unknown>>> {
  const resolved: Array<Record<string, unknown>> = []
  for (const block of blocks) {
    if (block !== null && typeof block === 'object' && block.type === 'image' && typeof block.path === 'string') {
      resolved.push({ ...block, path: await engine.resolveAssetPath(block.path, workspace) })
    } else {
      resolved.push(block)
    }
  }
  return resolved
}

async function resolveOpPaths(engine: WordEngine, ops: Array<Record<string, unknown>>, workspace: string): Promise<Array<Record<string, unknown>>> {
  const resolved: Array<Record<string, unknown>> = []
  for (const op of ops) {
    if (op !== null && typeof op === 'object' && op.op === 'insert_image' && typeof op.path === 'string') {
      resolved.push({ ...op, path: await engine.resolveAssetPath(op.path, workspace) })
    } else {
      resolved.push(op)
    }
  }
  return resolved
}

function sessionWorkspace(engine: WordEngine, exec: unknown): string {
  if (engine.settings.workspaceConfigured) return engine.settings.workspace
  const cwd = (exec as { agent?: { session?: { meta?: { cwd?: string } } } }).agent?.session?.meta?.cwd
  return typeof cwd === 'string' && cwd.length > 0 ? cwd : engine.settings.workspace
}

function toJson(ref: ImageAttachmentRef): AttachmentJson {
  return {
    attachmentId: String(ref.attachmentId),
    mediaType: ref.mediaType,
    bytes: ref.bytes,
    width: ref.width,
    height: ref.height,
    ...(ref.name === undefined ? {} : { name: ref.name }),
    ...(ref.originalDimensions === undefined
      ? {}
      : { originalDimensions: { width: ref.originalDimensions.width, height: ref.originalDimensions.height } }),
  }
}

const PATH = {
  type: 'string',
  required: true,
  description: 'Path to a .docx file. Relative paths resolve against the workspace; absolute paths are allowed only inside it unless allowOutsideWorkspace is on.',
} as const

const BLOCK = {
  type: 'object',
  additionalProperties: true,
  properties: {
    type: {
      type: 'string',
      required: true,
      enum: ['paragraph', 'heading', 'bullets', 'list', 'quote', 'code', 'table', 'image', 'page_break', 'new_page', 'toc'],
      description: 'paragraph (default), heading, bullets, list, quote, code, table, image, page_break/new_page, or toc.',
    },
    text: { type: 'string', description: 'Text for paragraph, heading, quote or code. **bold**, *italic* and `code` are recognised. Newlines split paragraphs; two or more list lines are detected automatically.' },
    level: { type: 'integer', description: 'Heading level 1-6, or list nesting level 1-9.' },
    style: { type: 'string', description: 'A Word style name, e.g. Title, Heading 1, List Bullet, Quote.' },
    items: { type: 'array', description: 'Strings, or objects with text/level/ordered/format, for bullets/list. Strings may include indentation and markers such as "1.", "a.", "(i)", "(A)" or "-".' },
    ordered: { type: 'boolean', description: 'list: true = numbered, false = bullets.' },
    numbering: { type: 'string', enum: ['legal', 'decimal', 'bullet'], description: 'Numbering scheme: legal (1., (a), (i), (1), (A)...), decimal, or bullet.' },
    rows: { type: 'array', description: 'Table cells: [[cell, ...], ...].' },
    data: { type: 'array', description: 'Alias for rows.' },
    caption: { type: 'string', description: 'Caption. Tables get it above, figures below.' },
    header: { type: 'boolean', description: 'Table: bold and shade the first row (default true).' },
    align: { type: 'string', enum: ['left', 'center', 'right', 'justify'], description: 'Paragraph or table alignment.' },
    path: { type: 'string', description: 'Image file path for image blocks.' },
    width: { type: 'string', description: 'Image width, e.g. "12cm" or "360pt".' },
    title: { type: 'string', description: 'TOC title.' },
    runs: { type: 'array', description: 'Explicit runs: [{text, bold?, italic?, size?, color?, font?}].' },
    size: { type: 'string', description: 'Font size, e.g. "12pt".' },
    bold: { type: 'boolean' },
    italic: { type: 'boolean' },
    underline: { type: 'boolean' },
    color: { type: 'string', description: '#RRGGBB.' },
    font: { type: 'string' },
    firstLineIndent: { type: 'string', description: 'First-line indent, e.g. "2ch" or "0.74cm".' },
    leftIndent: { type: 'string', description: 'Left indent, e.g. "1.27cm".' },
    rightIndent: { type: 'string', description: 'Right indent, e.g. "1.27cm".' },
    indentChars: { type: 'string', description: 'First-line indent in characters, e.g. "2ch".' },
    lineSpacing: { type: 'string', description: '1.5, "20pt" or "at least 20pt".' },
    spaceBefore: { type: 'string' },
    spaceAfter: { type: 'string' },
    keepWithNext: { type: 'boolean', description: 'Keep this paragraph on the same page as the next one.' },
    keepTogether: { type: 'boolean', description: 'Keep all lines of this paragraph on one page.' },
    widowControl: { type: 'boolean', description: 'Avoid single-line widows and orphans.' },
    pageBreakBefore: { type: 'boolean', description: 'Start this paragraph on a new page.' },
  },
} as const

const PAGE = {
  type: 'object',
  additionalProperties: false,
  properties: {
    paper: { type: 'string', enum: ['A4', 'A3', 'B5', 'Letter'] },
    orientation: { type: 'string', enum: ['portrait', 'landscape'] },
    top: { type: 'number', description: 'Top margin in cm.' },
    bottom: { type: 'number', description: 'Bottom margin in cm.' },
    left: { type: 'number', description: 'Left margin in cm.' },
    right: { type: 'number', description: 'Right margin in cm.' },
  },
} as const

const OP = {
  type: 'object',
  additionalProperties: true,
  properties: {
    op: {
      type: 'string',
      required: true,
      enum: [
        'insert_paragraphs',
        'insert_list',
        'set_text',
        'replace_text',
        'delete_range',
        'insert_table',
        'set_cell',
        'insert_image',
        'format_paragraph',
        'style_format',
        'page_setup',
        'header',
        'footer',
        'page_numbers',
        'page_break',
        'new_page',
        'add_toc',
        'update_fields',
      ],
    },
    para: { type: 'integer', description: 'Paragraph number from word_read, e.g. 3 for [p3].' },
    to: { type: 'integer', description: 'Last paragraph for delete_range.' },
    expect: { type: 'string', description: 'First words of the paragraph, checked before editing. Required for set_text and delete_range; "" means empty.' },
    after: { type: 'integer', description: 'Insert after this paragraph number; omit to append at the end.' },
    where: { type: 'string', enum: ['after', 'before', 'start', 'end'], description: 'Placement relative to after, or start/end of the document.' },
    text: { type: 'string', description: 'New text, cell text, header/footer text, or TOC title.' },
    find: { type: 'string', description: 'Literal text for replace_text.' },
    replace: { type: 'string', description: 'Replacement text for replace_text.' },
    all: { type: 'boolean', description: 'replace_text: replace every occurrence (default true).' },
    items: { type: 'array', description: 'insert_paragraphs: strings or objects with text/style/format fields. insert_list: strings or objects with text/level/ordered; strings may include markers and indentation.' },
    ordered: { type: 'boolean', description: 'insert_list: true = numbered, false = bullets.' },
    numbering: { type: 'string', enum: ['legal', 'decimal', 'bullet'], description: 'insert_list numbering scheme.' },
    rows: { type: 'array', description: 'insert_table: [[cell, ...], ...].' },
    data: { type: 'array', description: 'Alias for rows.' },
    caption: { type: 'string', description: 'Table or figure caption.' },
    header: { type: 'boolean' },
    table: { type: 'integer', description: 'Table number from word_read, e.g. 0 for [t0].' },
    row: { type: 'integer' },
    col: { type: 'integer' },
    path: { type: 'string', description: 'Image path for insert_image.' },
    width: { type: 'string', description: 'Image width, e.g. "12cm".' },
    style: { type: 'string', description: 'Style name for style_format or insert_list.' },
    level: { type: 'integer', description: 'List nesting level 1-9 in an insert_list item.' },
    paper: { type: 'string', enum: ['A4', 'A3', 'B5', 'Letter'] },
    orientation: { type: 'string', enum: ['portrait', 'landscape'] },
    top: { type: 'number' },
    bottom: { type: 'number' },
    left: { type: 'number' },
    right: { type: 'number' },
    align: { type: 'string', enum: ['left', 'center', 'right', 'justify'] },
    size: { type: 'string' },
    bold: { type: 'boolean' },
    italic: { type: 'boolean' },
    underline: { type: 'boolean' },
    color: { type: 'string' },
    font: { type: 'string' },
    firstLineIndent: { type: 'string' },
    leftIndent: { type: 'string' },
    rightIndent: { type: 'string' },
    indentChars: { type: 'string' },
    lineSpacing: { type: 'string' },
    spaceBefore: { type: 'string' },
    spaceAfter: { type: 'string' },
    keepWithNext: { type: 'boolean' },
    keepTogether: { type: 'boolean' },
    widowControl: { type: 'boolean' },
    pageBreakBefore: { type: 'boolean' },
    start: { type: 'integer', description: 'page_numbers: first page number.' },
    title: { type: 'string', description: 'add_toc: title above the contents.' },
  },
} as const

export const HELP = `Microsoft Word tools (file-level, macOS). Use them without opening Word; Word only needs to be installed for word_open.

Reading and numbering
- word_read returns every body paragraph as [p0], [p1], ... with its style, and tables as [t0], [t1], ...
- List paragraphs keep their nesting: bullets and numbers are indented under their parent level.
- Paragraph numbers are body paragraphs only; paragraphs inside table cells are not numbered.
- Read again after inserting or deleting paragraphs: the numbers shift.
- For destructive edits pass expect: the first words of the paragraph. If it does not match, the edit stops instead of changing the wrong text.

Lists and indentation
- bullet/list items are strings or {text, level?, ordered?, ...format}. level is 1-9; nested levels indent under the parent.
- Strings may carry their own structure: "- item" and "  - sub item" become bullets; "1. item", "a. item", "(i) item", "(1) item" and "(A) item" become numbered items at the matching level.
- numbering: legal = 1., (a), (i), (1), (A), (I)...; decimal = 1., 2., 3.; bullet = bullets at every level.
- A paragraph whose text contains two or more list lines is converted to a list automatically.
- word_read prints nested lists as indented Markdown markers.

word_create
- blocks are added in order. Types:
  paragraph {text, style?, align?, bold?, italic?, size?, color?, font?, firstLineIndent?, leftIndent?, rightIndent?, indentChars?, lineSpacing?, spaceBefore?, spaceAfter?, keepWithNext?, keepTogether?, widowControl?, pageBreakBefore?}
  heading {text, level: 1-6}
  bullets {items: [text | {text, level?, ...format}], numbering?}
  list {items: [...], ordered?: true, numbering?: legal|decimal|bullet}
  quote {text}
  code {text}
  table {rows: [[cell, ...], ...], caption?, header?, style?}
  image {path, width? ("12cm" or "360pt"), caption?}
  page_break {} or new_page {}
  toc {title?}
- Top-level page: {paper: A4|A3|B5|Letter, orientation: portrait|landscape, top/bottom/left/right in cm}
- header, footer, page_numbers, title, subtitle, author, subject, keywords, comments, overwrite, open.

word_edit ops
- insert_paragraphs {items, after?, where?: after|before|start|end}
- insert_list {items, ordered?, numbering?, style?, after?, where?} — same list syntax as bullets/list.
- set_text {para, expect, text} — keeps the paragraph style; refuses pictures/fields.
- replace_text {find, replace, all?} — literal, within one paragraph; formatting is preserved when the match is inside one run.
- delete_range {para, expect, to?} — removes paragraphs para..to.
- insert_table {rows, caption?, header?, after?, where?}
- set_cell {table, row, col, text}
- insert_image {path, width?, caption?, after?, where?}
- format_paragraph {para, expect?, bold?, italic?, underline?, size?, color?, font?, align?, firstLineIndent?, leftIndent?, rightIndent?, indentChars?, lineSpacing?, spaceBefore?, spaceAfter?, keepWithNext?, keepTogether?, widowControl?, pageBreakBefore?}
- style_format {style, font?, size?, bold?, italic?, color?, align?, firstLineIndent?, leftIndent?, rightIndent?, lineSpacing?, spaceBefore?, spaceAfter?, keepWithNext?, keepTogether?, widowControl?, pageBreakBefore?} — changes a style everywhere it is used.
- page_setup {paper?, orientation?, top?, bottom?, left?, right?}
- header {text}; footer {text}; page_numbers {align?, start?}
- page_break {para, expect} — starts that paragraph on a new page. Without para, inserts a page-break paragraph; use after?/where? to place it.
- new_page is an alias for page_break.
- add_toc {title?}; update_fields {}
- Ops run in order and stop at the first failure. The file is saved atomically; the previous version is .bak unless backup is off.

word_export
- Exports PDF. output defaults to the .docx path with .pdf. engine: auto (LibreOffice first, then Word), libreoffice, or word.
- Word export uses AppleScript and may show a macOS Automation permission prompt the first time.

word_render
- Exports PDF, then renders pages to PNG so you can inspect layout. pages: "1" (default), "all", "1-3,5".
- Uses the same PDF engine. Keep max_pages small; the result includes image attachments.

word_app / word_open
- word_app {action: "status", path?} reports Word and whether the file is open.
- word_app {action: "open", path} or word_open {path} brings the document to the front in Word.
- Editing works on the file even when Word is closed. If the file is open in Word with unsaved changes, those changes can overwrite yours: save or close it first.

word_help
- Call at the start of a Word task or when an operation is unclear.

Colours are #RRGGBB. Font sizes and lengths accept pt, cm, mm, in, or "2ch" for first-line indents.`;

const normalizeEngine = (value: unknown, fallback: PdfEngine): PdfEngine => {
  return value === 'libreoffice' || value === 'word' || value === 'none' ? value : fallback
}

async function openWarning(engine: WordEngine, path: string, signal: AbortSignal, label: string): Promise<string> {
  const open = await engine.isOpenInWord(path, signal)
  return open === true ? `\nWarning: ${label} is open in Word. Save or close it first, or your edits may be overwritten.` : ''
}

export function createTools(host: ToolHost): ToolDefinition[] {
  const { engine } = host
  const tools: ToolDefinition[] = []

  tools.push(defineTool({
    name: 'word_create',
    description: 'Create a .docx from blocks: headings, paragraphs, nested bullets/lists, quotes, code, tables, images, page breaks/new pages and a table of contents. Supports list indentation and legal numbering, typesetting fields, page size/margins, header, footer and page numbers. Writes atomically and backs up an existing file unless overwrite is false.',
    parameters: {
      path: PATH,
      blocks: { type: 'array', items: BLOCK, description: 'Content blocks in document order.' },
      page: PAGE,
      title: { type: 'string' },
      subtitle: { type: 'string' },
      author: { type: 'string' },
      subject: { type: 'string' },
      keywords: { type: 'string' },
      comments: { type: 'string' },
      header: { type: 'string' },
      footer: { type: 'string' },
      page_numbers: { type: 'boolean' },
      overwrite: { type: 'boolean', description: 'Replace an existing file (default false).' },
      backup: { type: 'boolean', description: 'Keep a .bak of the previous file (default from plugin config).' },
      open: { type: 'boolean', description: 'Open the new document in Word after writing it.' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      const workspace = sessionWorkspace(engine, exec)
      const path = await engine.resolvePath(args.path, 'create', workspace)
      const result = await engine.create(path, {
        blocks: await resolveBlockPaths(engine, (args.blocks ?? []) as Array<Record<string, unknown>>, workspace),
        page: args.page,
        title: args.title,
        subtitle: args.subtitle,
        author: args.author,
        subject: args.subject,
        keywords: args.keywords,
        comments: args.comments,
        header: args.header,
        footer: args.footer,
        page_numbers: args.page_numbers ?? false,
        overwrite: args.overwrite ?? false,
        backup: args.backup,
      }, exec.signal)
      let text = formatCreate(result)
      if (args.open) {
        try {
          await engine.openInWord(path, exec.signal)
          text += '\nOpened in Microsoft Word.'
        } catch (error) {
          if (error instanceof WordError && error.code === 'ABORTED') throw error
          text += `\nCould not open Word: ${error instanceof Error ? error.message : String(error)}`
        }
      }
      return { text }
    },
  }))

  tools.push(defineTool({
    name: 'word_read',
    description: 'Read a .docx as numbered Markdown: [p0], [p1] for body paragraphs (with style and list nesting) and [t0] for tables, plus outline, properties, header/footer and statistics. Read before word_edit; numbers shift after inserts or deletes.',
    parameters: {
      path: PATH,
      max_chars: { type: 'integer', description: 'Maximum Markdown characters (default from plugin config).' },
      include_header_footer: { type: 'boolean', description: 'Include the first section header and footer (default true).' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      const path = await engine.resolvePath(args.path, 'read', sessionWorkspace(engine, exec))
      const result = await engine.read(path, {
        maxChars: args.max_chars,
        includeHeaderFooter: args.include_header_footer,
      }, exec.signal)
      return { text: formatRead(result) }
    },
  }))

  tools.push(defineTool({
    name: 'word_edit',
    description: 'Edit an existing .docx with a batch of operations, in order: paragraphs, nested lists, tables, cells, images, typesetting, page setup, header/footer, page numbers, page breaks/new pages and TOC. Read with word_read first and pass expect on destructive operations. Stops at the first failed op, saves atomically, and keeps a .bak by default.',
    parameters: {
      path: PATH,
      ops: { type: 'array', items: OP, required: true, description: 'Operations in order; see word_help.' },
      backup: { type: 'boolean', description: 'Keep a .bak of the previous file (default from plugin config).' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      const workspace = sessionWorkspace(engine, exec)
      const path = await engine.resolvePath(args.path, 'write', workspace)
      const result: EditResult = await engine.edit(path, {
        ops: await resolveOpPaths(engine, args.ops as Array<Record<string, unknown>>, workspace),
        backup: args.backup,
      }, exec.signal)
      const warning = await openWarning(engine, path, exec.signal, basename(path))
      return { text: formatEdit(result) + warning }
    },
  }))

  tools.push(defineTool({
    name: 'word_export',
    description: 'Export a .docx to PDF. Default output is the same name with .pdf. engine: auto (LibreOffice first, then Word), libreoffice, or word. Word export uses AppleScript and may prompt for Automation permission.',
    parameters: {
      path: PATH,
      output: { type: 'string', description: 'PDF path; defaults to the .docx path with .pdf.' },
      engine: { type: 'string', enum: ['auto', 'libreoffice', 'word'], description: 'PDF engine (default from plugin config).' },
      overwrite: { type: 'boolean', description: 'Replace an existing PDF (default false).' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      const workspace = sessionWorkspace(engine, exec)
      const input = await engine.resolvePath(args.path, 'read', workspace)
      const output = await engine.resolvePath(
        args.output ?? input.replace(/\.docx$/i, '.pdf'),
        'pdf',
        workspace,
      )
      const result = await engine.exportPdf(
        input,
        output,
        normalizeEngine(args.engine, engine.settings.pdfEngine),
        args.overwrite ?? false,
        exec.signal,
      )
      const warning = await openWarning(engine, input, exec.signal, basename(input))
      return { text: formatExport(result) + warning }
    },
  }))

  tools.push(defineTool({
    name: 'word_render',
    description: 'Render pages of a .docx as images so you can inspect the layout. Exports PDF internally, then returns page PNGs. pages: "1" (default), "all", or "1-3,5". Keep max_pages small.',
    parameters: {
      path: PATH,
      pages: { type: 'string', description: 'Pages to render: "1", "all", or "1-3,5".' },
      width: { type: 'integer', description: 'Image width in pixels (default from plugin config).' },
      max_pages: { type: 'integer', description: 'Maximum pages to render.' },
      engine: { type: 'string', enum: ['auto', 'libreoffice', 'word'], description: 'PDF engine.' },
    },
    output: RENDER_OUTPUT,
    async execute(args, exec) {
      const document = await engine.resolvePath(args.path, 'read', sessionWorkspace(engine, exec))
      const result = await engine.renderPages(document, {
        pages: args.pages,
        width: args.width,
        maxPages: args.max_pages,
        engine: args.engine === undefined ? undefined : normalizeEngine(args.engine, engine.settings.pdfEngine),
      }, exec.signal)
      try {
        const images: AttachmentJson[] = []
        for (const page of result.pages) {
          const data = await readFile(page.path)
          const ref = await host.saveImage(data, `word-page-${page.page}.png`)
          images.push(toJson(ref))
        }
        return {
          text: formatRender(result, result.pages.length),
          image: images[0],
          more: images.length > 1 ? images.slice(1) : undefined,
        }
      } finally {
        await result.cleanup()
      }
    },
  }))

  tools.push(defineTool({
    name: 'word_app',
    description: 'Report whether Microsoft Word is installed and whether a document is open in it, or open a document in Word. Editing works on the file even when Word is closed.',
    parameters: {
      action: { type: 'string', enum: ['status', 'open'], description: 'status (default) or open.' },
      path: { type: 'string', description: 'Document path for status or open.' },
    },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      const action = args.action ?? 'status'
      if (action === 'open') {
        if (!args.path) throw new WordError('INVALID_ARGUMENT', 'word_app {action: "open"} needs path.')
        const path = await engine.resolvePath(args.path, 'read', sessionWorkspace(engine, exec))
        return { text: formatApp(await engine.openInWord(path, exec.signal)) }
      }
      const path = args.path === undefined ? undefined : await engine.resolvePath(args.path, 'read', sessionWorkspace(engine, exec))
      return { text: formatApp(await engine.appStatus(path, exec.signal)) }
    },
  }))

  tools.push(defineTool({
    name: 'word_open',
    description: 'Open a .docx in Microsoft Word and bring it to the front. Use this only when the user wants to see or edit the document in Word.',
    parameters: { path: PATH },
    output: TEXT_OUTPUT,
    async execute(args, exec) {
      const path = await engine.resolvePath(args.path, 'read', sessionWorkspace(engine, exec))
      return { text: formatApp(await engine.openInWord(path, exec.signal)) }
    },
  }))

  tools.push(defineTool({
    name: 'word_help',
    description: 'Show the Word plugin guide: read numbering, every create block and edit operation, export/render options and safety notes. Call at the start of a Word task.',
    parameters: {},
    output: TEXT_OUTPUT,
    async execute() {
      return { text: HELP }
    },
  }))

  return tools
}
