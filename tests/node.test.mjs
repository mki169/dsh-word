import assert from 'node:assert/strict'
import { constants as fsConstants } from 'node:fs'
import { access, mkdtemp, readFile, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import { resolveConfig } from '../lib/config.js'
import { WordEngine, WordError } from '../lib/engine.js'
import { apply } from '../lib/index.js'

const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=',
  'base64',
)
const LIBREOFFICE = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
const WORD = '/Applications/Microsoft Word.app'
const ALL_TOOLS = [
  'word_app',
  'word_create',
  'word_edit',
  'word_export',
  'word_help',
  'word_open',
  'word_read',
  'word_render',
]

async function exists(path) {
  try {
    await access(path, fsConstants.F_OK)
    return true
  } catch {
    return false
  }
}

async function setup(options = {}) {
  const root = await mkdtemp(join(tmpdir(), 'dsh-word-test-'))
  const engine = new WordEngine(resolveConfig({ workspace: root, ...options }), { warn: () => undefined })
  return { root, engine, cleanup: () => rm(root, { recursive: true, force: true }) }
}

test('plugin registers every word tool', () => {
  const names = []
  const ctx = {
    logger: { warn: () => undefined },
    attachments: {
      saveImage: async () => ({ attachmentId: 'test', mediaType: 'image/png', bytes: 1, width: 1, height: 1 }),
    },
    tools: {
      register: tool => {
        names.push(tool.name)
        return () => undefined
      },
    },
    effect: callback => {
      callback()
      return () => undefined
    },
  }
  apply(ctx, {})
  assert.deepEqual(names.sort(), ALL_TOOLS)
})

test('create, read and edit a document', async t => {
  const { root, engine, cleanup } = await setup()
  t.after(cleanup)
  const image = join(root, 'dot.png')
  await writeFile(image, PNG)
  const document = join(root, 'report.docx')
  const created = await engine.create(document, {
    title: 'Report',
    author: 'Analyst',
    header: 'Header text',
    footer: 'Footer text',
    page_numbers: true,
    blocks: [
      { type: 'heading', level: 1, text: 'Introduction' },
      { type: 'paragraph', text: 'This is **bold** and *italic*.' },
      { type: 'bullets', items: ['One', 'Two'] },
      { type: 'list', numbering: 'legal', items: ['1. First', '   a. Sub item', '2. Second'] },
      { type: 'table', caption: 'Table 1', rows: [['Name', 'Value'], ['A', '1']] },
      { type: 'image', path: image, caption: 'Figure 1' },
    ],
  })
  assert.ok(created.bytes > 0)

  const read = await engine.read(document)
  assert.match(read.markdown, /\[p0\] \(Title\) # Report/)
  assert.match(read.markdown, /\[p1\] \(Heading 1\) # Introduction/)
  assert.match(read.markdown, /\[t0\] table 2x2/)
  assert.match(read.markdown, /<image>/)
  assert.match(read.markdown, /\(a\) Sub item/)
  assert.equal(read.header, 'Header text')
  assert.match(read.footer, /^Footer text/)
  assert.equal(read.properties.title, 'Report')
  assert.equal(read.properties.author, 'Analyst')
  assert.equal(read.stats.tables, 1)
  assert.equal(read.stats.images, 1)

  const edited = await engine.edit(document, {
    ops: [
      { op: 'insert_paragraphs', after: 1, items: ['Inserted paragraph'] },
      { op: 'insert_list', after: 2, items: ['1. Step', '   a. Substep'] },
      { op: 'replace_text', find: 'bold', replace: 'strong' },
      { op: 'set_cell', table: 0, row: 1, col: 1, text: '42' },
      { op: 'format_paragraph', para: 2, expect: 'Inserted', bold: true, color: '#1F4E79' },
    ],
  })
  assert.equal(edited.changes.length, 5)
  assert.ok(await exists(`${document}.bak`))

  const updated = await engine.read(document)
  assert.match(updated.markdown, /\[p5\] This is \*\*strong\*\* and \*italic\*\./)
  assert.match(updated.markdown, /\(a\) Substep/)
  assert.match(updated.markdown, /\| A \| 42 \|/)
})

test('path policy keeps file access inside the workspace', async t => {
  const { root, engine, cleanup } = await setup()
  t.after(cleanup)
  const outside = await mkdtemp(join(tmpdir(), 'dsh-word-outside-'))
  t.after(() => rm(outside, { recursive: true, force: true }))
  await assert.rejects(
    engine.resolvePath(join(outside, 'outside.docx'), 'create'),
    error => error instanceof WordError && error.code === 'PATH_OUTSIDE_WORKSPACE',
  )
  const permissive = new WordEngine(resolveConfig({ workspace: root, allowOutsideWorkspace: true }), { warn: () => undefined })
  assert.equal(await permissive.resolvePath(join(outside, 'outside.docx'), 'create'), join(outside, 'outside.docx'))
})

test('libreoffice PDF export and page rendering', { skip: !await exists(LIBREOFFICE) }, async t => {
  const { root, engine, cleanup } = await setup({ pdfEngine: 'libreoffice' })
  t.after(cleanup)
  const document = join(root, 'export.docx')
  await engine.create(document, {
    blocks: [
      { type: 'heading', level: 1, text: 'Export' },
      { type: 'paragraph', text: 'LibreOffice export check.' },
      { type: 'list', numbering: 'legal', items: ['1. First', '   a. Nested', '2. Second'] },
      { type: 'bullets', items: ['Bullet', '  Sub bullet'] },
      { type: 'new_page' },
      { type: 'paragraph', text: 'Second page.' },
    ],
  })
  const pdf = join(root, 'export.pdf')
  const exported = await engine.exportPdf(document, pdf, 'libreoffice', false)
  assert.equal(exported.engine, 'libreoffice')
  assert.ok((await stat(pdf)).size > 0)
  assert.equal((await readFile(pdf)).subarray(0, 4).toString(), '%PDF')

  const rendered = await engine.renderPages(document, { pages: '1', width: 640, engine: 'libreoffice' })
  try {
    assert.ok(rendered.total_pages >= 2)
    assert.equal(rendered.pages.length, 1)
    assert.equal(rendered.pages[0].width, 640)
    assert.deepEqual((await readFile(rendered.pages[0].path)).subarray(0, 4), Buffer.from([0x89, 0x50, 0x4e, 0x47]))
  } finally {
    await rendered.cleanup()
  }
})

test('word PDF export', { skip: process.env.DSH_WORD_TEST_WORD !== '1' || !await exists(WORD) }, async t => {
  const { root, engine, cleanup } = await setup({ pdfEngine: 'word', timeoutMs: 180000 })
  t.after(cleanup)
  const document = join(root, 'word.docx')
  await engine.create(document, {
    blocks: [
      { type: 'heading', level: 1, text: 'Word Export' },
      { type: 'paragraph', text: 'Microsoft Word export check.' },
    ],
  })
  const pdf = join(root, 'word.pdf')
  const exported = await engine.exportPdf(document, pdf, 'word', false)
  assert.equal(exported.engine, 'word')
  assert.equal((await readFile(pdf)).subarray(0, 4).toString(), '%PDF')
})

test('export refuses to overwrite and disabled engines fail', async t => {
  const { root, engine, cleanup } = await setup()
  t.after(cleanup)
  const document = join(root, 'guard.docx')
  await engine.create(document, { blocks: [{ type: 'paragraph', text: 'Guard.' }] })
  const pdf = join(root, 'guard.pdf')
  await writeFile(pdf, 'existing')
  await assert.rejects(
    engine.exportPdf(document, pdf, 'libreoffice', false),
    error => error instanceof WordError && error.code === 'FILE_EXISTS',
  )
  await assert.rejects(
    engine.exportPdf(document, join(root, 'disabled.pdf'), 'none', false),
    error => error instanceof WordError && error.code === 'PDF_ENGINE_DISABLED',
  )
})
