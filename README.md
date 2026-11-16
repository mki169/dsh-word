# dsh-office-word

Microsoft Word tools for DeepSeek Harness, for macOS. The plugin edits `.docx`
files directly with `python-docx`; it does not need COM, a Windows machine, or a
running Word window. Word is used only when you ask to open a document.

This is the Word-only macOS counterpart of the Windows COM-based
`cli-anything-wps` and `dsh-office` projects.

## What it does

| Tool | Purpose |
|------|---------|
| `word_help` | Show the operation guide. Call it at the start of a Word task. |
| `word_create` | Build a `.docx` from headings, paragraphs, lists, quotes, code, tables, images, page breaks and a TOC. Set page size, margins, header, footer, page numbers and properties. |
| `word_read` | Return numbered Markdown: `[p0]`, `[p1]` for body paragraphs and `[t0]` for tables, plus outline, properties, header/footer and statistics. |
| `word_edit` | Apply a batch of operations in order: insert/set/replace/delete paragraphs, tables, cells, images, formatting, styles, page setup, header/footer, page numbers, page breaks, TOC. |
| `word_export` | Export PDF with LibreOffice (default) or Microsoft Word via AppleScript. |
| `word_render` | Export PDF and render pages to PNG images so the model can inspect the layout. |
| `word_app` | Report whether Word is installed and whether a document is open, or open a document in Word. |
| `word_open` | Open a document in Word and bring it to the front. |

## Requirements

- macOS 12 or later.
- DeepSeek Harness (`dsh`) with the web profile or another profile.
- Python 3.9+ with `python-docx`:

  ```bash
  python3 -m pip install --user python-docx
  ```

- For `word_render`, also `PyMuPDF`:

  ```bash
  python3 -m pip install --user pymupdf
  ```

- For PDF export, at least one of:
  - LibreOffice at `/Applications/LibreOffice.app/Contents/MacOS/soffice`
    (recommended: silent, no Word window); or
  - Microsoft Word at `/Applications/Microsoft Word.app`.

`word_create`, `word_read` and `word_edit` do not need LibreOffice or Word.
`word_open` needs Word.

## Install

From GitHub:

```bash
dsh plugin --profile web add https://github.com/mki169/dsh-word.git
```

From a local checkout:

```bash
cd dsh-word
npm install
npm run build
dsh plugin --profile web add "$(pwd)"
```

`lib/` is committed, so the GitHub install needs no build step. For local
development, run `npm run build` after changing the source. Restart the profile
after installing. After changing `cordis.patch.yml`, reinstall or restart dsh.

To use a different profile, replace `web` with its name (for example `tui`).

The bundle entry is `office-word` (see `cordis.patch.yml`). It is inserted into
the profile's plugin tree and registers the `word_*` tools.

## Configuration

The plugin entry accepts these options (also documented in the dsh settings
schema):

| Option | Default | Meaning |
|--------|---------|---------|
| `python` | `python3` | Python interpreter that runs the helper. |
| `workspace` | session cwd | Base directory for relative document paths. |
| `allowOutsideWorkspace` | `false` | Allow paths outside `workspace`. |
| `backup` | `true` | Copy the previous file to `.bak` before an edit. |
| `pdfEngine` | `auto` | `auto`, `libreoffice`, `word`, or `none`. |
| `sofficePath` | LibreOffice path | LibreOffice executable. |
| `wordPath` | `/Applications/Microsoft Word.app` | Word application path. |
| `timeoutMs` | `90000` | Timeout for one operation. |
| `maxReadChars` | `40000` | Maximum Markdown characters returned by `word_read`. |
| `renderWidth` | `1000` | Page image width in pixels. |
| `maxRenderPages` | `8` | Maximum pages rendered by one call. |

Example profile override:

```yaml
# ~/.dsh/profiles/web/plugins/office-word.yml or a --patch overlay
config:
  workspace: /Users/you/Documents
  pdfEngine: libreoffice
  backup: true
```

(The exact user-override file layout depends on the dsh profile version; the
settings page and `dsh --dump-config` show the active entry.)

## Typical flow

```text
word_help
word_read  {path: "report.docx"}
word_edit  {path: "report.docx", ops: [
  {op: "set_text", para: 4, expect: "Old heading", text: "New heading"},
  {op: "replace_text", find: "2025", replace: "2026"},
  {op: "insert_paragraphs", after: 9, items: ["A new paragraph."]},
  {op: "insert_list", after: 9, items: ["1. First", "   a. Nested"]},
  {op: "new_page", after: 12}
]}
word_render {path: "report.docx", pages: "1-2"}
word_export {path: "report.docx"}
```

Read the document again after inserts or deletes: paragraph numbers shift.
Pass `expect` on destructive operations; if the first words do not match, the
edit stops instead of changing the wrong paragraph.

## Lists, indentation and typesetting

Lists use real Word numbering, so they renumber correctly and keep their
hierarchy:

```json
{
  "type": "list",
  "numbering": "legal",
  "items": [
    "1. First item",
    "   a. Nested item",
    "      i. Roman sub-item",
    "         (1) Decimal sub-item",
    "            (A) Upper-alpha sub-item",
    "2. Second item"
  ]
}
```

- `numbering: "legal"` gives `1.` → `(a)` → `(i)` → `(1)` → `(A)` → `(I)`.
- `numbering: "decimal"` gives decimal numbers at every level.
- `bullets` and `{type: "list", ordered: false}` give bullet and sub-bullet
  levels.
- Items can also be objects with an explicit `level` (1-9) instead of relying
  on indentation or marker syntax.
- A `paragraph` whose text is entirely two or more list lines is parsed as a
  list automatically.
- `word_edit` supports `insert_list` for adding a nested list at a position.

`word_read` keeps the hierarchy in its Markdown, for example:

```text
[p0] (List Paragraph) 1. First item
[p1] (List Paragraph)   (a) Nested item
[p2] (List Paragraph)     (i) Roman sub-item
```

Typesetting fields include `leftIndent`, `rightIndent`, `firstLineIndent`,
`indentChars`, `lineSpacing`, `spaceBefore`, `spaceAfter`, `keepWithNext`,
`keepTogether`, `widowControl` and `pageBreakBefore`. Use `page_break` or
`new_page` to start a new page, and the `new_page` edit operation to insert a
page break between paragraphs.

## Safety and behaviour

- Paragraph numbers in `word_read` count **body paragraphs only**. Paragraphs
  inside table cells are not numbered.
- Edits are applied in order and stop at the first failed operation.
- Saves are atomic: the helper writes a temporary file and replaces the target
  with `os.replace`.
- With `backup: true`, the previous version is copied to `<name>.docx.bak`
  before saving.
- Relative paths—documents and image sources—resolve against `workspace`. When
  unset, the plugin uses the current dsh session cwd; setting `workspace` pins a
  base directory instead. With `allowOutsideWorkspace: false`, the plugin
  rejects paths outside it and resolves symlinks before checking.
- If the file is open in Word, `word_edit` and `word_export` warn you. Word may
  have unsaved changes that would overwrite the plugin's edits; save or close
  the document first.
- Only `.docx` is supported. Legacy `.doc` files must be saved as `.docx` first.
- `.bak` is overwritten on each edit; copy it elsewhere if you need history.

## PDF export and rendering

`auto` uses LibreOffice when it is installed, otherwise Word. LibreOffice runs
headless, so it does not disturb an open Word window. Word export uses
AppleScript:

1. copy the document into Word's sandbox temporary folder;
2. open it in Word;
3. `save as ... file format format PDF`;
4. close the temporary document and copy the PDF to the requested path.

The first Word export may show a macOS **Automation** permission prompt
(System Settings > Privacy & Security > Automation). Allow dsh to control
Microsoft Word. If macOS also asks for file or folder access, grant it and run
the export again. If the permission is denied, use LibreOffice or enable it.

`word_render` uses the same PDF pipeline and then PyMuPDF. The returned images
are stored as dsh attachments; temporary files are removed after the tool
returns.

## Limitations

- Word only. Excel/PowerPoint are intentionally out of scope.
- No live window editing or COM automation. The plugin edits the file, then
  Word shows the result when opened.
- No comments, track changes, footnotes/endnotes, equations, SmartArt, content
  controls or advanced fields beyond TOC/page numbers.
- Reading returns rendered Markdown, not a lossless DOM. Formatting is best
  effort; use `word_render` to check layout.
- Page count in `word_read` is an estimate from word count; actual pagination
  is done by Word/LibreOffice.
- Word export depends on Microsoft Word's AppleScript support and macOS
  Automation permissions; LibreOffice is the more predictable exporter.

## Development

```bash
npm install
npm test
```

`npm test` builds the plugin, runs the Python helper tests, and runs the Node
integration tests. The Node tests cover tool registration, create/read/edit,
path policy, LibreOffice PDF export and page rendering. LibreOffice tests are
skipped when LibreOffice is not installed. To include the Word export test:

```bash
npm run test:word
```

When `src/` changes, commit the rebuilt `lib/` too: GitHub installs use the
committed build.

The Python helper is `helper/word_helper.py`; it reads one JSON command from
stdin and writes one JSON result to stdout. The TypeScript host spawns it,
enforces path policy and timeouts, and exposes the dsh tools.

The helper can be exercised directly:

```bash
echo '{"command":"probe"}' | python3 helper/word_helper.py
```

## Troubleshooting

- **`PYTHON_DOCX_MISSING`**: install `python-docx` for the configured Python.
- **`PYMUPDF_MISSING`**: install `pymupdf` for `word_render`.
- **`PDF_ENGINE_UNAVAILABLE`**: install LibreOffice, install Word, or set
  `pdfEngine`.
- **`PARAGRAPH_MISMATCH`**: the document changed since your last `word_read`;
  read again and use the new `[pN]` numbers.
- **`PATH_OUTSIDE_WORKSPACE`**: pass a path inside the workspace or set
  `allowOutsideWorkspace: true`.
- **Word export hangs or fails**: grant the macOS Automation/file-access
  prompts, or set `pdfEngine: libreoffice`.
