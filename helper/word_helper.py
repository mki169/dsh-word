#!/usr/bin/env python3
import json
import os
import plistlib
import re
import shutil
import sys
import time
from datetime import datetime
from importlib.metadata import PackageNotFoundError, version as package_version
from pathlib import Path

VERSION = "0.1.0"
WORD_APP = Path("/Applications/Microsoft Word.app")
LIBREOFFICE = Path("/Applications/LibreOffice.app/Contents/MacOS/soffice")
PAPER_SIZES = {
    "A4": (21.0, 29.7),
    "A3": (29.7, 42.0),
    "B5": (17.6, 25.0),
    "Letter": (21.59, 27.94),
}


class HelperError(Exception):
    def __init__(self, code, message):
        super().__init__(message)
        self.code = code
        self.message = message


def emit(value):
    sys.stdout.write(json.dumps(value, ensure_ascii=False))
    sys.stdout.flush()


def fail(code, message):
    return {"ok": False, "error": {"code": code, "message": message}}


def pkg_version(name):
    try:
        return package_version(name)
    except PackageNotFoundError:
        return None


def word_version():
    info = WORD_APP / "Contents" / "Info.plist"
    if not info.exists():
        return None
    try:
        with info.open("rb") as handle:
            data = plistlib.load(handle)
        return data.get("CFBundleShortVersionString")
    except Exception:
        return None


def probe():
    try:
        import docx
        docx_version = getattr(docx, "__version__", None) or pkg_version("python-docx")
    except Exception:
        docx_version = None
    return {
        "helper": VERSION,
        "python": sys.version.split()[0],
        "python_docx": docx_version,
        "pymupdf": pkg_version("pymupdf"),
        "platform": sys.platform,
        "word_app": WORD_APP.exists(),
        "word_version": word_version(),
        "libreoffice": LIBREOFFICE.exists(),
    }


def require_docx():
    try:
        from docx import Document
        return Document
    except Exception as exc:
        raise HelperError(
            "PYTHON_DOCX_MISSING",
            "python-docx is required. Install it with: python3 -m pip install --user python-docx",
        ) from exc


def require_fitz():
    try:
        import fitz
        return fitz
    except Exception as exc:
        raise HelperError(
            "PYMUPDF_MISSING",
            "PyMuPDF is required for word_render. Install it with: python3 -m pip install --user pymupdf",
        ) from exc


def unsupported(path):
    suffix = Path(path).suffix.lower()
    if suffix == ".doc":
        raise HelperError(
            "UNSUPPORTED_FORMAT",
            "Legacy .doc files are not supported. Open the file in Word and save it as .docx first.",
        )
    raise HelperError("UNSUPPORTED_FORMAT", f"Only .docx files are supported, got {suffix or 'no extension'}.")


def load_document(path):
    target = Path(path)
    if not target.exists():
        raise HelperError("FILE_NOT_FOUND", f"File not found: {target}")
    if not target.is_file():
        raise HelperError("NOT_A_FILE", f"Not a file: {target}")
    if target.suffix.lower() != ".docx":
        unsupported(target)
    Document = require_docx()
    try:
        return Document(str(target))
    except Exception as exc:
        raise HelperError("INVALID_DOCX", f"Cannot open {target}: {exc}") from exc


def save_document(document, path, backup=True):
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temp = target.with_name(f".{target.name}.{os.getpid()}.{int(time.time() * 1000)}.tmp")
    backup_path = None
    try:
        document.save(str(temp))
        if target.exists():
            mode = target.stat().st_mode
            os.chmod(temp, mode)
            if backup:
                backup_target = Path(str(target) + ".bak")
                shutil.copy2(target, backup_target)
                backup_path = str(backup_target)
        os.replace(temp, target)
    finally:
        try:
            temp.unlink()
        except FileNotFoundError:
            pass
    return backup_path


def core_properties(document):
    props = document.core_properties

    def text(value):
        if value is None:
            return ""
        if isinstance(value, datetime):
            return value.isoformat()
        return str(value)

    return {
        "title": text(props.title),
        "author": text(props.author),
        "subject": text(props.subject),
        "keywords": text(props.keywords),
        "comments": text(props.comments),
        "last_modified_by": text(props.last_modified_by),
        "revision": int(props.revision or 0),
        "created": text(props.created),
        "modified": text(props.modified),
    }


def paragraph_index(document, paragraph):
    for index, candidate in enumerate(document.paragraphs):
        if candidate._p is paragraph._p:
            return index
    return None


def table_index(document, table):
    for index, candidate in enumerate(document.tables):
        if candidate._tbl is table._tbl:
            return index
    return None


def find_paragraph(document, index):
    if index is None:
        raise HelperError("INVALID_ARGUMENT", "para is required.")
    try:
        index = int(index)
    except (TypeError, ValueError):
        raise HelperError("INVALID_ARGUMENT", f"para must be an integer, got {index!r}.")
    paragraphs = document.paragraphs
    if index < 0 or index >= len(paragraphs):
        raise HelperError(
            "PARAGRAPH_NOT_FOUND",
            f"Paragraph {index} does not exist. The document has {len(paragraphs)} body paragraphs (0-{max(len(paragraphs) - 1, 0)}).",
        )
    return paragraphs[index]


def check_expect(paragraph, expect):
    if expect is None:
        return
    actual = paragraph.text.strip()
    expected = str(expect).strip()
    if expected == "":
        if actual != "":
            raise HelperError(
                "PARAGRAPH_MISMATCH",
                f"The paragraph is not empty; it starts with {actual[:80]!r}.",
            )
        return
    if not actual.startswith(expected):
        raise HelperError(
            "PARAGRAPH_MISMATCH",
            f"Paragraph text does not start with {expected[:80]!r}; it starts with {actual[:80]!r}. Read the document again.",
        )


def check_text_only(paragraph):
    xml = paragraph._p
    if xml.xpath(".//w:drawing") or xml.xpath(".//w:object") or xml.xpath(".//w:fldChar"):
        raise HelperError(
            "PARAGRAPH_HAS_OBJECT",
            "That paragraph contains a picture or a field. Edit it in Word instead of replacing the whole paragraph.",
        )


def parse_length(value, font_size=11.0, characters=False):
    from docx.shared import Cm, Inches, Mm, Pt

    if value is None:
        return None
    if isinstance(value, (int, float)):
        if characters:
            return Pt(float(value) * font_size)
        return Pt(float(value))
    raw = str(value).strip().lower()
    if raw == "":
        return None
    if raw.endswith("ch"):
        return Pt(float(raw[:-2]) * font_size)
    try:
        amount = float(raw)
        return Pt(amount * font_size) if characters else Pt(amount)
    except ValueError:
        pass
    match = re.fullmatch(r"([0-9.]+)\s*(pt|cm|mm|in|inch|inches|磅|厘米|毫米|英寸)", raw)
    if not match:
        raise HelperError("INVALID_ARGUMENT", f"Unsupported length {value!r}. Use a number or pt/cm/mm/in.")
    amount = float(match.group(1))
    unit = match.group(2)
    if unit in ("pt", "磅"):
        return Pt(amount)
    if unit in ("cm", "厘米"):
        return Cm(amount)
    if unit in ("mm", "毫米"):
        return Mm(amount)
    return Inches(amount)


def parse_color(value):
    from docx.shared import RGBColor

    if value is None:
        return None
    raw = str(value).strip().lstrip("#")
    if not re.fullmatch(r"[0-9a-fA-F]{6}", raw):
        raise HelperError("INVALID_ARGUMENT", f"Colour must be #RRGGBB, got {value!r}.")
    return RGBColor.from_string(raw.upper())


def apply_run_format(run, spec):
    from docx.shared import Pt

    if spec.get("font") is not None:
        run.font.name = str(spec["font"])
    if spec.get("latinFont") is not None:
        run.font.name = str(spec["latinFont"])
    if spec.get("size") is not None or spec.get("fontSize") is not None:
        run.font.size = parse_length(spec.get("size", spec.get("fontSize")))
    if spec.get("bold") is not None:
        run.font.bold = bool(spec["bold"])
    if spec.get("italic") is not None:
        run.font.italic = bool(spec["italic"])
    if spec.get("underline") is not None:
        run.font.underline = bool(spec["underline"])
    if spec.get("color") is not None:
        run.font.color.rgb = parse_color(spec["color"])
    if spec.get("superscript") is not None:
        run.font.superscript = bool(spec["superscript"])
    if spec.get("subscript") is not None:
        run.font.subscript = bool(spec["subscript"])


def apply_line_spacing(fmt, spacing):
    from docx.enum.text import WD_LINE_SPACING

    if isinstance(spacing, (int, float)):
        fmt.line_spacing = float(spacing)
        return
    raw = str(spacing).strip().lower()
    if raw.startswith("at least "):
        fmt.line_spacing = parse_length(raw[len("at least "):])
        fmt.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
        return
    if raw.startswith("exactly "):
        fmt.line_spacing = parse_length(raw[len("exactly "):])
        fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        return
    try:
        fmt.line_spacing = float(raw)
    except ValueError:
        fmt.line_spacing = parse_length(raw)


def apply_paragraph_format(paragraph, spec, style_font_size=11.0):
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

    alignment = spec.get("align") or spec.get("alignment")
    if alignment is not None:
        mapping = {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        key = str(alignment).lower()
        if key not in mapping:
            raise HelperError("INVALID_ARGUMENT", "align must be left, center, right or justify.")
        paragraph.alignment = mapping[key]

    fmt = paragraph.paragraph_format
    before = spec.get("spaceBefore", spec.get("space_before"))
    if before is not None:
        fmt.space_before = parse_length(before)
    after = spec.get("spaceAfter", spec.get("space_after"))
    if after is not None:
        fmt.space_after = parse_length(after)

    spacing = spec.get("lineSpacing", spec.get("line_spacing"))
    if spacing is not None:
        apply_line_spacing(fmt, spacing)

    indent = spec.get("firstLineIndent", spec.get("first_line_indent"))
    if indent is not None:
        fmt.first_line_indent = parse_length(indent, style_font_size, characters=True)
    indent_chars = spec.get("indentChars", spec.get("indent_chars"))
    if indent_chars is not None:
        fmt.first_line_indent = parse_length(indent_chars, style_font_size, characters=True)
    left = spec.get("leftIndent", spec.get("left_indent"))
    if left is not None:
        fmt.left_indent = parse_length(left)
    right = spec.get("rightIndent", spec.get("right_indent"))
    if right is not None:
        fmt.right_indent = parse_length(right)
    if spec.get("keepWithNext") is not None:
        fmt.keep_with_next = bool(spec["keepWithNext"])
    if spec.get("pageBreakBefore") is not None:
        fmt.page_break_before = bool(spec["pageBreakBefore"])


def add_runs(paragraph, text, formatting=None):
    formatting = formatting or {}
    pattern = re.compile(r"(\*\*.+?\*\*|\*.+?\*|`.+?`)", re.S)
    position = 0
    for match in pattern.finditer(text):
        if match.start() > position:
            run = paragraph.add_run(text[position:match.start()])
            apply_run_format(run, formatting)
        token = match.group(0)
        if token.startswith("**"):
            run = paragraph.add_run(token[2:-2])
            run.bold = True
            apply_run_format(run, formatting)
            run.bold = True
        elif token.startswith("`"):
            run = paragraph.add_run(token[1:-1])
            run.font.name = "Courier New"
            apply_run_format(run, formatting)
        else:
            run = paragraph.add_run(token[1:-1])
            run.italic = True
            apply_run_format(run, formatting)
            run.italic = True
        position = match.end()
    if position < len(text):
        run = paragraph.add_run(text[position:])
        apply_run_format(run, formatting)


def add_explicit_runs(paragraph, runs, formatting=None):
    for item in runs:
        if isinstance(item, str):
            run = paragraph.add_run(item)
            apply_run_format(run, formatting or {})
        elif isinstance(item, dict):
            run = paragraph.add_run(str(item.get("text", "")))
            apply_run_format(run, item)
        else:
            raise HelperError("INVALID_ARGUMENT", "runs items must be strings or objects.")


def style_exists(document, name):
    try:
        document.styles[name]
        return True
    except KeyError:
        return False


def add_caption(document, anchor_xml, text, position="before"):
    from docx.shared import Pt

    if not text:
        return None
    paragraph = document.add_paragraph()
    paragraph.alignment = 1
    run = paragraph.add_run(text)
    run.italic = True
    run.font.size = Pt(9)
    if position == "before":
        anchor_xml.addprevious(paragraph._p)
    else:
        anchor_xml.addnext(paragraph._p)
    return paragraph


def ensure_trailing_paragraph(document):
    body = document.element.body
    last = None
    for child in body.iterchildren():
        if child.tag.endswith("}sectPr"):
            continue
        last = child
    if last is not None and last.tag.endswith("}tbl"):
        paragraph = document.add_paragraph()
        last.addnext(paragraph._p)


def add_field(paragraph, instruction, placeholder=""):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction_node = OxmlElement("w:instrText")
    instruction_node.set(qn("xml:space"), "preserve")
    instruction_node.text = instruction
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    value = OxmlElement("w:t")
    value.text = placeholder
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instruction_node)
    run._r.append(separate)
    run._r.append(value)
    run._r.append(end)
    return run


def mark_fields_for_update(document):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    settings = document.settings.element
    if settings.find(qn("w:updateFields")) is None:
        node = OxmlElement("w:updateFields")
        node.set(qn("w:val"), "true")
        settings.append(node)


def iter_block_items(document):
    from docx.document import Document as DocumentClass
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    parent = document.element.body
    for child in parent.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, document)
        elif isinstance(child, CT_Tbl):
            yield Table(child, document)


def runs_markdown(paragraph):
    parts = []
    for run in paragraph.runs:
        text = run.text
        if text == "":
            continue
        if run.font.superscript:
            text = f"^{text}^"
        elif run.font.subscript:
            text = f"~{text}~"
        if run.bold:
            text = f"**{text}**"
        if run.italic:
            text = f"*{text}*"
        if run.underline:
            text = f"<u>{text}</u>"
        if run.font.name == "Courier New":
            text = f"`{text}`"
        parts.append(text)
    return "".join(parts) or paragraph.text


def heading_level(style_name):
    match = re.fullmatch(r"Heading ([1-6])", style_name or "")
    if match:
        return int(match.group(1))
    if style_name == "Title":
        return 1
    return None


def markdown_escape_cell(value):
    return value.replace("|", "\\|").replace("\n", "<br>")


def render_markdown(document, include_header_footer=True, max_chars=40000):
    lines = []
    outline = []
    paragraph_number = 0
    table_number = 0
    image_count = 0
    word_count = 0
    for block in iter_block_items(document):
        if block.__class__.__name__ == "Paragraph":
            style_name = block.style.name if block.style is not None else ""
            text = block.text or ""
            words = len(re.findall(r"\S+", text))
            word_count += words
            page_break = bool(block._p.xpath('.//w:br[@w:type="page"]')) or bool(block.paragraph_format.page_break_before)
            has_drawing = bool(block._p.xpath(".//w:drawing"))
            if has_drawing:
                image_count += len(block._p.xpath(".//w:drawing"))
            visible = runs_markdown(block)
            level = heading_level(style_name)
            if level is not None and text.strip() and style_name != "Title":
                outline.append({"index": paragraph_number, "level": level, "style": style_name, "text": text.strip()})
            prefix = f"[p{paragraph_number}]"
            if style_name and style_name != "Normal":
                prefix += f" ({style_name})"
            if level is not None and text.strip():
                body = "#" * level + " " + visible
            elif text.strip() == "" and page_break:
                body = "<page break>"
            elif text.strip() == "" and has_drawing:
                body = "<image>"
            elif text.strip() == "":
                body = "(empty paragraph)"
            elif style_name.startswith("List Bullet"):
                body = "- " + visible
            elif style_name.startswith("List Number"):
                body = "1. " + visible
            elif style_name == "Quote":
                body = "> " + visible
            else:
                body = visible
            if page_break and body != "<page break>":
                body = "<page break> " + body
            if has_drawing and body != "<image>":
                body += " <image>"
            lines.append(f"{prefix} {body}".rstrip())
            paragraph_number += 1
        else:
            rows = []
            for row in block.rows:
                rows.append([markdown_escape_cell(cell.text or "") for cell in row.cells])
            width = max((len(row) for row in rows), default=0)
            words = sum(len(re.findall(r"\S+", cell)) for row in rows for cell in row)
            word_count += words
            lines.append(f"[t{table_number}] table {len(rows)}x{width}")
            if rows:
                header = rows[0] + [""] * (width - len(rows[0]))
                lines.append("| " + " | ".join(header) + " |")
                lines.append("| " + " | ".join(["---"] * width) + " |")
                for row in rows[1:]:
                    padded = row + [""] * (width - len(row))
                    lines.append("| " + " | ".join(padded) + " |")
            table_number += 1

    header = ""
    footer = ""
    if include_header_footer and document.sections:
        section = document.sections[0]
        try:
            header = "\n".join(p.text for p in section.header.paragraphs if p.text).strip()
        except Exception:
            header = ""
        try:
            footer = "\n".join(p.text for p in section.footer.paragraphs if p.text).strip()
        except Exception:
            footer = ""

    markdown = "\n".join(lines)
    truncated = len(markdown) > max_chars
    if truncated:
        markdown = markdown[:max_chars] + f"\n... [truncated at {max_chars} characters]"
    return {
        "markdown": markdown,
        "outline": outline,
        "truncated": truncated,
        "header": header,
        "footer": footer,
        "words": word_count,
        "images": image_count,
    }


def doc_stats(document, path):
    words = 0
    for paragraph in document.paragraphs:
        words += len(re.findall(r"\S+", paragraph.text or ""))
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                words += len(re.findall(r"\S+", cell.text or ""))
    target = Path(path)
    return {
        "paragraphs": len(document.paragraphs),
        "tables": len(document.tables),
        "images": len(document.inline_shapes),
        "sections": len(document.sections),
        "words": words,
        "pages_estimate": max(1, round(words / 450)),
        "bytes": target.stat().st_size if target.exists() else 0,
    }


def read_command(payload):
    path = payload.get("path")
    if not path:
        raise HelperError("INVALID_ARGUMENT", "path is required.")
    document = load_document(path)
    max_chars = int(payload.get("max_chars") or payload.get("maxChars") or 40000)
    max_chars = max(2000, min(max_chars, 200000))
    rendered = render_markdown(
        document,
        include_header_footer=bool(payload.get("include_header_footer", payload.get("includeHeaderFooter", True))),
        max_chars=max_chars,
    )
    stats = doc_stats(document, path)
    stats["words"] = rendered["words"]
    stats["images"] = max(stats["images"], rendered["images"])
    return {
        "path": str(Path(path)),
        "stats": stats,
        "properties": core_properties(document),
        "markdown": rendered["markdown"],
        "outline": rendered["outline"],
        "truncated": rendered["truncated"],
        "header": rendered["header"],
        "footer": rendered["footer"],
    }


def apply_page_setup(document, spec):
    from docx.enum.section import WD_ORIENT
    from docx.shared import Cm

    if not spec:
        return
    paper = spec.get("paper")
    orientation = spec.get("orientation")
    for section in document.sections:
        if paper and str(paper) in PAPER_SIZES:
            width_cm, height_cm = PAPER_SIZES[str(paper)]
            if orientation and str(orientation).lower() == "landscape":
                width_cm, height_cm = height_cm, width_cm
            section.page_width = Cm(width_cm)
            section.page_height = Cm(height_cm)
        if orientation:
            key = str(orientation).lower()
            if key == "landscape":
                section.orientation = WD_ORIENT.LANDSCAPE
                if section.page_width < section.page_height:
                    section.page_width, section.page_height = section.page_height, section.page_width
            elif key == "portrait":
                section.orientation = WD_ORIENT.PORTRAIT
                if section.page_width > section.page_height:
                    section.page_width, section.page_height = section.page_height, section.page_width
            else:
                raise HelperError("INVALID_ARGUMENT", "orientation must be portrait or landscape.")
        for key in ("top", "bottom", "left", "right"):
            value = spec.get(key)
            if value is None:
                continue
            margin = float(value)
            setattr(section, f"{key}_margin", Cm(margin))


def container_paragraph(container):
    paragraphs = container.paragraphs
    if paragraphs:
        return paragraphs[0]
    return container.add_paragraph()


def apply_header_footer(document, header, footer, page_numbers):
    for section in document.sections:
        if header is not None:
            paragraph = container_paragraph(section.header)
            paragraph.text = str(header)
        if footer is not None:
            paragraph = container_paragraph(section.footer)
            paragraph.text = str(footer)
        if page_numbers:
            paragraph = container_paragraph(section.footer)
            if paragraph.text:
                paragraph.add_run("    ")
            else:
                paragraph.alignment = 1
            add_field(paragraph, "PAGE", "1")


def add_toc(document, block):
    from docx.enum.text import WD_ALIGN_PARAGRAPH

    title = block.get("title")
    if title:
        paragraph = document.add_paragraph()
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = paragraph.add_run(str(title))
        run.bold = True
    paragraph = document.add_paragraph()
    add_field(paragraph, 'TOC \\o "1-3" \\h \\z \\u', "Right-click and choose Update Field.")
    mark_fields_for_update(document)
    return paragraph


def table_alignment(block):
    value = block.get("align")
    if value is None:
        return None
    return value


def add_table(document, block):
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt

    data = block.get("rows") or block.get("data") or []
    if not data or not isinstance(data, list):
        raise HelperError("INVALID_ARGUMENT", "table needs rows: [[cell, ...], ...].")
    width = max(len(row) for row in data)
    table = document.add_table(rows=len(data), cols=width)
    cell_size = parse_length(block.get("size", 10.5))
    style = block.get("style") or "Table Grid"
    if style_exists(document, style):
        table.style = style
    header = bool(block.get("header", True))
    for row_index, row in enumerate(data):
        for col_index in range(width):
            value = row[col_index] if col_index < len(row) else ""
            cell = table.cell(row_index, col_index)
            cell.text = str(value)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = cell_size
                    if header and row_index == 0:
                        run.bold = True
            if header and row_index == 0:
                shading = OxmlElement("w:shd")
                shading.set(qn("w:fill"), str(block.get("headerFill", "D9E2F3")))
                cell._tc.get_or_add_tcPr().append(shading)
    align = table_alignment(block)
    if align is not None:
        mapping = {
            "left": WD_TABLE_ALIGNMENT.LEFT,
            "center": WD_TABLE_ALIGNMENT.CENTER,
            "right": WD_TABLE_ALIGNMENT.RIGHT,
        }
        key = str(align).lower()
        if key in mapping:
            table.alignment = mapping[key]
    return table


def place_after(anchor, element):
    anchor.addnext(element)


def place_before(anchor, element):
    anchor.addprevious(element)


def add_block(document, block):
    if not isinstance(block, dict):
        raise HelperError("INVALID_ARGUMENT", "Each block must be an object.")
    kind = str(block.get("type") or "paragraph")
    if kind == "heading":
        level = int(block.get("level", 1))
        level = max(1, min(level, 6))
        text = str(block.get("text", ""))
        paragraph = document.add_heading(text, level=level)
        apply_paragraph_format(paragraph, block)
        for run in paragraph.runs:
            apply_run_format(run, block)
        return paragraph
    if kind == "paragraph":
        text = str(block.get("text", ""))
        parts = text.split("\n") if "\n" in text and not block.get("runs") else [text]
        paragraphs = []
        for part in parts:
            paragraph = document.add_paragraph()
            style = block.get("style")
            if style and style_exists(document, str(style)):
                paragraph.style = str(style)
            runs = block.get("runs")
            if runs:
                add_explicit_runs(paragraph, runs, block)
            else:
                add_runs(paragraph, part, block)
            apply_paragraph_format(paragraph, block)
            paragraphs.append(paragraph)
        return paragraphs[0] if len(paragraphs) == 1 else paragraphs
    if kind == "bullets":
        items = block.get("items") or []
        paragraphs = []
        for item in items:
            if isinstance(item, dict):
                text = str(item.get("text", ""))
                item_format = item
            else:
                text = str(item)
                item_format = block
            paragraph = document.add_paragraph()
            if style_exists(document, "List Bullet"):
                paragraph.style = "List Bullet"
            add_runs(paragraph, text, item_format)
            apply_paragraph_format(paragraph, item_format)
            paragraphs.append(paragraph)
        return paragraphs
    if kind == "list":
        items = block.get("items") or []
        ordered = bool(block.get("ordered", True))
        style = "List Number" if ordered else "List Bullet"
        paragraphs = []
        for item in items:
            if isinstance(item, dict):
                text = str(item.get("text", ""))
                item_format = item
            else:
                text = str(item)
                item_format = block
            paragraph = document.add_paragraph()
            if style_exists(document, style):
                paragraph.style = style
            add_runs(paragraph, text, item_format)
            apply_paragraph_format(paragraph, item_format)
            paragraphs.append(paragraph)
        return paragraphs
    if kind == "quote":
        paragraph = document.add_paragraph()
        if style_exists(document, "Quote"):
            paragraph.style = "Quote"
        add_runs(paragraph, str(block.get("text", "")), block)
        apply_paragraph_format(paragraph, block)
        return paragraph
    if kind == "code":
        paragraph = document.add_paragraph()
        run = paragraph.add_run(str(block.get("text", "")))
        run.font.name = "Courier New"
        run.font.size = parse_length(block.get("size", 9))
        apply_paragraph_format(paragraph, block)
        return paragraph
    if kind == "table":
        caption = block.get("caption")
        table = add_table(document, block)
        if caption:
            add_caption(document, table._tbl, str(caption), "before")
        return table
    if kind == "image":
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Cm, Pt

        path = block.get("path")
        if not path:
            raise HelperError("INVALID_ARGUMENT", "image block needs path.")
        image_path = Path(str(path)).expanduser()
        if not image_path.exists():
            raise HelperError("IMAGE_NOT_FOUND", f"Image not found: {image_path}")
        width = block.get("width")
        kwargs = {}
        if width is not None:
            if isinstance(width, (int, float)):
                kwargs["width"] = Pt(float(width))
            else:
                kwargs["width"] = parse_length(width)
        document.add_picture(str(image_path), **kwargs)
        paragraph = document.paragraphs[-1]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        caption = block.get("caption")
        if caption:
            add_caption(document, paragraph._p, str(caption), "after")
        return paragraph
    if kind == "page_break":
        document.add_page_break()
        return document.paragraphs[-1]
    if kind == "toc":
        return add_toc(document, block)
    raise HelperError("INVALID_ARGUMENT", f"Unknown block type {kind!r}.")


def move_block(paragraph_or_table, anchor, where):
    element = paragraph_or_table._p if hasattr(paragraph_or_table, "_p") else paragraph_or_table._tbl
    if anchor is None:
        return
    if where == "before":
        anchor.addprevious(element)
    else:
        anchor.addnext(element)


def create_command(payload):
    path = payload.get("path")
    if not path:
        raise HelperError("INVALID_ARGUMENT", "path is required.")
    target = Path(path)
    if target.suffix.lower() != ".docx":
        unsupported(target)
    if target.exists() and not payload.get("overwrite", False):
        raise HelperError("FILE_EXISTS", f"{target} already exists. Pass overwrite: true to replace it.")
    Document = require_docx()
    document = Document()
    apply_page_setup(document, payload.get("page"))
    props = document.core_properties
    props.author = str(payload.get("author", ""))
    props.last_modified_by = ""
    for key in ("title", "subject", "keywords", "comments"):
        if payload.get(key) is not None:
            setattr(props, key, str(payload[key]))
    if payload.get("title"):
        title = document.add_paragraph()
        if style_exists(document, "Title"):
            title.style = "Title"
        title.add_run(str(payload["title"]))
    if payload.get("subtitle"):
        subtitle = document.add_paragraph()
        if style_exists(document, "Subtitle"):
            subtitle.style = "Subtitle"
        subtitle.add_run(str(payload["subtitle"]))
    blocks = payload.get("blocks") or []
    for block in blocks:
        add_block(document, block)
    ensure_trailing_paragraph(document)
    apply_header_footer(
        document,
        payload.get("header"),
        payload.get("footer"),
        bool(payload.get("page_numbers", payload.get("pageNumbers", False))),
    )
    backup_path = save_document(document, target, backup=bool(payload.get("backup", False)))
    return {
        "path": str(target),
        "bytes": target.stat().st_size,
        "backup": backup_path,
        "stats": doc_stats(document, target),
    }


def replace_in_paragraph(paragraph, find, replace, all_matches):
    if find == "":
        raise HelperError("INVALID_ARGUMENT", "find must not be empty.")
    count = 0
    for run in paragraph.runs:
        if find in run.text:
            if all_matches:
                count += run.text.count(find)
                run.text = run.text.replace(find, replace)
            elif count == 0:
                count += 1
                run.text = run.text.replace(find, replace, 1)
            if not all_matches and count:
                return count
    if count:
        return count
    if find in paragraph.text:
        if all_matches:
            count = paragraph.text.count(find)
        else:
            count = 1
        new_text = paragraph.text.replace(find, replace, -1 if all_matches else 1)
        paragraph.text = new_text
    return count


def edit_command(payload):
    path = payload.get("path")
    if not path:
        raise HelperError("INVALID_ARGUMENT", "path is required.")
    document = load_document(path)
    ops = payload.get("ops")
    if not isinstance(ops, list) or not ops:
        raise HelperError("INVALID_ARGUMENT", "ops must be a non-empty array.")
    changes = []
    for op_index, op in enumerate(ops):
        if not isinstance(op, dict):
            raise HelperError("INVALID_ARGUMENT", f"Op {op_index} must be an object.")
        name = op.get("op")
        try:
            detail = ""
            if name == "insert_paragraphs":
                items = op.get("items")
                if not isinstance(items, list) or not items:
                    raise HelperError("INVALID_ARGUMENT", "insert_paragraphs needs items: [...].")
                anchor = None
                where = str(op.get("where", "after" if op.get("after") is not None else "end"))
                if op.get("after") is None and where in ("start", "before"):
                    where = "start"
                at_start = where == "start"
                if op.get("after") is not None:
                    anchor = find_paragraph(document, op["after"])._p
                elif at_start:
                    body = document.element.body
                    anchor = next((child for child in body.iterchildren() if child.tag.endswith("}p") or child.tag.endswith("}tbl")), None)
                cursor = anchor
                inserted = []
                for item in items:
                    if isinstance(item, str):
                        item = {"text": item}
                    if not isinstance(item, dict):
                        raise HelperError("INVALID_ARGUMENT", "insert_paragraphs items must be strings or objects.")
                    text = str(item.get("text", ""))
                    text_parts = text.split("\n") if "\n" in text and not item.get("runs") else [text]
                    for part in text_parts:
                        paragraph = document.add_paragraph()
                        style = item.get("style")
                        if style and style_exists(document, str(style)):
                            paragraph.style = str(style)
                        if item.get("runs"):
                            add_explicit_runs(paragraph, item["runs"], item)
                        else:
                            add_runs(paragraph, part, item)
                        apply_paragraph_format(paragraph, item)
                        if cursor is not None:
                            if (where == "before" or at_start) and len(inserted) == 0:
                                place_before(cursor, paragraph._p)
                            else:
                                place_after(cursor, paragraph._p)
                        cursor = paragraph._p
                        inserted.append(paragraph)
                numbers = [paragraph_index(document, paragraph) for paragraph in inserted]
                detail = f"inserted {len(inserted)} paragraph(s) at {numbers}"
            elif name == "set_text":
                paragraph = find_paragraph(document, op.get("para"))
                check_expect(paragraph, op.get("expect"))
                check_text_only(paragraph)
                paragraph.clear()
                add_runs(paragraph, str(op.get("text", "")), op)
                apply_paragraph_format(paragraph, op)
                detail = f"rewrote paragraph {op.get('para')}"
            elif name == "replace_text":
                find = str(op.get("find", ""))
                replace = str(op.get("replace", ""))
                all_matches = bool(op.get("all", True))
                total = 0
                paragraphs = 0
                for paragraph in document.paragraphs:
                    if find in paragraph.text:
                        count = replace_in_paragraph(paragraph, find, replace, all_matches)
                        if count:
                            total += count
                            paragraphs += 1
                            if not all_matches:
                                break
                detail = f"replaced {total} occurrence(s) in {paragraphs} paragraph(s)"
            elif name == "delete_range":
                start = int(op.get("para"))
                end = int(op.get("to", start))
                if end < start:
                    raise HelperError("INVALID_ARGUMENT", "to must be greater than or equal to para.")
                paragraphs = document.paragraphs
                if start < 0 or end >= len(paragraphs):
                    raise HelperError("PARAGRAPH_NOT_FOUND", f"delete_range {start}-{end} is outside 0-{max(len(paragraphs) - 1, 0)}.")
                check_expect(paragraphs[start], op.get("expect"))
                elements = [paragraphs[index]._p for index in range(start, end + 1)]
                for element in elements:
                    parent = element.getparent()
                    if parent is not None:
                        parent.remove(element)
                body = document.element.body
                if not any(child.tag.endswith("}p") for child in body.iterchildren()):
                    document.add_paragraph()
                detail = f"deleted paragraphs {start}-{end}"
            elif name == "insert_table":
                table = add_table(document, op)
                if op.get("caption"):
                    add_caption(document, table._tbl, str(op["caption"]), "before")
                caption_element = table._tbl.getprevious() if op.get("caption") else None
                where = str(op.get("where", "end"))
                anchor = None
                if op.get("after") is not None:
                    anchor = find_paragraph(document, op["after"])._p
                if anchor is not None:
                    if where == "before":
                        anchor.addprevious(table._tbl)
                    else:
                        anchor.addnext(table._tbl)
                elif where in ("start", "before"):
                    first = next((child for child in document.element.body.iterchildren() if child.tag.endswith("}p") or child.tag.endswith("}tbl")), None)
                    if first is not None:
                        first.addprevious(table._tbl)
                if caption_element is not None:
                    table._tbl.addprevious(caption_element)
                ensure_trailing_paragraph(document)
                table_number = table_index(document, table)
                detail = f"inserted table {table_number} ({len(table.rows)}x{len(table.columns)})"
            elif name == "set_cell":
                table_number = int(op.get("table"))
                tables = document.tables
                if table_number < 0 or table_number >= len(tables):
                    raise HelperError("TABLE_NOT_FOUND", f"Table {table_number} does not exist.")
                table = tables[table_number]
                row = int(op.get("row"))
                col = int(op.get("col"))
                if row < 0 or row >= len(table.rows) or col < 0 or col >= len(table.columns):
                    raise HelperError("CELL_NOT_FOUND", f"Cell {row},{col} is outside table {table_number}.")
                table.cell(row, col).text = str(op.get("text", ""))
                detail = f"set cell {row},{col} of table {table_number}"
            elif name == "insert_image":
                image_path = Path(str(op.get("path", ""))).expanduser()
                if not image_path.exists():
                    raise HelperError("IMAGE_NOT_FOUND", f"Image not found: {image_path}")
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                from docx.shared import Pt

                kwargs = {}
                width = op.get("width")
                if width is not None:
                    kwargs["width"] = Pt(float(width)) if isinstance(width, (int, float)) else parse_length(width)
                document.add_picture(str(image_path), **kwargs)
                paragraph = document.paragraphs[-1]
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                caption_element = None
                if op.get("caption"):
                    caption_element = add_caption(document, paragraph._p, str(op["caption"]), "after")._p
                if op.get("after") is not None:
                    anchor = find_paragraph(document, op["after"])._p
                    if str(op.get("where", "after")) == "before":
                        anchor.addprevious(paragraph._p)
                        if caption_element is not None:
                            paragraph._p.addnext(caption_element)
                    else:
                        anchor.addnext(paragraph._p)
                        if caption_element is not None:
                            paragraph._p.addnext(caption_element)
                elif str(op.get("where", "end")) in ("start", "before"):
                    first = next((child for child in document.element.body.iterchildren() if child.tag.endswith("}p") or child.tag.endswith("}tbl")), None)
                    if first is not None:
                        first.addprevious(paragraph._p)
                        if caption_element is not None:
                            paragraph._p.addnext(caption_element)
                detail = f"inserted image at paragraph {paragraph_index(document, paragraph)}"
            elif name == "format_paragraph":
                paragraph = find_paragraph(document, op.get("para"))
                if op.get("expect") is not None:
                    check_expect(paragraph, op.get("expect"))
                apply_paragraph_format(paragraph, op)
                for run in paragraph.runs:
                    apply_run_format(run, op)
                detail = f"formatted paragraph {op.get('para')}"
            elif name == "style_format":
                style_name = op.get("style")
                if not style_name:
                    raise HelperError("INVALID_ARGUMENT", "style_format needs style.")
                try:
                    style = document.styles[str(style_name)]
                except KeyError:
                    raise HelperError("STYLE_NOT_FOUND", f"Style {style_name!r} does not exist.")
                if op.get("font") is not None:
                    style.font.name = str(op["font"])
                if op.get("size") is not None:
                    style.font.size = parse_length(op["size"])
                if op.get("bold") is not None:
                    style.font.bold = bool(op["bold"])
                if op.get("italic") is not None:
                    style.font.italic = bool(op["italic"])
                if op.get("color") is not None:
                    style.font.color.rgb = parse_color(op["color"])
                from docx.enum.style import WD_STYLE_TYPE
                if style.type == WD_STYLE_TYPE.PARAGRAPH:
                    fmt = style.paragraph_format
                    if op.get("align") is not None:
                        from docx.enum.text import WD_ALIGN_PARAGRAPH
                        mapping = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}
                        key = str(op["align"]).lower()
                        if key in mapping:
                            fmt.alignment = mapping[key]
                    if op.get("firstLineIndent") is not None:
                        fmt.first_line_indent = parse_length(op["firstLineIndent"], 11.0, characters=True)
                    if op.get("lineSpacing") is not None:
                        apply_line_spacing(fmt, op["lineSpacing"])
                    if op.get("spaceBefore") is not None:
                        fmt.space_before = parse_length(op["spaceBefore"])
                    if op.get("spaceAfter") is not None:
                        fmt.space_after = parse_length(op["spaceAfter"])
                    if op.get("pageBreakBefore") is not None:
                        fmt.page_break_before = bool(op["pageBreakBefore"])
                detail = f"updated style {style_name!r}"
            elif name == "page_setup":
                apply_page_setup(document, op)
                detail = "updated page setup"
            elif name == "header":
                for section in document.sections:
                    container_paragraph(section.header).text = str(op.get("text", ""))
                detail = "updated header"
            elif name == "footer":
                for section in document.sections:
                    container_paragraph(section.footer).text = str(op.get("text", ""))
                detail = "updated footer"
            elif name == "page_numbers":
                from docx.enum.text import WD_ALIGN_PARAGRAPH
                from docx.oxml import OxmlElement
                from docx.oxml.ns import qn

                mapping = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT}
                for section in document.sections:
                    paragraph = container_paragraph(section.footer)
                    align = mapping.get(str(op.get("align", "center")).lower(), WD_ALIGN_PARAGRAPH.CENTER)
                    if paragraph._p.xpath('.//w:instrText[contains(., "PAGE")]'):
                        paragraph.alignment = align
                    else:
                        if paragraph.text:
                            paragraph.add_run("    ")
                        paragraph.alignment = align
                        add_field(paragraph, "PAGE", str(op.get("start", 1)))
                    if op.get("start") is not None:
                        sect_pr = section._sectPr
                        page_type = sect_pr.find(qn("w:pgNumType"))
                        if page_type is None:
                            page_type = OxmlElement("w:pgNumType")
                            sect_pr.append(page_type)
                        page_type.set(qn("w:start"), str(int(op["start"])))
                detail = "updated page numbers"
            elif name == "page_break":
                paragraph = find_paragraph(document, op.get("para"))
                check_expect(paragraph, op.get("expect"))
                paragraph.paragraph_format.page_break_before = True
                detail = f"page break before paragraph {op.get('para')}"
            elif name == "add_toc":
                add_toc(document, op)
                detail = "inserted table of contents field"
            elif name == "update_fields":
                mark_fields_for_update(document)
                detail = "marked fields for update"
            else:
                raise HelperError("UNKNOWN_OPERATION", f"Unknown op {name!r}.")
            changes.append({"op": op_index, "action": str(name), "status": "ok", "detail": detail})
        except HelperError as exc:
            exc.message = f"Op {op_index} ({name}): {exc.message}"
            raise
        except Exception as exc:
            raise HelperError("OP_FAILED", f"Op {op_index} ({name}) failed: {type(exc).__name__}: {exc}") from exc
    backup_path = save_document(document, path, backup=bool(payload.get("backup", True)))
    return {
        "path": str(Path(path)),
        "backup": backup_path,
        "changes": changes,
        "stats": doc_stats(document, path),
    }


def parse_pages(spec, total):
    if spec is None or spec == "" or str(spec).lower() == "all":
        return list(range(1, total + 1))
    pages = set()
    for part in str(spec).split(","):
        part = part.strip()
        if part == "":
            continue
        if "-" in part:
            start, end = part.split("-", 1)
            start = int(start)
            end = int(end)
            if end < start:
                raise HelperError("INVALID_ARGUMENT", "Page ranges must ascend.")
            pages.update(range(start, end + 1))
        else:
            pages.add(int(part))
    if not pages:
        raise HelperError("INVALID_ARGUMENT", "No pages selected.")
    for page in pages:
        if page < 1 or page > total:
            raise HelperError("PAGE_NOT_FOUND", f"Page {page} is outside 1-{total}.")
    return sorted(pages)


def render_pdf_command(payload):
    pdf_path = payload.get("pdf")
    out_dir = payload.get("out_dir") or payload.get("outDir")
    if not pdf_path or not out_dir:
        raise HelperError("INVALID_ARGUMENT", "pdf and out_dir are required.")
    pdf = Path(pdf_path)
    if not pdf.exists():
        raise HelperError("FILE_NOT_FOUND", f"PDF not found: {pdf}")
    output = Path(out_dir)
    output.mkdir(parents=True, exist_ok=True)
    fitz = require_fitz()
    width = int(payload.get("width") or 1000)
    width = max(200, min(width, 2400))
    try:
        document = fitz.open(str(pdf))
    except Exception as exc:
        raise HelperError("INVALID_PDF", f"Cannot open PDF: {exc}") from exc
    try:
        total = document.page_count
        pages = parse_pages(payload.get("pages"), total)
        max_pages = int(payload.get("max_pages") or payload.get("maxPages") or 12)
        max_pages = max(1, min(max_pages, 40))
        pages = pages[:max_pages]
        rendered = []
        for page_number in pages:
            page = document.load_page(page_number - 1)
            zoom = width / page.rect.width if page.rect.width else 1.0
            matrix = fitz.Matrix(zoom, zoom)
            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            filename = f"page-{page_number:04d}.png"
            path = output / filename
            pixmap.save(str(path))
            rendered.append({
                "page": page_number,
                "path": str(path),
                "width": pixmap.width,
                "height": pixmap.height,
                "mime": "image/png",
            })
        return {"pdf": str(pdf), "total_pages": total, "pages": rendered}
    finally:
        document.close()


def main():
    try:
        raw = sys.stdin.read()
        payload = json.loads(raw or "{}")
    except Exception as exc:
        emit(fail("INVALID_INPUT", f"Cannot read JSON input: {exc}"))
        return 1
    command = payload.get("command")
    try:
        if command == "probe":
            result = probe()
        elif command == "read":
            result = read_command(payload)
        elif command == "create":
            result = create_command(payload)
        elif command == "edit":
            result = edit_command(payload)
        elif command == "render_pdf":
            result = render_pdf_command(payload)
        else:
            raise HelperError("UNKNOWN_COMMAND", f"Unknown command {command!r}.")
    except HelperError as exc:
        emit(fail(exc.code, exc.message))
        return 1
    except Exception as exc:
        emit(fail("HELPER_FAILED", f"{type(exc).__name__}: {exc}"))
        return 1
    emit({"ok": True, "result": result})
    return 0


if __name__ == "__main__":
    sys.exit(main())
