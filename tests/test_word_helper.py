import base64
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
HELPER = ROOT / "helper" / "word_helper.py"
PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII="
)


def call(payload):
    completed = subprocess.run(
        [sys.executable, str(HELPER)],
        input=json.dumps(payload, ensure_ascii=False),
        text=True,
        capture_output=True,
    )
    parsed = json.loads(completed.stdout)
    if not parsed.get("ok"):
        raise AssertionError(parsed.get("error"))
    return parsed["result"]


def call_error(payload):
    completed = subprocess.run(
        [sys.executable, str(HELPER)],
        input=json.dumps(payload, ensure_ascii=False),
        text=True,
        capture_output=True,
    )
    parsed = json.loads(completed.stdout)
    if parsed.get("ok"):
        raise AssertionError("expected an error")
    return parsed["error"]


class WordHelperTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def create(self, name="test.docx", **payload):
        path = self.root / name
        result = call({"command": "create", "path": str(path), **payload})
        return path, result

    def read(self, path, **options):
        return call({"command": "read", "path": str(path), **options})

    def test_probe(self):
        result = call({"command": "probe"})
        self.assertIsNotNone(result["python_docx"])
        self.assertEqual(result["platform"], "darwin")

    def test_create_and_read(self):
        path, created = self.create(
            title="Report",
            author="Analyst",
            blocks=[
                {"type": "heading", "level": 1, "text": "Introduction"},
                {"type": "paragraph", "text": "This is **bold** and *italic*."},
                {"type": "bullets", "items": ["One", "Two"]},
                {"type": "table", "caption": "Table 1", "rows": [["Name", "Value"], ["A", "1"]]},
            ],
            header="Header text",
            footer="Footer text",
        )
        self.assertGreater(created["bytes"], 0)
        result = self.read(path)
        self.assertIn("[p0] (Title) # Report", result["markdown"])
        self.assertIn("[p0] (Title) # Report", result["markdown"])
        self.assertIn("[p1] (Heading 1) # Introduction", result["markdown"])
        self.assertIn("**bold**", result["markdown"])
        self.assertIn("[t0] table 2x2", result["markdown"])
        self.assertIn("| Name | Value |", result["markdown"])
        self.assertEqual(result["header"], "Header text")
        self.assertEqual(result["footer"], "Footer text")
        self.assertEqual(result["outline"][0]["text"], "Introduction")
        self.assertEqual(result["properties"]["title"], "Report")
        self.assertEqual(result["properties"]["author"], "Analyst")

    def test_set_text_and_expect(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "Original text."}])
        error = call_error({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "set_text", "para": 0, "expect": "Wrong", "text": "New text."}],
        })
        self.assertEqual(error["code"], "PARAGRAPH_MISMATCH")
        result = call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "set_text", "para": 0, "expect": "Original", "text": "New text."}],
        })
        self.assertEqual(result["changes"][0]["action"], "set_text")
        self.assertIn("New text.", self.read(path)["markdown"])

    def test_insert_and_delete(self):
        path, _ = self.create(blocks=[
            {"type": "paragraph", "text": "One"},
            {"type": "paragraph", "text": "Two"},
            {"type": "paragraph", "text": "Three"},
        ])
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "insert_paragraphs", "after": 1, "items": ["Inserted"]}],
        })
        markdown = self.read(path)["markdown"]
        self.assertLess(markdown.index("[p1] Two"), markdown.index("[p2] Inserted"))
        self.assertLess(markdown.index("[p2] Inserted"), markdown.index("[p3] Three"))
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "delete_range", "para": 2, "expect": "Inserted"}],
        })
        markdown = self.read(path)["markdown"]
        self.assertNotIn("Inserted", markdown)
        self.assertIn("[p2] Three", markdown)

    def test_replace_text(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "Alpha beta alpha."}])
        result = call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "replace_text", "find": "alpha", "replace": "gamma"}],
        })
        self.assertEqual(result["changes"][0]["detail"], "replaced 1 occurrence(s) in 1 paragraph(s)")
        self.assertIn("Alpha beta gamma.", self.read(path)["markdown"])

    def test_table_caption_position_and_cell(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "Intro"}])
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{
                "op": "insert_table",
                "after": 0,
                "caption": "Table caption",
                "rows": [["A", "B"], ["1", "2"]],
            }],
        })
        markdown = self.read(path)["markdown"]
        self.assertLess(markdown.index("[p1] *Table caption*"), markdown.index("[t0] table 2x2"))
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "set_cell", "table": 0, "row": 1, "col": 1, "text": "9"}],
        })
        self.assertIn("| 1 | 9 |", self.read(path)["markdown"])

    def test_style_and_page_setup(self):
        path, _ = self.create(
            blocks=[{"type": "heading", "level": 1, "text": "Heading"}],
            page={"paper": "A4", "orientation": "landscape", "top": 2.0},
        )
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{
                "op": "style_format",
                "style": "Heading 1",
                "font": "Helvetica",
                "size": "16pt",
                "bold": True,
                "color": "#1F4E79",
            }],
        })
        from docx import Document
        document = Document(str(path))
        self.assertEqual(document.sections[0].orientation, 1)
        self.assertAlmostEqual(document.sections[0].top_margin.cm, 2.0, places=1)
        style = document.styles["Heading 1"]
        self.assertEqual(style.font.size.pt, 16)
        self.assertTrue(style.font.bold)

    def test_header_footer_page_numbers_and_toc(self):
        path, _ = self.create(
            blocks=[{"type": "heading", "level": 1, "text": "Chapter"}, {"type": "toc", "title": "Contents"}],
            header="Top",
            footer="Bottom",
            page_numbers=True,
        )
        document = self.read(path)
        self.assertIn("Contents", document["markdown"])
        self.assertIn("Right-click", document["markdown"])
        self.assertIn("Top", document["header"])
        self.assertIn("Bottom", document["footer"])
        self.assertIn("1", document["footer"])

    def test_image_block(self):
        image = self.root / "dot.png"
        image.write_bytes(PNG)
        path, _ = self.create(blocks=[{"type": "image", "path": str(image), "caption": "Figure 1"}])
        result = self.read(path)
        self.assertEqual(result["stats"]["images"], 1)
        self.assertIn("<image>", result["markdown"])
        self.assertIn("Figure 1", result["markdown"])

    def test_backup(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "One"}])
        original = path.read_bytes()
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "insert_paragraphs", "items": ["Two"]}],
            "backup": True,
        })
        backup = Path(str(path) + ".bak")
        self.assertTrue(backup.exists())
        self.assertEqual(backup.read_bytes(), original)

    def test_render_pdf(self):
        import fitz
        pdf = self.root / "sample.pdf"
        document = fitz.open()
        page = document.new_page(width=595, height=842)
        page.insert_text((72, 72), "Hello")
        document.save(str(pdf))
        document.close()
        result = call({
            "command": "render_pdf",
            "pdf": str(pdf),
            "out_dir": str(self.root / "pages"),
            "pages": "1",
            "width": 600,
        })
        self.assertEqual(result["total_pages"], 1)
        self.assertEqual(result["pages"][0]["width"], 600)
        self.assertTrue(Path(result["pages"][0]["path"]).exists())
        self.assertEqual(Path(result["pages"][0]["path"]).read_bytes()[:4], b"\x89PNG")

    def test_list_dict_items(self):
        path, _ = self.create(blocks=[{"type": "list", "items": [{"text": "One"}, "Two"]}])
        markdown = self.read(path)["markdown"]
        self.assertIn("1. One", markdown)
        self.assertIn("1. Two", markdown)

    def test_newlines_become_paragraphs(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "One\nTwo\nThree"}])
        markdown = self.read(path)["markdown"]
        self.assertIn("[p0] One", markdown)
        self.assertIn("[p1] Two", markdown)
        self.assertIn("[p2] Three", markdown)

    def test_insert_paragraphs_at_start(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "First"}])
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "insert_paragraphs", "where": "start", "items": ["Before"]}],
        })
        markdown = self.read(path)["markdown"]
        self.assertLess(markdown.index("[p0] Before"), markdown.index("[p1] First"))

    def test_insert_image_after_with_caption(self):
        image = self.root / "dot.png"
        image.write_bytes(PNG)
        path, _ = self.create(blocks=[
            {"type": "paragraph", "text": "Intro"},
            {"type": "paragraph", "text": "Outro"},
        ])
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{
                "op": "insert_image",
                "after": 0,
                "path": str(image),
                "caption": "Figure A",
                "width": "1cm",
            }],
        })
        markdown = self.read(path)["markdown"]
        self.assertLess(markdown.index("[p0] Intro"), markdown.index("<image>"))
        self.assertLess(markdown.index("<image>"), markdown.index("Figure A"))
        self.assertLess(markdown.index("Figure A"), markdown.index("[p3] Outro"))

    def test_code_and_table_size_strings(self):
        path, _ = self.create(blocks=[
            {"type": "table", "rows": [["A", "B"]], "size": "11pt"},
            {"type": "code", "text": "print(1)", "size": "9pt"},
        ])
        result = self.read(path)
        self.assertIn("`print(1)`", result["markdown"])
        self.assertEqual(result["stats"]["tables"], 1)

    def test_default_author_is_empty(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "Text"}])
        self.assertEqual(self.read(path)["properties"]["author"], "")

    def test_page_break_marker(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "First"}, {"type": "paragraph", "text": "Second"}])
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "page_break", "para": 1, "expect": "Second"}],
        })
        self.assertIn("<page break> Second", self.read(path)["markdown"])

    def test_page_numbers_keep_existing_text(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "Text"}], footer="Keep", page_numbers=True)
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "page_numbers", "align": "right"}],
        })
        footer = self.read(path)["footer"]
        self.assertIn("Keep", footer)
        self.assertEqual(footer.count("Keep"), 1)

    def test_page_numbers_preserve_footer(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "Text"}], footer="Keep")
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "page_numbers", "align": "right"}],
        })
        footer = self.read(path)["footer"]
        self.assertIn("Keep", footer)
        self.assertIn("1", footer)

    def test_unsupported_doc(self):
        path = self.root / "legacy.doc"
        path.write_bytes(b"not a doc")
        error = call_error({"command": "read", "path": str(path)})
        self.assertEqual(error["code"], "UNSUPPORTED_FORMAT")

    def test_all_block_types_and_properties(self):
        image = self.root / "dot.png"
        image.write_bytes(PNG)
        path, created = self.create(
            title="Document",
            subtitle="Subtitle",
            author="Author",
            subject="Subject",
            keywords="one, two",
            comments="Comment",
            header="Top",
            footer="Bottom",
            page_numbers=True,
            page={"paper": "B5", "orientation": "landscape", "top": 2.5, "bottom": 2.0, "left": 1.8, "right": 1.8},
            blocks=[
                {"type": "heading", "level": 2, "text": "Heading"},
                {"type": "paragraph", "text": "Body text"},
                {"type": "quote", "text": "Quoted text"},
                {"type": "code", "text": "print(1)"},
                {"type": "bullets", "items": ["Bullet one", "Bullet two"]},
                {"type": "list", "ordered": False, "items": ["Plain one"]},
                {"type": "table", "align": "center", "rows": [["A", "B"], ["1", "2"]]},
                {"type": "image", "path": str(image), "width": "1cm"},
                {"type": "page_break"},
                {"type": "toc", "title": "Contents"},
            ],
        )
        self.assertIsNone(created["backup"])
        result = self.read(path)
        self.assertEqual(result["properties"]["title"], "Document")
        self.assertEqual(result["properties"]["author"], "Author")
        self.assertEqual(result["properties"]["subject"], "Subject")
        self.assertEqual(result["properties"]["keywords"], "one, two")
        self.assertEqual(result["properties"]["comments"], "Comment")
        self.assertIn("[p0] (Title) # Document", result["markdown"])
        self.assertIn("[p1] (Subtitle) Subtitle", result["markdown"])
        self.assertIn("[p2] (Heading 2) ## Heading", result["markdown"])
        self.assertIn("> Quoted text", result["markdown"])
        self.assertIn("`print(1)`", result["markdown"])
        self.assertIn("- Bullet one", result["markdown"])
        self.assertIn("- Plain one", result["markdown"])
        self.assertIn("[t0] table 2x2", result["markdown"])
        self.assertIn("<image>", result["markdown"])
        self.assertIn("Contents", result["markdown"])
        self.assertEqual(result["outline"][0]["text"], "Heading")
        self.assertEqual(result["stats"]["tables"], 1)
        self.assertEqual(result["stats"]["images"], 1)
        self.assertEqual(result["header"], "Top")
        self.assertIn("Bottom", result["footer"])

        from docx import Document
        document = Document(str(path))
        section = document.sections[0]
        self.assertGreater(section.page_width, section.page_height)
        self.assertAlmostEqual(section.top_margin.cm, 2.5, places=1)
        self.assertAlmostEqual(section.left_margin.cm, 1.8, places=1)

    def test_edit_op_batch(self):
        image = self.root / "dot.png"
        image.write_bytes(PNG)
        path, _ = self.create(blocks=[
            {"type": "paragraph", "text": "One"},
            {"type": "heading", "level": 2, "text": "Heading"},
            {"type": "paragraph", "text": "Two old"},
            {"type": "paragraph", "text": "Three"},
        ])
        result = call({
            "command": "edit",
            "path": str(path),
            "ops": [
                {"op": "format_paragraph", "para": 0, "expect": "One", "bold": True, "italic": True, "underline": True,
                 "size": "12pt", "color": "#112233", "font": "Helvetica", "align": "center",
                 "firstLineIndent": "2ch", "lineSpacing": "1.5", "spaceBefore": "6pt", "spaceAfter": "6pt"},
                {"op": "insert_paragraphs", "after": 0, "items": [
                    "Between", {"text": "Styled", "style": "Heading 2"},
                ]},
                {"op": "set_text", "para": 4, "expect": "Two", "text": "Second"},
                {"op": "replace_text", "find": "Second", "replace": "2nd", "all": False},
                {"op": "insert_table", "after": 0, "caption": "Table caption", "rows": [["A", "B"], ["1", "2"]]},
                {"op": "set_cell", "table": 0, "row": 1, "col": 1, "text": "9"},
                {"op": "insert_image", "after": 0, "path": str(image), "caption": "Figure caption", "width": "1cm"},
                {"op": "style_format", "style": "Normal", "font": "Helvetica", "size": "10pt", "color": "#445566",
                 "align": "justify", "firstLineIndent": "1ch", "lineSpacing": "1.2", "spaceBefore": "3pt", "spaceAfter": "3pt"},
                {"op": "page_setup", "paper": "Letter", "orientation": "landscape", "top": 1.5, "left": 2.0},
                {"op": "header", "text": "Header text"},
                {"op": "footer", "text": "Footer text"},
                {"op": "page_numbers", "align": "right", "start": 3},
                {"op": "page_break", "para": 0, "expect": "One"},
                {"op": "add_toc", "title": "Contents"},
                {"op": "update_fields"},
            ],
        })
        self.assertEqual(len(result["changes"]), 15)
        self.assertTrue(Path(str(path) + ".bak").exists())
        document = self.read(path)
        self.assertIn("<page break>", document["markdown"])
        self.assertIn("*One*", document["markdown"])
        self.assertIn("2nd", document["markdown"])
        self.assertIn("| 1 | 9 |", document["markdown"])
        self.assertIn("Contents", document["markdown"])
        self.assertEqual(document["stats"]["tables"], 1)
        self.assertEqual(document["stats"]["images"], 1)
        self.assertEqual(document["header"], "Header text")
        self.assertIn("Footer text", document["footer"])
        self.assertIn("3", document["footer"])
        self.assertEqual([item["text"] for item in document["outline"]], ["Styled", "Heading"])

        from docx import Document
        loaded = Document(str(path))
        parsed = loaded.sections[0]
        self.assertGreater(parsed.page_width, parsed.page_height)
        self.assertAlmostEqual(parsed.top_margin.cm, 1.5, places=1)
        self.assertAlmostEqual(parsed.left_margin.cm, 2.0, places=1)
        self.assertTrue(loaded.paragraphs[0].paragraph_format.page_break_before)
        first = loaded.paragraphs[0].runs[0]
        self.assertTrue(first.bold)
        self.assertTrue(first.italic)
        self.assertTrue(first.underline)
        self.assertEqual(loaded.styles["Normal"].font.size.pt, 10)

    def test_read_truncation_and_header_toggle(self):
        path, _ = self.create(
            blocks=[{"type": "paragraph", "text": "word " * 2000}],
            header="Hidden header",
        )
        result = self.read(path, max_chars=2000, include_header_footer=False)
        self.assertTrue(result["truncated"])
        self.assertIn("[truncated at 2000 characters]", result["markdown"])
        self.assertEqual(result["header"], "")

    def test_file_exists_and_backup_control(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "Original"}])
        error = call_error({"command": "create", "path": str(path), "blocks": [{"type": "paragraph", "text": "Again"}]})
        self.assertEqual(error["code"], "FILE_EXISTS")
        call({"command": "create", "path": str(path), "overwrite": True,
              "backup": True, "blocks": [{"type": "paragraph", "text": "Again"}]})
        self.assertTrue(Path(str(path) + ".bak").exists())
        Path(str(path) + ".bak").unlink()
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "insert_paragraphs", "items": ["More"]}],
            "backup": False,
        })
        self.assertFalse(Path(str(path) + ".bak").exists())

    def test_edit_errors(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "Text"}])
        unknown = call_error({"command": "edit", "path": str(path), "ops": [{"op": "nope"}]})
        self.assertEqual(unknown["code"], "UNKNOWN_OPERATION")
        bad_color = call_error({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "format_paragraph", "para": 0, "color": "blue"}],
        })
        self.assertEqual(bad_color["code"], "INVALID_ARGUMENT")
        missing_para = call_error({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "set_text", "para": 99, "text": "x"}],
        })
        self.assertEqual(missing_para["code"], "PARAGRAPH_NOT_FOUND")
        missing_table = call_error({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "set_cell", "table": 0, "row": 0, "col": 0, "text": "x"}],
        })
        self.assertEqual(missing_table["code"], "TABLE_NOT_FOUND")
        image = self.root / "dot.png"
        image.write_bytes(PNG)
        image_path, _ = self.create("image.docx", blocks=[{"type": "image", "path": str(image)}])
        image_error = call_error({
            "command": "edit",
            "path": str(image_path),
            "ops": [{"op": "set_text", "para": 0, "text": "x"}],
        })
        self.assertEqual(image_error["code"], "PARAGRAPH_HAS_OBJECT")

    def test_render_page_ranges(self):
        import fitz
        pdf = self.root / "pages.pdf"
        document = fitz.open()
        for number in range(3):
            page = document.new_page(width=595, height=842)
            page.insert_text((72, 72), f"Page {number + 1}")
        document.save(str(pdf))
        document.close()
        result = call({
            "command": "render_pdf",
            "pdf": str(pdf),
            "out_dir": str(self.root / "rendered"),
            "pages": "1-3,2",
            "width": 400,
            "max_pages": 2,
        })
        self.assertEqual(result["total_pages"], 3)
        self.assertEqual([page["page"] for page in result["pages"]], [1, 2])
        error = call_error({
            "command": "render_pdf",
            "pdf": str(pdf),
            "out_dir": str(self.root / "rendered"),
            "pages": "4",
        })
        self.assertEqual(error["code"], "PAGE_NOT_FOUND")

    def test_nested_lists_and_numbering(self):
        path, _ = self.create(blocks=[
            {"type": "list", "numbering": "legal", "items": [
                "1. First",
                "   a. Alpha",
                "   b. Beta",
                "      i. Roman",
                "         (1) Decimal sub",
                "            (A) Upper alpha",
                "2. Second",
            ]},
            {"type": "bullets", "items": ["- Bullet", "  - Sub bullet", "    - Deeper"]},
        ])
        markdown = self.read(path)["markdown"]
        self.assertIn("[p0] (List Paragraph) 1. First", markdown)
        self.assertIn("[p1] (List Paragraph)   (a) Alpha", markdown)
        self.assertIn("[p2] (List Paragraph)   (a) Beta", markdown)
        self.assertIn("    (i) Roman", markdown)
        self.assertIn("      (1) Decimal sub", markdown)
        self.assertIn("        (A) Upper alpha", markdown)
        self.assertIn("[p6] (List Paragraph) 1. Second", markdown)
        self.assertIn("[p7] (List Paragraph) - Bullet", markdown)
        self.assertIn("[p8] (List Paragraph)   - Sub bullet", markdown)
        self.assertIn("[p9] (List Paragraph)     - Deeper", markdown)

        from docx import Document
        from docx.oxml.ns import qn
        document = Document(str(path))
        levels = [paragraph._p.pPr.numPr.ilvl.val for paragraph in document.paragraphs]
        self.assertEqual(levels, [0, 1, 1, 2, 3, 4, 0, 0, 1, 2])
        formats = {}
        for abstract in document.part.numbering_part.element.findall(qn("w:abstractNum")):
            for level in abstract.findall(qn("w:lvl")):
                index = int(level.get(qn("w:ilvl")))
                fmt = level.find(qn("w:numFmt")).get(qn("w:val"))
                formats.setdefault(index, set()).add(fmt)
        self.assertIn("decimal", formats[0])
        self.assertIn("lowerLetter", formats[1])
        self.assertIn("lowerRoman", formats[2])
        self.assertIn("decimal", formats[3])
        self.assertIn("upperLetter", formats[4])
        self.assertIn("upperRoman", formats[5])
        self.assertIn("bullet", formats[0])

    def test_auto_parse_list_paragraph(self):
        path, _ = self.create(blocks=[
            {"type": "paragraph", "text": "1. One\n2. Two\n   a. Sub"},
            {"type": "paragraph", "text": "1. Single line stays a paragraph"},
        ])
        markdown = self.read(path)["markdown"]
        self.assertIn("[p0] (List Paragraph) 1. One", markdown)
        self.assertIn("[p1] (List Paragraph) 1. Two", markdown)
        self.assertIn("[p2] (List Paragraph)   (a) Sub", markdown)
        self.assertIn("[p3] 1. Single line stays a paragraph", markdown)

    def test_insert_list_and_new_page(self):
        path, _ = self.create(blocks=[
            {"type": "paragraph", "text": "Intro"},
            {"type": "paragraph", "text": "Outro"},
        ])
        result = call({
            "command": "edit",
            "path": str(path),
            "ops": [
                {"op": "insert_list", "after": 0, "items": [
                    "1. Top", "   a. Sub", "   b. Sub two", "2. Top two",
                ]},
                {"op": "new_page", "after": 4},
            ],
        })
        self.assertEqual([change["action"] for change in result["changes"]], ["insert_list", "new_page"])
        markdown = self.read(path)["markdown"]
        self.assertIn("[p1] (List Paragraph) 1. Top", markdown)
        self.assertIn("[p2] (List Paragraph)   (a) Sub", markdown)
        self.assertIn("[p5] <page break>", markdown)
        self.assertIn("[p6] Outro", markdown)

    def test_insert_list_at_start_keeps_order(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "Existing"}])
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{"op": "insert_list", "where": "start", "items": ["1. One", "2. Two"]}],
        })
        markdown = self.read(path)["markdown"]
        self.assertLess(markdown.index("1. One"), markdown.index("1. Two"))
        self.assertLess(markdown.index("1. Two"), markdown.index("Existing"))

    def test_typesetting_flags(self):
        path, _ = self.create(blocks=[{"type": "paragraph", "text": "Typeset"}])
        call({
            "command": "edit",
            "path": str(path),
            "ops": [{
                "op": "format_paragraph", "para": 0, "expect": "Typeset",
                "leftIndent": "1cm", "rightIndent": "2cm",
                "keepWithNext": True, "keepTogether": True, "widowControl": False,
                "pageBreakBefore": True,
            }],
        })
        from docx import Document
        document = Document(str(path))
        fmt = document.paragraphs[0].paragraph_format
        self.assertAlmostEqual(fmt.left_indent.cm, 1.0, places=1)
        self.assertAlmostEqual(fmt.right_indent.cm, 2.0, places=1)
        self.assertTrue(fmt.keep_with_next)
        self.assertTrue(fmt.keep_together)
        self.assertFalse(fmt.widow_control)
        self.assertTrue(fmt.page_break_before)

    def test_read_list_styles(self):
        path, _ = self.create(blocks=[
            {"type": "paragraph", "text": "Styled bullet", "style": "List Bullet 2"},
            {"type": "paragraph", "text": "Styled number", "style": "List Number 3"},
        ])
        markdown = self.read(path)["markdown"]
        self.assertIn("  - Styled bullet", markdown)
        self.assertIn("    1. Styled number", markdown)


if __name__ == "__main__":
    unittest.main()
