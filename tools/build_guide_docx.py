"""Convert docs/user-guide.md to a Word document.

    python tools/build_guide_docx.py [source.md] [target.docx]

Handles the Markdown used in the guide: headings, paragraphs, bullet and numbered lists,
tables, fenced code blocks, block quotes, **bold** and `code`.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

INLINE = re.compile(r"(\*\*.+?\*\*|`[^`]+`)")


def add_inline(paragraph, text: str) -> None:
    for part in INLINE.split(text):
        if not part:
            continue
        if part.startswith("**"):
            paragraph.add_run(part[2:-2]).bold = True
        elif part.startswith("`"):
            run = paragraph.add_run(part[1:-1])
            run.font.name = "Consolas"
            run.font.size = Pt(10)
        else:
            paragraph.add_run(part)


def shade(cell, colour: str) -> None:
    props = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), colour)
    props.append(shd)


def add_table(doc, rows: list[list[str]]) -> None:
    table = doc.add_table(rows=len(rows), cols=len(rows[0]))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for r, row in enumerate(rows):
        for c, value in enumerate(row[:len(rows[0])]):
            cell = table.cell(r, c)
            paragraph = cell.paragraphs[0]
            add_inline(paragraph, value)
            for run in paragraph.runs:
                run.font.size = Pt(10)
                if r == 0:
                    run.bold = True
            if r == 0:
                shade(cell, "D9E2F3")
    doc.add_paragraph()


def cells(line: str) -> list[str]:
    return [c.strip() for c in line.strip().strip("|").split("|")]


def convert(source: Path, target: Path) -> Path:
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name, style.font.size = "Calibri", Pt(11)
    lines = source.read_text(encoding="utf-8").splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        if not stripped:
            i += 1
        elif stripped.startswith("```"):
            i += 1
            block = []
            while i < len(lines) and not lines[i].strip().startswith("```"):
                block.append(lines[i])
                i += 1
            i += 1
            paragraph = doc.add_paragraph()
            run = paragraph.add_run("\n".join(block))
            run.font.name, run.font.size = "Consolas", Pt(10)
        elif stripped.startswith("#"):
            level = len(stripped) - len(stripped.lstrip("#"))
            doc.add_heading(stripped.lstrip("#").strip().replace("`", ""), level=min(level, 3) - 1 if level == 1 else min(level, 3))
            i += 1
        elif stripped.startswith("|"):
            rows = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                if not re.fullmatch(r"[|\s:-]+", lines[i].strip()):
                    rows.append(cells(lines[i]))
                i += 1
            add_table(doc, rows)
        elif stripped.startswith(">"):
            paragraph = doc.add_paragraph()
            paragraph.paragraph_format.left_indent = Pt(18)
            text = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                text.append(lines[i].strip().lstrip(">").strip())
                i += 1
            add_inline(paragraph, " ".join(text))
            for run in paragraph.runs:
                run.font.color.rgb = RGBColor(0x9C, 0x00, 0x06)
        elif re.match(r"^[-*] ", stripped):
            add_inline(doc.add_paragraph(style="List Bullet"), stripped[2:])
            i += 1
        elif re.match(r"^\d+\. ", stripped):
            add_inline(doc.add_paragraph(style="List Number"), re.sub(r"^\d+\. ", "", stripped))
            i += 1
        else:
            add_inline(doc.add_paragraph(), stripped)
            i += 1
    target.parent.mkdir(parents=True, exist_ok=True)
    doc.save(target)
    return target


if __name__ == "__main__":
    root = Path(__file__).resolve().parent.parent
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else root / "docs" / "user-guide.md"
    dst = Path(sys.argv[2]) if len(sys.argv) > 2 else root / "dist" / "User guide.docx"
    print(f"Written {convert(src, dst)}")
