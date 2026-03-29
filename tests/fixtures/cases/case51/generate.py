#!/usr/bin/env python3
"""case51: Nested tables (tables inside table cells).

Tests table rendering when cells contain inner tables at various depths:
1. Simple 2-level nesting: outer 2x2 table with an inner 2x2 table in one cell
2. Inner table with different borders than outer table
3. 3-level nesting: table inside a table inside a table
4. Text before and after an inner table within the same cell
5. Multiple inner tables in the same cell

Usage:
    uv run tests/fixtures/cases/case51/generate.py
"""

import os
import re
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH

OUT = Path("tests/fixtures/cases/case51/input.docx")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def cell_xml(text, bold=False):
    """Simple table cell with text."""
    bold_xml = "<w:b/>" if bold else ""
    return (
        f'<w:tc>'
        f'<w:p><w:r><w:rPr>{bold_xml}<w:sz w:val="22"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        f'</w:tc>'
    )


def shaded_cell_xml(text, shading_color="E7E6E6"):
    """Table cell with background shading."""
    return (
        f'<w:tc>'
        f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{shading_color}"/></w:tcPr>'
        f'<w:p><w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr>'
        f'<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
        f'</w:tc>'
    )


def simple_table_xml(rows, col_widths_twips, border_color="000000", border_sz="4"):
    """Build a simple table XML string."""
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in col_widths_twips)

    border_xml = (
        f'<w:tblBorders>'
        f'<w:top w:val="single" w:sz="{border_sz}" w:color="{border_color}"/>'
        f'<w:left w:val="single" w:sz="{border_sz}" w:color="{border_color}"/>'
        f'<w:bottom w:val="single" w:sz="{border_sz}" w:color="{border_color}"/>'
        f'<w:right w:val="single" w:sz="{border_sz}" w:color="{border_color}"/>'
        f'<w:insideH w:val="single" w:sz="{border_sz}" w:color="{border_color}"/>'
        f'<w:insideV w:val="single" w:sz="{border_sz}" w:color="{border_color}"/>'
        f'</w:tblBorders>'
    )

    rows_xml = ""
    for row in rows:
        cells = "".join(row)
        rows_xml += f'<w:tr>{cells}</w:tr>'

    return (
        f'<w:tbl>'
        f'<w:tblPr>'
        f'<w:tblStyle w:val="TableGrid"/>'
        f'<w:tblW w:w="0" w:type="auto"/>'
        f'{border_xml}'
        f'<w:tblLook w:val="04A0"/>'
        f'</w:tblPr>'
        f'<w:tblGrid>{grid}</w:tblGrid>'
        f'{rows_xml}'
        f'</w:tbl>'
    )


def cell_with_nested_table(before_text, table_xml, after_text):
    """Cell containing text, then a nested table, then more text."""
    parts = '<w:tc>'
    if before_text:
        parts += (
            f'<w:p><w:r><w:rPr><w:sz w:val="22"/></w:rPr>'
            f'<w:t xml:space="preserve">{before_text}</w:t></w:r></w:p>'
        )
    parts += table_xml
    if after_text:
        parts += (
            f'<w:p><w:r><w:rPr><w:sz w:val="22"/></w:rPr>'
            f'<w:t xml:space="preserve">{after_text}</w:t></w:r></w:p>'
        )
    else:
        # Cell must end with a paragraph
        parts += '<w:p/>'
    parts += '</w:tc>'
    return parts


def generate():
    doc = Document()
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    doc.add_paragraph("NESTED_TABLE_PLACEHOLDER")

    tmp = tempfile.mktemp(suffix=".docx")
    doc.save(tmp)

    # --- Scenario 1: Simple 2-level nesting ---
    inner_table_1 = simple_table_xml(
        rows=[
            [shaded_cell_xml("Inner A1"), shaded_cell_xml("Inner A2")],
            [cell_xml("Inner B1"), cell_xml("Inner B2")],
        ],
        col_widths_twips=[2160, 2160],  # ~1.5" each
        border_color="4472C4",
        border_sz="6",
    )

    scenario1_outer = simple_table_xml(
        rows=[
            [cell_xml("Outer A1", bold=True), cell_xml("Outer A2", bold=True)],
            [
                cell_with_nested_table(
                    "Text before inner table:",
                    inner_table_1,
                    "Text after inner table."
                ),
                cell_xml("Outer B2 — this cell has no nested table, just text. "
                         "It should align with the cell containing the inner table."),
            ],
        ],
        col_widths_twips=[4680, 4680],  # ~3.25" each
        border_color="000000",
        border_sz="4",
    )

    # --- Scenario 2: Inner table with different border style ---
    inner_table_2 = simple_table_xml(
        rows=[
            [cell_xml("Red 1"), cell_xml("Red 2"), cell_xml("Red 3")],
            [cell_xml("Red 4"), cell_xml("Red 5"), cell_xml("Red 6")],
        ],
        col_widths_twips=[1440, 1440, 1440],
        border_color="C00000",
        border_sz="8",
    )

    scenario2_outer = simple_table_xml(
        rows=[
            [
                cell_xml("Header Left", bold=True),
                cell_xml("Header Right", bold=True),
            ],
            [
                cell_with_nested_table(None, inner_table_2, None),
                cell_xml("Plain cell beside nested table with thick red borders."),
            ],
        ],
        col_widths_twips=[4680, 4680],
        border_color="000000",
        border_sz="4",
    )

    # --- Scenario 3: 3-level nesting ---
    innermost = simple_table_xml(
        rows=[
            [cell_xml("Deep A"), cell_xml("Deep B")],
        ],
        col_widths_twips=[1080, 1080],
        border_color="006600",
        border_sz="4",
    )

    middle_table = simple_table_xml(
        rows=[
            [cell_xml("Mid Header", bold=True), cell_xml("Mid Header 2", bold=True)],
            [
                cell_with_nested_table("Contains deepest table:", innermost, None),
                cell_xml("Mid plain cell"),
            ],
        ],
        col_widths_twips=[2160, 2160],
        border_color="4472C4",
        border_sz="6",
    )

    scenario3_outer = simple_table_xml(
        rows=[
            [
                cell_xml("Outer single-column", bold=True),
            ],
            [
                cell_with_nested_table(
                    "This cell contains a middle table, which itself contains an innermost table:",
                    middle_table,
                    "Text after the middle table."
                ),
            ],
        ],
        col_widths_twips=[9360],
        border_color="000000",
        border_sz="4",
    )

    # --- Scenario 4: Multiple inner tables in one cell ---
    inner_a = simple_table_xml(
        rows=[
            [cell_xml("Table A: Row 1")],
            [cell_xml("Table A: Row 2")],
        ],
        col_widths_twips=[3600],
        border_color="4472C4",
        border_sz="4",
    )

    inner_b = simple_table_xml(
        rows=[
            [cell_xml("Table B: Col 1"), cell_xml("Table B: Col 2")],
        ],
        col_widths_twips=[1800, 1800],
        border_color="ED7D31",
        border_sz="4",
    )

    # Cell with two sequential inner tables
    multi_table_cell = (
        '<w:tc>'
        '<w:p><w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr>'
        '<w:t>First inner table (blue borders):</w:t></w:r></w:p>'
        f'{inner_a}'
        '<w:p><w:r><w:rPr><w:b/><w:sz w:val="22"/></w:rPr>'
        '<w:t>Second inner table (orange borders):</w:t></w:r></w:p>'
        f'{inner_b}'
        '<w:p/>'
        '</w:tc>'
    )

    scenario4_outer = simple_table_xml(
        rows=[
            [
                cell_xml("Left column", bold=True),
                cell_xml("Right column", bold=True),
            ],
            [
                multi_table_cell,
                cell_xml("This cell has no inner tables. The left cell contains "
                         "two separate inner tables stacked vertically."),
            ],
        ],
        col_widths_twips=[4680, 4680],
        border_color="000000",
        border_sz="4",
    )

    # Full body content
    body_xml = f"""
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Scenario 1: Simple Nested Table</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:rPr><w:sz w:val="22"/></w:rPr>
  <w:t xml:space="preserve">Outer 2x2 table with a 2x2 inner table (blue borders) in the bottom-left cell.</w:t></w:r>
</w:p>
{scenario1_outer}

<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Scenario 2: Different Border Styles</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:rPr><w:sz w:val="22"/></w:rPr>
  <w:t xml:space="preserve">Outer table (thin black borders) contains a 2x3 inner table with thick red borders.</w:t></w:r>
</w:p>
{scenario2_outer}

<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Scenario 3: Three-Level Nesting</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:rPr><w:sz w:val="22"/></w:rPr>
  <w:t xml:space="preserve">Outer table (black) contains a middle table (blue), which contains an innermost table (green).</w:t></w:r>
</w:p>
{scenario3_outer}

<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Scenario 4: Multiple Inner Tables in One Cell</w:t></w:r>
</w:p>
<w:p>
  <w:r><w:rPr><w:sz w:val="22"/></w:rPr>
  <w:t xml:space="preserve">Left cell contains two separate inner tables stacked vertically.</w:t></w:r>
</w:p>
{scenario4_outer}
"""

    # Post-process the ZIP
    with zipfile.ZipFile(tmp, "r") as zin:
        with zipfile.ZipFile(str(OUT), "w", zipfile.ZIP_DEFLATED) as zout:
            doc_xml = zin.read("word/document.xml").decode()

            doc_xml = re.sub(
                r'<w:p[^>]*><w:r><w:t>NESTED_TABLE_PLACEHOLDER</w:t></w:r></w:p>',
                lambda m: body_xml,
                doc_xml,
                count=1,
            )

            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, doc_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.unlink(tmp)
    print(f"Generated {OUT}")


if __name__ == "__main__":
    generate()
