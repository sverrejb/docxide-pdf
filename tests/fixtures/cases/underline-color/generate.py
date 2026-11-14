"""Generate underline-color: w:u w:color apart from the text colour

Tests:
- single, double, dotted, thick, wave and words underlines in red, green or
  blue under black or coloured text (Calibri 12)
- w:color="auto" under red text (the underline takes the text colour)
- w:themeColor="accent2" beside a w:color that disagrees with the theme
"""

import pathlib
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

CASES = [
    ("single", "FF0000", None, None),
    ("double", "00B050", "1F3864", None),
    ("dotted", "0070C0", None, None),
    ("thick", "C00000", "7F7F7F", None),
    ("wave", "7030A0", None, None),
    ("words", "FF0000", "0070C0", None),
    ("single", "auto", "FF0000", None),
    ("single", "000000", None, "accent2"),
]

doc = Document()
doc.styles["Normal"].paragraph_format.space_after = Pt(8)
for val, ucolor, text_color, theme in CASES:
    p = doc.add_paragraph()
    r = p.add_run(f"{val} underline in {theme or ucolor}")
    r.font.name, r.font.size = "Calibri", Pt(12)
    if text_color:
        r.font.color.rgb = RGBColor.from_string(text_color)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), val)
    u.set(qn("w:color"), ucolor)
    if theme:
        u.set(qn("w:themeColor"), theme)
    r._r.get_or_add_rPr().append(u)
doc.save(pathlib.Path(__file__).with_name("input.docx"))
