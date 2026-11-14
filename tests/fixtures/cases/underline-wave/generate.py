"""Generate underline-wave: w:u wave, wavyHeavy and wavyDouble

Word strokes a zigzag of 1pt half-waves (0.25pt line, 1pt for wavyHeavy)
whose phase follows the text's left edge, not the run; wavyDouble is two
thinner waves at the top and bottom of the single wave's band.

Tests:
- wave, wavyHeavy, wavyDouble in Calibri 11, Times New Roman 20, Arial 32
- a wave run that starts mid-line after plain text
- a wavy paragraph indented 7.25pt, so the phase anchor shows
"""

import pathlib
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


def underlined(p, text, font, size, val):
    r = p.add_run(text)
    r.font.name, r.font.size = font, Pt(size)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), val)
    r._r.get_or_add_rPr().append(u)


doc = Document()
doc.styles["Normal"].paragraph_format.space_after = Pt(6)
for font, size in [("Calibri", 11), ("Times New Roman", 20), ("Arial", 32)]:
    for val in ("wave", "wavyHeavy", "wavyDouble"):
        p = doc.add_paragraph()
        underlined(p, f"{val} {font}", font, size, val)
p = doc.add_paragraph()
r = p.add_run("Plain text, then ")
r.font.name, r.font.size = "Calibri", Pt(14)
underlined(p, "a wavy run mid-line", "Calibri", 14, "wave")
p = doc.add_paragraph()
p.paragraph_format.left_indent = Pt(7.25)
underlined(p, "Indented 7.25pt wavy paragraph", "Calibri", 14, "wavyHeavy")
doc.save(pathlib.Path(__file__).with_name("input.docx"))
