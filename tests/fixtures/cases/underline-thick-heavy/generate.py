"""Generate underline-thick-heavy: w:u thick and the five Heavy dash styles

Word draws them all with one thick line geometry (its top at half the single
underline's offset, about half as deep again as the single line's bottom),
the Heavy ones cut into the plain styles' dash lengths.

Tests:
- thick, dottedHeavy, dashedHeavy, dashLongHeavy, dashDotHeavy,
  dashDotDotHeavy in Calibri 11, Times New Roman 16 and Arial 24
- thick in Aptos 9 and 30, where the minimum thickness and the wide range show
"""

import pathlib
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

STYLES = ["thick", "dottedHeavy", "dashedHeavy", "dashLongHeavy", "dashDotHeavy", "dashDotDotHeavy"]


def underlined(p, text, font, size, val):
    r = p.add_run(text)
    r.font.name, r.font.size = font, Pt(size)
    u = OxmlElement("w:u")
    u.set(qn("w:val"), val)
    r._r.get_or_add_rPr().append(u)


doc = Document()
doc.styles["Normal"].paragraph_format.space_after = Pt(6)
for font, size in [("Calibri", 11), ("Times New Roman", 16), ("Arial", 24)]:
    for val in STYLES:
        p = doc.add_paragraph()
        lead = p.add_run(f"{val}: ")
        lead.font.name, lead.font.size = font, Pt(size)
        underlined(p, "heavy underlined text", font, size, val)
for size in (9, 30):
    p = doc.add_paragraph()
    underlined(p, f"Aptos {size} thick underline", "Aptos", size, "thick")
doc.save(pathlib.Path(__file__).with_name("input.docx"))
