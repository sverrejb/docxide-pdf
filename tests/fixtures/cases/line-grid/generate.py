"""Word draws baselines on a 0.25pt grid and stretches the page by one cell.

Each page sets one font, size and line rule over 40-odd one-line paragraphs,
so the lines land on every phase of the 0.25pt grid:

1. Arial 11pt single, A4, top margin 70.85pt: a 12.65pt line rounds up to
   12.75, the baseline keeps its descent above the rounded line foot.
2. Times New Roman 9pt single: a 10.35pt line rounds down to 10.25 and loses
   half the difference above the baseline.
3. Arial 11pt exactly 13.1pt.
4. Arial 11pt single on a 1000pt tall page, footer lines near its foot: the
   writer's stretch moves the lowest lines 0.25pt down.
5. Times New Roman 14pt at 1.5 lines: the 24.15pt line rounds up to 24.25
   and its 8.05pt leading down to 8, so the text box grows and the baseline
   drops a cell below the rounded ascent.
6. Arial 14pt at 1.25 lines: the text box rounds down, but with leading below
   the baseline keeps its rounded ascent rather than losing half the
   difference as at single spacing.
7. Times New Roman 10pt at least 13.85pt: the leading goes above the text,
   the 13.85pt line rounds down to 13.75 and its 2.35pt leading to 2.25, so
   the text box keeps 11.5 cells and the baseline loses nothing.
8. Numbered Arial 11pt single, top margin 70pt: each number sits on its
   line's rounded and stretched baseline.
9. Two Arial 11pt paragraphs bottom-aligned (w:vAlign), bottom margin 72pt:
   the alignment moves the lines before they are rounded and stretched.
10. Arial 11pt lines with inline pictures: a 1pt icon leaves the line to its
    text, a 20pt picture lowers the baseline and rounds it itself.
Every page has a two-line footer (Arial 8pt).
"""

import struct
import zipfile
import zlib
from pathlib import Path

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
     'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
     'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
     'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>
<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
<Default Extension="png" ContentType="image/png"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/dot.png"/>
</Relationships>"""

SETTINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {W}><w:defaultTabStop w:val="720"/><w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat></w:settings>"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/>
</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

NUMBERING = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering {W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>
<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"""

NUMBERED = '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'


def png_dot():
    def chunk(kind, data):
        return (struct.pack(">I", len(data)) + kind + data
                + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF))
    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 0, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"\x00\x00")) + chunk(b"IEND", b""))


def picture(n, points):
    emu = int(points * 12700)
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{emu}" cy="{emu}"/><wp:docPr id="{n}" name="Picture {n}"/>'
            f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="Picture {n}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu}" cy="{emu}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            f'</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


WORDS = ("grid baseline device cell rounds every line top while the layout keeps "
         "exact positions down the page and the writer stretches").split()


def run_props(font, half_points):
    return (f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/>')


def text_run(text, rpr):
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def paragraph(text, font, half_points, spacing="", sect="", runs=None):
    rpr = run_props(font, half_points)
    return (f'<w:p><w:pPr>{spacing}{sect}<w:rPr>{rpr}</w:rPr></w:pPr>'
            f'{runs or text_run(text, rpr)}</w:p>')


def sect_pr(w, h, top, bottom=1134, v_align=""):
    v_align = f'<w:vAlign w:val="{v_align}"/>' if v_align else ""
    return (f'<w:sectPr><w:footerReference w:type="default" r:id="rId3"/>'
            f'<w:pgSz w:w="{w}" w:h="{h}"/>'
            f'<w:pgMar w:top="{top}" w:right="1417" w:bottom="{bottom}" w:left="1417" '
            f'w:header="708" w:footer="567" w:gutter="0"/><w:cols w:space="720"/>{v_align}</w:sectPr>')


def line_text(page, i):
    words = [WORDS[(i * 3 + k + page) % len(WORDS)] for k in range(6)]
    return f"Page {page} line {i + 1:02d} " + " ".join(words)


PAGES = [
    # (font, half-points, spacing element, page w, page h, top margin, lines)
    ("Arial", 22, "", 11906, 16838, 1417, 50),
    ("Times New Roman", 18, "", 11906, 16838, 1400, 60),
    ("Arial", 22, '<w:spacing w:line="262" w:lineRule="exact"/>', 11906, 16838, 1417, 50),
    ("Arial", 22, "", 11906, 20000, 1417, 64),
    ("Times New Roman", 28, '<w:spacing w:line="360" w:lineRule="auto"/>', 11906, 16838, 1403, 28),
    ("Arial", 28, '<w:spacing w:line="300" w:lineRule="auto"/>', 11906, 16838, 1417, 34),
    ("Times New Roman", 20, '<w:spacing w:line="277" w:lineRule="atLeast"/>', 11906, 16838, 1403, 50),
]

body = []
for p, (font, hp, spacing, w, h, top, n) in enumerate(PAGES, 1):
    sect = sect_pr(w, h, top)
    for i in range(n):
        body.append(paragraph(line_text(p, i), font, hp, spacing, sect if i == n - 1 else ""))

for i in range(50):
    sect = sect_pr(11906, 16838, 1400) if i == 49 else ""
    body.append(paragraph(line_text(8, i), "Arial", 22, NUMBERED, sect))

body.append(paragraph(line_text(9, 0), "Arial", 22))
body.append(paragraph(line_text(9, 1), "Arial", 22, "",
                      sect_pr(11906, 16838, 1400, 1440, "bottom")))

rpr = run_props("Arial", 22)
for i in (i for i in range(50) if i % 3):
    text = line_text(10, i)
    runs = (text_run(text[:13], rpr) + picture(i, 1 if i % 3 == 1 else 20)
            + text_run(text[13:], rpr))
    body.append(paragraph(text, "Arial", 22, runs=runs))

body.append(sect_pr(11906, 16838, 1417))

document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document {W}><w:body>{"".join(body)}</w:body></w:document>')

footer = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr {W}>'
          + paragraph("Footer line one: grid baseline probe", "Arial", 16)
          + paragraph("Footer line two: the lowest text on the page", "Arial", 16)
          + '</w:ftr>')

out = Path(__file__).with_name("input.docx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for name, data in [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", RELS),
        ("word/_rels/document.xml.rels", DOC_RELS),
        ("word/document.xml", document),
        ("word/styles.xml", STYLES),
        ("word/settings.xml", SETTINGS),
        ("word/footer1.xml", footer),
        ("word/numbering.xml", NUMBERING),
        ("word/media/dot.png", png_dot()),
    ]:
        info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, data)
print(f"wrote {out}")
