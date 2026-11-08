"""Generate identical-parts-and-fonts: byte-identical media parts, one face in several styles

A document can store one picture in several media parts with identical bytes,
for instance after copying a picture between documents. And a face without
bold or italic variants, such as Symbol or Wingdings, is often used under
several styles: the bullets of a bold list item, a bold or italic Greek letter.

Tests:
- Two media parts, word/media/image1.png and word/media/image2.png, with
  byte-identical data, each drawn once at the same size
- Symbol bullets in a list whose items are regular, bold and italic
- Symbol text under regular, bold, italic and bold italic
- Wingdings text under regular, bold and italic
"""

import pathlib
import struct
import zipfile
import zlib

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = (
    f'xmlns:w="{W_NS}" xmlns:r="{R_NS}" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)
EMU = 12700


def png(width, height):
    """Four coloured quadrants with grain and a dark diagonal."""
    quadrants = [[(30, 90, 200), (240, 150, 40)], [(60, 170, 90), (200, 60, 90)]]
    seed = 12345
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            seed = (seed * 1103515245 + 12345) & 0x7FFFFFFF
            grain = (seed >> 16) % 25 - 12
            if abs(x * height - y * width) < width * 3:
                row += bytes((25, 25, 25))
            else:
                row += bytes(min(255, max(0, c + grain)) for c in quadrants[y * 2 // height][x * 2 // width])
        rows.append(bytes(row))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(b"".join(rows), 9)) + chunk(b"IEND", b""))


def inline(doc_id, rel_id, w_pt, h_pt):
    cx, cy = int(w_pt * EMU), int(h_pt * EMU)
    return (
        '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{doc_id}" name="Picture {doc_id}"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{doc_id}" name="image{doc_id}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="{rel_id}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        '</wp:inline></w:drawing></w:r>'
    )


def run(text, font=None, bold=False, italic=False):
    rpr = ""
    if font:
        rpr += f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    if bold:
        rpr += "<w:b/>"
    if italic:
        rpr += "<w:i/>"
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content, before=0):
    return f'<w:p><w:pPr><w:spacing w:before="{before * 20}"/></w:pPr>{content}</w:p>'


def bullet(text, bold=False, italic=False):
    mark = ("<w:b/>" if bold else "") + ("<w:i/>" if italic else "")
    ppr = f'<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:ind w:left="720" w:hanging="360"/>'
    return (f'<w:p><w:pPr>{ppr}<w:rPr>{mark}</w:rPr></w:pPr>'
            f'{run(text, bold=bold, italic=italic)}</w:p>')


def styled_lines(font, text, styles):
    return [
        para(run(f"{label}: ") + run(text, font, bold, italic), before=4)
        for label, bold, italic in styles
    ]


ALL_STYLES = [("Regular", False, False), ("Bold", True, False), ("Italic", False, True),
              ("Bold italic", True, True)]

body = [
    para(run("Two media parts with identical bytes, drawn the same way")),
    para(inline(1, "rIdImg1", 160, 80), before=6),
    para(inline(2, "rIdImg2", 160, 80), before=6),
    para(run("A list with Symbol bullets"), before=12),
    bullet("Regular list item"),
    bullet("Bold list item", bold=True),
    bullet("Italic list item", italic=True),
    para(run("Symbol text in several styles"), before=12),
    *styled_lines("Symbol", "abgdeqlmpsW", ALL_STYLES),
    para(run("Wingdings text in several styles"), before=12),
    *styled_lines("Wingdings", "JKL!\"#$%", ALL_STYLES[:3]),
]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="709" w:footer="709" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)

numbering = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="{W_NS}">'
    '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>'
    '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/>'
    '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>'
    '<w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl></w:abstractNum>'
    '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'
)

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W_NS}"><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)

settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W_NS}">'
    '<w:defaultTabStop w:val="708"/><w:compat><w:compatSetting w:name="compatibilityMode" '
    'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
)

content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    "</Types>"
)

REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
package_rels = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELS_NS}">'
    f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>'
)
document_rels = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELS_NS}">'
    f'<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId2" Type="{REL}/settings" Target="settings.xml"/>'
    f'<Relationship Id="rId3" Type="{REL}/numbering" Target="numbering.xml"/>'
    f'<Relationship Id="rIdImg1" Type="{REL}/image" Target="media/image1.png"/>'
    f'<Relationship Id="rIdImg2" Type="{REL}/image" Target="media/image2.png"/></Relationships>'
)

picture = png(320, 160)
out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/numbering.xml", numbering)
    z.writestr("word/media/image1.png", picture)
    z.writestr("word/media/image2.png", picture)
print(f"wrote {out}")
