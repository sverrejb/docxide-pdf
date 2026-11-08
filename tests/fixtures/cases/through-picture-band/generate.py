"""Generate through-picture-band: a page-anchored tight or through picture with no room beside it

A picture anchored to the page with tight or through wrapping (rectangular wrap
polygon) and no usable strip beside it keeps the text out of its whole band, as
square wrapping does, but a line reaching into it steps down by whole line
heights until it is clear, where square wrapping starts it right at the
picture's bottom. The anchor paragraph's own line steps too when it reaches in.

One section per case, each on its own page, exact line heights, the picture
150pt tall at 120pt from the page top (its band ends 270pt down):
1. wrapThrough, 500pt wide, 13pt lines: Line 5 steps 13 lines, to 277.7pt
2. wrapTight, 500pt wide, 15pt lines: Line 5 steps 11 lines, to 281.7pt
3. wrapThrough, 500pt wide, 13pt lines, a long paragraph whose first line ends
   above the picture: its second line steps 13 lines, to 277.7pt
4. As 3 with wrapSquare: the second line starts at 270pt
5. wrapTight, 467.7pt (column-wide), its wrap polygon around the left half
   only, at 121.7pt: the lines beside it sit right of the polygon
"""

import pathlib
import struct
import zipfile
import zlib

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)
SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr>"
)
POLYGON = (
    '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/>'
    '<wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>'
)
HALF = (
    '<wp:wrapPolygon edited="1"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/>'
    '<wp:lineTo x="10800" y="21600"/><wp:lineTo x="10800" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>'
)
WRAPS = {
    "through": f'<wp:wrapThrough wrapText="bothSides">{POLYGON}</wp:wrapThrough>',
    "tight": f'<wp:wrapTight wrapText="bothSides">{POLYGON}</wp:wrapTight>',
    "square": '<wp:wrapSquare wrapText="bothSides"/>',
    "tight-half": f'<wp:wrapTight wrapText="bothSides">{HALF}</wp:wrapTight>',
}
WORDS = "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu".split()


def placeholder_png(w=200, h=60):
    """A grey checkerboard, drawn here."""
    rows = b""
    for y in range(h):
        rows += b"\x00" + b"".join(
            bytes((v, v, v)) for v in (60 if (x // 10 + y // 10) % 2 else 150 for x in range(w))
        )

    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def picture(wrap, w_pt, h_pt, pid, top_emu=1524000):
    cx, cy = int(w_pt * 12700), int(h_pt * 12700)
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{pid}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{top_emu}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{WRAPS[wrap]}'
        f'<wp:docPr id="{pid}" name="Picture {pid}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{pid}" name="picture.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        "</wp:anchor></w:drawing></w:r>"
    )


def para(runs, line, sect=False):
    return f'<w:p><w:pPr><w:spacing w:line="{line}" w:lineRule="exact"/>{SECT if sect else ""}</w:pPr>{runs}</w:p>'


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


CASES = (
    ("through", 500, 260, 1524000), ("tight", 500, 300, 1524000),
    ("through", 500, 260, None), ("square", 500, 260, None), ("tight-half", 467.7, 260, 1545590),
)
LONG = " ".join(f"{WORDS[i % len(WORDS)]}{i}" for i in range(200))
body = []
for n, (wrap, width, line, top) in enumerate(CASES, 1):
    body += [para(run(f"Line {i}"), line) for i in range(1, 4)]
    if top is None:
        body.append(para(picture(wrap, width, 150, n) + run(LONG), line, sect=n < len(CASES)))
        continue
    body.append(para(picture(wrap, width, 150, n, top) + run("Line 4 holds the picture"), line))
    body += [para(run(f"Line {i}"), line, sect=i == 20 and n < len(CASES)) for i in range(5, 21)]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body) + SECT + "</w:body></w:document>"
)

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)

settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:defaultTabStop w:val="708"/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
    "</w:settings>"
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
    "</Types>"
)

rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
package_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>'
)
document_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/>'
    f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/picture.png"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/media/picture.png", placeholder_png())
print(f"wrote {out}")
