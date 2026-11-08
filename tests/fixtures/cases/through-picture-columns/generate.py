"""Generate through-picture-columns: page-anchored pictures with no room beside them in two columns

Two columns 216.2pt wide, exact 13pt lines. Line 1 holds a 150pt tall picture
anchored to the page; a long paragraph after it fills column one and runs on in
column two. A picture keeps lines out of its band only in the columns it leaves
no room in, and a paragraph continuing in a column starts below the band too
(stepping by whole lines for through):
1. wrapThrough, 500pt wide at the page top: Line 1 steps below it, and the
   paragraph's lines in column two start below it, stepped by lines from the
   column top
2. wrapSquare, column one wide, 300pt down: column one's lines clear it,
   column two's run on with no gap
3. As 2 with wrapTopAndBottom
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
    '<w:cols w:num="2" w:space="708"/></w:sectPr>'
)
POLYGON = (
    '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/>'
    '<wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>'
)
WRAPS = {
    "through": f'<wp:wrapThrough wrapText="bothSides">{POLYGON}</wp:wrapThrough>',
    "square": '<wp:wrapSquare wrapText="bothSides"/>',
    "topAndBottom": "<wp:wrapTopAndBottom/>",
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


COL_W = (11906 - 1418 - 1134 - 708) / 2 / 20
CASES = (
    ("through", 500, 720090),
    ("square", COL_W, 3810000), ("topAndBottom", COL_W, 3810000),
)
LONG = " ".join(f"{WORDS[i % len(WORDS)]}{i}" for i in range(380))
body = []
for n, (wrap, width, top) in enumerate(CASES, 1):
    body.append(para(picture(wrap, width, 150, n, top) + run("Line 1 holds the picture"), 260))
    body.append(para(run(LONG), 260, sect=n < len(CASES)))

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
