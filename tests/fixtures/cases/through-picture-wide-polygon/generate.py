"""Generate through-picture-wide-polygon: a picture whose edited wrap polygon reaches past it

A page-anchored tight or through picture whose rectangular wrap polygon is
wider than the picture leaves room by its polygon, not its frame: lines step
through its band by whole lines in every column the polygon leaves no room in,
though the picture itself leaves room. No text distances, 150pt tall, 13pt
lines. One section per case, each on its own page.

One 467.7pt column, the picture 388pt wide at the left margin, 120pt down
(its frame leaves 79.7pt), Line 4 holds it:
1. wrapThrough, polygon 1.25 times as wide (485pt): Line 5 steps 13 lines,
   to 277.7pt

Two 216.2pt columns, the picture 140pt wide at the left margin, 300pt down
(its frame leaves 76.2pt in column one), a long paragraph runs on into column
two:
2. wrapThrough, polygon 1.5 times as wide (210pt): column one steps past it,
   column two runs on
3. wrapThrough, polygon 3.33 times as wide (466.7pt, to 1.1pt short of
   column two's right edge): both columns step past it
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
WORDS = "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu".split()


def sect(cols):
    c = f'<w:cols w:num="{cols}" w:space="708"/>' if cols > 1 else ""
    return (
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
        f"{c}</w:sectPr>"
    )


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



def polygon(right):
    return (
        f'<wp:wrapPolygon edited="1"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/>'
        f'<wp:lineTo x="{right}" y="21600"/><wp:lineTo x="{right}" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>'
    )


def picture(wrap, w_pt, right, pid, top_emu):
    cx, cy = int(w_pt * 12700), 150 * 12700
    wrap_xml = f'<wp:wrap{wrap} wrapText="bothSides">{polygon(right)}</wp:wrap{wrap}>'
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        f'relativeHeight="{pid}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{top_emu}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap_xml}'
        f'<wp:docPr id="{pid}" name="Picture {pid}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{pid}" name="picture.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        "</wp:anchor></w:drawing></w:r>"
    )


def para(runs, sect_xml=""):
    return f'<w:p><w:pPr><w:spacing w:line="260" w:lineRule="exact"/>{sect_xml}</w:pPr>{runs}</w:p>'


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


CASES = (
    (1, "Through", 388, 27000), (2, "Through", 140, 32400), (2, "Through", 140, 72000),
)
LONG = " ".join(f"{WORDS[i % len(WORDS)]}{i}" for i in range(380))
body = []
for n, (cols, wrap, width, right) in enumerate(CASES, 1):
    s = "" if n == len(CASES) else sect(cols)
    if cols == 1:
        body += [para(run(f"Case {n} line {i}")) for i in range(1, 4)]
        body.append(para(picture(wrap, width, right, n, 1524000) + run("Line 4 holds the picture")))
        body += [para(run(f"Line {i}"), s if i == 20 else "") for i in range(5, 21)]
    else:
        body.append(para(picture(wrap, width, right, n, 3810000) + run(f"Case {n} line 1 holds the picture")))
        body.append(para(run(LONG), s))

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body) + sect(2) + "</w:body></w:document>"
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
