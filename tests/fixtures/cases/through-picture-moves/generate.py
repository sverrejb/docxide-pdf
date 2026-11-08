"""Generate through-picture-moves: paragraphs moved whole below a page-anchored picture at the page top

A 500pt wide, 150pt tall picture at the top margin (56.7pt) leaves no room
beside it. A paragraph that moves whole to the next column or page starts below
the picture there: by whole lines from the column top with wrapThrough, at the
picture's bottom with wrapSquare. One section per case, each on its own page.

Two columns, Line 1 holds the picture, fillers end column one one line short:
1. a keepNext heading whose successor can't follow moves with it: 212.7pt
2. a three-line paragraph with widow control moves: 212.7pt
3. a two-line keepLines paragraph with 17pt lines and wrapSquare moves to
   column two: 206.7pt
7. a two-line keepLines paragraph moves to column two, Line 1 also holding a
   page-wide through picture from 710pt to 780pt that the paragraph's second
   line would step through in column one: 212.7pt

One column, the picture on page two:
4. a keepLines paragraph at the foot of page one holding the picture moves to
   page two with it: page one keeps its lines at the top, page two's start at
   212.7pt
5. as 4, the picture held by the paragraph after it
6. as 5, a long paragraph without keepLines runs on to page two: its lines
   there start at 212.7pt
"""

import math
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
POLYGON = (
    '<wp:wrapPolygon edited="0"><wp:start x="0" y="0"/><wp:lineTo x="0" y="21600"/>'
    '<wp:lineTo x="21600" y="21600"/><wp:lineTo x="21600" y="0"/><wp:lineTo x="0" y="0"/></wp:wrapPolygon>'
)
WRAPS = {
    "through": f'<wp:wrapThrough wrapText="bothSides">{POLYGON}</wp:wrapThrough>',
    "square": '<wp:wrapSquare wrapText="bothSides"/>',
}
WORDS = "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu".split()
KEEP_LINES, KEEP_NEXT, WIDOW = "<w:keepLines/>", "<w:keepNext/>", "<w:widowControl/>"
TOP, BOTTOM = 56.7, 841.9 - 56.7


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


def picture(wrap, pid, y=56.7, h=150):
    cx, cy = 500 * 12700, round(h * 12700)
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{pid}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{round(y * 12700)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{WRAPS[wrap]}'
        f'<wp:docPr id="{pid}" name="Picture {pid}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{pid}" name="picture.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        "</wp:anchor></w:drawing></w:r>"
    )


def para(runs, line, sect_xml="", props=""):
    return (f'<w:p><w:pPr>{props}<w:spacing w:line="{line}" w:lineRule="exact"/>{sect_xml}</w:pPr>'
            f"{runs}</w:p>")


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def words(n):
    return " ".join(f"{WORDS[i % len(WORDS)]}{i}" for i in range(n))


CASES = (
    ("keepNext", "through", 260), ("widow", "through", 260), ("keepLines", "square", 340),
    ("page-own", "through", 260), ("page-next", "through", 260), ("page-split", "through", 260),
    ("keepLines-low", "through", 260),
)
body = []
for n, (case, wrap, line) in enumerate(CASES, 1):
    last = n == len(CASES)
    lh = line / 20
    if not case.startswith("page"):
        s = sect(2)
        first_top = TOP + math.ceil(150 / lh) * lh if wrap == "through" else TOP + 150
        fits = int((BOTTOM - first_top) / lh)
        pic = picture(wrap, n)
        if case == "keepLines-low":
            fits = int((710 - first_top) / lh)
            pic += picture(wrap, 100 + n, 710, 70)
        body.append(para(pic + run(f"Case {n} line 1 holds the picture"), line))
        body += [para(run(f"Filler {i}"), line) for i in range(2, fits)]
        if case.startswith("keepLines"):
            body.append(para(run("Kept lines " + words(9)), line, props=KEEP_LINES))
        elif case == "keepNext":
            body.append(para(run("Kept with next"), line, props=KEEP_NEXT))
            body.append(para(run("Next " + words(16)), line, props=WIDOW))
        else:
            body.append(para(run("Widow control " + words(16)), line, props=WIDOW))
    else:
        s = sect(1)
        fits = int((BOTTOM - TOP) / lh)
        body.append(para(run(f"Case {n} page one"), line))
        body += [para(run(f"Filler {i}"), line) for i in range(2, fits - 6 if case == "page-split" else fits)]
        pic = picture(wrap, n)
        if case == "page-own":
            body.append(para(pic + run("Kept lines " + words(22)), line, props=KEEP_LINES))
        else:
            if case == "page-next":
                body.append(para(run("Kept lines " + words(22)), line, props=KEEP_LINES))
            else:
                body.append(para(run("Split " + words(160)), line))
            pic = picture(wrap, n)
            body.append(para(pic + run("After 0 holds the picture"), line))
    body += [para(run(f"After {i}"), line, sect_xml="" if (last or i < 3) else s) for i in range(1, 4)]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body) + s + "</w:body></w:document>"
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
