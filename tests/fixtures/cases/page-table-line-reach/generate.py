"""Generate page-table-line-reach: lines reaching into a page-anchored floating table

A full-width floating table is placed against the page (tblpY from the page
top). Word moves every line whose box reaches the table top below the table,
including the anchor's own line and exact contact with the line bottom.

One page per variant (Intro, the table, then lines L1-L8, each 13pt exact
unless noted; Intro starts at 72pt, so L1 spans 85-98, L2 98-111, L3 111-124):
- table top 3pt into L1, the anchor paragraph
- table top exactly at L2's bottom
- single line spacing, table top 3pt into L3
- L3 has a wrapNone text box anchored 0pt below its paragraph: the box stays
  where L3 stood when the line moves below the table
- L3 is an inline canvas holding a rectangle and connector: both move

Spacing and inline-drawing reach (own space after counts before compat 15,
covered by page-table-line-reach-compat14):
- 24pt before carries L1 into the table, keeping its 24pt below the table
- L1 6pt after, L2 12pt before: L2 moves and starts 6pt below the table
- a text-anchored table 40pt below L1: L2 6pt after and L3 12pt before put
  L3 6pt below the table
- an 80pt inline text box or picture moves with its line
- an empty paragraph with 24pt before moves
"""

import pathlib
import struct
import zipfile
import zlib

W = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"'
)
EMU = 12700


def tw(pt):
    return int(round(pt * 20))


def para(text, rule="exact", page_break=False, drawing="", before=0, after=0):
    line = 'w:line="260" w:lineRule="exact"' if rule == "exact" else 'w:line="240" w:lineRule="auto"'
    pb = "<w:pageBreakBefore/>" if page_break else ""
    return (
        f'<w:p><w:pPr>{pb}<w:spacing w:before="{tw(before)}" w:after="{tw(after)}" {line}/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>{drawing}</w:p>'
    )


def sp_pr(w, h, fill):
    return (
        f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
    )


def text_box(w, h, text):
    return (
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp><wps:cNvSpPr txBox="1"/>{sp_pr(w, h, "FFD8D8")}'
        f'<wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
        '<wps:bodyPr wrap="square" lIns="45720" tIns="0" rIns="45720" bIns="0"><a:noAutofit/></wps:bodyPr>'
        '</wps:wsp></a:graphicData></a:graphic>'
    )


def canvas_with_line(w, h):
    half = w * EMU // 2
    return (
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">'
        f'<wpc:wpc><wpc:bg/><wpc:whole/><wps:wsp><wps:cNvPr id="20" name="Rect"/><wps:cNvSpPr/>'
        f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{half}" cy="{h * EMU}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="6080D0"/></a:solidFill>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>'
        f'<wps:wsp><wps:cNvPr id="21" name="Line"/><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="{half}" y="0"/>'
        f'<a:ext cx="{half}" cy="{h * EMU}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        '<a:ln w="19050"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>'
        '</wpc:wpc></a:graphicData></a:graphic>'
    )


def inline(graphic, w, h, n):
    return (
        f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{w * EMU}" cy="{h * EMU}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{n}" name="Drawing {n}"/>{graphic}'
        "</wp:inline></w:drawing></w:r>"
    )


def picture(w, h, n):
    return (
        f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{w * EMU}" cy="{h * EMU}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{n}" name="Picture {n}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr>'
        f'<pic:cNvPr id="{n}" name="square.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill>'
        '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId3"/>'
        '<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )


def square_png(size=8, rgb=(90, 140, 200)):
    raw = b"".join(b"\x00" + bytes(rgb) * size for _ in range(size))

    def chunk(kind, data):
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))

    header = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b"")


def anchored(graphic, w, h, n):
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="5" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{380 * EMU}</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{w * EMU}" cy="{h * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{n}" name="Drawing {n}"/>{graphic}</wp:anchor></w:drawing></w:r>'
    )


def table(y, overlap="never", h_anchor="page", width=595.3, v_anchor="page"):
    x = 'w:tblpXSpec="center"' if h_anchor == "margin" else 'w:tblpX="0"'
    overlap_xml = f'<w:tblOverlap w:val="{overlap}"/>' if overlap else ""
    return (
        f'<w:tbl><w:tblPr><w:tblpPr w:leftFromText="0" w:rightFromText="0" w:vertAnchor="{v_anchor}" '
        f'w:horzAnchor="{h_anchor}" {x} w:tblpY="{tw(y)}"/>{overlap_xml}'
        f'<w:tblW w:w="{tw(width)}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{tw(width)}"/></w:tblGrid>'
        f'<w:tr><w:trPr><w:trHeight w:val="800" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{tw(width)}" w:type="dxa"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tcPr>'
        f'<w:p><w:r><w:t>Floating table, top at {y}pt</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    )


variants = [
    (88, "exact", {}),
    (111, "exact", {}),
    (112.95, "auto", {}),
]
body = []
for n, (y, rule, opts) in enumerate(variants):
    label = " ".join(f"{k}={v}" for k, v in opts.items() if k != "width")
    body.append(para(f"Intro {n + 1}: top {y}pt, {rule} {label}".rstrip(), rule, n > 0))
    body.append(table(y, **opts))
    body += [para(f"L{k}", rule) for k in range(1, 9)]

drawings = [
    ("an anchored wrapNone box", "L3 ", anchored(text_box(90, 16, "Anchored box"), 90, 16, 3)),
    ("an inline canvas with a line", "", inline(canvas_with_line(60, 16), 60, 16, 4)),
]
for n, (label, text, drawing) in enumerate(drawings, len(variants) + 1):
    body.append(para(f"Intro {n}: top 113pt, L3 is {label}", "auto", True))
    body.append(table(113))
    body += [para(text, "auto", drawing=drawing) if k == 3 else para(f"L{k}", "auto") for k in range(1, 9)]

spacing_pages = [
    ("24pt before, top 114pt", "exact", table(114), [para("L1 before 24", before=24)]),
    ("L1 6pt after, L2 12pt before, top 114pt", "exact", table(114),
     [para("L1 after 6", after=6), para("L2 before 12", before=12)]),
    ("text-anchored 40pt, L2 6pt after, L3 12pt before", "exact", table(40, v_anchor="text"),
     [para("L1"), para("L2 after 6", after=6), para("L3 before 12", before=12)]),
    ("80pt inline box, top 120pt", "auto", table(120), [para("", "auto", drawing=inline(text_box(90, 80, "Tall box"), 90, 80, 5))]),
    ("80pt inline picture, top 120pt", "auto", table(120), [para("Picture ", "auto", drawing=picture(80, 80, 7))]),
    ("empty paragraph 24pt before, top 100pt", "exact", table(100), [para("", before=24)]),
]
for n, (label, rule, tbl, lines) in enumerate(spacing_pages, len(variants) + len(drawings) + 1):
    body.append(para(f"Intro {n}: {label}", rule, True))
    body.append(tbl)
    body += lines + [para(f"L{k}", rule) for k in range(len(lines) + 1, 9)]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
    'w:left="1440" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>'
)
styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)
settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/>'
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
    f'<Relationship Id="rId3" Type="{rel}/image" Target="media/square.png"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/media/square.png", square_png())
print(f"wrote {out}")
