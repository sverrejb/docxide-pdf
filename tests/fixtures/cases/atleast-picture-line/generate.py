"""Generate atleast-picture-line: a picture or inline text box on an at-least line

An at-least line spacing is a minimum, not a multiple: a picture taller than
it keeps its own height (no leading below it, whatever the paragraph mark's
size), a taller minimum adds its surplus above the picture, and a picture
shorter than the line sits at the line's foot. Word lays an inline text box
alone on its line out the same way.

Every case is a "Case" line, the test paragraph (spacing after 0) and an
"After" line, all 10pt Arial at single spacing.

Tests (case numbers kept from the fuller example set):
- 26.2pt picture under 15pt (mark 10pt), 24pt (mark 16.5pt) and 40pt
  (mark 10pt) minimums
- 26.2pt inline text box under 15 and 40pt (textAlignment center) and 24pt
  (mark 16.5pt) minimums
- 8pt picture and 8pt inline text box under a 15pt minimum
- 26.2pt picture between words of a text line under 15 and 40pt minimums
- 26.2pt picture with a 3pt shadow below it (effectExtent b) under a 40pt
  minimum: the picture's foot sits the effect extent above the line's foot

Section 2 (header, top margin 140pt):
- Header and a floating text box each hold text beside a 26.2pt picture
  under a 40pt minimum, then a 26.2pt picture alone under it
- An inline canvas (picture, text box, line) alone under a 15pt minimum
- Page fit, one case per page after an exact filler leaving R points: text
  alone under 24pt (R 23.5 moves), inline text box and canvas under 40pt
  (R 30 move)
- An inline canvas under a 40pt minimum beside a floating table

Section 3 (70pt page body, its own header):
- Header: a 40pt picture alone under a 40pt minimum, a 10.5pt one at single
  spacing
- Eight broken lines under 24pt without widow control: two per page
- Eight broken lines under 20pt, a footnote on line 5: three lines on a
  page, two on the one carrying the footnote
- Eight broken lines under 24pt, line 4 a 40pt picture alone: two per page
- A 50pt filler, then eight broken lines under 24pt: none fit the 20pt
  left, so they move and take two a page

Section 4 (10pt Arial under a 40pt minimum, its own header and footer):
- Body: text raised 1pt and text lowered 1pt alone on their lines
- Header, body, floating text box and table cell: a raised and a lowered
  run between plain text
- Footer: a 26.2pt picture alone under a 40pt minimum
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
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
)
EMU = 12700
W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'

SPACING = {
    "at least 15pt": 'w:line="300" w:lineRule="atLeast"',
    "at least 20pt": 'w:line="400" w:lineRule="atLeast"',
    "at least 24pt": 'w:line="480" w:lineRule="atLeast"',
    "at least 40pt": 'w:line="800" w:lineRule="atLeast"',
    "single": 'w:line="240" w:lineRule="auto"',
    "1.15 lines": 'w:line="276" w:lineRule="auto"',
}


def png(width, height):
    rows = b"".join(b"\x00" + bytes((90, 110, 160)) * width for _ in range(height))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(rows))
            + chunk(b"IEND", b""))


def pic_xml(n, width, height, x=0, shadow=0):
    cx, cy = int(width * EMU), int(height * EMU)
    effects = (f'<a:effectLst><a:outerShdw dist="{shadow * EMU}" dir="5400000" algn="t" rotWithShape="0">'
               '<a:srgbClr val="000000"/></a:outerShdw></a:effectLst>') if shadow else ""
    return (
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="block.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="{int(x * EMU)}" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{effects}</pic:spPr></pic:pic>'
    )


def picture(n, height, width=120, shadow=0):
    cx, cy = int(width * EMU), int(height * EMU)
    return (
        f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        + (f'<wp:effectExtent l="0" t="0" r="0" b="{shadow * EMU}"/>' if shadow else "")
        + f'<wp:docPr id="{n}" name="Picture {n}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        + pic_xml(n, width, height, shadow=shadow)
        + "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )


def box_wsp(width, height, label, x=0, fill="D9DEE8", left_inset=144000):
    cx, cy = int(width * EMU), int(height * EMU)
    return (
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="{int(x * EMU)}" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        '<a:ln w="6350"><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>'
        f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>{label}</w:p>'
        f'</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="{left_inset}" '
        'tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp>"
    )


def text_box(n, height):
    cx, cy = 300 * EMU, int(height * EMU)
    label = f'<w:r><w:rPr><w:b/></w:rPr><w:t>Box {n}</w:t></w:r>' if height > 15 else ""
    return (
        '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:docPr id="{n}" name="Text Box {n}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        + box_wsp(300, height, label)
        + "</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice></mc:AlternateContent></w:r>"
    )


def canvas(n, height):
    line = (
        '<wps:wsp><wps:cNvCnPr/><wps:spPr>'
        f'<a:xfrm><a:off x="{150 * EMU}" y="0"/><a:ext cx="{40 * EMU}" cy="{int(height * EMU)}"/></a:xfrm>'
        '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></wps:spPr>'
        "<wps:bodyPr/></wps:wsp>"
    )
    label = '<w:r><w:rPr><w:b/></w:rPr><w:t>Tb</w:t></w:r>'
    kids = pic_xml(n * 10, 50, height) + box_wsp(60, height, label, x=70, fill="E8C8C8", left_inset=91440) + line
    return (
        '<w:r><mc:AlternateContent><mc:Choice Requires="wpc"><w:drawing>'
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{200 * EMU}" cy="{int(height * EMU)}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{n}" name="Canvas {n}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">'
        f"<wpc:wpc><wpc:bg/><wpc:whole/>{kids}</wpc:wpc>"
        "</a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice></mc:AlternateContent></w:r>"
    )


def floating_box(n, paras_xml, x=320, y=150, w=200, h=130):
    return (
        '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>'
        f'<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{n}" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{x * EMU}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{y * EMU}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{w * EMU}" cy="{h * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{n}" name="Float {n}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="6350"><a:solidFill>'
        '<a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>'
        f'{paras_xml}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" '
        'tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice></mc:AlternateContent></w:r>"
    )


def text(t):
    return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'


def para(runs, spacing="single", mark_sz=20, align=None, before_ppr="", exact_twips=None):
    text_alignment = f'<w:textAlignment w:val="{align}"/>' if align else ""
    line = f'w:line="{exact_twips}" w:lineRule="exact"' if exact_twips else SPACING[spacing]
    return (f'<w:p><w:pPr>{before_ppr}<w:spacing w:after="0" {line}/>{text_alignment}'
            f'<w:rPr><w:sz w:val="{mark_sz}"/></w:rPr></w:pPr>{runs}</w:p>')


CASES = [
    (1, "picture", 26.2, "at least 15pt", 20, None),
    (4, "picture", 26.2, "at least 24pt", 33, None),
    (5, "picture", 26.2, "at least 40pt", 20, None),
    (7, "text box", 26.2, "at least 15pt", 20, "center"),
    (10, "text box", 26.2, "at least 24pt", 33, None),
    (11, "text box", 26.2, "at least 40pt", 20, "center"),
    (13, "picture", 8, "at least 15pt", 20, None),
    (15, "text box", 8, "at least 15pt", 20, None),
    (17, "picture in text", 26.2, "at least 15pt", 20, None),
    (19, "picture in text", 26.2, "at least 40pt", 20, None),
    (24, "shadowed picture", 26.2, "at least 40pt", 20, None),
]

body = []
for n, kind, height, spacing, mark_sz, align in CASES:
    label = f"{kind} {height:g}pt, {spacing}, mark {mark_sz / 2:g}pt" + (f", textAlignment {align}" if align else "")
    body.append(para(text(f"Case {n}: {label}")))
    if kind == "picture":
        obj = picture(n, height)
    elif kind == "shadowed picture":
        obj = picture(n, height, shadow=3)
    elif kind == "text box":
        obj = text_box(n, height)
    else:
        obj = text("Text before ") + picture(n, height) + text(" and after")
    body.append(para(obj, spacing, mark_sz, align))
    body.append(para(text(f"After {n}")))

PAGE = '<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="{top}" w:right="1134" w:bottom="850" ' \
    'w:left="1134" w:header="425" w:footer="425" w:gutter="0"/>'
body.append(f'<w:p><w:pPr><w:sectPr>{PAGE.format(top=850)}</w:sectPr></w:pPr></w:p>')


def beside(tag, n):
    return (para(text(f"{tag} ") + picture(n, 26.2, width=60) + text(" after"), "at least 40pt")
            + para(text(f"Next {tag}")))


def alone(tag, n):
    return para(picture(n, 26.2, width=60), "at least 40pt") + para(text(f"Next {tag}"))


header = beside("Header", 101) + alone("header picture", 102)
boxed = beside("Box", 111) + alone("box picture", 112)
body.append(para(text("Section 2: header, text box and body") + floating_box(120, boxed)))
body.append(para(text("Canvas 132, at least 15pt")))
body.append(para(canvas(132, 26.2), "at least 15pt"))
body.append(para(text("After canvas 132")))

BODY_TWIPS = 16838 - 2800 - 850
FIT = [(143, 23.5, "text", "at least 24pt"), (144, 30, "text box", "at least 40pt"),
       (145, 30, "canvas", "at least 40pt")]
for n, room, kind, spacing in FIT:
    filler = int(round(BODY_TWIPS - room * 20))
    body.append(para(text(f"Fit {n}: {kind}, {spacing}, {room:g}pt left"), before_ppr="<w:pageBreakBefore/>",
                     exact_twips=filler))
    if kind == "text":
        runs = text("Only text")
    elif kind == "text box":
        runs = text_box(n, 26.2)
    else:
        runs = canvas(n, 26.2)
    body.append(para(runs, spacing, before_ppr='<w:widowControl w:val="0"/>'))
    body.append(para(text(f"Marker {n}")))

FLOAT_TABLE = (
    '<w:tbl><w:tblPr><w:tblpPr w:leftFromText="142" w:rightFromText="142" w:vertAnchor="text" '
    'w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="1"/><w:tblW w:w="2000" w:type="dxa"/>'
    '<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>'
    '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>'
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr>'
    f'{para(text("Floating table"))}</w:tc></w:tr></w:tbl>'
)
body.append(para(text("Canvas 151 beside a floating table, at least 40pt"), before_ppr="<w:pageBreakBefore/>"))
body.append(FLOAT_TABLE + para(canvas(151, 26.2), "at least 40pt"))
body.append(para(text("After canvas 151")))
body.append(f'<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/>'
            f'{PAGE.format(top=2800)}</w:sectPr></w:pPr></w:p>')

NO_WIDOWS = '<w:widowControl w:val="0"/>'
FOOTNOTE_REF = '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>'


def broken_lines(tag, spacing, line=None, page_break=True, count=8):
    line = line or (lambda k: text(f"{tag} line {k}"))
    runs = "<w:r><w:br/></w:r>".join(line(k) for k in range(1, count + 1))
    return para(runs, spacing, before_ppr=("<w:pageBreakBefore/>" if page_break else "") + NO_WIDOWS)


body.append(broken_lines("Split", "at least 24pt", page_break=False))
body.append(broken_lines("Note", "at least 20pt",
                         lambda k: text(f"Note line {k}") + (FOOTNOTE_REF if k == 5 else "")))
body.append(broken_lines("Pic", "at least 24pt",
                         lambda k: picture(301, 40, width=60) if k == 4 else text(f"Pic line {k}")))
body.append(para(text("Filler using 50pt"), before_ppr="<w:pageBreakBefore/>", exact_twips=1000))
body.append(broken_lines("Move", "at least 24pt", page_break=False))
body.append(para(text("Band filler using 50pt"), before_ppr="<w:pageBreakBefore/>", exact_twips=1000))
band = floating_box(302, para(text("Page-top band")), x=57, y=100, w=480, h=24).replace(
    "<wp:wrapNone/>", "<wp:wrapTopAndBottom/>")
body.append(broken_lines("Band move", "at least 24pt",
                         lambda k: (band if k == 1 else "") + text(f"Band move line {k}"), page_break=False))
body.append(para(text("Spaced filler using 50pt"), before_ppr="<w:pageBreakBefore/>",
                 exact_twips=1000).replace('w:after="0"', 'w:after="240"'))
spaced_band = floating_box(303, para(text("Spaced page-top band")), x=57, y=100, w=480, h=24).replace(
    "<wp:wrapNone/>", "<wp:wrapTopAndBottom/>")
body.append(broken_lines("Spaced move", "at least 24pt",
                         lambda k: (spaced_band if k == 1 else "") + text(f"Spaced move line {k}"),
                         page_break=False))
SMALL_PAGE = PAGE.replace('w:h="16838"', 'w:h="4250"')
body.append(f'<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHdr2"/>'
            f'{SMALL_PAGE.format(top=2000)}</w:sectPr></w:pPr></w:p>')
body.append(para(text("Picture clearance")))
body.append(para(text("Picture filler using 100pt"), before_ppr="<w:pageBreakBefore/>", exact_twips=2000))
picture_band = floating_box(304, para(text("Picture band")), x=57, y=130, w=480, h=24).replace(
    "<wp:wrapNone/>", "<wp:wrapTopAndBottom/>")
body.append(broken_lines("Picture move", "at least 24pt",
                         lambda k: (picture_band + picture(305, 40, width=60) if k == 1 else "")
                         + text(f"Picture move line {k}"), page_break=False))
PICTURE_PAGE = PAGE.replace('w:h="16838"', 'w:h="5250"')
body.append(f'<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHdr2"/>'
            f'{PICTURE_PAGE.format(top=2000)}</w:sectPr></w:pPr></w:p>')
body.append(para(text("Continuation filler using 50pt"), before_ppr="<w:pageBreakBefore/>", exact_twips=1000))
body.append(broken_lines("Continuation", "at least 24pt", page_break=False, count=3))
continuation_band = floating_box(306, para(text("Continuation band")), x=57, y=100, w=480, h=24).replace(
    "<wp:wrapNone/>", "<wp:wrapTopAndBottom/>")
body.append(para(continuation_band + text("After continuation")))
body.append(para(text("Hanging filler using 50pt"), before_ppr="<w:pageBreakBefore/>", exact_twips=1000))
hanging_box = floating_box(307, para(text("Hanging box")), x=57, y=24, w=480, h=12).replace(
    '<wp:positionV relativeFrom="page">', '<wp:positionV relativeFrom="paragraph">').replace(
    "<wp:wrapNone/>", "<wp:wrapTopAndBottom/>")
body.append(para(hanging_box + text("Hanging moved anchor"), "at least 24pt", before_ppr=NO_WIDOWS))
body.append(para(text("After hanging box")))
body.append(para(text("Pinned filler using 50pt"), before_ppr="<w:pageBreakBefore/>", exact_twips=1000))
pinned_band = floating_box(308, para(text("Pinned page band")), x=57, y=100, w=480, h=24).replace(
    "<wp:wrapNone/>", "<wp:wrapTopAndBottom/>")
pinned_box = floating_box(309, para(text("Pinned drawing")), x=350, y=0, w=120, h=12).replace(
    '<wp:positionV relativeFrom="page">', '<wp:positionV relativeFrom="paragraph">')
body.append(para(pinned_band + pinned_box + text("Pinned moved anchor"), "at least 24pt",
                 before_ppr=NO_WIDOWS))
body.append(para(text("After pinned anchor")))
body.append(f'<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rIdHdr2"/>'
            f'{SMALL_PAGE.format(top=2000)}</w:sectPr></w:pPr></w:p>')


def positioned(tag, half_points):
    rpr = f'<w:rPr><w:position w:val="{half_points}"/></w:rPr>' if half_points else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{tag}</w:t></w:r>'


def positioned_lines(where, after="", wholly=False):
    whole = (para(positioned(f"{where} raised 1pt", 2), "at least 40pt")
             + para(positioned(f"{where} lowered 1pt", -2), "at least 40pt")) if wholly else ""
    return (whole
            + para(text(f"{where} plain ") + positioned("raised 1pt", 2) + text(" mixed"), "at least 40pt")
            + para(text(f"{where} plain ") + positioned("lowered 1pt", -2) + text(" mixed"), "at least 40pt")
            + para(text(f"Next {where}") + after))


CELL_TABLE = (
    '<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblBorders>'
    '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>'
    '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/>'
    '</w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
    '<w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>'
    f'{positioned_lines("Cell")}</w:tc></w:tr></w:tbl>'
)
body.append(positioned_lines("Body", floating_box(401, positioned_lines("Box"), x=330, y=250, w=200, h=250),
                             wholly=True))
body.append(CELL_TABLE + para(text("After cell")))

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr3"/>'
    f'<w:footerReference w:type="default" r:id="rIdFtr"/>{PAGE.format(top=4800)}</w:sectPr>'
    "</w:body></w:document>"
)
header2 = (para(picture(201, 40, width=60), "at least 40pt") + para(picture(202, 10.5, width=60))
           + para(text("Next header 3")))
header3 = positioned_lines("Header")
footer = para(picture(501, 26.2, width=60), "at least 40pt") + para(text("Footer end"))
footnotes = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes {W}>'
    '<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
    '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p>'
    '</w:footnote><w:footnote w:id="1"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr>'
    '<w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Footnote on line 5</w:t></w:r></w:p></w:footnote>'
    "</w:footnotes>"
)

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    "</w:styles>"
)

settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" '
    'w:val="15"/></w:compat></w:settings>'
)

content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    '<Override PartName="/word/header3.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
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
    f'<Relationship Id="rIdHdr" Type="{rel}/header" Target="header1.xml"/>'
    f'<Relationship Id="rIdHdr2" Type="{rel}/header" Target="header2.xml"/>'
    f'<Relationship Id="rIdHdr3" Type="{rel}/header" Target="header3.xml"/>'
    f'<Relationship Id="rIdFtr" Type="{rel}/footer" Target="footer1.xml"/>'
    f'<Relationship Id="rIdFn" Type="{rel}/footnotes" Target="footnotes.xml"/>'
    f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/block.png"/></Relationships>'
)
header_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/block.png"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/header1.xml",
               f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>{header}</w:hdr>')
    z.writestr("word/_rels/header1.xml.rels", header_rels)
    z.writestr("word/header2.xml",
               f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>{header2}</w:hdr>')
    z.writestr("word/_rels/header2.xml.rels", header_rels)
    z.writestr("word/header3.xml",
               f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>{header3}</w:hdr>')
    z.writestr("word/footer1.xml",
               f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {NS}>{footer}</w:ftr>')
    z.writestr("word/_rels/footer1.xml.rels", header_rels)
    z.writestr("word/footnotes.xml", footnotes)
    z.writestr("word/media/block.png", png(40, 9))
print(f"wrote {out}")
