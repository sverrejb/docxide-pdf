"""Generate picture-part-reuse: one media part drawn by many pictures

A document often shows one picture several times: a logo in the header and
the body, the same photo cropped differently, a grayscale copy. All of these
pictures point at one media part, word/media/picture.png.

Tests:
- Header: inline picture beside header text
- Body: plain inline picture, twice at the same size
- Table: the picture in two cells, cropped by a:srcRect differently
  (right half removed; top and bottom quarter removed)
- Inline picture recoloured to grayscale (a:grayscl)
- Text box holding the picture inline
- Floating picture with square wrapping
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
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
)
EMU = 12700


def png(width, height):
    """Four coloured quadrants with grain and a dark diagonal: shows crops, recolour and size."""
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


def pic_graphic(doc_id, w_pt, h_pt, src_rect="", blip_extra=""):
    cx, cy = int(w_pt * EMU), int(h_pt * EMU)
    return (
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{doc_id}" name="picture.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        f'<pic:blipFill><a:blip r:embed="rIdImg">{blip_extra}</a:blip><a:srcRect{src_rect}/>'
        '<a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
    )


def inline(doc_id, w_pt, h_pt, src_rect="", blip_extra=""):
    cx, cy = int(w_pt * EMU), int(h_pt * EMU)
    return (
        '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:docPr id="{doc_id}" name="Picture {doc_id}"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        f'{pic_graphic(doc_id, w_pt, h_pt, src_rect, blip_extra)}</wp:inline></w:drawing></w:r>'
    )


def anchor(doc_id, x_pt, y_pt, w_pt, h_pt, wrap, graphic):
    cx, cy = int(w_pt * EMU), int(h_pt * EMU)
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{251658240 + doc_id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:posOffset>{int(x_pt * EMU)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{int(y_pt * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap}'
        f'<wp:docPr id="{doc_id}" name="Shape {doc_id}"/>{graphic}</wp:anchor></w:drawing></w:r>'
    )


def text_box_graphic(w_pt, h_pt, content):
    return (
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(w_pt * EMU)}" cy="{int(h_pt * EMU)}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="202020"/></a:solidFill></a:ln></wps:spPr>'
        f'<wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t">'
        '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>'
    )


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content, before=0):
    return f'<w:p><w:pPr><w:spacing w:before="{before * 20}"/></w:pPr>{content}</w:p>'


def cell(content, width):
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{content}</w:tc>'


BORDERS = "".join(
    f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
    for side in ("top", "left", "bottom", "right", "insideH", "insideV")
)

table = (
    '<w:tbl><w:tblPr><w:tblW w:w="7200" w:type="dxa"/>'
    f'<w:tblBorders>{BORDERS}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
    '<w:tblGrid><w:gridCol w:w="3600"/><w:gridCol w:w="3600"/></w:tblGrid>'
    '<w:tr>'
    + cell(para(run("Right half cropped")) + para(inline(4, 80, 80, ' r="50000"')), 3600)
    + cell(para(run("Top and bottom quarter cropped")) + para(inline(5, 160, 40, ' t="25000" b="25000"')), 3600)
    + "</w:tr></w:tbl>"
)

header = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>'
    + para(inline(1, 80, 40) + run("  Header beside the picture"))
    + "</w:hdr>"
)

square = '<wp:wrapSquare wrapText="bothSides"/>'
body = [
    para(run("One media part drawn by every picture on this page")),
    para(inline(2, 160, 80) + run(" ") + inline(3, 160, 80), before=6),
    para(run("The same picture cropped differently in two table cells:"), before=6),
    table,
    para(run("Grayscale copy:"), before=6),
    para(inline(6, 160, 80, blip_extra="<a:grayscl/>")),
    para(run("A text box holding the picture inline:"), before=6),
    para(anchor(7, 0, 4, 200, 90, "<wp:wrapTopAndBottom/>",
                text_box_graphic(200, 90, para(run("Text box")) + para(inline(8, 120, 50))))),
    para(anchor(9, 333.6, 0, 120, 60, square, pic_graphic(9, 120, 60))
         + run("A floating copy of the picture sits at the right margin with square wrapping. "
               "This paragraph flows around it on the left, line after line, until the text "
               "has passed the bottom of the picture and the lines take the full column width "
               "again. The floating copy uses the same media part as every other picture on "
               "this page, in the header, the table cells and the text box."), before=12),
]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="2268" w:right="1417" w:bottom="1134" w:left="1417" w:header="709" w:footer="709" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W_NS}"><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
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
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
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
    f'<Relationship Id="rIdH" Type="{REL}/header" Target="header1.xml"/>'
    f'<Relationship Id="rIdImg" Type="{REL}/image" Target="media/picture.png"/></Relationships>'
)
header_rels = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELS_NS}">'
    f'<Relationship Id="rIdImg" Type="{REL}/image" Target="media/picture.png"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/header1.xml", header)
    z.writestr("word/_rels/header1.xml.rels", header_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/media/picture.png", png(320, 160))
print(f"wrote {out}")
