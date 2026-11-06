"""Generate textbox-oversized-inline-image: inline pictures in fixed-size text boxes

A small text box holds one centred paragraph with a zero width space and an
inline picture larger than the box.
Word (online export, Word for Mac behavior) then:

- breaks the line after the zero width space when the picture is wider than
  the text area, and starts a line wider than the measure at the left inset
  whatever the paragraph alignment;
- centres or bottom-aligns (bodyPr anchor ctr/b) the real content height,
  pictures included, and starts overflowing content at the top inset;
- sizes a picture line as the picture, the descent of visible text on it (a
  zero width space has none) and the leading multiple line spacing gives the
  font (a picture-only paragraph gets its mark font's leading);
- clips the box's content to its height inside the top and bottom insets; a
  standalone box clips pictures to its text area, a group's box to its
  left and right edges.

Pages: oversized picture in standalone boxes, the same in groups (anchor
t/ctr/b by row, jc left/center/right by column); small pictures bottom-anchored
(no text / zero width space / text by row, line 240/271/360 by column);
wide, tall and text overflow boxes (anchor t/ctr/b by column).
"""

import pathlib
import struct
import zipfile
import zlib

EMU = 12700
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
      'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"')
WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="20"/></w:rPr>'
BOX_W, BOX_H = 106.6, 141.5
COLS, ROWS = (60, 240, 420), (70, 310, 570)
BIG = (223.1, 223.1)


def placeholder_png(n=60):
    """A framed checkerboard, a different colour per quadrant, so a clip shows."""
    colours = ((200, 0, 0), (0, 0, 200), (0, 140, 0), (0, 0, 0))
    rows = []
    for y in range(n):
        row = bytearray([0])
        for x in range(n):
            if min(x, y, n - 1 - x, n - 1 - y) < 3:
                row += bytes((0, 0, 0))
            elif (x // 4 + y // 4) % 2:
                row += bytes(colours[(y >= n // 2) * 2 + (x >= n // 2)])
            else:
                row += bytes((255, 255, 255))
        rows.append(bytes(row))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", n, n, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b""))


ids = iter(range(100, 10000))


def emu(v):
    return int(round(v * EMU))


def inline_picture(w, h):
    i = next(ids)
    ext = f'cx="{emu(w)}" cy="{emu(h)}"'
    return ('<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent {ext}/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{i}" name="Picture {i}"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="code.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext {ext}/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing>')


LEADS = {
    "none": '<w:t xml:space="preserve"></w:t>',
    "zwsp": '<w:t xml:space="preserve"></w:t><w:t xml:space="preserve">​</w:t>',
    "text": '<w:t xml:space="preserve">Code</w:t>',
}


def picture_paragraph(pic, jc="center", line=271, lead="zwsp"):
    ppr = f'<w:pPr><w:spacing w:line="{line}" w:lineRule="auto"/><w:jc w:val="{jc}"/>{RPR}</w:pPr>'
    return f'<w:p>{ppr}<w:r>{RPR}{LEADS[lead]}{inline_picture(*pic)}</w:r></w:p>'


def text_paragraph():
    text = " ".join(f"Access line {i}." for i in range(1, 31))
    return f'<w:p><w:pPr>{RPR}</w:pPr><w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def shape(x, y, w, h, anchor, story=None, fill=None, child=False):
    i = next(ids)
    name = f'<wps:cNvPr id="{i}" name="Shape {i}"/>' if child else ""
    nv = '<wps:cNvSpPr txBox="1"/>' if story else "<wps:cNvSpPr/>"
    paint = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    txbx = f"<wps:txbx><w:txbxContent>{story}</w:txbxContent></wps:txbx>" if story else ""
    return (f"<wps:wsp>{name}{nv}<wps:spPr>"
            f'<a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{paint}<a:ln><a:noFill/></a:ln></wps:spPr>'
            f'{txbx}<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" '
            f'rIns="91440" bIns="45720" anchor="{anchor}" anchorCtr="0" upright="1"><a:noAutofit/>'
            "</wps:bodyPr></wps:wsp>")


def anchored(x, y, w, h, uri, graphic):
    i = next(ids)
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            f'relativeHeight="{i}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{emu(x)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{emu(y)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{emu(w)}" cy="{emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapNone/><wp:docPr id="{i}" name="Box {i}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="{uri}">{graphic}</a:graphicData></a:graphic>'
            "</wp:anchor></w:drawing></w:r>")


def text_box(x, y, story, anchor="ctr", group=False):
    if not group:
        return anchored(x, y, BOX_W, BOX_H, WPS_URI, shape(0, 0, BOX_W, BOX_H, anchor, story))
    pad = 10
    gw, gh = BOX_W + 2 * pad, BOX_H + 2 * pad
    size = f'cx="{emu(gw)}" cy="{emu(gh)}"'
    group_xml = ('<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
                 f'<a:ext {size}/><a:chOff x="0" y="0"/><a:chExt {size}/></a:xfrm></wpg:grpSpPr>'
                 + shape(0, 0, gw, gh, "ctr", fill="EEEEEE", child=True)
                 + shape(pad, pad, BOX_W, BOX_H, anchor, story, child=True) + "</wpg:wgp>")
    return anchored(x - pad, y - pad, gw, gh, WPG_URI, group_xml)


def label(text):
    return f'<w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r>'


ANCHORS, JCS = ("t", "ctr", "b"), ("left", "center", "right")


def oversized(group):
    return [text_box(COLS[c], ROWS[r], picture_paragraph(BIG, jc=jc), anchor, group)
            for r, anchor in enumerate(ANCHORS) for c, jc in enumerate(JCS)]


small = [text_box(COLS[c], ROWS[r], picture_paragraph((50, 50), line=line, lead=lead), "b")
         for r, lead in enumerate(("none", "zwsp", "text")) for c, line in enumerate((240, 271, 360))]

mixed = ([text_box(COLS[c], ROWS[0], picture_paragraph((223.1, 60)), anchor) for c, anchor in enumerate(ANCHORS)]
         + [text_box(COLS[c], ROWS[1], picture_paragraph((60, 223.1)), anchor) for c, anchor in enumerate(ANCHORS)]
         + [text_box(COLS[c], ROWS[2], text_paragraph(), anchor) for c, anchor in enumerate(ANCHORS)])

pages = [
    ("Oversized picture after a zero width space: anchor t/ctr/b by row, jc left/center/right by column", oversized(False)),
    ("The same boxes in groups", oversized(True)),
    ("50pt picture, bottom-anchored: no text/zero width space/text by row, line 240/271/360 by column", small),
    ("Wide picture, tall picture, overflowing text: anchor t/ctr/b by column", mixed),
]

body = "".join(
    f'<w:p>{"<w:r><w:br w:type=\"page\"/></w:r>" if n else ""}{label(title)}{"".join(boxes)}</w:p>'
    for n, (title, boxes) in enumerate(pages))

document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="567" w:right="1134" w:bottom="567" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>'
            "</w:sectPr></w:body></w:document>")

styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          "</w:docDefaults></w:styles>")

content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    "</Types>")

rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
package_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>')
document_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/code.png"/></Relationships>')

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/media/code.png", placeholder_png())
print(f"wrote {out}")
