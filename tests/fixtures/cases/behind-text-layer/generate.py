"""Generate behind-text-layer: behind-text drawings anchored after the text they lie under

Word (online export, Word for Mac behavior, compat 15) paints every behindDoc
drawing of a page in one layer under all of the page's body text, whatever
paragraph or table cell anchors it:

- a light box, a picture and a picture-filled box anchored in a later
  paragraph, or in the table's last cell, stay under the earlier text;
- paragraph shading and cell shading cover a behind box anchored after them;
- behind drawings stack among themselves by relativeHeight, not anchor order
  (a dark box anchored earlier but higher stays on top of a pale one);
- the layer lies over the header: a body box hides header text;
- a wrapNone box in front (behindDoc 0) still covers the text.

Page 1 anchors its drawings in body paragraphs, page 2 in the last cell of a
table (layoutInCell: offsets from the cell), page 3 in a row that breaks onto
page 4 and in a table nested in it.
"""

import pathlib
import struct
import sys
import zipfile
import zlib

EMU = 12700
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"')
WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"
RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="20"/></w:rPr>'

LIGHT, DARK, PALE = "FFE699", "7F7F7F", "DDEBF7"
SHADE = "E2EFDA"


def stripes_png(n=64):
    """Light stripes (above the ink threshold) framed in black."""
    rows = []
    for y in range(n):
        row = bytearray([0])
        for x in range(n):
            if min(x, y, n - 1 - x, n - 1 - y) < 2:
                row += bytes((0, 0, 0))
            elif (x // 8) % 2:
                row += bytes((255, 220, 220))
            else:
                row += bytes((220, 235, 255))
        rows.append(bytes(row))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", n, n, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b""))


ids = iter(range(100, 10000))


def emu(v):
    return int(round(v * EMU))


def wsp(w, h, fill=None, picture=False, story=True):
    if picture:
        paint = ('<a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rIdImg"/><a:srcRect/>'
                 '<a:stretch><a:fillRect/></a:stretch></a:blipFill>')
    else:
        paint = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
    txbx = (f'<wps:txbx><w:txbxContent><w:p><w:pPr>{RPR}</w:pPr></w:p></w:txbxContent></wps:txbx>'
            if story else "")
    return (f'<wps:wsp><wps:cNvSpPr/><wps:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{paint}<a:ln><a:noFill/></a:ln></wps:spPr>'
            f'{txbx}<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" '
            'rIns="91440" bIns="45720" anchor="t" anchorCtr="0" upright="1"><a:noAutofit/>'
            "</wps:bodyPr></wps:wsp>")


def pic(w, h):
    ext = f'cx="{emu(w)}" cy="{emu(h)}"'
    return ('<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="stripes.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext {ext}/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>')


def anchored(x, y, w, h, z, behind, uri, graphic, h_from="page", v_from="page"):
    i = next(ids)
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            f'relativeHeight="{z}" behindDoc="{int(behind)}" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="{h_from}"><wp:posOffset>{emu(x)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="{v_from}"><wp:posOffset>{emu(y)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{emu(w)}" cy="{emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapNone/><wp:docPr id="{i}" name="Drawing {i}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="{uri}">{graphic}</a:graphicData></a:graphic>'
            "</wp:anchor></w:drawing></w:r>")


def box(x, y, w, h, z, behind=True, fill=LIGHT, picture=False, story=True, **pos):
    return anchored(x, y, w, h, z, behind, WPS_URI, wsp(w, h, fill, picture, story), **pos)


def picture(x, y, w, h, z, behind=True, **pos):
    return anchored(x, y, w, h, z, behind, PIC_URI, pic(w, h), **pos)


def run(text):
    return f'<w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(text, drawings="", shade=None, br=False):
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>' if shade else ""
    brk = '<w:r><w:br w:type="page"/></w:r>' if br else ""
    return f'<w:p><w:pPr>{shd}{RPR}</w:pPr>{brk}{run(text)}{drawings}</w:p>'


LINE = "The quick brown fox jumps over the lazy dog and keeps running along the river bank. "


def text(n, tag):
    return f"{tag}: " + LINE * n


def cell(w, content, shade=None):
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>' if shade else ""
    return f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>{shd}</w:tcPr>{content}</w:tc>'


def table(rows, width=4819):
    grid = "".join(f'<w:gridCol w:w="{w}"/>' for w in (width, width))
    borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                      for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    body = "".join(f"<w:tr>{''.join(cell(width, c, s) for c, s in r)}</w:tr>" for r in rows)
    return (f'<w:tbl><w:tblPr><w:tblW w:w="{2 * width}" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
            f'<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>{grid}</w:tblGrid>{body}</w:tbl>')


# Page 1: drawings anchored in the last paragraph lie on the text above them.
page1 = [
    para(text(3, "Under a light behind box")),
    para(text(3, "Shaded paragraph over a dark behind box"), shade=SHADE),
    para(text(3, "Under a behind picture")),
    para(text(1, "Stacking: the dark box anchored here is higher"),
         box(330, 380, 120, 60, 300, fill=DARK)),
    para(text(3, "Under a picture-filled behind box")),
    para(text(3, "Under a light box in front"),
         box(60, 72, 250, 40, 200)
         + box(400, 120, 150, 40, 210, fill=DARK)
         + picture(200, 150, 200, 30, 220)
         + box(380, 400, 120, 60, 100, fill=PALE)
         + box(300, 225, 200, 30, 230, picture=True)
         + box(60, 265, 250, 20, 240, behind=False)
         + box(60, 32, 200, 20, 250)),
    para("Behind a later paragraph's drawings: none."),
]

# Page 2: drawings anchored in the last cell lie on the cells before it.
# Keep the empty second-row slot so later anchors retain their offsets.
cell_drawings = (box(-230, -120, 200, 30, 400)
                 + box(180, -100, 100, 40, 410, fill=DARK)
                 + picture(-220, -50, 180, 25, 420)
                 + box(-220, 15, 200, 20, 430, picture=True)
                 + box(20, 30, 150, 15, 450))
page2 = [
    para("Table whose last cell anchors behind drawings", br=True),
    table([
        [(para(text(3, "Cell 1 under a light box")), None),
         (para(text(3, "Shaded cell over a dark box")), SHADE)],
        [(para(text(3, "Cell 3 under a picture")), None),
         (para(""), None)],
        [(para(text(2, "Cell 4 under a picture-filled box")), None),
         (para(text(2, "Cell 5 anchors the drawings"), cell_drawings), None)],
    ]),
    para(text(2, "After the table")),
]

# Pages 3-4: a row that breaks across the pages; the second cell anchors
# behind pictures in its first paragraph and in a nested table.
nested = table([[(para("Nested cell anchors a behind picture", picture(-250, 20, 200, 30, 510)), None),
                 (para("Nested"), None)]], width=2200)
page3 = [
    para("A row that breaks across pages", br=True),
    table([[(para(text(60, "Long cell under two pictures")), None),
            (para("Anchors a behind picture", picture(-180, 8, 150, 20, 500)) + nested + para(""), None)]]),
]

body = "".join(page1 + page2 + page3)

document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}'
            '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="567" w:footer="567" w:gutter="0"/>'
            "</w:sectPr></w:body></w:document>")

styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          "</w:docDefaults></w:styles>")

header = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>'
          + para("Header text under a light behind box anchored in the body: " + LINE) + "</w:hdr>")

settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            "</w:compat></w:settings>")

content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
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
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/>'
    f'<Relationship Id="rIdHdr" Type="{rel}/header" Target="header1.xml"/>'
    f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/stripes.png"/></Relationships>')

out = pathlib.Path(sys.argv[1]) if len(sys.argv) > 1 else pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/header1.xml", header)
    z.writestr("word/media/stripes.png", stripes_png())
print(f"wrote {out}")
