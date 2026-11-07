"""Generate text-box-picture-fill-group: picture fills in groups, scaled shapes and
under header lines

A picture-filled text box (a:blipFill) draws its picture at the size the
shape is shown, under a foreground line anchored with it. Examples (Word for
Mac, online export):

- Header: a behind-text picture-filled text box and a line in front of it,
  both anchored in the header paragraph: the line stays on top.
- A group whose child space is a tenth of its extent: a 20 x 20 child text
  box filled with a fine checkerboard prints 200 x 200 and as sharp as at
  that size; next to it a text box whose a:xfrm says 20 x 20 in a 200 x
  200pt wp:extent, the same.
- A group filled with a picture (grpSpPr a:blipFill) and three children
  taking it with a:grpFill: an outlined shape without text, a text box with
  text and a shape without outline or text.
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
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"'
)
EMU = 12700
WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"


def png(n=400, cell=8):
    """A checkerboard of `cell` px squares, a different colour per quadrant."""
    colours = [[(30, 90, 200), (240, 150, 40)], [(60, 170, 90), (200, 60, 90)]]
    rows = []
    for y in range(n):
        row = bytearray([0])
        for x in range(n):
            dark = (x // cell + y // cell) % 2
            row += bytes(colours[y * 2 // n][x * 2 // n] if dark else (250, 250, 250))
        rows.append(bytes(row))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    ihdr = struct.pack(">IIBBBBB", n, n, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b""))


BLIP = ('<a:blipFill dpi="0" rotWithShape="0"><a:blip r:embed="rIdImg"/><a:srcRect/>'
        '<a:stretch><a:fillRect/></a:stretch></a:blipFill>')
ids = iter(range(10, 1000))


def xfrm(x, y, w, h, ch=None):
    child = "" if ch is None else f'<a:chOff x="0" y="0"/><a:chExt cx="{int(ch[0] * EMU)}" cy="{int(ch[1] * EMU)}"/>'
    return (f'<a:xfrm><a:off x="{int(x * EMU)}" y="{int(y * EMU)}"/>'
            f'<a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/>{child}</a:xfrm>')


def wsp(x, y, w, h, fill, text=None, outline=None, child=False):
    i = next(ids)
    name = f'<wps:cNvPr id="{i}" name="Shape {i}"/>' if child else ""
    ln = (f'<a:ln w="25400"><a:solidFill><a:srgbClr val="{outline}"/></a:solidFill></a:ln>'
          if outline else "<a:ln><a:noFill/></a:ln>")
    body = ""
    if text is not None:
        r = f'<w:r><w:rPr><w:b/></w:rPr><w:t>{text}</w:t></w:r>' if text else ""
        body = (f'<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr>{r}</w:p>'
                "</w:txbxContent></wps:txbx>")
    nv = '<wps:cNvSpPr txBox="1"/>' if text is not None else "<wps:cNvSpPr/>"
    return (f"<wps:wsp>{name}{nv}<wps:spPr>{xfrm(x, y, w, h)}"
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill}{ln}</wps:spPr>{body}'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" '
            'bIns="45720" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp>')


def line_wsp(w):
    return ('<wps:wsp><wps:cNvCnPr/><wps:spPr>'
            f'{xfrm(0, 0, w, 0)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/>'
            '<a:ln w="38100"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></wps:spPr>'
            "<wps:bodyPr/></wps:wsp>")


def anchored(x, y, w, h, uri, graphic, v_from="paragraph", behind=False):
    i = next(ids)
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        f'relativeHeight="{251658240 + i}" behindDoc="{int(behind)}" locked="0" layoutInCell="1" '
        'allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:posOffset>{int(x * EMU)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="{v_from}"><wp:posOffset>{int(y * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{int(w * EMU)}" cy="{int(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:wrapNone/><wp:docPr id="{i}" name="Drawing {i}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic><a:graphicData uri="{uri}">{graphic}</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    )


def group(w, h, ch, fill, shapes):
    return (f"<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>{xfrm(0, 0, w, h, ch)}{fill}</wpg:grpSpPr>"
            + "".join(shapes) + "</wpg:wgp>")


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content, before=0):
    return f'<w:p><w:pPr><w:spacing w:before="{before * 20}"/></w:pPr>{content}</w:p>'


GRP_FILL = "<a:grpFill/>"

header = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>'
    + para(anchored(250, -4, 200, 45, WPS_URI, wsp(0, 0, 200, 45, BLIP, ""), behind=True)
           + anchored(230, 18, 240, 0, WPS_URI, line_wsp(240))
           + run("Header text beside a picture-filled box under a line"))
    + "</w:hdr>"
)

scaled = (anchored(0, 10, 200, 200, WPG_URI,
                   group(200, 200, (20, 20), "", [wsp(0, 0, 20, 20, BLIP, "Scaled", child=True)]))
          + anchored(240, 10, 200, 200, WPS_URI, wsp(0, 0, 20, 20, BLIP, "Small xfrm")))
grouped = anchored(0, 10, 440, 140, WPG_URI, group(440, 140, (440, 140), BLIP, [
    wsp(0, 0, 140, 140, GRP_FILL, outline="202020", child=True),
    wsp(150, 0, 140, 140, GRP_FILL, "Group fill", outline="202020", child=True),
    wsp(300, 0, 140, 140, GRP_FILL, child=True),
]))

body = [
    para(run("Scaled picture fills: a group child and a small a:xfrm")),
    para(scaled, before=6),
    *[para("") for _ in range(17)],
    para(run("A group's picture fill taken by its children (a:grpFill)")),
    para(grouped, before=6),
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

content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
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
    z.writestr("word/styles.xml", styles)
    z.writestr("word/header1.xml", header)
    z.writestr("word/_rels/header1.xml.rels", header_rels)
    z.writestr("word/media/picture.png", png())
print(f"wrote {out}")
