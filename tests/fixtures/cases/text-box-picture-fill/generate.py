"""Generate text-box-picture-fill: text boxes filled with a picture

A wps text box (wps:txbx) whose shape properties carry an a:blipFill instead
of a solid fill draws that picture as its background, stretched over the
shape and clipped to its outline, with the text box's text and outline on
top. A letterhead logo is often a text box like this holding one empty
paragraph.

Tests:
- Header: picture-filled text box with one empty paragraph, positioned
  relativeFrom="line", white outline
- Rectangle text box with text over the picture and a dark outline
- Rounded rectangle text box with text: the picture is clipped to the corners
- Ellipse text box with one empty paragraph and no outline
- Rectangle text box whose picture is cropped by a:srcRect (left and right 25%)
- Rectangle text box with a shape style (wps:style fillRef): the picture
  fill replaces the style's fill
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
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
)
EMU = 12700


def png(width, height):
    """Four coloured quadrants and a dark diagonal: shows stretch, crop and flips."""
    quadrants = [[(30, 90, 200), (240, 150, 40)], [(60, 170, 90), (200, 60, 90)]]
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            if abs(x * height - y * width) < width * 3:
                row += bytes((25, 25, 25))
            else:
                row += bytes(quadrants[y * 2 // height][x * 2 // width])
        rows.append(bytes(row))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b""))


STYLE = (
    '<wps:style><a:lnRef idx="2"><a:srgbClr val="7030A0"/></a:lnRef>'
    '<a:fillRef idx="1"><a:srgbClr val="FF00FF"/></a:fillRef>'
    '<a:effectRef idx="0"><a:srgbClr val="000000"/></a:effectRef>'
    '<a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></wps:style>'
)


def text_box_graphic(w_pt, h_pt, geom, text, outline, src_rect="", style=""):
    run = f'<w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:t>{text}</w:t></w:r>' if text else ""
    ln = (f'<a:ln w="12700"><a:solidFill><a:srgbClr val="{outline}"/></a:solidFill></a:ln>'
          if outline else "<a:ln><a:noFill/></a:ln>")
    return (
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(w_pt * EMU)}" cy="{int(h_pt * EMU)}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>'
        f'<a:blipFill dpi="0" rotWithShape="0"><a:blip r:embed="rIdImg"/><a:srcRect{src_rect}/>'
        f'<a:stretch><a:fillRect/></a:stretch></a:blipFill>{ln}</wps:spPr>{style}'
        f'<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr>{run}</w:p></w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t">'
        '<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>'
    )


def anchored(doc_id, x_pt, y_pt, w_pt, h_pt, v_from, graphic):
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{251658240 + doc_id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:posOffset>{int(x_pt * EMU)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="{v_from}"><wp:posOffset>{int(y_pt * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{int(w_pt * EMU)}" cy="{int(h_pt * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'<wp:wrapNone/><wp:docPr id="{doc_id}" name="Text Box {doc_id}"/>{graphic}</wp:anchor></w:drawing></w:r>'
    )


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content, before=0):
    return f'<w:p><w:pPr><w:spacing w:before="{before * 20}"/></w:pPr>{content}</w:p>'


header = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>'
    + para(anchored(1, 250, -4, 200, 45, "line", text_box_graphic(200, 45, "rect", "", "FFFFFF"))
           + run("Header text beside a picture-filled text box"))
    + "</w:hdr>"
)

body = [
    para(run("Text boxes filled with a picture")),
    para(anchored(2, 0, 10, 200, 80, "paragraph",
                  text_box_graphic(200, 80, "rect", "Text over the picture", "202020"))
         + anchored(3, 240, 10, 200, 80, "paragraph",
                    text_box_graphic(200, 80, "roundRect", "Rounded corners", None)),
         before=12),
    *[para("") for _ in range(7)],
    para(anchored(4, 0, 10, 200, 80, "paragraph", text_box_graphic(200, 80, "ellipse", "", None))
         + anchored(5, 240, 10, 200, 80, "paragraph",
                    text_box_graphic(200, 80, "rect", "Cropped", "202020", ' l="25000" r="25000"')),
         before=12),
    *[para("") for _ in range(7)],
    para(anchored(6, 0, 10, 200, 80, "paragraph",
                  text_box_graphic(200, 80, "rect", "Style fill replaced", "202020", style=STYLE)),
         before=12),
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
    z.writestr("word/media/picture.png", png(240, 120))
print(f"wrote {out}")
