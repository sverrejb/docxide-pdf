"""Generate header-para-anchor-above-space-before: header drawings anchored to a
paragraph with space before

A drawing positioned relativeFrom="paragraph" in a header counts its offset
from the paragraph's top above its own space before, as in the body: below
the previous paragraph's space after, but above the space this paragraph adds
on top of it. A letterhead whose first header paragraph has 36pt before keeps
its logo 36pt higher than the text line.

Each page is a section with its own header. On pages 1-6 the anchor
paragraph carries a picture, a filled text box and a line, all
relativeFrom="paragraph":
- page 1: first header paragraph, 36pt before
- page 2: first header paragraph, 36pt before, negative offset
- page 3: second paragraph, 24pt before, previous paragraph no space after
- page 4: second paragraph, 24pt before, previous paragraph 12pt after
- page 5: second paragraph, 12pt before, previous paragraph 24pt after
- page 6: second paragraph, 18pt before from its style
- page 7: first paragraph, 36pt before, a top-and-bottom picture at the
  anchor: the text moves below the picture and keeps its 36pt before there
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
    """A blue frame with a diagonal band: easy to place by eye and by ink."""
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            edge = x < 3 or y < 3 or x >= width - 3 or y >= height - 3
            band = abs(x * height - y * width) < width * 4
            row += bytes((20, 60, 160) if edge or band else (200, 220, 245))
        rows.append(bytes(row))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b""))


def anchor(doc_id, x_pt, y_pt, w_pt, h_pt, graphic, wrap="<wp:wrapNone/>"):
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{251658240 + doc_id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:posOffset>{int(x_pt * EMU)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{int(y_pt * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{int(w_pt * EMU)}" cy="{int(h_pt * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'{wrap}<wp:docPr id="{doc_id}" name="Drawing {doc_id}"/>'
        f'<a:graphic>{graphic}</a:graphic></wp:anchor></w:drawing></w:r>'
    )


def picture(doc_id, x_pt, y_pt, wrap="<wp:wrapNone/>"):
    w, h = 90, 40
    return anchor(doc_id, x_pt, y_pt, w, h, wrap=wrap, graphic=(
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
        f'<pic:nvPicPr><pic:cNvPr id="{doc_id}" name="Picture {doc_id}"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData>'
    ))


def text_box(doc_id, x_pt, y_pt, label, wrap="<wp:wrapNone/>"):
    w, h = 110, 40
    return anchor(doc_id, x_pt, y_pt, w, h, wrap=wrap, graphic=(
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
        '<wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F4D7A1"/></a:solidFill>'
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="A05A00"/></a:solidFill></a:ln></wps:spPr>'
        f'<wps:txbx><w:txbxContent><w:p><w:r><w:t>{label}</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
        '<wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:noAutofit/></wps:bodyPr>'
        '</wps:wsp></a:graphicData>'
    ))


def line(doc_id, x_pt, y_pt):
    w = 150
    return anchor(doc_id, x_pt, y_pt, w, 0, graphic=(
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
        '<wps:cNvCnPr/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w * EMU}" cy="0"/></a:xfrm>'
        '<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/>'
        '<a:ln w="28575"><a:solidFill><a:srgbClr val="2E7D32"/></a:solidFill></a:ln></wps:spPr>'
        '<wps:bodyPr/></wps:wsp></a:graphicData>'
    ))


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content, before=0, after=0, style=None):
    style_xml = f'<w:pStyle w:val="{style}"/>' if style else ""
    spacing = "" if style else f'<w:spacing w:before="{before * 20}" w:after="{after * 20}"/>'
    return f"<w:p><w:pPr>{style_xml}{spacing}</w:pPr>{content}</w:p>"


TOP_BOTTOM = '<wp:wrapTopAndBottom/>'


def drawings(base_id, offset, kind):
    if kind == "picture below":
        return picture(base_id, 0, offset, TOP_BOTTOM)
    return (picture(base_id, 0, offset) + text_box(base_id + 1, 110, offset, "Box")
            + line(base_id + 2, 240, offset))


# (label, header paragraphs before the anchor, anchor before, anchor style,
#  offset, drawings)
VARIANTS = [
    ("first paragraph, before 36pt", [], 36, None, 4, "all"),
    ("first paragraph, before 36pt, offset -12pt", [], 36, None, -12, "all"),
    ("second paragraph, before 24pt, previous after 0", [0], 24, None, 4, "all"),
    ("second paragraph, before 24pt, previous after 12pt", [12], 24, None, 4, "all"),
    ("second paragraph, before 12pt, previous after 24pt", [24], 12, None, 4, "all"),
    ("second paragraph, before 18pt from its style", [0], 0, "Letterhead", 4, "all"),
    ("first paragraph, before 36pt, top-and-bottom picture", [], 36, None, 0, "picture below"),
]

headers = []
for i, (label, prev, before, style, offset, kind) in enumerate(VARIANTS):
    paras = [para(run("Previous header paragraph"), after=a) for a in prev]
    anchor_text = run(f"Anchor paragraph: {label}.")
    paras.append(para(drawings(100 + i * 10, offset, kind) + anchor_text, before=before, style=style))
    headers.append(
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>{"".join(paras)}</w:hdr>'
    )


def sect(i):
    return (
        f'<w:sectPr><w:headerReference w:type="default" r:id="rIdH{i}"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="3402" w:right="1417" w:bottom="1134" w:left="1417" w:header="709" w:footer="709" w:gutter="0"/>'
        "</w:sectPr>"
    )


body = []
for i, (label, *_rest) in enumerate(VARIANTS):
    text = run(f"Page {i + 1}: {label}. Body text starts at the top margin.")
    if i < len(VARIANTS) - 1:
        body.append(f"<w:p><w:pPr>{sect(i)}</w:pPr>{text}</w:p>")
    else:
        body.append(para(text))
document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body) + sect(len(VARIANTS) - 1) + "</w:body></w:document>"
)

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W_NS}"><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    '</w:docDefaults><w:style w:type="paragraph" w:styleId="Letterhead"><w:name w:val="Letterhead"/>'
    '<w:pPr><w:spacing w:before="360" w:after="0"/></w:pPr></w:style></w:styles>'
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
    + "".join(
        f'<Override PartName="/word/header{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        for i in range(len(VARIANTS))
    )
    + "</Types>"
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
    + "".join(
        f'<Relationship Id="rIdH{i}" Type="{REL}/header" Target="header{i}.xml"/>'
        for i in range(len(VARIANTS))
    )
    + "</Relationships>"
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
    z.writestr("word/settings.xml", settings)
    z.writestr("word/media/picture.png", png(180, 80))
    for i, header in enumerate(headers):
        z.writestr(f"word/header{i}.xml", header)
        z.writestr(f"word/_rels/header{i}.xml.rels", header_rels)
print(f"wrote {out}")
