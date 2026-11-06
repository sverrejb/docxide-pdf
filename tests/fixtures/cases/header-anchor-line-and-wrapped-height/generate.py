"""Generate header-anchor-line-and-wrapped-height: line-relative header drawings
and header text wrapped beside a picture

A header drawing positioned relativeFrom="paragraph" hangs from above the
paragraph's own space before. One positioned relativeFrom="line" in the
first line hangs from the same place: the first line's box starts above the
space before (below the previous paragraph's space after). And a header line
wraps beside a square-wrapped picture its line box overlaps, so a header
paragraph beside a picture takes more lines than at full width; the body
starts below all of them (Word for Mac behavior, compatibilityMode 15).

Each page is a section with its own header:
- page 1: first header paragraph, 36pt before, a picture, a filled text box
  and a line relativeFrom="line" 4pt down
- page 2: the same, -12pt
- page 3: second paragraph, 24pt before, previous paragraph 12pt after, 4pt
  down from the line
- pages 4-7: a long header paragraph beside a square-wrapped picture, the
  body 22pt below the header margin, so the header text pushes it down:
  picture left / right, paragraph 36pt before (the raised 40pt picture
  overlaps its first line by 4pt) / no space before (a 34pt picture beside
  three lines, ending above the body)
- page 8: a centred logo beside empty lines, then a long paragraph beside a
  right picture: the logo, with room on both sides, is above the paragraph
  and does not keep its lines at full width
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


def anchor(doc_id, h_pos, v_from, y_pt, w_pt, h_pt, graphic, wrap="<wp:wrapNone/>"):
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{251658240 + doc_id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column">{h_pos}</wp:positionH>'
        f'<wp:positionV relativeFrom="{v_from}"><wp:posOffset>{int(y_pt * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{int(w_pt * EMU)}" cy="{int(h_pt * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
        f'{wrap}<wp:docPr id="{doc_id}" name="Drawing {doc_id}"/>'
        f'<a:graphic>{graphic}</a:graphic></wp:anchor></w:drawing></w:r>'
    )


def offset(x_pt):
    return f"<wp:posOffset>{int(x_pt * EMU)}</wp:posOffset>"


def picture(doc_id, h_pos, v_from, y_pt, wrap="<wp:wrapNone/>", h=40, w=90):
    return anchor(doc_id, h_pos, v_from, y_pt, w, h, wrap=wrap, graphic=(
        '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
        f'<pic:nvPicPr><pic:cNvPr id="{doc_id}" name="Picture {doc_id}"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData>'
    ))


def text_box(doc_id, v_from, y_pt):
    w, h = 110, 40
    return anchor(doc_id, offset(110), v_from, y_pt, w, h, graphic=(
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
        '<wps:cNvSpPr txBox="1"/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F4D7A1"/></a:solidFill>'
        '<a:ln w="12700"><a:solidFill><a:srgbClr val="A05A00"/></a:solidFill></a:ln></wps:spPr>'
        '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Box</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
        '<wps:bodyPr lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"><a:noAutofit/></wps:bodyPr>'
        '</wps:wsp></a:graphicData>'
    ))


def line(doc_id, v_from, y_pt):
    w = 150
    return anchor(doc_id, offset(240), v_from, y_pt, w, 0, graphic=(
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>'
        '<wps:cNvCnPr/><wps:spPr>'
        f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w * EMU}" cy="0"/></a:xfrm>'
        '<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:noFill/>'
        '<a:ln w="28575"><a:solidFill><a:srgbClr val="2E7D32"/></a:solidFill></a:ln></wps:spPr>'
        '<wps:bodyPr/></wps:wsp></a:graphicData>'
    ))


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content, before=0, after=0):
    return f'<w:p><w:pPr><w:spacing w:before="{before * 20}" w:after="{after * 20}"/></w:pPr>{content}</w:p>'


SQUARE = '<wp:wrapSquare wrapText="bothSides"/>'
LONG = ("Header text beside the picture: these words wrap in the room the picture leaves, then"
        " continue at the full width of the header below it, line after line, until the paragraph"
        " ends here.")


def anchored_header(i, v_from, y, before, prev_after=None):
    base = 100 + i * 10
    drawings = (picture(base, offset(0), v_from, y) + text_box(base + 1, v_from, y)
                + line(base + 2, v_from, y))
    paras = [] if prev_after is None else [para(run("Previous header paragraph"), after=prev_after)]
    label = f"Anchor paragraph: relativeFrom {v_from}, {y}pt, {before}pt before."
    return paras + [para(drawings + run(label), before=before)]


def wrapped_header(i, side, before):
    h_pos = f"<wp:align>{side}</wp:align>"
    pic = picture(100 + i * 10, h_pos, "paragraph", 0, SQUARE, h=40 if before else 34)
    return [para(pic + run(LONG), before=before)]


def logo_above_wrapped_header(i):
    logo = picture(100 + i * 10, "<wp:align>center</wp:align>", "paragraph", 0, SQUARE)
    side = picture(101 + i * 10, "<wp:align>right</wp:align>", "paragraph", 0, SQUARE, h=80, w=180)
    return [para(logo)] + [para("")] * 4 + [para(side + run(f"{LONG} {LONG}"))]


PAGES = [
    ("line-relative drawings 4pt down, 36pt before", anchored_header(0, "line", 4, 36), 3402),
    ("line-relative drawings -12pt, 36pt before", anchored_header(1, "line", -12, 36), 3402),
    ("line-relative drawings, second paragraph", anchored_header(2, "line", 4, 24, 12), 3402),
    ("picture left, 36pt before", wrapped_header(3, "left", 36), 1149),
    ("picture right, 36pt before", wrapped_header(4, "right", 36), 1149),
    ("picture left, no space before", wrapped_header(5, "left", 0), 1149),
    ("picture right, no space before", wrapped_header(6, "right", 0), 1149),
    ("centred logo above, picture right", logo_above_wrapped_header(7), 1149),
]

headers = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>{"".join(p)}</w:hdr>'
           for _, p, _ in PAGES]


def sect(i):
    return (
        f'<w:sectPr><w:headerReference w:type="default" r:id="rIdH{i}"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        f'<w:pgMar w:top="{PAGES[i][2]}" w:right="1417" w:bottom="1134" w:left="1417" w:header="709" '
        'w:footer="709" w:gutter="0"/></w:sectPr>'
    )


body = []
for i, (label, *_rest) in enumerate(PAGES):
    text = run(f"Page {i + 1}: {label}. Body text starts below the header.")
    if i < len(PAGES) - 1:
        body.append(f"<w:p><w:pPr>{sect(i)}</w:pPr>{text}</w:p>")
    else:
        body.append(para(text))
document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body) + sect(len(PAGES) - 1) + "</w:body></w:document>"
)

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W_NS}"><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    '</w:docDefaults></w:styles>'
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
        for i in range(len(PAGES))
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
        for i in range(len(PAGES))
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
