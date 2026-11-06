"""Generate underlined-picture-places: underlined inline pictures outside a body paragraph

An underlined inline picture is underlined in the run's colour and keeps the run
font's descent below it wherever it sits, as in a body paragraph (fixture
underlined-inline-picture). Each picture is 100 x 30pt, its run Tahoma 16pt.

Header: a paragraph of two underlined 250pt pictures wrapping onto two lines,
then "Header after pictures"; the same with the paragraph mark in the
pictures' font (whose last line loses its descent when not underlined); a table
whose first cell holds an underlined picture, then "Header after table".
Body, one label line before and after each case:
1. A two-row table: an underlined picture alone in the first cell, "Row 2"
   below it.
2. A text box at the page's right holding an underlined picture alone in its
   paragraph and "Box after" below it.
3. A VML (w:pict) inline picture, underlined.
4. A picture underlined in red (the run's colour).
5. A hidden (vanish) underlined picture alone in its paragraph, then a hidden
   picture without underline.
6. A shape-fitted text box holding only an underlined 100pt picture: as tall
   as the picture and the run's descent.
7. A hidden picture alone in a 72pt run: its paragraph is a line of its 10pt
   mark in a table cell, a text box and the header too.
8. Shape-fitted boxes as 6 whose paragraph is exactly 12pt: as tall as the
   line, underlined or not.
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
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:o="urn:schemas-microsoft-com:office:office"'
)
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 12700


def png():
    """A grey square, drawn here."""
    w = h = 40
    rows = b"".join(b"\x00" + bytes((150, 150, 150)) * w for _ in range(h))
    chunk = lambda t, d: struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(rows))
        + chunk(b"IEND", b"")
    )


ids = iter(range(1, 999))


def rpr(u=True, color=None, vanish=False, size=16):
    return (
        '<w:rPr><w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma"/>'
        + (f'<w:color w:val="{color}"/>' if color else "")
        + ('<w:vanish/>' if vanish else "")
        + f'<w:sz w:val="{size * 2}"/>'
        + ('<w:u w:val="single"/>' if u else "")
        + "</w:rPr>"
    )


def pic(w_pt=100, h_pt=30, **fmt):
    n = next(ids)
    cx, cy = int(w_pt * EMU), int(h_pt * EMU)
    return (
        f'<w:r>{rpr(**fmt)}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{n}" name="Picture {n}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
        f'<pic:nvPicPr><pic:cNvPr id="{n}" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
        "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )


def vml_pic():
    return (
        f'<w:r>{rpr()}<w:pict><v:shape id="vml{next(ids)}" type="#_x0000_t75" style="width:100pt;height:30pt">'
        '<v:imagedata r:id="rIdImg" o:title=""/></v:shape></w:pict></w:r>'
    )


def para(runs="", text=None, mark="", rule="auto"):
    if text is not None:
        runs = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
    return f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="{rule}"/>{mark}</w:pPr>{runs}</w:p>'


def table(first_cell):
    borders = "".join(
        f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for s in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    cell = lambda body: f'<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>{body}</w:tc>'
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="3000" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
        '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>'
        f"<w:tr>{cell(para(first_cell))}</w:tr><w:tr>{cell(para(text='Row 2'))}</w:tr></w:tbl>"
    )


def text_box(x=400, y=300, content=None, fit=False):
    w, h = 160 * EMU, 120 * EMU
    n = next(ids)
    content = content or para(pic()) + para(text="Box after")
    fit = "<a:spAutoFit/>" if fit else "<a:noAutofit/>"
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{x * EMU}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{y * EMU}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{w}" cy="{h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{n}" name="Box {n}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
        f"<wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx>"
        f'<wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0">{fit}</wps:bodyPr></wps:wsp>'
        "</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )


HIDDEN_BIG = pic(vanish=True, u=False, size=72)
header = (
    para(text="Header top")
    + para(pic(250) + pic(250))
    + para(text="Header after pictures")
    + para(pic(250) + pic(250), mark=rpr())
    + para(text="Header after Tahoma mark")
    + table(pic())
    + para(text="Header after table")
    + para(HIDDEN_BIG)
    + para(text="Header after hidden picture")
)
body = (
    para(text="Case 1: table")
    + table(pic())
    + para(text="Case 2: text box" )
    + para(text_box() + '<w:r><w:t xml:space="preserve">Anchor of the box</w:t></w:r>')
    + para(text="Case 3: VML picture")
    + para(vml_pic())
    + para(text="Case 4: red underline")
    + para(pic(color="FF0000"))
    + para(text="Case 5: hidden underlined picture")
    + para(pic(vanish=True))
    + para(text="Hidden plain picture")
    + para(pic(u=False, vanish=True))
    + para(text="Case 6: shape-fitted box")
    + para(
        text_box(300, 450, para(pic(h_pt=100)), fit=True)
        + '<w:r><w:t xml:space="preserve">Anchor of the box</w:t></w:r>'
    )
    + para(text="Case 7: hidden 72pt picture run")
    + table(HIDDEN_BIG)
    + para(text_box(400, 600, para(HIDDEN_BIG) + para(text="Box after")) + '<w:r><w:t xml:space="preserve">Anchor</w:t></w:r>')
    + para(text="Case 8: shape-fitted boxes, exact lines")
    + para(
        text_box(60, 760, para(pic(h_pt=100), rule="exact"), fit=True)
        + text_box(230, 760, para(pic(h_pt=100, u=False), rule="exact"), fit=True)
        + '<w:r><w:t xml:space="preserve">Anchor of the exact boxes</w:t></w:r>'
    )
    + para(text="End")
)
sect = (
    '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="1418" w:bottom="1134" w:left="1418" w:header="567" w:footer="567" w:gutter="0"/>'
    "</w:sectPr>"
)
document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}{sect}</w:body></w:document>'
)
header_xml = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>{header}</w:hdr>'
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
override = lambda part, kind: (
    f'<Override PartName="/word/{part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>'
)
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
    + override("document.xml", "document.main")
    + override("styles.xml", "styles")
    + override("settings.xml", "settings")
    + override("header1.xml", "header")
    + "</Types>"
)
relationships = lambda items: (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + "".join(f'<Relationship Id="{rid}" Type="{REL}/{kind}" Target="{target}"/>' for rid, kind, target in items)
    + "</Relationships>"
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", relationships([("rId1", "officeDocument", "word/document.xml")]))
    z.writestr("word/document.xml", document)
    z.writestr(
        "word/_rels/document.xml.rels",
        relationships(
            [
                ("rIdStyles", "styles", "styles.xml"),
                ("rIdSettings", "settings", "settings.xml"),
                ("rIdHdr", "header", "header1.xml"),
                ("rIdImg", "image", "media/p.png"),
            ]
        ),
    )
    z.writestr("word/_rels/header1.xml.rels", relationships([("rIdImg", "image", "media/p.png")]))
    z.writestr("word/header1.xml", header_xml)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/media/p.png", png())
print(f"wrote {out}")
