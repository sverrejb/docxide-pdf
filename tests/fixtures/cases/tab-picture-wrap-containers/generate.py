"""Generate tab-picture-wrap-containers: tab-picture-wrap's trailing tab line in
table cells, text boxes and headers

A picture that does not fit after a tab starts the next line, and a tab after
it that no longer fits leaves a blank line sized by the tab's font. That line
also holds the paragraph mark, so a larger mark makes it taller, wherever the
paragraph is laid out (tab-picture-wrap covers body paragraphs).

Every container is 468pt wide like the body column and holds "Before", a
banner paragraph (tab, 480pt picture, tab, space; 10.5pt runs) and "After":
- page 1: a table cell (T1), mark 14pt; an empty header
- page 2: a text box (B1) and the header (H1), both with a 14pt mark
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
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
)
COL_TWIPS = 11906 - 1418 - 1134


def placeholder_png(w=200, h=20):
    """A grey checkerboard, drawn here."""
    rows = b""
    for y in range(h):
        rows += b"\x00" + b"".join(
            bytes((v, v, v)) for v in (60 if (x // 10 + y // 10) % 2 else 150 for x in range(w))
        )

    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def picture(w_pt, h_pt, pid):
    cx, cy = int(w_pt * 12700), int(h_pt * 12700)
    return (
        f'<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{pid}" name="Banner {pid}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{pid}" name="banner.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        "</wp:inline></w:drawing></w:r>"
    )


TAB = '<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:tab/></w:r>'
SPACE = '<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>'


def marker(text):
    return f'<w:p><w:pPr><w:spacing w:line="260" w:lineRule="exact"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def banner(pid, mark=None):
    mark_xml = f'<w:rPr><w:sz w:val="{mark}"/></w:rPr>' if mark else ""
    return (
        '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1660"/><w:tab w:val="center" w:pos="4025"/></w:tabs>'
        f"{mark_xml}</w:pPr>{TAB}{picture(480, 30, pid)}{TAB}{SPACE}</w:p>"
    )


def group(label, pid, mark):
    return marker(f"Before {label}") + banner(pid, mark) + marker(f"After {label}")


def table(content):
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="{COL_TWIPS}" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{COL_TWIPS}"/></w:tblGrid><w:tr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{COL_TWIPS}" w:type="dxa"/></w:tcPr>{content}</w:tc></w:tr></w:tbl>'
    )


def text_box(content, shape_id):
    cx, cy = COL_TWIPS * 635, 3000000
    return (
        '<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        f'relativeHeight="{shape_id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        '<wp:positionV relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{shape_id}" name="Text Box {shape_id}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
        f'<a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t">'
        "<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>"
    )


def sect(header_rid):
    return (
        f'<w:sectPr><w:headerReference w:type="default" r:id="{header_rid}"/><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="3402" w:right="1134" w:bottom="1134" w:left="1418" w:header="284" w:footer="708" w:gutter="0"/>'
        "</w:sectPr>"
    )


body = (
    table(group("T1", 2, 28))
    + f'<w:p><w:pPr>{sect("rIdH1")}</w:pPr></w:p>'
    + text_box(group("B1", 4, 28), 5)
)
# Keep separate headers so the retained header case stays on page 2.
headers = {"rIdH1": "<w:p/>", "rIdH2": group("H1", 7, 28)}

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + body + sect("rIdH2") + "</w:body></w:document>"
)

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
        for i in (1, 2)
    )
    + "</Types>"
)

rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
package_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>'
)
image_rel = f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/banner.png"/>'
document_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/>'
    f'<Relationship Id="rIdH1" Type="{rel}/header" Target="header1.xml"/>'
    f'<Relationship Id="rIdH2" Type="{rel}/header" Target="header2.xml"/>'
    f"{image_rel}</Relationships>"
)
header_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{image_rel}</Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    for i, rid in enumerate(headers, 1):
        z.writestr(f"word/header{i}.xml",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>{headers[rid]}</w:hdr>')
        z.writestr(f"word/_rels/header{i}.xml.rels", header_rels)
    z.writestr("word/media/banner.png", placeholder_png())
print(f"wrote {out}")
