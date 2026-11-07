"""Generate shape-outline-dash-cells: dashed text box outlines anchored in table cells

Text boxes anchored in a table cell stroke their a:ln/a:prstDash pattern as
body text boxes do, in multiples of the line width, a rectangle starting at
its bottom-left corner.

Tests:
- A cell holding a dash text box at 1pt
- A dashDot rounded rectangle and a solid box for comparison
"""

import pathlib
import zipfile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = (
    f'xmlns:w="{W_NS}" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)
WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
EMU = 12700


def wsp(w, h, dash, width_pt, label, geom):
    return (
        f'<a:graphic><a:graphicData uri="{WPS_URI}"><wps:wsp><wps:cNvSpPr txBox="1"/>'
        f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        f'<a:ln w="{int(width_pt * EMU)}"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>'
        f'<a:prstDash val="{dash}"/><a:miter lim="800000"/></a:ln></wps:spPr>'
        f'<wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="16"/></w:rPr>'
        f'<w:t xml:space="preserve">{label}</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
        f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" '
        f'anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>'
    )


def anchored(idx, dash, width_pt, geom="rect"):
    w, h = int(130 * EMU), int(45 * EMU)
    return (
        f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{idx}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="column"><wp:posOffset>{int(6 * EMU)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{int(16 * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{w}" cy="{h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{idx}" name="Shape {idx}"/><wp:cNvGraphicFramePr/>'
        f'{wsp(w, h, dash, width_pt, f"{dash} {width_pt:g}pt", geom)}</wp:anchor></w:drawing></w:r>'
    )


def cell(idx, dash, width_pt, geom="rect"):
    return (
        '<w:tc><w:tcPr><w:tcW w:w="3118" w:type="dxa"/></w:tcPr>'
        f'<w:p>{anchored(idx, dash, width_pt, geom)}<w:r><w:t>Cell {idx}</w:t></w:r></w:p></w:tc>'
    )


rows = [
    [("dash", 1.0), ("dashDot", 1.0, "roundRect"), ("solid", 1.0)],
]

table = (
    '<w:tbl><w:tblPr><w:tblW w:w="9354" w:type="dxa"/><w:tblBorders>'
    + "".join(f'<w:{s} w:val="single" w:sz="4" w:color="A0A0A0"/>' for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    + '</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>'
    + '<w:gridCol w:w="3118"/>' * 3
    + "</w:tblGrid>"
    + "".join(
        '<w:tr><w:trPr><w:trHeight w:val="1600" w:hRule="exact"/></w:trPr>'
        + "".join(cell(r * 3 + c + 1, *spec) for c, spec in enumerate(row))
        + "</w:tr>"
        for r, row in enumerate(rows)
    )
    + "</w:tbl>"
)

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "<w:p><w:r><w:t>Text boxes anchored in table cells</w:t></w:r></w:p>"
    + table
    + "<w:p/>"
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
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
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
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
    f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
print(f"wrote {out}")
