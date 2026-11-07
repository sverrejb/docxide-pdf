"""Generate shape-outline-dash: a:prstDash on text box and shape outlines

Word strokes a shape outline with the dash pattern of its a:ln/a:prstDash, in
multiples of the line width.

Tests:
- Every non-solid preset dash on a floating text box, at 1pt
- dash at 0.5pt and 3pt (the pattern scales with the width)
- An inline text box with a dashed 1pt outline (a framed note, as in letters)
- A dashed ellipse without text
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

DASHES = ["dot", "dash", "lgDash", "dashDot", "lgDashDot", "lgDashDotDot",
          "sysDash", "sysDot", "sysDashDot", "sysDashDotDot"]


def wsp(w, h, dash, width_pt, label, geom="rect"):
    text = (
        f'<wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="16"/></w:rPr>'
        f'<w:t xml:space="preserve">{label}</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
        if label else ""
    )
    return (
        f'<a:graphic><a:graphicData uri="{WPS_URI}"><wps:wsp><wps:cNvSpPr txBox="1"/>'
        f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="{geom}"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        f'<a:ln w="{int(width_pt * EMU)}"><a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>'
        f'<a:prstDash val="{dash}"/><a:miter lim="800000"/></a:ln></wps:spPr>'
        f'{text}<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" '
        f'anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>'
    )


def anchored(idx, x_pt, y_pt, w_pt, h_pt, dash, width_pt, label, geom="rect"):
    w, h = int(w_pt * EMU), int(h_pt * EMU)
    return (
        f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{idx}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="margin"><wp:posOffset>{int(x_pt * EMU)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="margin"><wp:posOffset>{int(y_pt * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{w}" cy="{h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{idx}" name="Shape {idx}"/><wp:cNvGraphicFramePr/>'
        f'{wsp(w, h, dash, width_pt, label, geom)}</wp:anchor></w:drawing></w:r>'
    )


def inline(idx, w_pt, h_pt, dash, width_pt, label):
    w, h = int(w_pt * EMU), int(h_pt * EMU)
    return (
        f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{w}" cy="{h}"/><wp:effectExtent l="0" t="0" r="12700" b="12700"/>'
        f'<wp:docPr id="{idx}" name="Note {idx}"/><wp:cNvGraphicFramePr/>'
        f'{wsp(w, h, dash, width_pt, label)}</wp:inline></w:drawing></w:r>'
    )


shapes = []
for i, dash in enumerate(DASHES):
    col, row = i % 3, i // 3
    shapes.append(anchored(i + 1, col * 160, row * 70, 140, 50, dash, 1.0, f"{dash} 1pt"))
extra = [("dash", 0.5), ("dash", 3.0)]
for j, (dash, width) in enumerate(extra):
    shapes.append(anchored(20 + j, (j % 2) * 240, 290 + (j // 2) * 70, 220, 50, dash, width, f"{dash} {width}pt"))
shapes.append(anchored(30, 330, 430, 140, 80, "lgDashDot", 1.5, "", "ellipse"))

body = (
    f'<w:p>{"".join(shapes)}</w:p>'
    + "".join("<w:p/>" for _ in range(30))
    + f'<w:p>{inline(40, 300, 60, "dash", 1.0, "Inline note with a dashed outline")}</w:p>'
)

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + body
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W_NS}"><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
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
