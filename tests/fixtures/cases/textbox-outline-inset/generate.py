"""Generate textbox-outline-inset: a text box's outline pushes its text inward

Word lays a text box's text out inside its body insets (lIns/tIns/rIns/bIns)
measured from the inner edge of the outline, so the text moves in by half
the outline width on every side; a box without an outline keeps the plain
insets. An invisible outline (a:ln with a width and noFill) moves the text in
the same way.

Tests:
- 6pt outline with default insets, zero insets, and right/bottom alignment
- Invisible 8pt outline with right-aligned, bottom-anchored text
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


def box(idx, x_pt, y_pt, w_pt, h_pt, line_pt, label, insets=True, jc="left", anchor="t", hidden_line_pt=0):
    w, h = int(w_pt * EMU), int(h_pt * EMU)
    ln = (
        f'<a:ln w="{int(line_pt * EMU)}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
        if line_pt
        else f'<a:ln w="{int(hidden_line_pt * EMU)}"><a:noFill/></a:ln>' if hidden_line_pt
        else "<a:ln><a:noFill/></a:ln>"
    )
    ins = ('lIns="91440" tIns="45720" rIns="91440" bIns="45720"' if insets
           else 'lIns="0" tIns="0" rIns="0" bIns="0"')
    return (
        f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{idx}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="margin"><wp:posOffset>{int(x_pt * EMU)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="margin"><wp:posOffset>{int(y_pt * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{w}" cy="{h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{idx}" name="Box {idx}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic><a:graphicData uri="{WPS_URI}"><wps:wsp><wps:cNvSpPr txBox="1"/>'
        f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>{ln}</wps:spPr>'
        f'<wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="{jc}"/></w:pPr>'
        f'<w:r><w:t xml:space="preserve">{label}</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
        f'<wps:bodyPr rot="0" vert="horz" wrap="square" {ins} anchor="{anchor}" anchorCtr="0">'
        f'<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'
    )


# Keep the retained boxes at their original page positions.
shapes = [
    box(1, 0, 300, 150, 55, 6, "Text 6pt"),
    box(2, 165, 300, 150, 55, 6, "No insets 6pt", insets=False),
    box(3, 330, 300, 150, 55, 6, "Right 6pt", jc="right", anchor="b"),
    box(4, 330, 460, 150, 55, 0, "Invisible 8pt", hidden_line_pt=8, jc="right", anchor="b"),
]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    f'<w:p>{"".join(shapes)}</w:p>'
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
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
