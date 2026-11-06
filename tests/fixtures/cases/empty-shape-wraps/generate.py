"""Generate empty-shape-wraps: an empty, unfilled, unlined shape still wraps text

Letterhead templates reserve the letterhead area with a page-wide text box that
has no text, no fill and no outline. Nothing of it prints, but Word still wraps
the body around it: the first lines start below it.

One section per case, each on its own page, the box 100pt tall at the page top:
1. Top and bottom wrap, invisible box
2. Square wrap, invisible box (as wide as the page: no room beside it)
3. No wrap, invisible box: the text starts at the top margin
4. Square wrap, invisible box at the left margin sized 100% of the text width
   (wp14:pctWidth) over a stale 40% extent: no room beside it, text below
"""

import pathlib
import zipfile

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="wp14"'
)
TEXT_WIDTH_EMU = (11906 - 1134 - 1418) * 635
WRAPS = {
    "tb": "<wp:wrapTopAndBottom/>",
    "square": '<wp:wrapSquare wrapText="bothSides"/>',
    "none": "<wp:wrapNone/>",
}
SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="567" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr>"
)


def box(wrap, shape_id, pct=None):
    cx, cy = 7560000, 1270000
    h_pos = '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>'
    size_rel = ""
    if pct is not None:
        cx = TEXT_WIDTH_EMU * 40 // 100
        h_pos = '<wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        size_rel = (
            f'<wp14:sizeRelH relativeFrom="margin"><wp14:pctWidth>{pct * 1000}</wp14:pctWidth></wp14:sizeRelH>'
            '<wp14:sizeRelV relativeFrom="page"><wp14:pctHeight>0</wp14:pctHeight></wp14:sizeRelV>'
        )
    return (
        '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
        f'relativeHeight="{shape_id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/>' + h_pos +
        '<wp:positionV relativeFrom="page"><wp:align>top</wp:align></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{WRAPS[wrap]}'
        f'<wp:docPr id="{shape_id}" name="Letterhead {shape_id}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
        '<a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent/></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t">'
        "<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>" + size_rel + "</wp:anchor></w:drawing></w:r>"
    )


def para(runs, sect=False):
    return f'<w:p><w:pPr><w:spacing w:line="260" w:lineRule="exact"/>{SECT if sect else ""}</w:pPr>{runs}</w:p>'


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


body = []
CASES = (("tb", None), ("square", None), ("none", None), ("square", 100))
for n, (wrap, pct) in enumerate(CASES, 1):
    body.append(para(box(wrap, n, pct) + run("Line 1")))
    body += [para(run(f"Line {i}"), sect=i == 8) for i in range(2, 9)]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body) + SECT + "</w:body></w:document>"
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
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
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
    f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
print(f"wrote {out}")
