"""Generate anchor-below-own-hanging-box: a band pushes an anchor into its own box

A paragraph anchors two topAndBottom textboxes: a band at the page top and a
paragraph-relative box hanging some distance below the paragraph. The band
pushes the paragraph down; the hanging box stays where the paragraph stood.
When the pushed line now reaches into that box, Word puts the line below the
box too, and the following paragraphs follow it.

One section per variant:
- Offset 30pt, 100pt band: the pushed line reaches into the box (goes below it)
- Offset 50pt, 100pt band: the pushed line fits above the box (stays)
- Offset 30pt, empty anchor paragraph (goes below it)
"""

import pathlib
import zipfile

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'
)
EMU = 12700
SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
    'w:left="1440" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'
)


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content, sect=False):
    ppr = '<w:spacing w:before="0" w:after="0" w:line="260" w:lineRule="exact"/>' + (SECT if sect else "")
    return f"<w:p><w:pPr>{ppr}</w:pPr>{content}</w:p>"


def textbox(i, y, h, w, rel_v, text, fill):
    """A topAndBottom textbox `y` pt below the page top or its paragraph."""
    return (
        f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        f'relativeHeight="{i}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="{rel_v}"><wp:posOffset>{y * EMU}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{w * EMU}" cy="{h * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/>'
        f'<wp:docPr id="{i}" name="Box {i}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>'
        f'<wps:txbx><w:txbxContent><w:p>{run(text)}</w:p></w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )


variants = [
    (100, 30, "Anchor pushed into its box, offset 30pt"),
    (100, 50, "Anchor above its box, offset 50pt"),
    (100, 30, ""),
]
body = []
for n, (band, offset, text) in enumerate(variants):
    boxes = (
        textbox(2 * n + 1, 0, band, 451, "page", f"Page band {band}pt", "D9E2F3")
        + textbox(2 * n + 2, offset, 100, 300, "paragraph", f"Box {offset}pt below the anchor", "E2EFD9")
    )
    body.append(para(boxes + (run(text) if text else "")))
    body.append(para(run("First paragraph after the anchor")))
    body.append(para(run("Second paragraph after the anchor"), sect=n < len(variants) - 1))

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body) + SECT + "</w:body></w:document>"
)

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)
settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/>'
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
