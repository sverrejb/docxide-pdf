"""Generate line-anchor-overrides: where line-relative drawings hang when the flow moves

Drawings positioned relativeFrom="line" follow their paragraph below a band
that pushed it (fixture line-anchor-follows-push). Each case in its own section
and page, exact 12pt lines from the 56.7pt top margin:
1. A band 100 to 200pt down pushes the anchor below it. Its topAndBottom
   line box 50pt below the line (30pt tall) sits at 250pt and the lines after
   it skip 250 to 280pt; a topAndBottom paragraph box with the same offset
   stays inside the band.
2. Lines filling the page, then a floating table 1pt above its anchor's
   text that does not fit and moves whole to the next page with its anchor,
   which holds a line box and a paragraph box at offset 0: the paragraph box
   hangs from the page top, the line box from the line below the table.
"""

import pathlib
import zipfile

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 12700
SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="1418" w:bottom="1134" w:left="1418" w:header="567" w:footer="567" w:gutter="0"/>'
    "</w:sectPr>"
)
ids = iter(range(1, 1000))


def para(inner="", text="", before=0, sect=""):
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return (f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="0" w:line="240" w:lineRule="exact"/>{sect}</w:pPr>'
            f"{inner}{run}</w:p>")


def shape(rel_v, off_pt, x_pt, w_pt, h_pt, wrap, fill, label="", rel_h="column"):
    n = next(ids)
    cx, cy = int(w_pt * EMU), int(h_pt * EMU)
    text = f"<wps:txbx><w:txbxContent>{para(text=label)}</w:txbxContent></wps:txbx>" if label else ""
    sp_attr = ' txBox="1"' if label else ""
    line = '<a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>' if label else "<a:ln><a:noFill/></a:ln>"
    return (
        "<w:r><w:drawing>"
        f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{n}" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="{rel_h}"><wp:posOffset>{int(x_pt * EMU)}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="{rel_v}"><wp:posOffset>{int(off_pt * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap}'
        f'<wp:docPr id="{n}" name="Shape {n}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp><wps:cNvSpPr{sp_attr}/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>{line}</wps:spPr>'
        f'{text}<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="36000" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )


def boxes(off, wrap="<wp:wrapNone/>", h=20, paragraph_box=True, line_box=True):
    out = shape("line", off, 0, 90, h, wrap, "C0C0C0", "line") if line_box else ""
    if paragraph_box:
        out += shape("paragraph", off, 150, 90, h, wrap, "C0C0C0", "paragraph")
    return out


def band():
    return shape("page", 100, 0, 400, 100, "<wp:wrapTopAndBottom/>", "E0E0E0", rel_h="margin")


def text_table(y):
    borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ("top", "left", "bottom", "right"))
    return (
        '<w:tbl><w:tblPr><w:tblpPr w:leftFromText="180" w:rightFromText="180" w:vertAnchor="text" '
        f'w:horzAnchor="margin" w:tblpY="{y}"/><w:tblW w:w="9050" w:type="dxa"/>'
        f'<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9050"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="2000" w:hRule="exact"/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="9050" w:type="dxa"/></w:tcPr>{para(text="Floating table")}</w:tc></w:tr></w:tbl>'
    )


def tail(case, sect, n=8):
    return "".join(para(text=f"Case {case} line {i}") for i in range(5, 5 + n)) + para(text=f"Case {case} end", sect=sect)


top_bottom = "<wp:wrapTopAndBottom/>"
cases = [
    para(band(), "Case 1 line 1") + para(text="Case 1 line 2") + para(text="Case 1 line 3")
    + para(boxes(50, top_bottom, h=30), "Case 1 anchor") + tail(1, SECT),
    "".join(para(text=f"Case 2 fill {i}") for i in range(1, 57))
    + text_table(-20) + para(boxes(0), "Case 2 anchor") + tail(2, ""),
]

document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
            + "".join(cases) + SECT + "</w:body></w:document>")
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
    '<Default Extension="xml" ContentType="application/xml"/>'
    + override("document.xml", "document.main") + override("styles.xml", "styles") + override("settings.xml", "settings")
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
    z.writestr("word/_rels/document.xml.rels", relationships([("rId1", "styles", "styles.xml"), ("rId2", "settings", "settings.xml")]))
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
print(f"wrote {out}")
