"""Generate line-anchor-follows-push: line-relative drawings follow a pushed paragraph

A paragraph whose first line would reach into a band (a page-anchored
topAndBottom shape) or a page-anchored floating table with no room beside it
goes below it. Its drawings positioned relativeFrom="paragraph" stay where the
paragraph stood; those positioned relativeFrom="line" go with the paragraph and
count from its new top (above its own space before, which stays below the
band). Each case holds a "line" box at the left and a "paragraph" box at 150pt,
both wrapNone with the same offset, in its own section and page; exact 12pt
lines from the 56.7pt top margin, the band 100 to 200pt down:
1. Pushed below the band, offset -30pt: line box at 170pt, paragraph box 62.7.
2. 24pt space before, pushed below the band, offset 0: line box at 200pt.
3. Pushed below a floating table (200 to 260pt, 150pt above and 15pt below it
   kept free), offset -20pt: the line box at 255.5pt (line box only).
4. As 1, the line box a group (wrapNone) of a text box and a picture joined
   by a connector: the group at 170pt, the connector with it at 180pt.
5. As 1, a topAndBottom group of a 30pt text box 50pt below the line: at
   250pt, the lines after it skip 250 to 280pt.
6. As 1, the line box a lone straight connector: at 170pt.
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
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
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


def boxes(off, paragraph_box=True):
    out = shape("line", off, 0, 90, 20, "<wp:wrapNone/>", "C0C0C0", "line")
    if paragraph_box:
        out += shape("paragraph", off, 150, 90, 20, "<wp:wrapNone/>", "C0C0C0", "paragraph")
    return out


def band():
    return shape("page", 100, 0, 400, 100, "<wp:wrapTopAndBottom/>", "E0E0E0", rel_h="margin")


def floating_table():
    borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ("top", "left", "bottom", "right"))
    return (
        '<w:tbl><w:tblPr><w:tblpPr w:topFromText="3000" w:bottomFromText="300" w:vertAnchor="page" '
        'w:horzAnchor="margin" w:tblpY="4000"/><w:tblW w:w="9050" w:type="dxa"/>'
        f'<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9050"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="1200" w:hRule="exact"/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="9050" w:type="dxa"/></w:tcPr>{para(text="Floating table")}</w:tc></w:tr></w:tbl>'
    )


def png():
    """A grey square, drawn here."""
    w = h = 20
    rows = b"".join(b"\x00" + bytes((150, 150, 150)) * w for _ in range(h))
    chunk = lambda t, d: struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(rows))
        + chunk(b"IEND", b"")
    )


def group(off, wrap, box_h, with_picture):
    n = next(ids)
    w, h = 130, box_h
    emu = lambda v: int(v * EMU)
    xfrm = lambda x, y, cx, cy, ch=False: (
        f'<a:xfrm><a:off x="{emu(x)}" y="{emu(y)}"/><a:ext cx="{emu(cx)}" cy="{emu(cy)}"/>'
        + (f'<a:chOff x="0" y="0"/><a:chExt cx="{emu(cx)}" cy="{emu(cy)}"/>' if ch else "")
        + "</a:xfrm>"
    )
    box = (
        f'<wps:wsp><wps:cNvPr id="{n}1" name="Group box {n}"/><wps:cNvSpPr txBox="1"/>'
        f'<wps:spPr>{xfrm(0, 0, 90, box_h)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="C0C0C0"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
        f"<wps:txbx><w:txbxContent>{para(text='group')}</w:txbxContent></wps:txbx>"
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="36000" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr></wps:wsp>'
    )
    picture = (
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}2" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr>{xfrm(110, 0, 20, 20)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
    ) if with_picture else ""
    connector = (
        f'<wps:wsp><wps:cNvPr id="{n}3" name="Group line {n}"/><wps:cNvCnPr/>'
        f'<wps:spPr>{xfrm(90, 10, 20, 0)}<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom>'
        '<a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>'
    ) if with_picture else ""
    return (
        "<w:r><w:drawing>"
        f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{n}" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="line"><wp:posOffset>{emu(off)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{emu(w)}" cy="{emu(h)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap}'
        f'<wp:docPr id="{n}" name="Group {n}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">'
        f"<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr>{xfrm(0, 0, w, h, ch=True)}</wpg:grpSpPr>{box}{picture}{connector}</wpg:wgp>"
        "</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )


def connector(off):
    n = next(ids)
    cx = int(90 * EMU)
    return (
        "<w:r><w:drawing>"
        f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{n}" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="line"><wp:posOffset>{int(off * EMU)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{cx}" cy="0"/><wp:effectExtent l="0" t="12700" r="0" b="12700"/><wp:wrapNone/>'
        f'<wp:docPr id="{n}" name="Line {n}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="0"/></a:xfrm>'
        '<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="25400"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
        "<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )


def tail(case, sect):
    return "".join(para(text=f"Case {case} line {i}") for i in range(5, 8)) + para(text=f"Case {case} end", sect=sect)


cases = [
    para(band(), "Case 1 line 1") + para(text="Case 1 line 2") + para(text="Case 1 line 3")
    + para(boxes(-30), "Case 1 anchor") + tail(1, SECT),
    para(band(), "Case 2 line 1") + para(text="Case 2 line 2")
    + para(boxes(0), "Case 2 anchor", before=480) + tail(2, SECT),
    floating_table() + para(boxes(-20, paragraph_box=False), "Case 3 anchor") + tail(3, SECT),
    para(band(), "Case 4 line 1") + para(text="Case 4 line 2") + para(text="Case 4 line 3")
    + para(group(-30, "<wp:wrapNone/>", 20, True), "Case 4 anchor") + tail(4, SECT),
    para(band(), "Case 5 line 1") + para(text="Case 5 line 2") + para(text="Case 5 line 3")
    + para(group(50, "<wp:wrapTopAndBottom/>", 30, False), "Case 5 anchor") + tail(5, SECT),
    para(band(), "Case 6 line 1") + para(text="Case 6 line 2") + para(text="Case 6 line 3")
    + para(connector(-30), "Case 6 anchor") + tail(6, ""),
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
    '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
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
    z.writestr("word/_rels/document.xml.rels", relationships([("rId1", "styles", "styles.xml"), ("rId2", "settings", "settings.xml"), ("rIdImg", "image", "media/p.png")]))
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/media/p.png", png())
print(f"wrote {out}")
