"""Generate autofit-textbox-wrap: text wraps around an auto-fitted text box's fitted height

A text box whose body is set to resize to its text (a:spAutoFit) is drawn as
tall as its text and insets need, whatever height its stored extent says, and
Word wraps the text after it around that fitted height too. A box without
auto-fit keeps its extent.

Each case: a label, an empty paragraph anchoring a 450pt-wide box 20pt below
it holding one line of text (fitted height 18.7pt), then "After n":
1. wrapSquare, extent 60pt, spAutoFit: After 1 starts below the 18.7pt box.
2. wrapSquare, extent 10pt, spAutoFit: the box grows to 18.7pt, After 2 below it.
3. wrapSquare, extent 60pt, noAutofit: After 3 starts below the 60pt box.
"""
import pathlib, zipfile
out = pathlib.Path(__file__).parent
W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
     'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
     'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
     'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
EMU = 12700
ids = iter(range(1, 999))


def para(inner="", text="", before=0):
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>{inner}{run}</w:p>'


def box(ext_h, fit, off=20, label="Box text"):
    n = next(ids)
    cx, cy = 450 * EMU, int(ext_h * EMU)
    fit_xml = "<a:spAutoFit/>" if fit else "<a:noAutofit/>"
    return ('<w:r><w:drawing>'
            f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{n}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{int(off * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/>'
            f'<wp:docPr id="{n}" name="Text Box {n}"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
            f'<a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="DDDDDD"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>'
            f'<wps:txbx><w:txbxContent>{para(text=label)}</w:txbxContent></wps:txbx>'
            f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t">{fit_xml}</wps:bodyPr>'
            '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


cases = [("square", 60, True), ("square", 10, True), ("square", 60, False)]
body = ""
for n, (wrap, h, fit) in enumerate(cases, 1):
    body += para(text=f"Case {n}: {wrap} extent {h} fit {fit}") + para(box(h, fit)) + para(text=f"After {n}") + para()
doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1418" w:bottom="1134" w:left="1418" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>')
styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>')
settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
ov = lambda p, t: f'<Override PartName="/word/{p}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{t}+xml"/>'
ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      + ov("document.xml", "document.main") + ov("styles.xml", "styles") + ov("settings.xml", "settings") + "</Types>")
rels = lambda items: ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                      + "".join(f'<Relationship Id="{i}" Type="{REL}/{t}" Target="{g}"/>' for i, t, g in items) + "</Relationships>")
with zipfile.ZipFile(out / "input.docx", "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", ct)
    z.writestr("_rels/.rels", rels([("rId1", "officeDocument", "word/document.xml")]))
    z.writestr("word/document.xml", doc)
    z.writestr("word/_rels/document.xml.rels", rels([("rId1", "styles", "styles.xml"), ("rId2", "settings", "settings.xml")]))
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
print(f"wrote {out / 'input.docx'}")
