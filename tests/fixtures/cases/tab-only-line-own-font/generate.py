"""Generate tab-only-line-own-font: each tab-only line sized by what ends it

A line holding only an underlined tab is sized by the character that ends
it, its ascent included: its w:br's font, or the paragraph mark's on the
last line, whatever sizes the paragraph's other lines (Word for Mac behavior,
online export): under a 20pt and an 11pt break two such lines sit 18.1pt and
9.96pt below their tops (F), beside text lines too. The tab's underline keeps
the tab run's own offset and thickness. Text boxes do the same (C, F), and a
box paragraph of only a 20pt tab run (H20) is as tall as its 11pt mark where
a bottom anchor or a shape that fits its text measures it.

Under at-least spacing the line keeps the boxed baseline of the font that
ends it (K, N, O; M's 12pt minimum is below Arial 11's natural line). A
tab that only spaces of another size follow up to the break or the mark
underlines at its own size (P, Q).

Each case sits between two 11pt lines; C..J vary the break, mark and text
sizes (A = Arial, T = Times New Roman, pt). After the body one page repeats C
and F in top-anchored text boxes, the next holds H20 and P in bottom-anchored
boxes and H20 in an autofit box.
"""

import pathlib
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = (f'xmlns:w="{W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"')
FN = {"A": "Arial", "T": "Times New Roman", "C": "Calibri"}


def rpr(spec, u=False):
    f, size = FN[spec[0]], int(spec[1:]) * 2
    return (f'<w:rPr><w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}"/><w:sz w:val="{size}"/>'
            + ('<w:u w:val="single"/>' if u else "") + "</w:rPr>")


def tab(s="A11"):
    return f'<w:r>{rpr(s, True)}<w:tab/></w:r>'


def br(s="A11"):
    return f'<w:r>{rpr(s)}<w:br/></w:r>'


def text(t, s="A11"):
    return f'<w:r>{rpr(s)}<w:t xml:space="preserve">{t}</w:t></w:r>'


def para(content, mark="A11", before=0, spacing=""):
    return (f'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="3000"/></w:tabs>'
            f'<w:spacing w:before="{before}"{spacing}/>{rpr(mark)}</w:pPr>{content}</w:p>')


def at_least(twips):
    return f' w:line="{twips}" w:lineRule="atLeast"'


CASES = {
    "C": (text("Name", "A20") + br("A11") + tab(), "A11"),
    "E": (tab() + br("T20") + text("Name"), "A11"),
    "F": (tab() + br("A20") + tab() + br("A11") + tab(), "A16"),
    "I": (text("Name") + br() + tab() + br("A20") + text("Name"), "A11"),
    "J": (text("Name") + br() + tab(), "T20"),
    "K": (tab(), "A11", at_least(600)),
    "M": (tab(), "A11", at_least(240)),
    "N": (tab() + br("A20") + text("Name"), "A11", at_least(600)),
    "O": (text("Name", "A20") + br() + tab(), "A11", at_least(600)),
    "P": (tab() + text(" ", "A20") + br() + text("Name"), "A11"),
    "Q": (tab() + text("  ", "A20"), "A11"),
}
BOX_ONLY = {"H20": (tab("A20"), "A11")}


def case(name):
    content, mark, *spacing = (CASES | BOX_ONLY)[name]
    return [para(text(f"{name} above")), para(content, mark, 0, *spacing), para(text(f"{name} below"))]


def textbox(idx, x_pt, y_pt, paras, anchor="t", fit="<a:noAutofit/>"):
    emu = lambda pt: int(pt * 12700)
    cx, cy = emu(240), emu(200)
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{idx}" '
            'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{emu(x_pt)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{emu(y_pt)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{cx}" cy="{cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="{idx}" name="Box {idx}"/><wp:cNvGraphicFramePr/>'
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="3175"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></wps:spPr>'
            f'<wps:txbx><w:txbxContent>{"".join(paras)}</w:txbxContent></wps:txbx>'
            f'<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="{anchor}" anchorCtr="0">{fit}</wps:bodyPr>'
            '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


body = []
for name in CASES:
    body += case(name)
body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
boxes = []
for i, name in enumerate(["C", "F"]):
    x, y = 40 + (i % 2) * 270, 60 + (i // 2) * 230
    boxes.append(textbox(i + 1, x, y, case(name)))
body.append(f'<w:p>{"".join(boxes)}</w:p>')
body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
boxes = []
for i, (name, anchor, fit) in enumerate([("H20", "b", "<a:noAutofit/>"), ("H20", "t", "<a:spAutoFit/>"),
                                         ("P", "b", "<a:noAutofit/>")]):
    x, y = 40 + (i % 2) * 270, 60 + (i // 2) * 230
    boxes.append(textbox(i + 11, x, y, case(name), anchor, fit))
body.append(f'<w:p>{"".join(boxes)}</w:p>')

document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(body)}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>'
            "</w:sectPr></w:body></w:document>")
styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W}"><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          "</w:docDefaults></w:styles>")
settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W}"><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            "</w:compat></w:settings>")
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    "</Types>")
rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
package_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>')
document_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/></Relationships>')
out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
print(f"wrote {out}")
