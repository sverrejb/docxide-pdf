"""Generate ul-trail-space-runs and ul-trail-space-bidi

Both set w:compat/w:ulTrailSpace, under which Word underlines the underlined
spaces a line ends in (see ul-trail-space-on).

ul-trail-space-runs: each trailing space keeps its own run's underline.
- An underlined space followed by two plain ones, and two plain spaces
  followed by an underlined one (Word leaves the gap)
- The same before a manual break
- Red then blue, single then double, Arial then 16pt Times New Roman spaces
- Underlined spaces before an inline picture too wide for the rest of the
  line, from one run and after plain spaces
- Justified lines ending in underlined spaces of other colours, before a
  manual break and wrapping: the spaces neither take a share of the
  stretch nor show past the right margin
- Lines holding only underlined spaces, before the paragraph mark, a manual
  break, centred and in a table cell: underlined from the
  line's start, as high as the paragraph mark makes the line; then a dozen
  such paragraphs of different lengths, left and centred, and a paragraph
  of ten such lines ended by manual breaks
- Underlined spaces after a plain tab, alone, centred and after the last
  manual break of a tabbed paragraph: underlined from the tab stop
- A justified line beside a picture wrapping both sides that ends in an
  underlined space: the text right of the picture still reaches its edge

ul-trail-space-bidi: bidi ordering and w:bidi paragraphs.
- Left-to-right paragraphs holding Hebrew and ending in underlined spaces,
  left and centred: the spaces neither shift the text nor
  pass the right margin
- w:bidi paragraphs of w:rtl Hebrew ending in underlined spaces, aligned to
  their start, left and centre, before a manual break, mixed with plain
  spaces, wrapping and justified: the spaces are underlined left of the
  text, never past the left margin
- The same after a tab, and a w:bidi line of underlined spaces only,
  alone, after a plain tab and after the last manual break of a tabbed
  paragraph
"""

import io
import pathlib
import sys
import zipfile

from PIL import Image, ImageDraw

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
     'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"')
U = '<w:u w:val="single"/>'
RED = '<w:color w:val="C00000"/>'
BLUE = '<w:color w:val="0070C0"/>'
TIMES = '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="32"/>'
COUR = '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="24"/>'
DBL = '<w:u w:val="double"/>'
RTL = "<w:rtl/>"
BIDI = "<w:bidi/>"
HEB = "שלום עולם"


def run(text, extra=U):
    return f'<w:r><w:rPr>{extra}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


def br(extra=U):
    return f"<w:r><w:rPr>{extra}</w:rPr><w:br/></w:r>"


def tab(extra=U):
    return f"<w:r><w:rPr>{extra}</w:rPr><w:tab/></w:r>"


def cell(content):
    return ('<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblBorders>'
            '<w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/></w:tblBorders></w:tblPr>'
            '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/>'
            f'</w:tcPr>{content}</w:tc></w:tr></w:tbl>')


def para(runs, jc=None, extra=""):
    j = f'<w:jc w:val="{jc}"/>' if jc else ""
    return f'<w:p><w:pPr>{extra}{j}</w:pPr>{"".join(runs)}</w:p>'


def picture(width, height, ident):
    w, h = round(width * 12700), round(height * 12700)
    return (f'<w:r><w:drawing><wp:inline><wp:extent cx="{w}" cy="{h}"/>'
            f'<wp:docPr id="{ident}" name="Picture {ident}"/>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{ident}" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


def floating(width, height, x, ident):
    w, h, ox = round(width * 12700), round(height * 12700), round(x * 12700)
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            'relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="column"><wp:posOffset>{ox}</wp:posOffset></wp:positionH>'
            '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{w}" cy="{h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapSquare wrapText="bothSides"/><wp:docPr id="{ident}" name="Picture {ident}"/>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{ident}" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>')


runs_body = [
    para([run("Underlined word "), run("  ", "")]),
    para([run("Underlined word"), run("  ", ""), run(" ")]),
    para([run("Underlined then plain "), run("  ", ""), br(""), run("after the break")]),
    para([run("Red word ", U + RED), run("   ", U + BLUE)]),
    para([run("Single ", U), run("   ", DBL)]),
    para([run("Arial word ", U), run("   ", U + TIMES)]),
    para([run("One two three "), run(" ", U + RED), run(" ", U + BLUE), br(), run("after the break")], "both"),
    para([r for i in range(30) for r in (run(f"word{i}"), run(" ", U + RED), run(" ", U + BLUE))], "both"),
    para([run("     ")]),
    para([run("     "), br(), run("after the break")]),
    para([run("     ")], "center"),
    cell(para([run("     ")])),
    *[para([run(" " * (4 + i % 5))], "center" if i % 3 == 1 else None) for i in range(12)],
    para([r for i in range(10) for r in (run(" " * (3 + i % 4)), br())] + [run("     ")]),
    para([run("Caption wraps   "), picture(430, 40, 1)]),
    para([run("Caption plain"), run("   ", ""), run(" "), picture(420, 40, 2)]),
    para([tab(""), run("     ")]),
    para([tab(""), run("     ")], "center"),
    para([run("Text", ""), tab(""), run("word", ""), br(""), run("     ")]),
    para([floating(100, 60, 100, 3), run("aaaa bbbb cccc dddd ffff", COUR), run(" ", U + COUR),
          run("eeeeeeeeeee", COUR)], "both", '<w:ind w:right="3354"/>'),
]

wrap = (HEB + " ") * 24
bidi_body = [
    para([run("Text " + HEB + "   ")]),
    para([run("Text " + HEB + "   ")], "center"),
    para([run(HEB + "   ", U + RTL)], None, BIDI),
    para([run(HEB + "   ", U + RTL)], "left", BIDI),
    para([run(HEB + "   ", U + RTL)], "center", BIDI),
    para([run(HEB + "   ", U + RTL), br(U + RTL), run(HEB, U + RTL)], None, BIDI),
    para([run(HEB, U + RTL), run("  ", RTL), run(" ", U + RTL)], None, BIDI),
    para([run(wrap, U + RTL)], None, BIDI),
    para([run(wrap, U + RTL)], "both", BIDI),
    para([run(HEB, U + RTL), tab(U + RTL), run(HEB + "   ", U + RTL)], None, BIDI),
    para([run(HEB, U + RTL), tab(U + RTL), run(HEB + "   ", U + RTL), br(U + RTL), run(HEB, U + RTL)], None, BIDI),
    para([run(HEB, U + RTL), tab(U + RTL), run(HEB + "   ", U + RTL)], "center", BIDI),
    para([run("     ", U + RTL)], None, BIDI),
    para([tab(RTL), run("     ", U + RTL)], None, BIDI),
    para([run(HEB, U + RTL), tab(RTL), run(HEB, U + RTL), br(RTL), run("     ", U + RTL)], None, BIDI),
]

variants = {"ul-trail-space-runs": runs_body, "ul-trail-space-bidi": bidi_body}

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)

settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
    '<w:defaultTabStop w:val="709"/><w:compat><w:ulTrailSpace/>'
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    "</w:compat></w:settings>"
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
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/>'
    f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/image1.png"/></Relationships>'
)


def placeholder_png():
    img = Image.new("RGB", (420, 40), (210, 225, 240))
    ImageDraw.Draw(img).rectangle([0, 0, 419, 39], outline=(40, 70, 120), width=3)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


cases = pathlib.Path(__file__).parent.parent
for name, body in variants.items():
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
        + "".join(body)
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
        "</w:sectPr></w:body></w:document>"
    )
    out = cases / name / "input.docx"
    out.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", package_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", document_rels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
        z.writestr("word/media/image1.png", placeholder_png())
    print(f"wrote {out}", file=sys.stderr)
