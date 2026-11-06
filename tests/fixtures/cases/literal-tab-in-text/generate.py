"""Generate literal-tab-in-text: TAB characters inside w:t

Some generators (template engines that write the text verbatim) put a TAB
character straight into w:t instead of a w:tab element. Word lays it out
exactly like w:tab: it jumps to the next tab stop.

Tests:
- Literal TAB against a custom left stop, as in "Label:<TAB>Value"
- Several literal TABs in one w:t against left, center, right and decimal stops
- Literal TAB at the start and at the end of a w:t
- Literal TABs falling on the default tab stops (no custom stops)
- Literal TAB and w:tab mixed in one paragraph
- Literal TAB in a table cell
- Hidden (w:vanish) text with a literal TAB and a w:tab takes no space
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def run(text, bold=False, hidden=False):
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else "<w:rPr><w:vanish/></w:rPr>" if hidden else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(runs, tabs=(), before=0):
    stops = "".join(f'<w:tab w:val="{v}" w:pos="{p}"/>' for v, p in tabs)
    tabs_xml = f"<w:tabs>{stops}</w:tabs>" if tabs else ""
    return f'<w:p><w:pPr>{tabs_xml}<w:spacing w:before="{before}"/></w:pPr>{"".join(runs)}</w:p>'


body = [
    para([run("Literal TAB characters inside w:t", bold=True)]),
    para([run("Check-in:\tWednesday, 15 January 2025, 10:00")], [("left", 1418)], 240),
    para([run("Item\tCentered\tRight\t12.50")],
         [("left", 2000), ("center", 5000), ("right", 7500), ("decimal", 9000)], 240),
    para([run("\tLeading TAB")], [("left", 1418)], 240),
    para([run("Trailing TAB\t"), run("after the run")], [("left", 4000)]),
    para([run("A\tB\tC\tD\tE")], (), 240),
    para([run("Literal:\tone"), '<w:r><w:tab/><w:t>element</w:t></w:r>', run("\tliteral again")],
         [("left", 2000), ("left", 4500), ("left", 7000)], 240),
    para([run("Visible"), run("\thidden\t", hidden=True),
          '<w:r><w:rPr><w:vanish/></w:rPr><w:tab/></w:r>', run(" text\tafter")], [("left", 2835)], 240),
]

cell = lambda xml: f'<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>{xml}</w:tc>'
table = (
    '<w:tbl><w:tblPr><w:tblW w:w="4500" w:type="dxa"/><w:tblBorders>'
    + "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    + '</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/></w:tblGrid>'
    + "<w:tr>" + cell(para([run("Egg\tOil\tVinegar")], [("left", 1200), ("left", 2400)]))
    + "</w:tr></w:tbl>"
)

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
    + "".join(body) + para([], (), 240) + table
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)

settings = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/></w:settings>'

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
