"""Generate header-footer-border-groups: bordered paragraphs join in headers and footers

Adjacent header or footer paragraphs with the same borders and box edges form
one bordered group, as in the body: one top rule above the first, one bottom
rule below the last, the between border (if any) between them, side borders
running unbroken through the group.

Tests:
- Footer: a signature box of four boxed paragraphs, the middle two shaded
- Footer: a boxed paragraph with a between border after them (its own group)
- Footer: two boxed paragraphs that both set a between border (rule between)
- Header: three paragraphs with only a bottom border (one rule under the last)
- Header: a boxed group of two with different indents but the same box edges
- Body and header: boxed paragraphs where only some set a between border
  (the group still joins; a between rule follows each paragraph that sets it)
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def bdr(sides):
    return "<w:pBdr>" + "".join(
        f'<w:{s} w:val="single" w:sz="6" w:space="1" w:color="000000"/>' for s in sides
    ) + "</w:pBdr>"


def para(text, sides=(), shd=None, ind="", tabs=()):
    stops = "".join(f'<w:tab w:val="left" w:pos="{p}"/>' for p in tabs)
    tabs_xml = f"<w:tabs>{stops}</w:tabs>" if tabs else ""
    shd_xml = f'<w:shd w:val="pct10" w:color="auto" w:fill="auto"/>' if shd else ""
    runs = "".join(
        "<w:r><w:tab/></w:r>" if part == "\t" else f'<w:r><w:t xml:space="preserve">{part}</w:t></w:r>'
        for part in text
    )
    return f'<w:p><w:pPr>{bdr(sides) if sides else ""}{shd_xml}{tabs_xml}{ind}</w:pPr>{runs}</w:p>'


box = ("top", "left", "bottom", "right")
boxb = box + ("between",)
cols = (2835, 7088)
footer_paras = [
    para([]),
    para(["Text labels", "\t", "Sample set 0123456789", "\t", "Page 1"], box, tabs=cols),
    para(["\t", ".", "\t", "\t"], box, shd=True, tabs=cols),
    para(["\t", "\t", "\t"], box, shd=True, tabs=cols),
    para(["First text column ", "\t", "Middle column", "\t", "Final"], boxb, tabs=cols),
    para(["\t", "\t", "\t"], box, tabs=(851,) + cols),
    para([]),
    para(["Grouped with a between rule"], boxb),
    para(["second paragraph of that group"], boxb),
]
header_paras = [
    para(["Sample header group, aligned text 001"], ["bottom"]),
    para(["Examples of grouped text rows"], ["bottom"]),
    para(["Sample text 001"], ["bottom"]),
    para([]),
    para(["Boxed, left 1134"], box, ind='<w:ind w:left="1134" w:right="1134"/>'),
    para(["Boxed, left 1701 hanging 567"], box, ind='<w:ind w:left="1701" w:right="1134" w:hanging="567"/>'),
    para([]),
    para(["Mixed A: no between"], box),
    para(["Mixed B: between"], boxb),
    para(["Mixed C: no between"], box),
]
mixed = [
    para(["Body mixed A: no between"], box),
    para(["Body mixed B: between"], boxb),
    para(["Body mixed C: no between"], box),
    para(["Body mixed D: no between"], box),
    para([]),
]

filler = "<w:p><w:r><w:t>Body text of the page between the header and the footer.</w:t></w:r></w:p>"
document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W} {R}><w:body>'
    + filler
    + "".join(mixed)
    + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId4"/>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="3402" w:right="1134" w:bottom="3969" w:left="1418" w:header="567" w:footer="284" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)
header = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{"".join(header_paras)}</w:hdr>'
footer = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {W}>{"".join(footer_paras)}</w:ftr>'

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)

settings = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/></w:settings>'

ct = "application/vnd.openxmlformats-officedocument.wordprocessingml"
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    f'<Override PartName="/word/document.xml" ContentType="{ct}.document.main+xml"/>'
    f'<Override PartName="/word/styles.xml" ContentType="{ct}.styles+xml"/>'
    f'<Override PartName="/word/settings.xml" ContentType="{ct}.settings+xml"/>'
    f'<Override PartName="/word/header1.xml" ContentType="{ct}.header+xml"/>'
    f'<Override PartName="/word/footer1.xml" ContentType="{ct}.footer+xml"/>'
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
    f'<Relationship Id="rId3" Type="{rel}/header" Target="header1.xml"/>'
    f'<Relationship Id="rId4" Type="{rel}/footer" Target="footer1.xml"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/header1.xml", header)
    z.writestr("word/footer1.xml", footer)
print(f"wrote {out}")
