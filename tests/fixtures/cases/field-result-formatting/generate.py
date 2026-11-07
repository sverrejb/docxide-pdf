"""Generate field-result-formatting: a field's result keeps its run formatting

PAGE, NUMPAGES and STYLEREF results are laid out with the full formatting of
their runs, as any text: underlined, struck through, scaled (w:w), spaced
(w:spacing), in small caps and so on.

Tests:
- Header: an underlined "- Seite PAGE von NUMPAGES -" line after a tab
- Header: right-aligned "Seite PAGE von NUMPAGES" in bold at w:w 85
- Header: STYLEREF to Heading 1 with small caps, underline and strike
- Footer: PAGE with w:w 150 and 3pt character spacing, NUMPAGES double struck
- Body: PAGE and NUMPAGES with dotted underline, PAGE scaled to 50%
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def rpr(extra=""):
    return f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>{extra}</w:rPr>'


def text(t, extra=""):
    return f'<w:r>{rpr(extra)}<w:t xml:space="preserve">{t}</w:t></w:r>'


def field(instr, cached, extra=""):
    return (
        f'<w:r>{rpr(extra)}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rpr(extra)}<w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
        f'<w:r>{rpr(extra)}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rpr(extra)}<w:t>{cached}</w:t></w:r>'
        f'<w:r>{rpr(extra)}<w:fldChar w:fldCharType="end"/></w:r>'
    )


def tab(extra=""):
    return f'<w:r>{rpr(extra)}<w:tab/></w:r>'


def para(runs, ppr=""):
    return f'<w:p><w:pPr>{ppr}</w:pPr>{"".join(runs)}</w:p>'


u = '<w:u w:val="single"/>'
header_paras = [
    para([text("Guide to the summer lake walk", u), tab(u), text("- Page ", u), field("PAGE", "1", u),
          text(" of ", u), field("NUMPAGES", "1", u), text(" -", u)],
         '<w:tabs><w:tab w:val="left" w:pos="6804"/></w:tabs>'),
    para([text("Page ", '<w:w w:val="85"/>'), field("PAGE", "1", '<w:b/><w:w w:val="85"/>'),
          text(" of ", '<w:w w:val="85"/>'), field("NUMPAGES", "1", '<w:b/><w:w w:val="85"/>')],
         '<w:jc w:val="right"/>'),
    para([text("Section: "), field('STYLEREF "Heading 1"', "x", '<w:smallCaps/><w:u w:val="single"/><w:strike/>'),
          text(" (continued)")]),
]
footer_paras = [
    para([text("Wide page number "), field("PAGE", "1", '<w:w w:val="150"/><w:spacing w:val="60"/>'),
          text(" and a struck total "), field("NUMPAGES", "1", '<w:dstrike/>'), text(" end")]),
]

heading = lambda t: para([text(t)], '<w:pStyle w:val="Heading1"/>')
filler = "The guide covered the route, the ferry timetable and the best places to stay along the coast. " * 3
body = [
    heading("Arriving at the harbour"),
    para([text(filler)]),
    para([text("This is page "), field("PAGE", "1", '<w:u w:val="dotted"/>'), text(" of "),
          field("NUMPAGES", "1", '<w:u w:val="dotted"/>'), text(", and a narrow "),
          field("PAGE", "1", '<w:w w:val="50"/>'), text(" page number.")]),
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>',
    heading("Walks along the north cove"),
    para([text(filler)]),
]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W} {R}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId4"/>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1701" w:right="1134" w:bottom="1701" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)
header = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{"".join(header_paras)}</w:hdr>'
footer = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {W}>{"".join(footer_paras)}</w:ftr>'

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>'
    "</w:styles>"
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
