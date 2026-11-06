"""Generate section-pages-field: SECTIONPAGES in footers and body

SECTIONPAGES is the page count of the section the field is on. Template
engines often leave a stale cached result in it; Word recomputes it.

Tests:
- Footer "Page PAGE of SECTIONPAGES (NUMPAGES in total)" with stale cached results
- Three sections of 2, 1 and 3 pages (next-page section breaks)
- Page numbering restarting in the third section
- A \\* ROMAN switch on SECTIONPAGES
- SECTIONPAGES in the body text
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' \
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def text(t):
    return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'


def field(instr, cached):
    return ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r><w:instrText xml:space="preserve"> {instr} </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r><w:t>{cached}</w:t></w:r>'
            '<w:r><w:fldChar w:fldCharType="end"/></w:r>')


def para(*runs, jc=None):
    ppr = f'<w:pPr><w:jc w:val="{jc}"/></w:pPr>' if jc else ""
    return f"<w:p>{ppr}{''.join(runs)}</w:p>"


def page_break():
    return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'


def sect_pr(footer_id, restart=False):
    pg_num = '<w:pgNumType w:start="1"/>' if restart else ""
    return (f'<w:sectPr><w:footerReference w:type="default" r:id="{footer_id}"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>'
            f"{pg_num}</w:sectPr>")


def section_end(footer_id, restart=False):
    return f"<w:p><w:pPr>{sect_pr(footer_id, restart)}</w:pPr></w:p>"


body = [
    para(text("Section 1, page 1. This section has "), field("SECTIONPAGES", "9"), text(" pages.")),
    page_break(),
    para(text("Section 1, page 2.")),
    section_end("rId10"),
    para(text("Section 2, its only page. SECTIONPAGES here: "), field("SECTIONPAGES", "9")),
    section_end("rId11"),
    para(text("Section 3, page 1, numbering restarts.")),
    page_break(),
    para(text("Section 3, page 2.")),
    page_break(),
    para(text("Section 3, page 3. SECTIONPAGES in roman: "), field("SECTIONPAGES \\* ROMAN", "IX")),
]

footer_plain = para(
    text("Page "), field("PAGE \\* Arabic \\* MERGEFORMAT", "5"), text(" of "),
    field("SECTIONPAGES \\* Arabic \\* MERGEFORMAT", "2"), text(" ("),
    field("NUMPAGES", "7"), text(" in total)"), jc="right")
footer_roman = para(
    text("Page "), field("PAGE", "5"), text(" of "), field("SECTIONPAGES \\* ROMAN", "II"), jc="center")

document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            + "".join(body) + sect_pr("rId12", restart=True) + "</w:body></w:document>")


def footer(p):
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr {W}>{p}</w:ftr>'


styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + "".join(f'<Override PartName="/word/footer{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' for i in (1, 2, 3))
    + "</Types>"
)
package_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>'
)
document_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId10" Type="{REL}/footer" Target="footer1.xml"/>'
    f'<Relationship Id="rId11" Type="{REL}/footer" Target="footer2.xml"/>'
    f'<Relationship Id="rId12" Type="{REL}/footer" Target="footer3.xml"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/footer1.xml", footer(footer_plain))
    z.writestr("word/footer2.xml", footer(footer_plain))
    z.writestr("word/footer3.xml", footer(footer_roman))
print(f"wrote {out}")
