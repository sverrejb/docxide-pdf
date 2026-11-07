"""Generate section-pages-continuous: SECTIONPAGES around a continuous section break

A page where a continuous section starts belongs to both sections: Word
counts it for each. The footer of that page shows the count of the section
the page starts in; a body field shows its own section's count.

Tests:
- Section 1 on pages 1-2, section 2 starting mid page 2 and running to page 4
- SECTIONPAGES in the footer of every page (2, 2, 3, 3)
- SECTIONPAGES in the body of section 2 on the shared page
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


def sect_pr(footer_id):
    return (f'<w:sectPr><w:footerReference w:type="default" r:id="{footer_id}"/>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1417" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>'
            "</w:sectPr>")


def section_end(footer_id):
    return f"<w:p><w:pPr>{sect_pr(footer_id)}</w:pPr></w:p>"


body = [
    para(text("Section 1 starts on page 1.")),
    "".join(para(text(f"Section 1 line {i}.")) for i in range(30)),
    page_break(),
    para(text("Section 1 on page 2; section 2 starts below on this page.")),
    section_end("rId10"),
    para(text("Section 2 starts mid page (continuous break). Its SECTIONPAGES: "), field("SECTIONPAGES", "9")),
    "".join(para(text(f"Section 2 line {i}.")) for i in range(150)),
    para(text("Section 2 ends.")),
]

footer_plain = para(
    text("Page "), field("PAGE \\* Arabic \\* MERGEFORMAT", "5"), text(" of "),
    field("SECTIONPAGES \\* Arabic \\* MERGEFORMAT", "2"), text(" ("),
    field("NUMPAGES", "7"), text(" in total)"), jc="right")

last_sect_pr = sect_pr("rId11").replace("<w:pgSz", '<w:type w:val="continuous"/><w:pgSz')
document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
            + "".join(body) + last_sect_pr + "</w:body></w:document>")


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
    + "".join(f'<Override PartName="/word/footer{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' for i in (1, 2))
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
    f'<Relationship Id="rId11" Type="{REL}/footer" Target="footer2.xml"/></Relationships>'
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
print(f"wrote {out}")
