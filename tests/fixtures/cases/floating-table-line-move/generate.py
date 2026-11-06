"""Generate floating-table-line-move: where a line meeting a page-anchored floating table lands

A floating table anchored to the page with no room beside it keeps text out of
its wrap area (the table plus topFromText above and bottomFromText below): a
line whose box reaches into it goes below it. Which line that is depends on
where the paragraph finally sits, so it is judged where its lines are placed.

One section per case, exact 13pt lines, the table 34.25pt tall with 20pt
bottomFromText:
1. A five-line keepLines paragraph whose third line reaches 8pt into a table
   near the page foot: the paragraph moves to the next page whole, its lines
   one under another
2. A five-line paragraph with widow control whose last line reaches 8pt into a
   table near the page foot and finds no room below it: three lines stay, the
   last two start the next page one under another
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TOP, LINE = 1134, 260
SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    f'<w:pgMar w:top="{TOP}" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr>"
)


def para(texts, sect=False, keep_lines=False, widow=False):
    ppr = "<w:keepLines/>" if keep_lines else ""
    ppr += "<w:widowControl/>" if widow else ""
    ppr += f'<w:spacing w:line="{LINE}" w:lineRule="exact"/>'
    runs = "<w:r><w:br/></w:r>".join(f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>' for t in texts)
    return f'<w:p><w:pPr>{ppr}{SECT if sect else ""}</w:pPr>{runs}</w:p>'


def table(wrap_top):
    """A 500pt wide table at x 65.25pt, its top wrap_top twips from the page top."""
    borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                      for s in ("top", "left", "bottom", "right"))
    return (
        f'<w:tbl><w:tblPr><w:tblpPr w:bottomFromText="400" w:vertAnchor="page" w:horzAnchor="page"'
        f' w:tblpX="1305" w:tblpY="{wrap_top}"/><w:tblW w:w="10000" w:type="dxa"/>'
        f'<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="10000"/></w:tblGrid>'
        '<w:tr><w:trPr><w:trHeight w:val="685" w:hRule="exact"/></w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:w="10000" w:type="dxa"/></w:tcPr>'
        '<w:p><w:pPr><w:spacing w:line="200" w:lineRule="exact"/></w:pPr><w:r><w:t>Floating table</w:t></w:r></w:p>'
        "</w:tc></w:tr></w:tbl>"
    )


def fillers(n, label):
    return [para([f"{label} filler {i}"]) for i in range(1, n + 1)]


FIVE = [f"Paragraph line {i}" for i in range(1, 6)]
body = (
    [table(TOP + 49 * LINE + 3 * LINE - 8 * 20)] + fillers(49, "Case 1")
    + [para(FIVE, keep_lines=True), para(["After"], sect=True)]
    + [table(TOP + 48 * LINE + 5 * LINE - 8 * 20)] + fillers(48, "Case 2")
    + [para(FIVE, widow=True), para(["After"])]
)

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
    + "".join(body) + SECT + "</w:body></w:document>"
)

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)

settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
    "</w:settings>"
)

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
