"""Generate floating-table-line-reach: lines meeting a page-anchored floating table from above

A floating table anchored to the page with no room beside it (wider than the
column) keeps text out of its wrap area: the table plus topFromText above and
bottomFromText below. Word moves a line whose box reaches into that area,
however little, below the area with no gap; the lines above it stay.

One section per case, each on its own page:
1. An empty paragraph indented into the left margin, 4pt into the area; the
   37pt between the margin and the table holds no empty line
2. A multi-line paragraph (exact 13pt lines) whose second line reaches 8pt in
3. A multi-line paragraph (single spacing) whose third line reaches 2pt in
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TOP, LINE = 1134, 260
TEXT = "Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod tempor " * 4


def para(text="", line=LINE, rule="exact", ind_left=None, sect=False):
    ind = f'<w:ind w:left="{ind_left}"/>' if ind_left is not None else ""
    sect_pr = SECT if sect else ""
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f'<w:p><w:pPr><w:spacing w:line="{line}" w:lineRule="{rule}"/>{ind}{sect_pr}</w:pPr>{run}</w:p>'


def table(wrap_top, top_from):
    """A 500pt wide, 34.25pt tall table at x 65.25pt; wrap_top in twips from the page top."""
    tf = f' w:topFromText="{top_from}"' if top_from else ""
    borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                      for s in ("top", "left", "bottom", "right"))
    return (
        f'<w:tbl><w:tblPr><w:tblpPr{tf} w:bottomFromText="400" w:vertAnchor="page" w:horzAnchor="page"'
        f' w:tblpX="1305" w:tblpY="{wrap_top + top_from}"/><w:tblW w:w="10000" w:type="dxa"/>'
        f'<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="10000"/></w:tblGrid>'
        '<w:tr><w:trPr><w:trHeight w:val="685" w:hRule="exact"/></w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:w="10000" w:type="dxa"/></w:tcPr>'
        f'{para("Floating table", line=200)}</w:tc></w:tr></w:tbl>'
    )


def lines(first, n):
    return [para(f"Line {i}", sect=i == first + n - 1) for i in range(first, first + n)]


SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    f'<w:pgMar w:top="{TOP}" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr>"
)
line3_top = TOP + 2 * LINE
body = (
    [table(line3_top + LINE - 4 * 20, 1000), para("Line 1"), para("Line 2"), para(ind_left=-850)] + lines(4, 6)
    + [table(TOP + 3 * LINE - 8 * 20, 1000), para("Line 1"), para(TEXT), para("After", sect=True)]
    + [table(TOP + 3 * LINE - 2 * 20, 1000), para("Line 1"), para(TEXT, line=240, rule="auto"), para("After")]
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
