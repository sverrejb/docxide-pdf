"""Generate autofit-cell-widths

An AutoFit table of auto width (tblW auto, tblLayout autofit) whose cells
name their widths in twips: Word lays the columns out from the cells' tcW,
each column as wide as its widest cell, and ignores a saved tblGrid that
disagrees, wider or narrower. Only a word longer than its cell widens a
column. (Cells too wide for the page together are squeezed to it; a
Word-saved grid holds that result, so they are left out here.)

Tests (each table's grid disagrees with its cells):
- A a grid wider in the first column, narrower in the others
- B a cell narrower than its unbreakable word
- C two rows with different cell widths per column
- D a second row of auto-width cells
- E narrow cells whose text wraps
"""

import pathlib
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def p(text):
    return f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


B = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            for s in ["top", "left", "bottom", "right", "insideH", "insideV"])
AUTO = '<w:tblW w:w="0" w:type="auto"/>'


def table(grid, tcw, texts, tblw=AUTO, layout=""):
    g = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    rows = ""
    for r, row in enumerate(texts):
        widths = tcw[r] if isinstance(tcw[0], list) else tcw
        cells = ""
        for i, t in enumerate(row):
            w = widths[i]
            tw = f'<w:tcW w:w="{w}" w:type="dxa"/>' if w is not None else '<w:tcW w:w="0" w:type="auto"/>'
            cells += f'<w:tc><w:tcPr>{tw}</w:tcPr>{p(t)}</w:tc>'
        rows += f'<w:tr>{cells}</w:tr>'
    return (f'<w:tbl><w:tblPr>{tblw}{layout}<w:tblBorders>{B}</w:tblBorders>'
            f'<w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>{g}</w:tblGrid>{rows}</w:tbl>' + p(""))


H = [["Bike colours", "Red", "Tan", "Natural"]]
T3 = [["One", "Two", "Three"]]
cases = [
    ("A", [3474, 1882, 1883, 2105], [2372, 2304, 2305, 2305], H),
    ("B", [3000, 3000, 3000], [600, 2000, 2500], [["Unbreakablewordhere", "Two", "Three"]]),
    ("C", [3000, 3000, 3000], [[1500, 2000, 2500], [2500, 1500, 1000]], T3 + T3),
    ("D", [3474, 1882, 1883, 2105], [[2372, 2304, 2305, 2305], [None, None, None, None]], H + H),
    ("E", [2000, 2000, 2000], [1000, 1000, 1000], [["A much longer cell text that wraps", "Two", "Three"]]),
]
body = ""
for c in cases:
    name, grid, tcw, texts = c[:4]
    body += p(name) + table(grid, tcw, texts, *c[4:])

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>' + body
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1000" w:right="1134" w:bottom="1000" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)
styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)
settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
    '<w:defaultTabStop w:val="709"/><w:compat>'
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    "</w:compat></w:settings>"
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
print(f"wrote {out}", file=sys.stderr)
