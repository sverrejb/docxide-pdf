"""Generate floating-table-stack: several page-anchored floating tables before the first paragraph

Three floating tables sit at the top of the body: a wide table whose
topFromText reserves an area above the text, a narrower table inside that
area and a narrow table in the left margin. Word wraps the text around all of them: the narrow table inside the
wide one's area and the one in the margin take nothing more, so the text goes
below the wide table's area, whatever order the tables come in.

One section per order, each on its own page:
1. wide, narrow, margin
2. narrow, wide, margin
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TOP, LINE = 1134, 260


def para(text="", line=LINE, rule="exact", ind_left=None, sect=False):
    ind = f'<w:ind w:left="{ind_left}"/>' if ind_left is not None else ""
    sect_pr = SECT if sect else ""
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f'<w:p><w:pPr><w:spacing w:line="{line}" w:lineRule="{rule}"/>{ind}{sect_pr}</w:pPr>{run}</w:p>'


def table(x, y, width, height, text, top_from=0, bottom_from=0):
    """A floating table anchored to the page; positions and sizes in twips."""
    tf = f' w:topFromText="{top_from}"' if top_from else ""
    bf = f' w:bottomFromText="{bottom_from}"' if bottom_from else ""
    borders = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                      for s in ("top", "left", "bottom", "right"))
    return (
        f'<w:tbl><w:tblPr><w:tblpPr{tf}{bf} w:vertAnchor="page" w:horzAnchor="page"'
        f' w:tblpX="{x}" w:tblpY="{y}"/><w:tblW w:w="{width}" w:type="dxa"/>'
        f'<w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid>'
        f'<w:tr><w:trPr><w:trHeight w:val="{height}" w:hRule="exact"/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>'
        f'{para(text, line=200)}</w:tc></w:tr></w:tbl>'
    )


def lines(n):
    return [para(f"Line {i}", sect=i == n) for i in range(1, n + 1)]


SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    f'<w:pgMar w:top="{TOP}" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr>"
)
wide = table(1305, 3000, 10000, 685, "Wide", top_from=1000, bottom_from=400)
narrow = table(1305, 2400, 4000, 400, "Narrow")
margin = table(228, 2600, 454, 2000, "M")
body = [wide, narrow, margin] + lines(20) + [narrow, wide, margin] + lines(20)

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
