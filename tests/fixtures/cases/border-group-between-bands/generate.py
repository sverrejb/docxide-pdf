"""Generate border-group-between-bands: the band between two paragraphs of a border group

Where a border group goes on, Word leaves the next paragraph's between
`space`, draws this paragraph's between border, then its own `space`. The
shading above that border is this paragraph's, the shading below it the next
one's, so two equally shaded paragraphs show no gap. Table cells reserve the
same band.

Tests:
- Body: boxed pairs with between borders, both shaded alike (space 8)
- Body: pairs with different shading, or only one of them shaded
- Body: a pair where only the second paragraph sets a between border, and
  one whose between borders differ in width and space
- Header: the same shading pairs as in the body
- Table cells: bottom-bordered pairs where one or both set a between border
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
YELLOW, BLUE = "FFD966", "9DC3E6"


def bdr(sides, space, between_space, between_sz):
    out = ""
    for s in sides:
        sp, sz = (between_space, between_sz) if s == "between" else (space, 6)
        out += f'<w:{s} w:val="single" w:sz="{sz}" w:space="{sp}" w:color="000000"/>'
    return f"<w:pBdr>{out}</w:pBdr>"


def para(text, sides=(), fill=None, space=1, between_space=None, between_sz=6):
    b = bdr(sides, space, space if between_space is None else between_space, between_sz) if sides else ""
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    return f'<w:p><w:pPr>{b}{shd}</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


box = ("top", "left", "bottom", "right")
boxb = box + ("between",)
low = ("bottom",)
lowb = ("bottom", "between")


def pairs(tag):
    return [
        para(f"{tag} same shading, between space 8", boxb, YELLOW, between_space=8),
        para(f"{tag} second paragraph", boxb, YELLOW, between_space=8),
        para(""),
        para(f"{tag} yellow above the between border", boxb, YELLOW, between_space=8),
        para(f"{tag} blue below it", boxb, BLUE, between_space=8),
        para(""),
        para(f"{tag} unshaded", boxb, between_space=8),
        para(f"{tag} blue below the between border", boxb, BLUE, between_space=8),
        para(""),
    ]


body = pairs("Body") + [
    para("Body shaded yellow, no between border", box, YELLOW),
    para("Body yellow, between space 8", boxb, YELLOW, between_space=8),
    para(""),
    para("Body yellow, between 3pt space 8", boxb, YELLOW, between_space=8, between_sz=24),
    para("Body yellow, between space 2", boxb, YELLOW, between_space=2),
    para(""),
    para("Body shaded yellow", boxb, YELLOW, between_space=8),
    para("Body unshaded below the between border", boxb, between_space=8),
    para(""),
]


def table(rows):
    tr = "".join(
        '<w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>' + "".join(r) + "</w:tc></w:tr>"
        for r in rows
    )
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/><w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:color="FF0000"/><w:bottom w:val="single" w:sz="4" w:color="FF0000"/>'
        '<w:insideH w:val="single" w:sz="4" w:color="FF0000"/></w:tblBorders></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>' + tr + "</w:tbl>"
    )


rows = [
    [para("Cell A: bottom only"), para("A", low), para("B with between 8", lowb, between_space=8), para("after")],
    [para("Cell B: between 8 first"), para("A with between 8", lowb, between_space=8), para("B", low), para("after")],
    [
        para("Cell C: differing between borders"),
        para("A between 3pt space 8", lowb, between_space=8, between_sz=24),
        para("B between space 2", lowb, between_space=2),
        para("after"),
    ],
    [
        para("Cell D: three paragraphs"),
        para("A between 4", lowb, between_space=4),
        para("B", low),
        para("C between 12", lowb, between_space=12),
        para("after"),
    ],
]

split = para("</w:t><w:br/><w:t>".join(f"Split line {i}" for i in range(1, 61)),
             boxb, YELLOW, between_space=8).replace(
                 "<w:pPr>", '<w:pPr><w:pageBreakBefore/><w:widowControl w:val="0"/>', 1)
following = para("Blue after split", boxb, BLUE, between_space=8)

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W} {R}><w:body>'
    + "".join(body)
    + table(rows)
    + para("")
    + split + following
    + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="4536" w:right="1134" w:bottom="1134" w:left="1418" w:header="567" w:footer="284" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)
header = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{"".join(pairs("Header"))}</w:hdr>'

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)

ct = "application/vnd.openxmlformats-officedocument.wordprocessingml"
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    f'<Override PartName="/word/document.xml" ContentType="{ct}.document.main+xml"/>'
    f'<Override PartName="/word/styles.xml" ContentType="{ct}.styles+xml"/>'
    f'<Override PartName="/word/header1.xml" ContentType="{ct}.header+xml"/>'
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
    f'<Relationship Id="rId3" Type="{rel}/header" Target="header1.xml"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/header1.xml", header)
print(f"wrote {out}")
