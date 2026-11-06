"""Generate border-group-between-rule-fill: shading next to a between border

Where a border group goes on, Word fills the gap above the between border
with the upper paragraph's shading and the border's `space` below it with the
lower paragraph's, but never the border's own band: a thick single border
keeps its full width and a double border's gap stays white, whichever of the
two paragraphs is shaded.

Tests (no paragraph spacing):
- Body and header: a 6pt single between border, both paragraphs green, space 0
- Body and header: a 4.5pt double between border, green over pink, space 1
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
GREEN, PINK = "A9D18E", "F4B6C2"


def para(text, fill, sz, space, val="single"):
    sides = ""
    for s in ("top", "left", "bottom", "right", "between"):
        z, v = (sz, val) if s == "between" else (6, "single")
        sides += f'<w:{s} w:val="{v}" w:sz="{z}" w:space="{space}" w:color="1F3864"/>'
    shd = f'<w:shd w:val="clear" w:color="auto" w:fill="{fill}"/>' if fill else ""
    return f'<w:p><w:pPr><w:pBdr>{sides}</w:pBdr>{shd}</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'


def gap():
    return "<w:p/>"


VARIANTS = [
    ("both green, 6pt, space 0", GREEN, GREEN, 48, 0, "single"),
    ("green over pink, double 4.5pt", GREEN, PINK, 36, 1, "double"),
]


def pairs(tag, variants):
    out = []
    for label, upper, lower, sz, space, val in variants:
        out += [para(f"{tag}: {label}", upper, sz, space, val), para(f"{tag}: lower paragraph", lower, sz, space, val), gap()]
    return out


document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W} {R}><w:body>'
    + "".join(pairs("Body", VARIANTS))
    + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="5000" w:right="1134" w:bottom="1134" w:left="1418" w:header="454" w:footer="284" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)
header = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{"".join(pairs("Header", VARIANTS))}</w:hdr>'

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="18"/><w:szCs w:val="18"/></w:rPr></w:rPrDefault>'
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
