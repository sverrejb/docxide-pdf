"""Generate paragraph-border-indents: paragraph borders follow the indents

Word draws a paragraph's border box around the indented text, not the
column: the box starts at the left indent (or at the first line when that
hangs further out) and ends at the right indent.

Tests:
- Bottom rule with a left indent, a right indent, both
- Bottom rule with a hanging indent (box starts at the hanging first line)
- Bottom rule with a positive first-line indent (box starts at the left indent)
- Full box with left/right indents, with and without a hanging indent
- Negative left and right indents (box reaches into the margins)
- Two adjacent boxes with the same edges but different indents (one group)
- Two adjacent boxes with different edges (two groups)
- Shading, alone and inside a box, follows the same indents
- A bottom rule with left/right indents in the header
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def bdr(sides, sz=6, space=1):
    return "<w:pBdr>" + "".join(
        f'<w:{s} w:val="single" w:sz="{sz}" w:space="{space}" w:color="000000"/>' for s in sides
    ) + "</w:pBdr>"


def para(text, sides=(), left=None, right=None, hanging=None, first=None, after=120, shd=None):
    ind = ""
    attrs = "".join(
        f' w:{k}="{v}"'
        for k, v in (("left", left), ("right", right), ("hanging", hanging), ("firstLine", first))
        if v is not None
    )
    if attrs:
        ind = f"<w:ind{attrs}/>"
    pbdr = bdr(sides) if sides else ""
    if shd:
        pbdr += f'<w:shd w:val="clear" w:color="auto" w:fill="{shd}"/>'
    return (
        f'<w:p><w:pPr>{pbdr}<w:spacing w:after="{after}"/>{ind}</w:pPr>'
        f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


box = ("top", "left", "bottom", "right")
body = [
    para("Bottom rule, left indent 1440", ["bottom"], left=1440),
    para("Bottom rule, right indent 2880", ["bottom"], right=2880),
    para("Bottom rule, left 2552 right 4819 hanging 1418", ["bottom"], left=2552, right=4819, hanging=1418),
    para("Bottom rule, left 1134 first line 567", ["bottom"], left=1134, first=567),
    para("Box, left 1440 right 1440", box, left=1440, right=1440),
    para("Box, left 2160 hanging 1080 right 720: the first line hangs out of the text column", box,
         left=2160, hanging=1080, right=720),
    para("Box, negative indents left -567 right -567", box, left=-567, right=-567),
    para("Group: left 1134", box, left=1134, right=1134, after=0),
    para("Group: left 1701 hanging 567, same box start", box, left=1701, hanging=567, right=1134, after=0),
    para("Group: left 1134 again", box, left=1134, right=1134),
    para("Two groups: left 851", ["bottom", "top"], left=851, after=0),
    para("Two groups: left 1418", ["bottom", "top"], left=1418),
    para("Bottom rule, hanging 720 with no left indent", ["bottom"], hanging=720, left=0),
    para("Shading only, left 1440 right 2880", left=1440, right=2880, shd="BFBFBF"),
    para("Shading only, left 1701 hanging 851", left=1701, hanging=851, shd="BFBFBF"),
    para("Shaded box, left 1134 right 1134", box, left=1134, right=1134, shd="D9D9D9"),
]

header_body = para("Header with a bottom rule, left 1134 right 2268", ["bottom"], left=1134, right=2268, after=0)

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W} {R}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)

header = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{header_body}</w:hdr>'

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)

settings = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/></w:settings>'

content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
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
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/>'
    f'<Relationship Id="rId3" Type="{rel}/header" Target="header1.xml"/></Relationships>'
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
print(f"wrote {out}")
