"""Generate page-table-line-reach-compat14: a one-line paragraph's space after reaching a floating table

The compat 14 twin of page-table-line-reach's spacing pages. Before compat 15
Word counts a one-line paragraph's space after into what reaches a floating
table: the paragraph moves below the table when its line and space after
reach the table top (ending on it included). Compat 15 ignores the space
after (page-table-line-reach).

One page per variant (Intro, the table, then lines, 13pt exact; Intro starts
at 72pt, so L1 spans 85-98):
- the anchor L1 with 20pt after, table top 118pt: exact contact moves L1
- a text-anchored table 40pt below its anchor L1, L2 with 20pt after (moves)
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def tw(pt):
    return int(round(pt * 20))


def para(text, page_break=False, before=0, after=0):
    pb = "<w:pageBreakBefore/>" if page_break else ""
    return (
        f'<w:p><w:pPr>{pb}<w:spacing w:before="{tw(before)}" w:after="{tw(after)}" w:line="260" w:lineRule="exact"/>'
        f'</w:pPr><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>'
    )


def table(y, v_anchor="page"):
    return (
        f'<w:tbl><w:tblPr><w:tblpPr w:leftFromText="0" w:rightFromText="0" w:vertAnchor="{v_anchor}" '
        f'w:horzAnchor="page" w:tblpX="0" w:tblpY="{tw(y)}"/><w:tblOverlap w:val="never"/>'
        '<w:tblW w:w="11906" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="11906"/></w:tblGrid>'
        '<w:tr><w:trPr><w:trHeight w:val="800" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="11906" w:type="dxa"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tcPr>'
        f'<w:p><w:r><w:t>Floating table, top at {y}pt ({v_anchor})</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'
    )


pages = [
    ("L1 20pt after, top 118pt", table(118), [para("L1 after 20", after=20)]),
    ("text-anchored 40pt, L2 20pt after", table(40, "text"), [para("L1"), para("L2 after 20", after=20)]),
]
body = []
for n, (label, tbl, lines) in enumerate(pages, 1):
    body.append(para(f"Intro {n}: {label}", n > 1))
    body.append(tbl)
    body += lines + [para(f"L{k}") for k in range(len(lines) + 1, 9)]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
    'w:left="1440" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>'
)
styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)
settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>'
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
