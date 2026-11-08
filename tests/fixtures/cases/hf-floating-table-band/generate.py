"""Generate hf-floating-table-band: body text keeps out of header and footer floating tables

A floating table in a header or footer, placed against the page or the margin,
that leaves no room beside it in the text column is a band the body text keeps
out of, as it keeps out of a body floating table's: a line reaching into the
band (topFromText and bottomFromText included) goes below it, and when the
band reaches the bottom margin the line goes to the next page.

One section per case, each starting a new page, Arial 10pt on exact 12pt lines:
1. Footer table 200pt below the top margin (vertAnchor margin), 10pt above and
   15pt below it: Line 28 goes below it, to 465.5pt.
2. A first-page footer holds a table 732.6pt down, 65.2pt tall, topFromText
   28.35pt (a band from 704.25pt), the other pages' footer only text:
   justified five-line paragraphs, one of them split at the band.
The footers holding a table end in an empty paragraph: Word moves a footer's
own lines below its floating table, which this fixture does not test.
"""

import pathlib
import zipfile

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WORDS = (
    "The workshop unpacked the newest machines and agreed to display the samples "
    "before the next market so that every visitor can study the drawings in time "
    "and send questions to the maker in writing."
)


def para(text="", jc=None):
    jc_xml = f'<w:jc w:val="{jc}"/>' if jc else ""
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="exact"/>{jc_xml}</w:pPr>{run}</w:p>'


def floating_table(y, height, top_from=0, bottom_from=0, v_anchor="page"):
    borders = "".join(
        f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for side in ("top", "left", "bottom", "right")
    )
    return (
        f'<w:tbl><w:tblPr><w:tblpPr w:topFromText="{top_from}" w:bottomFromText="{bottom_from}" '
        f'w:vertAnchor="{v_anchor}" w:horzAnchor="page" w:tblpX="1305" w:tblpY="{y}"/>'
        f'<w:tblW w:w="8901" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="8901"/></w:tblGrid>'
        f'<w:tr><w:trPr><w:trHeight w:val="{height}" w:hRule="exact"/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="8901" w:type="dxa"/></w:tcPr>{para("Floating table")}</w:tc></w:tr></w:tbl>'
    )


BOTTOM_567 = floating_table(14652, 1304, top_from=567)
MID_MARGIN = floating_table(4000, 1000, top_from=200, bottom_from=300, v_anchor="margin")

# (body, header, first-page footer, footer)
CASES = [
    ([para(f"Case 1 line {i}") for i in range(1, 31)], None, None, MID_MARGIN + para()),
    (
        [para(f"Case 2 line {i}") for i in range(1, 40)]
        + [para(f"Paragraph {i}: {WORDS} {WORDS}", jc="both") for i in range(1, 6)],
        None,
        BOTTOM_567 + para(),
        para("Footer text"),
    ),
]

parts = {}
rels = [("rIdStyles", "styles", "styles.xml"), ("rIdSettings", "settings", "settings.xml")]
body = []
for n, (paras, header, first_footer, footer) in enumerate(CASES, 1):
    refs = ""
    for kind, ref_type, xml in (("header", "default", header), ("footer", "first", first_footer), ("footer", "default", footer)):
        if xml is None:
            continue
        name = f"{kind}{n}{ref_type}"
        tag = "hdr" if kind == "header" else "ftr"
        parts[f"word/{name}.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{tag} {NS}>{xml}</w:{tag}>',
            kind,
        )
        rels.append((f"rId{name}", kind, f"{name}.xml"))
        refs += f'<w:{kind}Reference w:type="{ref_type}" r:id="rId{name}"/>'
    title_pg = "<w:titlePg/>" if first_footer else ""
    sect = (
        f'<w:sectPr>{refs}<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1701" w:bottom="1276" w:left="2155" w:header="284" w:footer="823" w:gutter="0"/>'
        f"{title_pg}</w:sectPr>"
    )
    if n < len(CASES):
        paras[-1] = paras[-1].replace("</w:pPr>", f"{sect}</w:pPr>", 1)
        body += paras
    else:
        body += paras + [sect]

document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(body)}</w:body></w:document>'
styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)
settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:defaultTabStop w:val="708"/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
    "</w:settings>"
)
override = lambda part, kind: (
    f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>'
)
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    + override("word/document.xml", "document.main")
    + override("word/styles.xml", "styles")
    + override("word/settings.xml", "settings")
    + "".join(override(part, kind) for part, (_, kind) in parts.items())
    + "</Types>"
)
relationships = lambda items: (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + "".join(f'<Relationship Id="{rid}" Type="{REL}/{kind}" Target="{target}"/>' for rid, kind, target in items)
    + "</Relationships>"
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", relationships([("rId1", "officeDocument", "word/document.xml")]))
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", relationships(rels))
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    for part, (xml, _) in parts.items():
        z.writestr(part, xml)
print(f"wrote {out}")
