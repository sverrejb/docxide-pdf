"""Generate hf-floating-table-band-flow: header table bands across lines, columns and sections

A header's floating table placed against the page that leaves no room beside it
is a band body lines keep out of (fixture hf-floating-table-band), line by line
and on every page the header is drawn on. Arial 10pt on exact 12pt lines, the
header table 250 to 300pt down (5000 twips, exact 1000 twips row):
1. One long justified paragraph from the page top: its line reaching 250pt
   goes to 300pt, on the first page and again on the next, which it runs on to.
2. Two columns: the lines of both columns skip the band.
3. A section without the table, then a continuous section whose header has
   it: the second section's lines on the first page run through 250 to 300pt
   (that page shows the first section's header) and skip it on the next.
4. Two columns, a header table over column one only and a footer table over
   column one reaching past the bottom margin: column one's lines skip the
   header band and end above the footer band, column two's run through both.
5. Footer bands 625-725pt (a table with 60pt bottomFromText) and 675-805pt
   (a table with 10pt topFromText) crossing the 778pt bottom margin: the page
   ends above 625pt.
6. Two columns in a continuous section after 5, whose header has a table over
   each column from the top margin past the bottom margin: the lines run on
   below 5's on its page, then each page takes one line, its top on the bottom
   margin of column one.
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


def floating_table(y, height, x=1305, width=8901, dist=""):
    borders = "".join(
        f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        for side in ("top", "left", "bottom", "right")
    )
    return (
        f'<w:tbl><w:tblPr><w:tblpPr {dist} w:vertAnchor="page" w:horzAnchor="page" w:tblpX="{x}" w:tblpY="{y}"/>'
        f'<w:tblW w:w="{width}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{width}"/></w:tblGrid>'
        f'<w:tr><w:trPr><w:trHeight w:val="{height}" w:hRule="exact"/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/></w:tcPr>{para("Floating table")}</w:tc></w:tr></w:tbl>'
    )


BAND = floating_table(5000, 1000) + para("Header text")
COLS = '<w:cols w:num="2" w:space="425"/>'
LONG = " ".join([WORDS] * 14)
COLUMN_ONE = {"x": 2155, "width": 3812}
CROSSING = (
    floating_table(12500, 800, dist='w:bottomFromText="1200"')
    + floating_table(13700, 2400, dist='w:topFromText="200"')
    + para()
)

# (body paragraphs, header, section break type, columns, footer)
CASES = [
    ([para(f"Case 1. {LONG} {LONG} {LONG}", jc="both"), para("Case 1 end")], BAND, "nextPage", "", None),
    ([para(f"Case 2 line {i}") for i in range(1, 81)], BAND, "nextPage", COLS, None),
    ([para(f"Case 3a line {i}") for i in range(1, 6)], para("Plain header"), "nextPage", "", None),
    ([para(f"Case 3b line {i}") for i in range(1, 91)], BAND, "continuous", "", None),
    (
        [para(f"Case 4 line {i}") for i in range(1, 161)],
        floating_table(5000, 1000, **COLUMN_ONE) + para("Header text"),
        "nextPage",
        COLS,
        floating_table(13800, 2000, **COLUMN_ONE) + para(),
    ),
    ([para(f"Case 5 line {i}") for i in range(1, 81)], para("Plain header"), "nextPage", "", CROSSING),
    (
        [para(f"Case 6 line {i}") for i in range(1, 32)],
        floating_table(1134, 14700, **COLUMN_ONE) + floating_table(1134, 14700, x=6392, width=3812) + para(),
        "continuous",
        COLS,
        para(),
    ),
]

parts = {}
rels = [("rIdStyles", "styles", "styles.xml"), ("rIdSettings", "settings", "settings.xml")]
body = []
for n, (paras, header, kind, cols, footer) in enumerate(CASES, 1):
    refs = ""
    for part, tag, content in (("header", "hdr", header), ("footer", "ftr", footer)):
        if content is None:
            continue
        name = f"{part}{n}"
        parts[f"word/{name}.xml"] = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{tag} {NS}>{content}</w:{tag}>',
            part,
        )
        rels.append((f"rId{name}", part, f"{name}.xml"))
        refs += f'<w:{part}Reference w:type="default" r:id="rId{name}"/>'
    sect = (
        f'<w:sectPr>{refs}<w:type w:val="{kind}"/>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1134" w:right="1701" w:bottom="1276" w:left="2155" w:header="284" w:footer="823" w:gutter="0"/>'
        f"{cols}</w:sectPr>"
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
