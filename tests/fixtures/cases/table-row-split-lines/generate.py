"""Generate table-row-split-lines

Where Word breaks a table row that does not fit in the room left on the page,
and how tall its pieces are: an at-least trHeight, 5pt top and bottom cell
margins, single borders, Times New Roman 12pt
at 1.15 lines, Word 2013+ layout (compatibilityMode 15).

Each scenario is a section of its own: a spacer that leaves room for a few
lines of the row (none for a row at the page top), the row, then a
paragraph after the table that ends the section.

Tests:
- a 45pt at-least row split 2/2: both pieces are 45pt plus the cell margins
- a 30pt at-least row split 2/1 without widow control: the one-line
  continuation is 30pt plus the cell margins
- a 300pt at-least row taller than the 260pt page body, twenty lines from
  the page top: each piece fills the page, the five-line last one included
- the same row with three lines: one piece that fills the page, no empty
  continuation
- the three-line row with room for five lines above it: it moves to the
  next page, as its minimum does not fit the room left
- the twenty-line 300pt row nested in a zero-margin outer cell: each piece
  fills the page, as the unnested row's do
- the three-line 300pt row nested: one piece that fills the page; the outer
  cell's end mark adds nothing, so the next paragraph opens the next page
- the three-line 300pt row with a footnote on its first line: the row fills
  the page and the note goes to the next page
- a repeating header row over the twenty-line 300pt row: the header stays
  alone on the first page and does not repeat above the row's pieces
- a repeating header row, a two-line row and a three-line 200pt row: the
  200pt row moves to the next page under the repeated header
- the three-line 300pt row nested two deep: one piece that fills the page;
  the wrappers' end marks open the next page with empty pieces, no second
  piece of the row
- a section ending with a three-line 228pt row whose note does not fit: the
  note fills a page of its own under the section's first-page header, and
  the next section starts after it
- the document ending with that row, citing a footnote and an endnote: the
  notes fill a last page of their own, the endnote at its top under the
  default header
"""

import pathlib
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P = "http://schemas.openxmlformats.org/package/2006/relationships"
PITCH = 13.8 * 1.15
SECTION = (
    '<w:sectPr><w:pgSz w:w="8000" w:h="6000"/><w:pgMar w:top="400" w:bottom="400" '
    'w:left="400" w:right="400" w:header="200" w:footer="200"/></w:sectPr>'
)

SCENARIOS = [
    (4, 3, 1, True, 900),
    (3, 2, 1, False, 600),
    (20, None, 1, True, 6000),
    (3, None, 1, True, 6000),
    (3, 5, 1, True, 6000),
]


def para(words, ppr=""):
    runs = "".join(f'<w:r><w:t xml:space="preserve">{w} </w:t></w:r>' for w in words)
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>"


def cell(width, content):
    mar = '<w:tcMar><w:top w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/></w:tcMar>'
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{mar}</w:tcPr>{content}</w:tc>'


def scenario(n, lines, room, col2, widow, height):
    rows = [row(n, lines, height, col2=col2, widow=widow)]
    spacer = (
        ""
        if room is None
        else f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{round((260 - (10 + room * PITCH + 4)) * 20)}" w:lineRule="exact"/></w:pPr></w:p>'
    )
    return (
        spacer
        + table(rows)
        + para([f"After{n}", "table"], SECTION)
    )


def row(n, lines, height, header=False, note=None, endnote=None, col2=0, widow=True):
    wc = "" if widow else '<w:widowControl w:val="0"/>'
    c1 = para([f"S{n}Line{i + 1}abcdefgh" for i in range(lines)], wc)
    c2 = para([f"Mark{i + 1}abc" for i in range(col2)], wc) if col2 else "<w:p/>"
    refs = "".join(
        f'<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:{kind}Reference w:id="{ref}"/></w:r>'
        for kind, ref in (("footnote", note), ("endnote", endnote))
        if ref
    )
    c1 = c1.replace("</w:t></w:r>", "</w:t></w:r>" + refs, 1)
    trpr = (f'<w:trHeight w:val="{height}"/>' if height else "") + ("<w:tblHeader/>" if header else "")
    return f'<w:tr><w:trPr>{trpr}</w:trPr>{cell(2400, c1)}{cell(1600, c2)}{cell(2400, "<w:p/>")}</w:tr>'


def table(rows):
    borders = "".join(
        f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
        for s in ("top", "left", "bottom", "right", "insideH", "insideV")
    )
    return (
        f'<w:tbl><w:tblPr><w:tblW w:w="6400" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
        '<w:tblLayout w:type="fixed"/></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="1600"/><w:gridCol w:w="2400"/></w:tblGrid>'
        f'{"".join(rows)}</w:tbl>'
    )


def nested(inner):
    return (
        '<w:tbl><w:tblPr><w:tblW w:w="6600" w:type="dxa"/><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/>'
        '<w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="6600"/></w:tblGrid>'
        f'<w:tr><w:tc><w:tcPr><w:tcW w:w="6600" w:type="dxa"/></w:tcPr>{inner}<w:p/></w:tc></w:tr></w:tbl>'
    )


def generate(path):
    body = "".join(scenario(n + 1, *s) for n, s in enumerate(SCENARIOS))
    more = [
        nested(table([row(6, 20, 6000)])),
        nested(table([row(7, 3, 6000)])),
        table([row(8, 3, 6000, note=2)]),
        table([row(9, 1, None, header=True), row(9, 20, 6000)]),
        table([row(10, 1, None, header=True), row(10, 2, None), row(10, 3, 4000)]),
        nested(nested(table([row(11, 3, 6000)]))),
        table([row(12, 3, 4560, note=3)]),
        table([row(13, 3, 4560, note=4, endnote=2)]),
    ]
    headed = SECTION.replace(
        "<w:sectPr>",
        '<w:sectPr><w:headerReference w:type="default" r:id="h1"/><w:headerReference w:type="first" r:id="h2"/>',
    ).replace("</w:sectPr>", "<w:titlePg/></w:sectPr>")
    last = len(SCENARIOS) + len(more)
    for n, t in enumerate(more, len(SCENARIOS) + 1):
        body += t + para([f"After{n}", "table"], "" if n == last else headed if n == last - 1 else SECTION)
    body += headed
    styles = (
        f'<w:styles xmlns:w="{W}"><w:docDefaults><w:rPrDefault><w:rPr>'
        '<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>'
        '<w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
        '<w:widowControl/><w:spacing w:after="0" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        '</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
        '<w:name w:val="Normal"/></w:style></w:styles>'
    )
    settings = (
        f'<w:settings xmlns:w="{W}"><w:compat><w:compatSetting w:name="compatibilityMode" '
        'w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>'
    )
    sep = '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:{}/></w:r></w:p>'

    def notes(kind, texts):
        body = "".join(
            f'<w:{kind} w:id="{i}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:{kind}Ref/></w:r>'
            f'<w:r><w:t xml:space="preserve"> {t}</w:t></w:r></w:p></w:{kind}>'
            for i, t in texts.items()
        )
        return (
            f'<w:{kind}s xmlns:w="{W}">'
            f'<w:{kind} w:type="separator" w:id="0">{sep.format("separator")}</w:{kind}>'
            f'<w:{kind} w:type="continuationSeparator" w:id="1">{sep.format("continuationSeparator")}</w:{kind}>'
            f"{body}</w:{kind}s>"
        )

    def header(text):
        return (
            f'<w:hdr xmlns:w="{W}"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:rPr><w:sz w:val="12"/></w:rPr><w:t>{text}</w:t></w:r></w:p></w:hdr>'
        )

    ct = (
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        '<Override PartName="/word/header2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        "</Types>"
    )
    parts = {
        "[Content_Types].xml": ct,
        "_rels/.rels": f'<Relationships xmlns="{P}"><Relationship Id="d" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>',
        "word/document.xml": f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}</w:body></w:document>',
        "word/_rels/document.xml.rels": (
            f'<Relationships xmlns="{P}"><Relationship Id="s" Type="{R}/styles" Target="styles.xml"/>'
            f'<Relationship Id="t" Type="{R}/settings" Target="settings.xml"/>'
            f'<Relationship Id="f" Type="{R}/footnotes" Target="footnotes.xml"/>'
            f'<Relationship Id="e" Type="{R}/endnotes" Target="endnotes.xml"/>'
            f'<Relationship Id="h1" Type="{R}/header" Target="header1.xml"/>'
            f'<Relationship Id="h2" Type="{R}/header" Target="header2.xml"/></Relationships>'
        ),
        "word/styles.xml": styles,
        "word/settings.xml": settings,
        "word/footnotes.xml": notes("footnote", {2: "Note8", 3: "Note12", 4: "Note13"}),
        "word/endnotes.xml": notes("endnote", {2: "Endnote13"}),
        "word/header1.xml": header("DefaultHead"),
        "word/header2.xml": header("FirstHead"),
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, xml in parts.items():
            z.writestr(name, xml)


if __name__ == "__main__":
    generate(pathlib.Path(__file__).with_name("input.docx"))
