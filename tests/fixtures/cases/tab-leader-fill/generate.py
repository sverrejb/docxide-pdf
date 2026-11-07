"""Generate tab-leader-fill: where Word draws a tab's leader characters

Word draws each leader glyph with its device advance (288 dpi, 0.25pt, the
glyph's width rounded up from 0.46 px) and puts the glyphs on the page's
grid of that pitch, counted from x = 0. Every slot that starts at or after
the tab's start and ends by the tab's end gets one glyph, so the leader is
flush with neither the text before it nor the text after it. The leader
takes the tab character's own font, size and colour.

Tests:
- Dot, hyphen, underscore, heavy and middle-dot leaders
- Left, right, centre and decimal stops
- Proportional (Arial, Georgia, Times New Roman) and monospace (Courier New)
  fonts, 8 to 20pt
- Table-of-contents entries: right-aligned page numbers after dot leaders
- Text before the leader ending at varied, fractional positions
- Gaps holding one or two leader glyphs only
- Signature lines: underscore leaders on tabs with nothing after them
- A tab whose run differs in size, colour, weight or slant from the text around it
- A leader in a justified paragraph, on a wrapped line and on the last line
- Centred and right-aligned paragraphs and a left indent
- A gap narrower than the leader glyph that still holds one device slot
- Distributed lines: only the text after the last tab spreads, leaders keep
  their pitch, and a line with nothing to spread is centred, a tab at its
  end included
- Centred and right-aligned lines ending in a leaderless tab, before a break
  or at the paragraph end: the tab's extent counts
- Right-to-left distributed paragraphs spread only their spaces and centre
  a line without any, tabs included
- A header whose leader is drawn in a face (bold Courier New) used nowhere else
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def rpr(font="Arial", size=10, extra=""):
    half = int(size * 2)
    return (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{half}"/><w:szCs w:val="{half}"/>{extra}</w:rPr>')


def run(text, font="Arial", size=10, extra=""):
    if text == "\t":
        return f"<w:r>{rpr(font, size, extra)}<w:tab/></w:r>"
    return f'<w:r>{rpr(font, size, extra)}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(runs, tabs=(), jc=None, ind="", after=60, bidi=False):
    stops = "".join(f'<w:tab w:val="{v}" w:leader="{l}" w:pos="{p}"/>' for v, l, p in tabs)
    tabs_xml = f"<w:tabs>{stops}</w:tabs>" if tabs else ""
    jc_xml = f'<w:jc w:val="{jc}"/>' if jc else ""
    bidi_xml = "<w:bidi/>" if bidi else ""
    return f'<w:p><w:pPr>{bidi_xml}{tabs_xml}<w:spacing w:after="{after}"/>{ind}{jc_xml}</w:pPr>{"".join(runs)}</w:p>'


def line(before, after, tab, font="Arial", size=10):
    return para([run(before, font, size), run("\t", font, size), run(after, font, size)], [tab])


body = [para([run("Leader kinds and stop alignments", size=12, extra="<w:b/>")])]
for leader in ["dot", "hyphen", "underscore", "heavy", "middleDot"]:
    body.append(line(f"{leader} left", "after", ("left", leader, 4100)))
    body.append(line(f"{leader} right", "Page 7", ("right", leader, 6300)))
    body.append(line(f"{leader} centre", "centred", ("center", leader, 5200)))
    body.append(line(f"{leader} decimal", "12.50", ("decimal", leader, 7400)))

body.append(para([run("Fonts and sizes", size=12, extra="<w:b/>")], after=120))
for font, size, leader in [
    ("Arial", 8, "dot"), ("Arial", 11, "underscore"), ("Arial", 14, "dot"), ("Arial", 20, "underscore"),
    ("Georgia", 9, "hyphen"), ("Georgia", 13, "dot"), ("Times New Roman", 10, "underscore"),
    ("Times New Roman", 16, "dot"), ("Courier New", 10, "dot"), ("Courier New", 13, "underscore"),
    ("Courier New", 18, "hyphen"),
]:
    body.append(line(f"{font} {size}", "42", ("right", leader, 8400), font, size))

body.append(para([run("Contents", size=12, extra="<w:b/>")], after=120))
toc = [("1", "Introduction", "1"), ("2.1", "Fixed parts", "4"),
       ("4", "Appendix with a much longer title than the others", "123")]
for num, title, page in toc:
    indent = '<w:ind w:left="440"/>' if "." in num else ""
    body.append(para([run(num), run("\t"), run(title), run("\t"), run(page)],
                     [("left", "none", 1100 if "." in num else 660), ("right", "dot", 9060)], ind=indent))

body.append(para([run("Text before the leader", size=12, extra="<w:b/>")], after=120))
for text in ["i", "il", "W", "Wm", "Mill", "lllll", "W.i.W", "Fifteen letters"]:
    body.append(line(text, "end", ("left", "dot", 3400)))
for k in range(6):
    body.append(line("Narrow" + "l" * k, "9", ("right", "underscore", 1500 + k * 23)))

body.append(para([run("Signature lines", size=12, extra="<w:b/>")], after=120))
body.append(para([run("\t"), run("\t"), run("\t")],
                 [("right", "underscore", 2880), ("left", "none", 3960), ("right", "underscore", 9072)]))
body.append(para([run("\t"), run("\t"), run("End")], [("left", "dot", 4000), ("right", "hyphen", 8000)]))

body.append(para([run("Tab formatting and paragraphs", size=12, extra="<w:b/>")], after=120))
body.append(para([run("Small tab", size=16), run("\t", size=9), run("12", size=16)], [("right", "dot", 7000)]))
body.append(para([run("Coloured tab"), run("\t", extra='<w:color w:val="C00000"/>'), run("12")],
                 [("right", "underscore", 7000)]))
body.append(para([run("Bold tab"), run("\t", extra="<w:b/>"), run("12")], [("right", "dot", 7000)]))
body.append(para([run("Italic tab", font="Times New Roman"), run("\t", font="Times New Roman", extra="<w:i/>"),
                  run("12", font="Times New Roman")], [("right", "underscore", 7000)]))
body.append(para([run("Centred paragraph"), run("\t"), run("12")], [("right", "dot", 5000)], jc="center"))
body.append(para([run("Right-aligned paragraph"), run("\t"), run("12")], [("right", "dot", 5000)], jc="right"))
body.append(para([run("Indented"), run("\t"), run("12")], [("right", "dot", 8000)], ind='<w:ind w:left="567"/>'))
for words in (0, 14):
    body.append(para([run("Justified"), run("\t"), run("middle"), run("\t"),
                      run("then " + "some words " * words + "to the end")],
                     [("left", "dot", 3000), ("right", "underscore", 8000)], jc="both"))

body.append(para([run("Narrow gaps and distributed lines", size=12, extra="<w:b/>")], after=120))
body.append(para([run("\t")], [("left", "underscore", 120)], ind='<w:ind w:left="12"/>', after=0))
narrow = '<w:ind w:right="5754"/>'
for text, tabs in [("A\tB", [("left", "dot", 1440)]), ("A\tB", [("right", "dot", 2880)]),
                   ("A\tBC", [("left", "dot", 1440)]), ("AB\tCD", [("left", "none", 1440)]),
                   ("Name\t", [("left", "dot", 1440)]), ("A\tB\tC D", [("left", "dot", 1440), ("left", "hyphen", 2880)]),
                   ("Name\t", [("left", "none", 1440)]), ("Name\t", [("right", "none", 2880)])]:
    body.append(para([run(t, "Courier New") for t in text.replace("\t", "|\t|").split("|") if t],
                     tabs, jc="distribute", ind=narrow))
for jc in ("center", "right"):
    body.append(para([run("Name"), run("\t")], [("left", "none", 1440)], jc=jc, ind=narrow))
    body.append(para([run("Name"), run("\t"), '<w:r><w:br/></w:r>', run("XY")], [("left", "none", 1440)], jc=jc, ind=narrow))
for text in ("ABCD", "AB CD"):
    body.append(para([run(text, "Courier New")], jc="distribute", ind='<w:ind w:left="5754"/>', bidi=True))
body.append(para([run("A", "Courier New"), run("\t", "Courier New"), run("A", "Courier New")],
                 [("left", "dot", 1440)], jc="distribute", bidi=True))
body.append(para([run("Word one"), run("\t"), run("end two")], [("left", "dot", 1440)], jc="distribute", ind=narrow))

header = para([run("Header"), run("\t", "Courier New", extra="<w:b/>"), run("1")], [("right", "dot", 6000)])

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W} {R}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
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
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
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
    '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
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

out = pathlib.Path(__file__).with_name("input.docx")
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/header1.xml", f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W} {R}>{header}</w:hdr>')
print(f"wrote {out}")
