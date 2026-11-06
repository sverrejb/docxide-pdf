"""Generate ul-trail-space-on and ul-trail-space-off

w:compat/w:ulTrailSpace ("underline trailing spaces") makes Word underline
the spaces a line ends in. Without it a line's underline stops at its last
visible character. The -on fixture exercises trailing-space underlining;
-off keeps the single-space and tabbed-line controls without ulTrailSpace.

Tests in -on:
- One trailing space before the paragraph mark
- A trailing space before a manual line break, left aligned
- Underlined text wrapping at spaces, left aligned
- Centred and right-aligned lines ending in spaces
- An underlined trailing space after a non-underlined word
- A tabbed line ending in spaces
- A trailing space followed by an empty run
"""

import pathlib
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
U = '<w:u w:val="single"/>'


def run(text, extra=U):
    return f'<w:r><w:rPr>{extra}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


BR = f"<w:r><w:rPr>{U}</w:rPr><w:br/></w:r>"
TAB = f"<w:r><w:rPr>{U}</w:rPr><w:tab/></w:r>"


def para(runs, jc="left", extra=""):
    return f'<w:p><w:pPr>{extra}<w:jc w:val="{jc}"/></w:pPr>{"".join(runs)}</w:p>'


wrap = ("Underlined text that is long enough to wrap at the spaces between its words, "
        "so that every line but the last ends at a space where the line breaks. ")
body = [
    para([run("Shelf B1: Item 1 ")]),
    para([run("Before a break "), BR, run("after the break ")]),
    para([run(wrap * 2)]),
    para([run("Centred with trailing spaces   ")], "center"),
    para([run("Right aligned with trailing spaces   ")], "right"),
    para([run("Plain word", ""), run(" ")]),
    para([run("Label:"), TAB, run("tabbed value  ")],
         extra='<w:tabs><w:tab w:val="left" w:pos="2835"/></w:tabs>'),
    para([run("Trailing space then an empty run "), '<w:r><w:t xml:space="preserve"></w:t></w:r>']),
]

off_body = [
    para([run("Shelf B1: Item 1 ")]),
    para([run("Label:"), TAB, run("tabbed value  ")],
         extra='<w:tabs><w:tab w:val="left" w:pos="2835"/></w:tabs>'),
]

variants = {"ul-trail-space-on": (True, body), "ul-trail-space-off": (False, off_body)}


styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="200" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)


def settings(on):
    flag = "<w:ulTrailSpace/>" if on else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
        f'<w:defaultTabStop w:val="709"/><w:compat>{flag}'
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

cases = pathlib.Path(__file__).parent.parent
for name, (on, body) in variants.items():
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
        + "".join(body)
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
        "</w:sectPr></w:body></w:document>"
    )
    out = cases / name / "input.docx"
    out.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", package_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", document_rels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings(on))
    print(f"wrote {out}", file=sys.stderr)
