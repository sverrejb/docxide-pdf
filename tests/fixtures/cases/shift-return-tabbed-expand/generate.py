"""Generate shift-return-tabbed-expand and shift-return-tabbed-noexpand

A justified line that ends in a manual line break (w:br) is stretched to the
margins unless the document sets w:compat/w:doNotExpandShiftReturn. Word
applies the setting to lines with tabs too: with it, a tabbed line ending in
a break keeps its natural spacing.

The -expand fixture keeps the hanging-indent "Day 1)" control.
The -noexpand fixture sets doNotExpandShiftReturn and also covers two tabs
before text and a wrapped tabbed paragraph whose last line ends in a break.
"""

import pathlib
import sys
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def run(text, extra=""):
    return f'<w:r><w:rPr>{extra}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


TAB = "<w:r><w:tab/></w:r>"
BR = "<w:r><w:br/></w:r>"


def para(runs, ind="", tabs=()):
    stops = "".join(f'<w:tab w:val="left" w:pos="{p}"/>' for p in tabs)
    tabs_xml = f"<w:tabs>{stops}</w:tabs>" if tabs else ""
    return f'<w:p><w:pPr>{tabs_xml}{ind}<w:jc w:val="both"/></w:pPr>{"".join(runs)}</w:p>'


hang = '<w:ind w:left="992" w:hanging="992"/>'
b = "<w:b/>"
body = [
    para([run("Day 1)", b), TAB, run("Arrival at the harbour ", b + '<w:u w:val="single"/>'), BR,
          run("Description of the first stop.")], hang, [993]),
    para([TAB, TAB, run("Two tabs then a few words "), BR, run("next line")], "", [1134, 2268]),
    para([run("Note:"), TAB, run("this paragraph is long enough to wrap onto a second line in the justified "
                                 "text column before the manual line break at the end of it "), BR,
          run("last line.")], hang, [993]),
]

variants = {"shift-return-tabbed-expand": False, "shift-return-tabbed-noexpand": True}


def document(noexpand):
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>'
        + "".join(body if noexpand else body[:1])
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1417" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
        "</w:sectPr></w:body></w:document>"
    )


styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="240" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)


def settings(noexpand):
    compat = "<w:doNotExpandShiftReturn/>" if noexpand else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
        f'<w:defaultTabStop w:val="709"/><w:compat>{compat}'
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
for name, noexpand in variants.items():
    out = cases / name / "input.docx"
    out.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", package_rels)
        z.writestr("word/document.xml", document(noexpand))
        z.writestr("word/_rels/document.xml.rels", document_rels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings(noexpand))
    print(f"wrote {out}", file=sys.stderr)
