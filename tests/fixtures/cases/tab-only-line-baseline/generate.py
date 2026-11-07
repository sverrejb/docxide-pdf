"""Generate tab-only-line-baseline: underlined tabs on lines with no text

A form's signature line is often a paragraph holding one underlined tab. No
glyph sizes that line, so Word sizes it by the paragraph mark's font, and
its baseline sits that font's ascent below the line top like a text line's
(Word for Mac behavior, online export): the tab's underline hangs from it at
the tab run's own underline offset and thickness, whatever size the tab run
is. Columns: Arial 11pt, Times New Roman 12pt, Calibri 14pt marks; rows: a
tab-only line at single spacing, at 1.5 lines and with a 20pt tab run.
"""

import pathlib
import zipfile

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
FONTS = (("Arial", 22), ("Times New Roman", 24), ("Calibri", 28))


def rpr(font, size, u=None):
    under = f'<w:u w:val="{u}"/>' if u else ""
    return (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{size}"/>{under}</w:rPr>')


def para(content, mark, before=0, line=240):
    return (f'<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="4800"/></w:tabs>'
            f'<w:spacing w:before="{before}" w:line="{line}" w:lineRule="auto"/>{mark}</w:pPr>{content}</w:p>')


def text(t, r):
    return f'<w:r>{r}<w:t xml:space="preserve">{t}</w:t></w:r>'


def tab(r):
    return f'<w:r>{r}<w:tab/></w:r>'


def section(font, size):
    plain, under = rpr(font, size), rpr(font, size, "single")
    mark = rpr(font, size, "dotted")
    cases = [
        ("tab only", para(tab(under), mark)),
        ("at 1.5 lines", para(tab(under), mark, line=360)),
        ("20pt tab run", para(tab(rpr(font, 40, "single")), mark)),
    ]
    out = []
    for name, p in cases:
        out.append(para(text(f"{font} {size // 2}pt, {name}:", plain), plain, before=120))
        out.append(p)
        out.append(para(text("Line below", plain), plain))
    return out


body = []
for n, (font, size) in enumerate(FONTS):
    if n:
        body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
    body.extend(section(font, size))

document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(body)}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>'
            "</w:sectPr></w:body></w:document>")
styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          "</w:docDefaults></w:styles>")
settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            "</w:compat></w:settings>")
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
    "</Types>")
rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
package_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>')
document_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/></Relationships>')
out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
print(f"wrote {out}")
