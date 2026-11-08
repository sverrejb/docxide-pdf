"""A lowered run in a top-aligned line grows it by one font unit less.

With w:textAlignment="top" Word hangs the glyphs from the line top, so a run
lowered by w:position drops into the font's external leading and one font
unit before the line grows (Word for Mac: Calibri 11pt lowered 0.5pt steps
13.922, not 13.928; Calibri 14pt 17.583; Arial 11pt, whose 67-unit leading
goes first, 12.784). Each block is twelve lines, half of them empty
paragraphs whose mark carries the position, so the unit adds up.

1. Calibri 11pt lowered 0.5pt.
2. Calibri 14pt lowered 0.5pt.
3. Calibri 11pt lowered 1pt.
4. Arial 11pt lowered 0.5pt.
5. Times New Roman 11pt lowered 3pt.
6. Control: Calibri 11pt lowered 0.5pt, baseline alignment.
"""

import zipfile
from pathlib import Path

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
</Types>"""

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOC_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
</Relationships>"""

SETTINGS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings {W}><w:defaultTabStop w:val="720"/><w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
</w:compat></w:settings>"""

STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>
<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/>
</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
</w:styles>"""

SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" '
        'w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')


def block(n, font, half_points, position, top=True):
    align = '<w:textAlignment w:val="top"/>' if top else ''
    rpr = (f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
           f'<w:position w:val="{position}"/><w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/>')
    out = ''
    for k in range(12):
        run = f'<w:r><w:rPr>{rpr}</w:rPr><w:t>Block {n} line {k + 1}</w:t></w:r>' if k % 2 == 0 else ''
        out += f'<w:p><w:pPr>{align}<w:rPr>{rpr}</w:rPr></w:pPr>{run}</w:p>'
    return out + '<w:p/>'


BLOCKS = [
    ('Calibri', 22, -1),
    ('Calibri', 28, -1),
    ('Calibri', 22, -2),
    ('Arial', 22, -1),
    ('Times New Roman', 22, -6),
    ('Calibri', 22, -1, False),
]

body = ''.join(block(n, *b) for n, b in enumerate(BLOCKS, 1)) + SECT
DOCUMENT = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:document {W}><w:body>{body}</w:body></w:document>')

out = Path(__file__).with_name('input.docx')
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in [('[Content_Types].xml', CONTENT_TYPES), ('_rels/.rels', RELS),
                       ('word/_rels/document.xml.rels', DOC_RELS), ('word/document.xml', DOCUMENT),
                       ('word/styles.xml', STYLES), ('word/settings.xml', SETTINGS)]:
        info = zipfile.ZipInfo(name, (1980, 1, 1, 0, 0, 0))
        info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, data)
