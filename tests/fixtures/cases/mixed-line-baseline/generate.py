"""A mixed-font line under a multiple draws its baseline on the grid by its glyph box.

Under a multiple above single, a line whose runs come from different fonts
steps by its own pitch; Word for Mac rounds that line's baseline as if its
natural height were its runs' highest ascent to their deepest descent, not
the paragraph's natural height scaled by the multiple. Each block repeats
four mixed lines at 1.15, 1.5, 2, 2.5 and 3 lines; the fifth paragraph wraps,
so its mixed first line steps alone.
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
<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/>
</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>"""

SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" '
        'w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')


def run(font, half_points, text):
    return (f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
            f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/></w:rPr>'
            f'<w:t xml:space="preserve">{text}</w:t></w:r>')


MIXES = [
    run('Verdana', 24, 'Verdana 12 ') + run('Courier New', 24, 'beside Courier New 12'),
    run('Verdana', 32, 'Verdana 16 ') + run('Courier New', 24, 'beside Courier New 12'),
    run('Verdana', 24, 'Verdana 12 ') + run('Courier New', 32, 'beside Courier New 16'),
    run('Arial', 20, 'Arial 10 ') + run('Georgia', 28, 'beside Georgia 14'),
]

WRAP = (run('Verdana', 32, 'Verdana 16 ') + run('Courier New', 24, 'then Courier New 12 ')
        + run('Courier New', 24, 'that runs on and on across the page until it wraps onto a second line '
              'of Courier New alone'))


def para(content, line):
    return (f'<w:p><w:pPr><w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/></w:pPr>'
            f'{content}</w:p>')


body = ''
for line in (276, 360, 480, 600, 720):
    body += ''.join(para(m, line) for m in MIXES) + para(WRAP, line)
body += SECT
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
