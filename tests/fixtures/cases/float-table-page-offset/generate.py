"""A page or margin anchored floating table sits one twip above its tblpY.

Word for Mac puts a table anchored to the page or the margin 1 twip (0.05pt)
above a positive tblpY and never above the page top; a text anchor is exact
(probes: tblpY 1, 3, 946, 957 put the first cell line 0, 0.1, 47.25 and
47.8pt down the page, tblpY -100 at the page top). Each table sits beside
its own page's text, which shows its line on the page grid; the tables are
narrow, so the body runs on unchanged beside them.

1. Page anchor, tblpY 1 (top of the page).
2. Page anchor, tblpY 946.
3. Page anchor, tblpY 957 with a 3pt top border.
4. Margin anchor, tblpY 3.
5. Page anchor, tblpY -100.
6. Text anchor, tblpY 1 (exact).
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
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="en-US"/>
</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>"""

SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" '
        'w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')


def para(text, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def floating_table(n, anchor, y, border=''):
    borders = f'<w:tblBorders>{border}</w:tblBorders>' if border else ''
    cell = (f'<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="2400"/></w:tcPr>'
            f'{para(f"Table {n} first")}{para(f"Table {n} second")}</w:tc>')
    return (f'<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="{anchor}" w:horzAnchor="page" w:tblpX="8400" '
            f'w:tblpY="{y}" w:leftFromText="180" w:rightFromText="180"/><w:tblW w:type="dxa" w:w="2400"/>'
            f'{borders}<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid>'
            f'<w:tr>{cell}</w:tr></w:tbl>')


def case(n, anchor, y, border=''):
    label = para(f'Case {n} body line one', '<w:pageBreakBefore/>' if n > 1 else '')
    return (label + floating_table(n, anchor, y, border)
            + ''.join(para(f'Case {n} body line {k}') for k in ('two', 'three', 'four')))


CASES = [
    ('page', 1),
    ('page', 946),
    ('page', 957, '<w:top w:val="single" w:sz="24" w:space="0" w:color="000000"/>'),
    ('margin', 3),
    ('page', -100),
    ('text', 1),
]

body = ''.join(case(n, *c) for n, c in enumerate(CASES, 1)) + SECT
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
