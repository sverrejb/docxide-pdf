"""A merged cell's bottom border does not grow the rows above its last.

Each table has two at-least 28.35pt rows (trHeight 567) whose first cell is
merged down both rows; the table's bottom border reaches the merged cell
only at the second row's foot. Word keeps the first row at its 28.35pt, as
if the merged cell had no bottom border (Word for Mac: a 1.5pt double
bottom under a cell spanning two rows).

1. Double 0.5pt bottom border (1.5pt band).
2. Single 1.5pt bottom border.
3. Single 3pt bottom border, the merged cell spanning three rows.
4. Control: double bottom border, no merged cell.
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
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US"/>
</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>"""

SECT = ('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" '
        'w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>')


def para(text):
    return f'<w:p><w:r><w:t>{text}</w:t></w:r></w:p>'


def cell(width, text, merge=''):
    return f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>{merge}</w:tcPr>{para(text)}</w:tc>'


def table(n, bottom, rows, merged):
    borders = ('<w:tblBorders><w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
               '<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
               f'<w:bottom {bottom} w:space="0" w:color="000000"/>'
               '<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
               '<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
               '<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/></w:tblBorders>')
    out = (f'<w:tbl><w:tblPr><w:tblW w:w="6000" w:type="dxa"/>{borders}'
           '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/>'
           '<w:gridCol w:w="4000"/></w:tblGrid>')
    for r in range(rows):
        if merged:
            first = cell(2000, f'Logo {n}' if r == 0 else '',
                         '<w:vMerge w:val="restart"/>' if r == 0 else '<w:vMerge/>')
        else:
            first = cell(2000, f'Cell {n}.{r + 1}')
        out += ('<w:tr><w:trPr><w:trHeight w:val="567"/></w:trPr>'
                + first + cell(4000, f'Row {n}.{r + 1} text') + '</w:tr>')
    return out + '</w:tbl>' + para(f'After table {n}') + para('')


TABLES = [
    ('w:val="double" w:sz="4"', 2, True),
    ('w:val="single" w:sz="12"', 2, True),
    ('w:val="single" w:sz="24"', 3, True),
    ('w:val="double" w:sz="4"', 2, False),
]

body = ''.join(table(n, *t) for n, t in enumerate(TABLES, 1)) + SECT
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
