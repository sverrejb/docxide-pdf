"""The line after a clearing break below a floating table is the mark's.

A paragraph of `<w:br w:clear="all"/>` runs after a floating table resumes
below the table: each further break's line takes the break run's height and
the line after the last break takes the paragraph mark's, whatever the runs
(Word for Mac: 12pt breaks under a 20pt mark resume with a 23pt line, 20pt
breaks under a 12pt mark step 23pt and then 13.8pt). The table is 60pt tall
(at-least row) and 1 twip below the anchor; one case per page.

1. One 12pt break, Times New Roman 12pt mark (no run draws the mark's size).
2. One 12pt break under a 20pt mark.
3. One 20pt break under a 12pt mark.
4. Two 20pt breaks under a 12pt mark.
5. Two 12pt breaks under a 20pt mark.
6. Text and a 12pt break under a 20pt mark.
7. Two breaks at 1.25 lines.
8. One Arial 12pt break.
9. One 12pt break under a 20pt Verdana mark (no run draws the mark's font).
10. A 12pt clearing break, then a 20pt break, under a 12pt mark.
11. A 20pt clearing break, then two 12pt breaks, under a 12pt mark.
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
<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="en-US"/>
</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>
</w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
</w:styles>"""

SECT = ('<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" '
        'w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>')


def rpr(font=None, half_points=None):
    out = ''
    if font:
        out += f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>'
    if half_points:
        out += f'<w:sz w:val="{half_points}"/><w:szCs w:val="{half_points}"/>'
    return out


def para(text, ppr=''):
    return f'<w:p><w:pPr>{ppr}</w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def floating_table(n):
    return ('<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="text" w:horzAnchor="margin" w:tblpY="1" '
            'w:leftFromText="180" w:rightFromText="180"/><w:tblW w:type="dxa" w:w="4000"/>'
            '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid>'
            '<w:tr><w:trPr><w:trHeight w:val="1200" w:hRule="atLeast"/></w:trPr><w:tc><w:tcPr>'
            f'<w:tcW w:type="dxa" w:w="4000"/></w:tcPr>{para(f"Table {n}")}</w:tc></w:tr></w:tbl>')


def case(n, breaks=1, break_hp=None, mark_hp=None, text='', line=240, font=None,
         mark_font=None, sizes=None):
    spacing = f'<w:spacing w:after="0" w:line="{line}" w:lineRule="auto"/>'
    clear_each = sizes is None
    sizes = sizes or [break_hp] * breaks
    brs = ''.join(f'<w:r><w:rPr>{rpr(font, hp)}</w:rPr><w:br w:type="textWrapping"'
                  f'{" w:clear=\"all\"" if clear_each or i == 0 else ""}/></w:r>'
                  for i, hp in enumerate(sizes))
    lead = f'<w:r><w:rPr>{rpr(font)}</w:rPr><w:t>{text}</w:t></w:r>' if text else ''
    clearing = (f'<w:p><w:pPr>{spacing}<w:rPr>{rpr(mark_font or font, mark_hp)}</w:rPr></w:pPr>'
                f'{lead}{brs}</w:p>')
    label = para(f'Case {n}', '<w:pageBreakBefore/>' if n > 1 else '')
    return (label + floating_table(n) + clearing
            + para(f'Body {n} resumes below the table') + para(f'Next {n} line'))


CASES = [
    dict(),
    dict(mark_hp=40),
    dict(break_hp=40),
    dict(breaks=2, break_hp=40),
    dict(breaks=2, mark_hp=40),
    dict(text='Text', mark_hp=40),
    dict(breaks=2, line=300),
    dict(font='Arial'),
    dict(mark_hp=40, mark_font='Verdana'),
    dict(sizes=[24, 40]),
    dict(sizes=[40, 24, 24]),
]

body = ''.join(case(n, **c) for n, c in enumerate(CASES, 1)) + SECT
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
