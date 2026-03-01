"""Create case17: paragraph shading, full borders, and run highlighting."""

import zipfile
import os
import shutil
from lxml import etree

NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
}

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'

def w(tag):
    return f'{{{W}}}{tag}'

def make_run(text, bold=False, italic=False, highlight=None, color=None, font_size=None, font_name=None):
    """Create a w:r element with optional formatting."""
    r = etree.SubElement(etree.Element('dummy'), w('r'))
    rpr = etree.SubElement(r, w('rPr'))
    has_props = False

    if font_name:
        rf = etree.SubElement(rpr, w('rFonts'))
        rf.set(w('ascii'), font_name)
        rf.set(w('hAnsi'), font_name)
        has_props = True
    if font_size:
        sz = etree.SubElement(rpr, w('sz'))
        sz.set(w('val'), str(int(font_size * 2)))  # half-points
        has_props = True
    if bold:
        etree.SubElement(rpr, w('b'))
        has_props = True
    if italic:
        etree.SubElement(rpr, w('i'))
        has_props = True
    if color:
        c = etree.SubElement(rpr, w('color'))
        c.set(w('val'), color)
        has_props = True
    if highlight:
        h = etree.SubElement(rpr, w('highlight'))
        h.set(w('val'), highlight)
        has_props = True

    if not has_props:
        r.remove(rpr)

    t = etree.SubElement(r, w('t'))
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    return r

def make_paragraph(runs=None, style=None, shading=None, borders=None, space_before=None, space_after=None, alignment=None):
    """Create a w:p element.

    shading: dict with 'fill' (hex color), optional 'val' (e.g. 'clear')
    borders: dict with sides ('top','bottom','left','right') each having 'val','sz','color','space'
    """
    p = etree.Element(w('p'))
    ppr = etree.SubElement(p, w('pPr'))
    has_ppr = False

    if style:
        ps = etree.SubElement(ppr, w('pStyle'))
        ps.set(w('val'), style)
        has_ppr = True

    if alignment:
        jc = etree.SubElement(ppr, w('jc'))
        jc.set(w('val'), alignment)
        has_ppr = True

    if space_before is not None or space_after is not None:
        spacing = etree.SubElement(ppr, w('spacing'))
        if space_before is not None:
            spacing.set(w('before'), str(space_before))
        if space_after is not None:
            spacing.set(w('after'), str(space_after))
        has_ppr = True

    if shading:
        shd = etree.SubElement(ppr, w('shd'))
        shd.set(w('val'), shading.get('val', 'clear'))
        shd.set(w('color'), 'auto')
        shd.set(w('fill'), shading['fill'])
        has_ppr = True

    if borders:
        pbdr = etree.SubElement(ppr, w('pBdr'))
        for side in ['top', 'left', 'bottom', 'right']:
            if side in borders:
                b = borders[side]
                el = etree.SubElement(pbdr, w(side))
                el.set(w('val'), b.get('val', 'single'))
                el.set(w('sz'), str(b.get('sz', 4)))
                el.set(w('space'), str(b.get('space', 1)))
                el.set(w('color'), b.get('color', 'auto'))
        has_ppr = True

    if not has_ppr:
        p.remove(ppr)

    if runs:
        for r in runs:
            p.append(r)

    return p


def build_document():
    body = etree.Element(w('body'))

    # --- Heading ---
    body.append(make_paragraph(
        runs=[make_run('Paragraph Shading, Borders & Highlighting', font_size=16, bold=True, color='2E74B5')],
        style='Heading1',
        space_after=200,
    ))

    # --- 1. Paragraph with yellow shading ---
    body.append(make_paragraph(
        runs=[
            make_run('Note: ', bold=True),
            make_run('This paragraph has a yellow background to simulate a note or callout box. Paragraph shading is controlled by the w:shd element in paragraph properties.'),
        ],
        shading={'fill': 'FFFFCC'},
        space_after=200,
    ))

    # --- 2. Paragraph with all-4-side borders ---
    border_def = {'val': 'single', 'sz': 4, 'space': 4, 'color': '000000'}
    body.append(make_paragraph(
        runs=[make_run('This paragraph has borders on all four sides, creating a box effect. Only bottom borders were previously supported.')],
        borders={'top': border_def, 'bottom': border_def, 'left': border_def, 'right': border_def},
        space_after=200,
    ))

    # --- 3. Borders + shading combined (callout box) ---
    body.append(make_paragraph(
        runs=[
            make_run('Warning: ', bold=True, color='CC0000'),
            make_run('This is a warning box with both a light red background and dark red borders. This pattern is common in technical documentation for important notices.'),
        ],
        shading={'fill': 'FFEEEE'},
        borders={
            'top': {'val': 'single', 'sz': 8, 'space': 4, 'color': 'CC0000'},
            'bottom': {'val': 'single', 'sz': 8, 'space': 4, 'color': 'CC0000'},
            'left': {'val': 'single', 'sz': 8, 'space': 4, 'color': 'CC0000'},
            'right': {'val': 'single', 'sz': 8, 'space': 4, 'color': 'CC0000'},
        },
        space_after=200,
    ))

    # --- 4. Normal paragraph with highlighted runs ---
    body.append(make_paragraph(
        runs=[
            make_run('This paragraph contains '),
            make_run('yellow highlighted text', highlight='yellow'),
            make_run(' and also '),
            make_run('cyan highlighted text', highlight='cyan'),
            make_run(' mixed with normal text. Highlighting uses the w:highlight element on individual runs.'),
        ],
        space_after=200,
    ))

    # --- 5. Green info box (left border only, like a blockquote) ---
    body.append(make_paragraph(
        runs=[
            make_run('Tip: ', bold=True, color='2E7D32'),
            make_run('This uses a thick left border with light green shading, a common pattern for tip or info boxes in documentation.'),
        ],
        shading={'fill': 'E8F5E9'},
        borders={
            'left': {'val': 'single', 'sz': 24, 'space': 8, 'color': '2E7D32'},
        },
        space_after=200,
    ))

    # --- 6. Blue info box with all borders ---
    body.append(make_paragraph(
        runs=[
            make_run('Info: ', bold=True, color='1565C0'),
            make_run('A blue-themed information box. Background shading combined with matching colored borders creates a professional look for callouts.'),
        ],
        shading={'fill': 'E3F2FD'},
        borders={
            'top': {'val': 'single', 'sz': 6, 'space': 4, 'color': '1565C0'},
            'bottom': {'val': 'single', 'sz': 6, 'space': 4, 'color': '1565C0'},
            'left': {'val': 'single', 'sz': 6, 'space': 4, 'color': '1565C0'},
            'right': {'val': 'single', 'sz': 6, 'space': 4, 'color': '1565C0'},
        },
        space_after=200,
    ))

    # --- 7. Multiple highlight colors ---
    body.append(make_paragraph(
        runs=[
            make_run('Highlighting comes in many colors: '),
            make_run('yellow', highlight='yellow'),
            make_run(', '),
            make_run('green', highlight='green'),
            make_run(', '),
            make_run('cyan', highlight='cyan'),
            make_run(', '),
            make_run('magenta', highlight='magenta'),
            make_run(', '),
            make_run('red', highlight='red'),
            make_run(', and '),
            make_run('darkYellow', highlight='darkYellow'),
            make_run('. Each uses a different w:highlight value.'),
        ],
        space_after=200,
    ))

    # --- 8. Gray code block style ---
    body.append(make_paragraph(
        runs=[make_run('fn main() {\n    println!("Hello, world!");\n}', font_name='Courier New', font_size=10)],
        shading={'fill': 'F5F5F5'},
        borders={
            'top': {'val': 'single', 'sz': 4, 'space': 4, 'color': 'CCCCCC'},
            'bottom': {'val': 'single', 'sz': 4, 'space': 4, 'color': 'CCCCCC'},
            'left': {'val': 'single', 'sz': 4, 'space': 4, 'color': 'CCCCCC'},
            'right': {'val': 'single', 'sz': 4, 'space': 4, 'color': 'CCCCCC'},
        },
        space_after=200,
    ))

    # --- 9. Final normal paragraph ---
    body.append(make_paragraph(
        runs=[make_run('This final paragraph has no special formatting. It verifies that normal text layout resumes correctly after paragraphs with shading, borders, and highlighting.')],
    ))

    # --- Section properties ---
    sect_pr = etree.SubElement(body, w('sectPr'))
    pg_sz = etree.SubElement(sect_pr, w('pgSz'))
    pg_sz.set(w('w'), '12240')
    pg_sz.set(w('h'), '15840')
    pg_mar = etree.SubElement(sect_pr, w('pgMar'))
    pg_mar.set(w('top'), '1440')
    pg_mar.set(w('right'), '1440')
    pg_mar.set(w('bottom'), '1440')
    pg_mar.set(w('left'), '1440')
    pg_mar.set(w('header'), '720')
    pg_mar.set(w('footer'), '720')
    pg_mar.set(w('gutter'), '0')
    doc_grid = etree.SubElement(sect_pr, w('docGrid'))
    doc_grid.set(w('linePitch'), '360')

    return body


def build_styles():
    """Minimal styles.xml with Normal and Heading1."""
    styles = etree.Element(w('styles'))

    # docDefaults
    doc_defaults = etree.SubElement(styles, w('docDefaults'))
    rpr_default = etree.SubElement(doc_defaults, w('rPrDefault'))
    rpr = etree.SubElement(rpr_default, w('rPr'))
    rf = etree.SubElement(rpr, w('rFonts'))
    rf.set(w('ascii'), 'Aptos')
    rf.set(w('hAnsi'), 'Aptos')
    rf.set(w('eastAsia'), 'Aptos')
    sz = etree.SubElement(rpr, w('sz'))
    sz.set(w('val'), '24')  # 12pt

    ppr_default = etree.SubElement(doc_defaults, w('pPrDefault'))
    ppr = etree.SubElement(ppr_default, w('pPr'))
    spacing = etree.SubElement(ppr, w('spacing'))
    spacing.set(w('after'), '160')
    spacing.set(w('line'), '278')
    spacing.set(w('lineRule'), 'auto')

    # Normal style
    normal = etree.SubElement(styles, w('style'))
    normal.set(w('type'), 'paragraph')
    normal.set(w('default'), '1')
    normal.set(w('styleId'), 'Normal')
    name = etree.SubElement(normal, w('name'))
    name.set(w('val'), 'Normal')

    # Heading1 style
    h1 = etree.SubElement(styles, w('style'))
    h1.set(w('type'), 'paragraph')
    h1.set(w('styleId'), 'Heading1')
    name = etree.SubElement(h1, w('name'))
    name.set(w('val'), 'heading 1')
    h1_ppr = etree.SubElement(h1, w('pPr'))
    h1_spacing = etree.SubElement(h1_ppr, w('spacing'))
    h1_spacing.set(w('before'), '240')
    h1_rpr = etree.SubElement(h1, w('rPr'))
    h1_rf = etree.SubElement(h1_rpr, w('rFonts'))
    h1_rf.set(w('ascii'), 'Aptos')
    h1_rf.set(w('hAnsi'), 'Aptos')
    h1_sz = etree.SubElement(h1_rpr, w('sz'))
    h1_sz.set(w('val'), '32')  # 16pt
    h1_b = etree.SubElement(h1_rpr, w('b'))
    h1_color = etree.SubElement(h1_rpr, w('color'))
    h1_color.set(w('val'), '2E74B5')

    return styles


def create_docx(output_path):
    """Build a DOCX by hand as a ZIP file."""
    body = build_document()
    styles_el = build_styles()

    # Wrap body in w:document
    doc_el = etree.Element(w('document'))
    doc_el.append(body)

    doc_xml = etree.tostring(doc_el, xml_declaration=True, encoding='UTF-8', standalone=True)
    styles_xml = etree.tostring(styles_el, xml_declaration=True, encoding='UTF-8', standalone=True)

    content_types = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>'''

    rels = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''

    doc_rels = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', content_types)
        zf.writestr('_rels/.rels', rels)
        zf.writestr('word/_rels/document.xml.rels', doc_rels)
        zf.writestr('word/document.xml', doc_xml)
        zf.writestr('word/styles.xml', styles_xml)

    print(f'Created {output_path}')


if __name__ == '__main__':
    create_docx('tests/fixtures/cases/case17/input.docx')
