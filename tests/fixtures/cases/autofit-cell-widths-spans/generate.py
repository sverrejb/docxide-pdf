"""Generate autofit-cell-widths-spans and autofit-cell-widths-nested

AutoFit tables of auto width (tblW auto) whose cells name their widths in
twips, as in autofit-cell-widths: Word lays the columns out from the cells.

autofit-cell-widths-spans (top-level tables):
- N1 uniform cells against an asymmetric grid: columns take the cells' widths
- S1-S4 a merged cell wider than the columns it spans widens the last of
  them by the difference: equal and unequal columns, two merged cells over
  different columns, and a merged cell in the first row
- S5-S7 columns without a cell of their own: a column only merged cells
  cover starts at zero; one a row leaves empty (gridBefore) starts at its grid width
- R1-R2 bidiVisual: the merged cell widens the last column in reading order
- C1-C9 a cell narrower than its picture or longest word: single and merged
  cells, equal and empty columns, columns already wide enough, and columns
  with different longest words

autofit-cell-widths-nested (inside a fixed 300pt cell):
- E1-E2 one and two rows with a grid wider than the cells
- E3 two rows whose grid agrees with the cells
- E4 a merged cell wider than its columns
- E5 text longer than its cell wraps in it
"""

import io
import pathlib
import sys
import zipfile

from PIL import Image, ImageDraw

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
     'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"')


def p(text):
    r = text if text.startswith("<w:r>") else f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'
    return f'<w:p><w:pPr><w:spacing w:after="0"/></w:pPr>{r}</w:p>'


def picture(width, height, ident):
    w, h = round(width * 12700), round(height * 12700)
    return (f'<w:r><w:drawing><wp:inline><wp:extent cx="{w}" cy="{h}"/>'
            f'<wp:docPr id="{ident}" name="Picture {ident}"/>'
            '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:nvPicPr><pic:cNvPr id="{ident}" name="Picture"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


B = ''.join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
            for s in ["top", "left", "bottom", "right", "insideH", "insideV"])
AUTO = '<w:tblW w:w="0" w:type="auto"/>'


def table(grid, rows, tblw=AUTO, layout="", bidi=False):
    """rows: list of rows, each a list of (text, tcw, span); ("before"|"after", n) leaves n columns empty."""
    g = "".join(f'<w:gridCol w:w="{w}"/>' for w in grid)
    body = ""
    for row in rows:
        cells = ""
        tr_pr = "".join(f'<w:grid{c[0].title()} w:val="{c[1]}"/>' for c in row if c[0] in ("before", "after"))
        if tr_pr:
            cells += f"<w:trPr>{tr_pr}</w:trPr>"
        for text, w, span in (c for c in row if c[0] not in ("before", "after")):
            tw = f'<w:tcW w:w="{w}" w:type="dxa"/>' if w is not None else '<w:tcW w:w="0" w:type="auto"/>'
            gs = f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
            content = text if text.startswith("<w:tbl>") else p(text)
            if text.startswith("<w:tbl>"):
                content += p("")
            cells += f'<w:tc><w:tcPr>{tw}{gs}</w:tcPr>{content}</w:tc>'
        body += f'<w:tr>{cells}</w:tr>'
    layout += "<w:bidiVisual/>" if bidi else ""
    return (f'<w:tbl><w:tblPr>{tblw}{layout}<w:tblBorders>{B}</w:tblBorders>'
            f'<w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid>{g}</w:tblGrid>{body}</w:tbl>')


def simple(grid, tcws, texts=None):
    rows = []
    for r, widths in enumerate(tcws):
        rows.append([((texts or ["One", "Two", "Three", "Four"])[i], w, 1) for i, w in enumerate(widths)])
    return table(grid, rows)


def host(inner):
    return table([6000], [[(inner, 6000, 1)]], '<w:tblW w:w="6000" w:type="dxa"/>',
                 '<w:tblLayout w:type="fixed"/>')


spans = [
    ("N1", simple([1000, 2000], [[2000, 2000]])),
    ("S1", table([2000, 2000], [[("One", 2000, 1), ("Two", 2000, 1)], [("Span", 6000, 2)]])),
    ("S2", table([2000, 2000], [[("One", 1000, 1), ("Two", 3000, 1)], [("Span", 6000, 2)]])),
    ("S3", table([2000, 2000, 2000], [[("One", 2000, 1), ("Two", 2000, 1), ("Three", 2000, 1)],
                                      [("Span", 6000, 2), ("Last", 2000, 1)],
                                      [("First", 2000, 1), ("Span", 6000, 2)]])),
    ("S4", table([2000, 2000], [[("Span", 6000, 2)], [("One", 2000, 1), ("Two", 2000, 1)]])),
    ("S5", table([1000, 1000], [[("Span", 6000, 2)]])),
    ("S6", table([1000, 4000, 1000], [[("Span", 3000, 2), ("Three", 1000, 1)],
                                       [("One", 1000, 1), ("Span", 5000, 2)]])),
    ("S7", table([1000, 2000, 2000], [[("before", 1), ("Two", 2000, 1), ("Three", 2000, 1)],
                                       [("Span", 4000, 2), ("Three", 2000, 1)]])),
    ("R1", table([2000, 2000], [[("One", 1000, 1), ("Two", 3000, 1)], [("Span", 6000, 2)]], bidi=True)),
    ("R2", table([1000, 1000, 1000], [[("Span", 3000, 2), ("Three", 1000, 1)],
                                      [("One", 1000, 1), ("Span", 5000, 2)]], bidi=True)),
    ("C1", table([3000], [[(picture(120, 30, 5), 1000, 1)]])),
    ("C2", table([3000, 3000], [[(picture(120, 30, 1), 1000, 2)]])),
    ("C3", table([2000, 2000], [[("One", 1000, 1), ("Two", 1000, 1)], [(picture(120, 30, 2), 1000, 2)]])),
    ("C4", table([2000, 2000], [[("One", 600, 1), ("Two", 600, 1)], [("Unbreakablelongword", 1000, 2)]])),
    ("C5", table([2000, 2000, 2000], [[("One", 400, 1), ("Two", 800, 1), ("Three", 1600, 1)],
                                      [(picture(120, 30, 6), 1000, 3)]])),
    ("C6", table([2000, 2000], [[("Wordwordw", 1200, 1), ("Ab", 1200, 1)], [(picture(200, 30, 7), 1000, 2)]])),
    ("C7", table([2000, 2000], [[("Ab cd ef gh", 1200, 1), ("Ab", 1200, 1)], [(picture(200, 30, 8), 1000, 2)]])),
    ("C8", table([2000, 2000], [[("Ab", 400, 1), ("Ab", 1600, 1)], [(picture(200, 30, 9), 1000, 2)]])),
    ("C9", table([2000, 2000], [[("", 400, 1), ("", 1600, 1)], [(picture(200, 30, 10), 1000, 2)]])),
]
nested = [
    ("E1", host(simple([2000, 2000], [[1200, 1600]]))),
    ("E2", host(simple([2000, 2000], [[1200, 1600], [1200, 1600]]))),
    ("E3", host(simple([1200, 1600], [[1200, 1600], [1200, 1600]]))),
    ("E4", host(table([2000, 2000], [[("One", 1200, 1), ("Two", 1600, 1)], [("Span", 4000, 2)]]))),
    ("E5", host(simple([2000, 2000], [[1200, 1600]], ["A longer text that wraps here", "Two"]))),
]
variants = {"autofit-cell-widths-spans": spans, "autofit-cell-widths-nested": nested}

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)
settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}>'
    '<w:defaultTabStop w:val="709"/><w:compat>'
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    "</w:compat></w:settings>"
)
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
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
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/>'
    f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/image1.png"/></Relationships>'
)


def placeholder_png():
    img = Image.new("RGB", (240, 60), (210, 225, 240))
    ImageDraw.Draw(img).rectangle([0, 0, 239, 59], outline=(40, 70, 120), width=3)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

cases_dir = pathlib.Path(__file__).parent.parent
for name, cases in variants.items():
    body = "".join(p(label) + tbl + p("") for label, tbl in cases)
    document = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>' + body
        + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1000" w:right="1134" w:bottom="1000" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
        "</w:sectPr></w:body></w:document>"
    )
    out = cases_dir / name / "input.docx"
    out.parent.mkdir(exist_ok=True)
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", package_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", document_rels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
        z.writestr("word/media/image1.png", placeholder_png())
    print(f"wrote {out}", file=sys.stderr)
