"""Generate float-table-band-before: floating tables acting on the text above them

Word places a page's floating tables before it lays out the page's lines. A
full-width page-anchored table keeps earlier paragraphs out of its rows too,
and a paragraph pushed below it is pushed on below any later anchor's
topAndBottom band it now meets. A floating table between a paragraph and a
later anchor does not stop that anchor's band from moving the paragraph.

One section per variant, except the linked mark and continuous sections:
- Paragraph, a full-width table lower on the page, then an anchor whose page
  band covers the paragraph (moves below the band)
- Empty paragraph, a full-width table at the page top, then an anchor whose
  page band starts just below the table (moves below the table and the band)
- Paragraph with page break before at the top of its page, above a page-top
  table (moves below the table, no blank page)
- Inline text box above a later anchor's page band (moves below the band)
- Inline canvas of a rectangle and connector above a page-top table (both
  move below the table with the line)
- A paragraph with 24pt after, then one with page break before and 12pt
  before above a page-top table: on the next page on the table's bottom;
  with 10pt after instead, 2pt below it (only the space before beyond the
  space after above it stays, as below a band)
- A section ending in a connector-only section mark with page break before,
  then a continuous section whose second paragraph has page break before:
  the mark takes its line and connector on a new page, and the later break
  still starts a page
- A paragraph with 24pt after above a page-top table: the table keeps its
  place, the paragraph moves below it
- A line with page break before and 28pt before above a page-top table
  60-100pt down: below the table with its 28pt before (at the page top it
  would lose them and run into the table)
- A paragraph with 24pt after, then one with page break before holding a
  topAndBottom box 40pt below it: the lines after it step below the box
  where it is drawn (the space after is gone with the break)
"""

import pathlib
import zipfile

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"'
)
SECT_CONTINUOUS = (
    '<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" '
    'w:right="1440" w:bottom="1440" w:left="1440" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'
)
EMU = 12700
SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" '
    'w:left="1440" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr>'
)
PAGE_W = 595.3


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(content="", line=13, sect=False, before=0, after=0, page_break=False):
    """`line`: an exact line height in pt, or None for single spacing."""
    exact = f' w:line="{int(line * 20)}" w:lineRule="exact"' if line is not None else ""
    ppr = (
        ("<w:pageBreakBefore/>" if page_break else "")
        + f'<w:spacing w:before="{before * 20}" w:after="{after * 20}"{exact}/>'
        + (SECT if sect is True else sect or "")
    )
    return f"<w:p><w:pPr>{ppr}</w:pPr>{content}</w:p>"


def band(i, y, h, text):
    """A full-width topAndBottom textbox `y` pt below the page top."""
    return (
        f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        f'relativeHeight="{i}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{y * EMU}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{451 * EMU}" cy="{h * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/>'
        f'<wp:docPr id="{i}" name="Band {i}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{451 * EMU}" cy="{h * EMU}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="D9E2F3"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>'
        f'<wps:txbx><w:txbxContent><w:p>{run(text)}</w:p></w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )


def hanging(i, off, h, text):
    """A full-width topAndBottom textbox `off` pt below its paragraph's top."""
    return (
        f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
        f'relativeHeight="{i}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{off * EMU}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{451 * EMU}" cy="{h * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapTopAndBottom/>'
        f'<wp:docPr id="{i}" name="Hang {i}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{451 * EMU}" cy="{h * EMU}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFE699"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>'
        f'<wps:txbx><w:txbxContent><w:p>{run(text)}</w:p></w:txbxContent></wps:txbx>'
        '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )


def table(y, h, w, x=0, v_anchor="page", text="Floating table", rows=1, tbl_w=None):
    tbl_w = tbl_w or f'w:w="{int(w * 20)}" w:type="dxa"'
    return (
        f'<w:tbl><w:tblPr><w:tblpPr w:leftFromText="0" w:rightFromText="0" w:vertAnchor="{v_anchor}" '
        f'w:horzAnchor="page" w:tblpX="{int(x * 20)}" w:tblpY="{int(y * 20)}"/><w:tblOverlap w:val="never"/>'
        f'<w:tblW {tbl_w}/><w:tblLayout w:type="fixed"/>'
        '<w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{int(w * 20)}"/></w:tblGrid>'
        + (
            f'<w:tr><w:trPr><w:trHeight w:val="{int(h * 20)}" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="{int(w * 20)}" w:type="dxa"/>'
            '<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/></w:tcPr>'
            f'<w:p>{run(text)}</w:p></w:tc></w:tr>'
        )
        * rows
        + "</w:tbl>"
    )


def inline_box(i, text):
    """An inline 90x16pt text box."""
    return (
        f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{90 * EMU}" cy="{16 * EMU}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{i}" name="Inline {i}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{90 * EMU}" cy="{16 * EMU}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFD8D8"/></a:solidFill>'
        '<a:ln w="6350"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr>'
        f'<wps:txbx><w:txbxContent><w:p>{run(text)}</w:p></w:txbxContent></wps:txbx>'
        '<wps:bodyPr wrap="square" lIns="45720" tIns="0" rIns="45720" bIns="0"><a:noAutofit/></wps:bodyPr>'
        "</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )


def connector_shape(i, x, w, h):
    return (
        f'<wps:wsp><wps:cNvPr id="{i}" name="Line {i}"/><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="{x * EMU}" y="0"/>'
        f'<a:ext cx="{w * EMU}" cy="{h * EMU}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        '<a:ln w="19050"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>'
    )


def inline_canvas(i):
    """An inline 60x16pt canvas: a rectangle and a connector line."""
    rect = (
        f'<wps:wsp><wps:cNvPr id="{i + 100}" name="Rect"/><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
        f'<a:ext cx="{30 * EMU}" cy="{16 * EMU}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:solidFill><a:srgbClr val="6080D0"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>'
    )
    return (
        f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{60 * EMU}" cy="{16 * EMU}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{i}" name="Canvas {i}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas">'
        f'<wpc:wpc><wpc:bg/><wpc:whole/>{rect}{connector_shape(i + 200, 30, 30, 16)}</wpc:wpc>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )


def line_anchor(i):
    """A wrapNone 200pt connector line 20pt below its paragraph."""
    return (
        f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="{i}" '
        'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="paragraph"><wp:posOffset>{20 * EMU}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{200 * EMU}" cy="0"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{i}" name="Line {i}"/><wp:cNvGraphicFramePr/>'
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
        f"{connector_shape(i, 0, 200, 0)}</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
    )


variants = [
    [para(run("Above a later band")), table(300, 40, PAGE_W),
     para(band(1, 60, 60, "Page band 60-120pt") + run("Anchor of the table and the band"))],
    [para(), table(0, 100, PAGE_W),
     para(band(2, 105, 100, "Page band 105-205pt") + run("Anchor of the table and the band"), line=10.5)],
    [para(run("Page break before, above a page-top table"), page_break=True, line=None), table(0, 100, PAGE_W),
     para(run("Anchor of the table"))],
    [para(inline_box(5, "Inline box"), line=None), para(band(6, 60, 60, "Page band 60-120pt") + run("Anchor of the band"))],
    [para(inline_canvas(7), line=None), table(0, 100, PAGE_W), para(run("Anchor of the table"))],
    [para(run("24pt after"), after=24), para(run("Page break before, 12pt before"), page_break=True, before=12, line=None),
     table(0, 100, PAGE_W), para(run("Anchor of the table"))],
    [para(run("10pt after"), after=10), para(run("Page break before, 12pt before"), page_break=True, before=12, line=None),
     table(0, 100, PAGE_W), para(run("Anchor of the table"))],
    ("mark", [para(run("Above a page-top table")), table(0, 100, PAGE_W), para(run("Anchor of the table")),
              para(line_anchor(9), page_break=True, sect=True, line=None)]),
    ("continuous", [para(run("Continuous section")), para(run("Page break before"), page_break=True, line=None),
                    para(run("After the break"))]),
    [para(run("24pt after, above a page-top table"), after=24), table(0, 100, PAGE_W), para(run("Anchor of the table"))],
    [para(run("No space after")), para(run("Page break before, 28pt before"), page_break=True, before=28),
     table(60, 40, PAGE_W)],
    [para(band(10, 700, 30, "Page band 700pt") + run("24pt after"), after=24),
     para(run("Page break before, a box hanging 40pt below") + hanging(11, 40, 30, "Hanging box"), page_break=True)]
    + [para(run(f"Following line {k}")) for k in range(1, 5)],
]
body = []
for n, variant in enumerate(variants):
    kind, blocks = variant if isinstance(variant, tuple) else ("page", variant)
    if kind == "mark":
        body += blocks
        continue
    sect = n < len(variants) - 1 and (SECT_CONTINUOUS if kind == "continuous" else True)
    body += blocks + [para(run("Next paragraph"), sect=sect)]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body) + SECT + "</w:body></w:document>"
)

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)
settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
    "</w:settings>"
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

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
print(f"wrote {out}")
