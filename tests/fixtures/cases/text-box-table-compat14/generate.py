"""Generate text-box-table-compat14: text box tables under Word 2010 layout
(compatibilityMode 14).

Before Word 2013 a left-aligned table with a style or cell margins is
outdented by its first cell's left margin, so the cell text lines up with the
paragraphs; Word does the same inside a text box, measured from the box's
text area. Without cell margins there is nothing to outdent.

Tests (250 x 150pt boxes, Word compat 14):
- TableGrid table with the default 5.4pt cell margins
- tblCellMar left/right 0 with cell shading
- 3pt borders
- tblInd 360
"""

import pathlib
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = (
    f'xmlns:w="{W}" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
)
EMU = 12700
TWO = [["Name:", "Ada"], ["Date:", "15 January 2025"]]
_ids = iter(range(1, 1000))


def para(text="", bold=False, jc=None, before=0, after=0):
    ppr = f'<w:spacing w:before="{before}" w:after="{after}"/>' + (f'<w:jc w:val="{jc}"/>' if jc else "")
    rpr = "<w:rPr><w:b/></w:rPr>" if bold else ""
    run = f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p><w:pPr>{ppr}</w:pPr>{run}</w:p>"


def borders(sz=4, val="single"):
    return "<w:tblBorders>" + "".join(
        f'<w:{s} w:val="{val}" w:sz="{sz}" w:space="0" w:color="000000"/>'
        for s in ("top", "left", "bottom", "right", "insideH", "insideV")) + "</w:tblBorders>"


def table(grid, rows, width=("auto", 0), style="TableGrid", jc=None, ind=None, bdr="", layout=None,
          cell_mar=None, shade=None, row_h=None, v_align=None):
    """Rows of cell texts over `grid` (twips); each cell names its grid width."""
    pr = f'<w:tblStyle w:val="{style}"/>' if style else ""
    pr += f'<w:tblW w:w="{width[1]}" w:type="{width[0]}"/>'
    pr += f'<w:jc w:val="{jc}"/>' if jc else ""
    pr += f'<w:tblInd w:w="{ind}" w:type="dxa"/>' if ind is not None else ""
    pr += bdr
    pr += f'<w:tblLayout w:type="{layout}"/>' if layout else ""
    if cell_mar:
        pr += "<w:tblCellMar>" + "".join(
            f'<w:{k} w:w="{v}" w:type="dxa"/>' for k, v in cell_mar.items()) + "</w:tblCellMar>"
    out = f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>" + "".join(
        f'<w:gridCol w:w="{c}"/>' for c in grid) + "</w:tblGrid>"
    for row in rows:
        out += "<w:tr>" + (f'<w:trPr><w:trHeight w:val="{row_h}"/></w:trPr>' if row_h else "")
        for i, text in enumerate(row):
            tc = f'<w:tcW w:w="{grid[i]}" w:type="dxa"/>'
            tc += f'<w:shd w:val="clear" w:color="auto" w:fill="{shade}"/>' if shade else ""
            tc += f'<w:vAlign w:val="{v_align}"/>' if v_align else ""
            out += f"<w:tc><w:tcPr>{tc}</w:tcPr>{para(text, bold=i == 0)}</w:tc>"
        out += "</w:tr>"
    return out + "</w:tbl>"


def shape(story, w, h, off=(0, 0), anchor="t", autofit="noAutofit", ins=(91440, 45720, 91440, 45720),
          fill="EEEEEE", line="808080", prst="rect"):
    """A wps shape (sizes in EMU); `story` None for one without text."""
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    line_xml = (f'<a:ln w="6350"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line
                else "<a:ln><a:noFill/></a:ln>")
    txbx = f"<wps:txbx><w:txbxContent>{story}</w:txbxContent></wps:txbx>" if story is not None else ""
    l, t, r, b = ins
    nv = '<wps:cNvSpPr txBox="1"/>' if story is not None else "<wps:cNvSpPr/>"
    return (f'<wps:wsp>{nv}<wps:spPr>'
            f'<a:xfrm><a:off x="{off[0]}" y="{off[1]}"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="{prst}"><a:avLst/></a:prstGeom>{fill_xml}{line_xml}</wps:spPr>{txbx}'
            f'<wps:bodyPr wrap="square" lIns="{l}" tIns="{t}" rIns="{r}" bIns="{b}" anchor="{anchor}">'
            f"<a:{autofit}/></wps:bodyPr></wps:wsp>")


def drawing(inner, requires, uri):
    return (f'<w:r><mc:AlternateContent><mc:Choice Requires="{requires}"><w:drawing>{inner}</w:drawing>'
            f'</mc:Choice></mc:AlternateContent></w:r>').replace("URI", uri)


WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
WPG_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"


def anchored(x, y, w, h, graphic, requires="wps"):
    """Anchored at (x, y) pt from the page corner, w x h pt."""
    n = next(_ids)
    uri = WPS_URI if requires == "wps" else WPG_URI
    return drawing(
        f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{n}" '
        f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
        f'<wp:positionH relativeFrom="page"><wp:posOffset>{x * EMU}</wp:posOffset></wp:positionH>'
        f'<wp:positionV relativeFrom="page"><wp:posOffset>{y * EMU}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{w * EMU}" cy="{h * EMU}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
        f'<wp:docPr id="{n}" name="Shape {n}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic><a:graphicData uri="URI">{graphic}</a:graphicData></a:graphic></wp:anchor>',
        requires, uri)


def box(x, y, w, h, story, **kw):
    return anchored(x, y, w, h, shape(story, w * EMU, h * EMU, **kw))


def group(ext, child_ext, children):
    return (f'<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
            f'<a:ext cx="{ext[0]}" cy="{ext[1]}"/><a:chOff x="0" y="0"/>'
            f'<a:chExt cx="{child_ext[0]}" cy="{child_ext[1]}"/></a:xfrm></wpg:grpSpPr>{children}</wpg:wgp>')


def inline_group(ext, child_ext, children):
    n = next(_ids)
    return drawing(
        f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{ext[0]}" cy="{ext[1]}"/>'
        f'<wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="{n}" name="Group {n}"/><wp:cNvGraphicFramePr/>'
        f'<a:graphic><a:graphicData uri="URI">{group(ext, child_ext, children)}</a:graphicData></a:graphic>'
        "</wp:inline>", "wpg", WPG_URI)


def vml_box(x, y, w, h, story):
    n = next(_ids)
    return (f'<w:r><w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" '
            f'path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/>'
            f'<v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>'
            f'<v:shape id="Box{n}" type="#_x0000_t202" style="position:absolute;margin-left:{x}pt;'
            f'margin-top:{y}pt;width:{w}pt;height:{h}pt;z-index:{n};mso-position-horizontal-relative:page;'
            f'mso-position-vertical-relative:page" fillcolor="#eeeeee" strokecolor="#808080" strokeweight=".5pt">'
            f'<v:textbox><w:txbxContent>{story}</w:txbxContent></v:textbox></v:shape></w:pict></w:r>')


def write(path, body, compat):
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}'
                '<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" '
                'w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>'
                "</w:body></w:document>")
    styles = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W}"><w:docDefaults>'
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/>'
        '<w:szCs w:val="20"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
        '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
        '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>'
        '<w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" '
        'w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
        '</w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
        f'<w:basedOn w:val="TableNormal"/><w:tblPr>{borders()}</w:tblPr></w:style></w:styles>')
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W}">'
                '<w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" '
                f'w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ct = ("application/vnd.openxmlformats-officedocument.wordprocessingml")
    parts = {
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            f'<Override PartName="/word/document.xml" ContentType="{ct}.document.main+xml"/>'
            f'<Override PartName="/word/styles.xml" ContentType="{ct}.styles+xml"/>'
            f'<Override PartName="/word/settings.xml" ContentType="{ct}.settings+xml"/></Types>'),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>'),
        "word/_rels/document.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>'
            f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/></Relationships>'),
        "word/document.xml": document,
        "word/styles.xml": styles,
        "word/settings.xml": settings,
    }
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            z.writestr(name, data)
    print(f"wrote {path}")


COMPAT = 14
W_, H_ = 250, 150
boxes = [
    (para("default margins") + table([1800, 1800], TWO, ("dxa", 3600)) + para("after"), {}),
    (para("cellMar 0") + table([1800, 1800], TWO, ("dxa", 3600), cell_mar={"left": 0, "right": 0},
                               shade="DDEBF7") + para("after"), {}),
    (para("3pt borders") + table([1800, 1800], TWO, ("dxa", 3600), bdr=borders(sz=24)) + para("after"), {}),
    (para("tblInd 360") + table([1800, 1800], TWO, ("dxa", 3600), ind=360) + para("after"), {}),
]
runs = "".join(box(40 + 280 * (i % 2), 40 + 180 * (i // 2), W_, H_, story, **kw)
               for i, (story, kw) in enumerate(boxes))
BODY = f"<w:p>{runs}</w:p>"

write(pathlib.Path(__file__).parent / "input.docx", BODY, COMPAT)
