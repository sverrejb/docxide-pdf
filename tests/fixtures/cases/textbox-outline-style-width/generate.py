"""Generate textbox-outline-style-width: a drawn text box outline without its
own w takes the shape style's line width, for the stroke and the inset.

Word draws an `a:ln` that has a solidFill but no w with the width of the
theme line its wps:style lnRef names (idx 1-3, here 2, 6 and 12pt), and moves
the text in by half that width on every side, like an invisible outline. With
lnRef idx 0 or no style the line is Word's default 0.75pt, but the text keeps
the plain insets. An explicit w (also w="0", a hairline) overrides the style
for both.

Tests (150 x 50pt boxes; left-aligned top-anchored text):
- solidFill without w, lnRef 0 (control) and lnRef 3
- solidFill with w 0 (control) and 2pt, lnRef 3
"""
import pathlib
import struct
import zipfile
import zlib

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = (
    f'xmlns:w="{W}" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"'
)
EMU = 12700
WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
_ids = iter(range(1, 10000))


def png(w, h, color=(200, 60, 40), border=(0, 0, 0)):
    """A w x h RGB PNG: `color` with a 2px `border` and a diagonal."""
    rows = []
    for y in range(h):
        row = bytearray([0])
        for x in range(w):
            edge = x < 2 or y < 2 or x >= w - 2 or y >= h - 2 or abs(x * h - y * w) < max(w, h)
            row += bytes(border if edge else color)
        rows.append(bytes(row))
    raw = zlib.compress(b"".join(rows))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", raw) + chunk(b"IEND", b""))


def run(text, bold=False, size=None):
    rpr = ("<w:b/>" if bold else "") + (f'<w:sz w:val="{size}"/>' if size else "")
    rpr = f"<w:rPr>{rpr}</w:rPr>" if rpr else ""
    return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'


def para(text="", bold=False, jc=None, before=0, after=0, extra=""):
    ppr = f'<w:spacing w:before="{before}" w:after="{after}"/>' + (f'<w:jc w:val="{jc}"/>' if jc else "")
    body = run(text, bold) if text else ""
    return f"<w:p><w:pPr>{ppr}</w:pPr>{body}{extra}</w:p>"


def borders_xml(tag, spec):
    """spec: dict side -> (sz eighths, color) or None for nil."""
    out = ""
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        if side not in spec:
            continue
        v = spec[side]
        if v is None:
            out += f'<w:{side} w:val="nil"/>'
        else:
            out += f'<w:{side} w:val="single" w:sz="{v[0]}" w:space="0" w:color="{v[1]}"/>'
    return f"<w:{tag}>{out}</w:{tag}>" if out else ""


def table(grid, rows, width=("auto", 0), jc=None, layout=None, tbl_borders=None, style=None,
          cell_mar=None, ind=None, row_h=None):
    """rows: list of rows; each a list of cells. A cell is a str or a dict with
    keys text, content (raw xml), borders, shade, vmerge, span, w, valign."""
    pr = f'<w:tblStyle w:val="{style}"/>' if style else ""
    pr += f'<w:tblW w:w="{width[1]}" w:type="{width[0]}"/>'
    pr += f'<w:jc w:val="{jc}"/>' if jc else ""
    pr += f'<w:tblInd w:w="{ind}" w:type="dxa"/>' if ind is not None else ""
    if tbl_borders is not None:
        pr += borders_xml("tblBorders", tbl_borders)
    pr += f'<w:tblLayout w:type="{layout}"/>' if layout else ""
    if cell_mar:
        pr += "<w:tblCellMar>" + "".join(
            f'<w:{k} w:w="{v}" w:type="dxa"/>' for k, v in cell_mar.items()) + "</w:tblCellMar>"
    out = f"<w:tbl><w:tblPr>{pr}</w:tblPr><w:tblGrid>" + "".join(
        f'<w:gridCol w:w="{c}"/>' for c in grid) + "</w:tblGrid>"
    for row in rows:
        out += "<w:tr>" + (f'<w:trPr><w:trHeight w:val="{row_h}"/></w:trPr>' if row_h else "")
        col = 0
        for cell in row:
            if isinstance(cell, str):
                cell = {"text": cell}
            span = cell.get("span", 1)
            w = cell.get("w", sum(grid[col:col + span]))
            tc = f'<w:tcW w:w="{w}" w:type="dxa"/>'
            tc += f'<w:gridSpan w:val="{span}"/>' if span > 1 else ""
            if "vmerge" in cell:
                tc += '<w:vMerge w:val="restart"/>' if cell["vmerge"] == "restart" else "<w:vMerge/>"
            if cell.get("borders"):
                tc += borders_xml("tcBorders", cell["borders"])
            if cell.get("shade"):
                tc += f'<w:shd w:val="clear" w:color="auto" w:fill="{cell["shade"]}"/>'
            if cell.get("valign"):
                tc += f'<w:vAlign w:val="{cell["valign"]}"/>'
            content = cell.get("content") or para(cell.get("text", ""))
            out += f"<w:tc><w:tcPr>{tc}</w:tcPr>{content}</w:tc>"
            col += span
        out += "</w:tr>"
    return out + "</w:tbl>"


def shape(story, w, h, anchor="t", autofit="noAutofit", ins=(91440, 45720, 91440, 45720),
          fill="EEEEEE", line="808080", line_w=6350, txbx_id=None, linked=None, style="", ln_xml=None):
    """A wps shape (sizes in EMU). linked=(id, seq) for a follow-on box."""
    fill_xml = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else "<a:noFill/>"
    if ln_xml is None:
        ln_xml = (f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line}"/></a:solidFill></a:ln>' if line
                  else "<a:ln><a:noFill/></a:ln>")
    if linked:
        txbx = f'<wps:linkedTxbx id="{linked[0]}" seq="{linked[1]}"/>'
    elif story is not None:
        idattr = f' id="{txbx_id}"' if txbx_id is not None else ""
        txbx = f"<wps:txbx{idattr}><w:txbxContent>{story}</w:txbxContent></wps:txbx>"
    else:
        txbx = ""
    l, t, r, b = ins
    return (f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{fill_xml}{ln_xml}</wps:spPr>{style}{txbx}'
            f'<wps:bodyPr wrap="square" lIns="{l}" tIns="{t}" rIns="{r}" bIns="{b}" anchor="{anchor}">'
            f"<a:{autofit}/></wps:bodyPr></wps:wsp>")


def anchored(x, y, w, h, graphic, rel=("page", "page"), layout_in_cell=1, uri=WPS_URI, requires="wps"):
    n = next(_ids)
    return (f'<w:r><mc:AlternateContent><mc:Choice Requires="{requires}"><w:drawing>'
            f'<wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="{n}" '
            f'behindDoc="0" locked="0" layoutInCell="{layout_in_cell}" allowOverlap="1"><wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="{rel[0]}"><wp:posOffset>{int(x * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="{rel[1]}"><wp:posOffset>{int(y * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(w * EMU)}" cy="{int(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
            f'<wp:docPr id="{n}" name="Shape {n}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="{uri}">{graphic}</a:graphicData></a:graphic></wp:anchor>'
            f'</w:drawing></mc:Choice></mc:AlternateContent></w:r>')


def box(x, y, w, h, story, rel=("page", "page"), layout_in_cell=1, **kw):
    return anchored(x, y, w, h, shape(story, int(w * EMU), int(h * EMU), **kw), rel, layout_in_cell)


def _pic(rid, w, h):
    n = next(_ids)
    return (f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{n}" name="Picture {n}"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{int(w * EMU)}" cy="{int(h * EMU)}"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>')


PIC_URI = "http://schemas.openxmlformats.org/drawingml/2006/picture"


def inline_pic(rid, w, h):
    n = next(_ids)
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent cx="{int(w * EMU)}" cy="{int(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{n}" name="Picture {n}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="{PIC_URI}">{_pic(rid, w, h)}</a:graphicData></a:graphic>'
            f'</wp:inline></w:drawing></w:r>')


def floating_pic(rid, x, y, w, h, rel_h="column", rel_v="paragraph", wrap="<wp:wrapNone/>"):
    n = next(_ids)
    return (f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" '
            f'relativeHeight="{n}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            f'<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="{rel_h}"><wp:posOffset>{int(x * EMU)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="{rel_v}"><wp:posOffset>{int(y * EMU)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{int(w * EMU)}" cy="{int(h * EMU)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>{wrap}'
            f'<wp:docPr id="{n}" name="Picture {n}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="{PIC_URI}">{_pic(rid, w, h)}</a:graphicData></a:graphic>'
            f'</wp:anchor></w:drawing></w:r>')


THEME_TMPL = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Probe"><a:themeElements>
<a:clrScheme name="Probe"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F2A44"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="3366AA"/></a:accent1><a:accent2><a:srgbClr val="CC5522"/></a:accent2>
<a:accent3><a:srgbClr val="888888"/></a:accent3><a:accent4><a:srgbClr val="DDAA00"/></a:accent4>
<a:accent5><a:srgbClr val="4488CC"/></a:accent5><a:accent6><a:srgbClr val="66AA44"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>
<a:fontScheme name="Probe"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Probe"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst>LINES</a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle>
<a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>"""


def theme(line_widths_emu):
    lines = "".join(
        f'<a:ln w="{w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
        f'<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' for w in line_widths_emu)
    return THEME_TMPL.replace("LINES", lines)


STYLES = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W}"><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/>'
    '<w:szCs w:val="20"/><w:lang w:val="en-US"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>'
    '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr>'
    '<w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" '
    'w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>'
    '</w:tblPr></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/>'
    '<w:basedOn w:val="TableNormal"/><w:tblPr><w:tblBorders>'
    + "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
              for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    + "</w:tblBorders></w:tblPr></w:style></w:styles>")


def write(path, body, compat=15, header=None, theme_xml=None, images=None,
          page=(12240, 15840), margins=(1440, 1440, 1440, 1440)):
    """images: dict rId -> png bytes (shared by document and header)."""
    images = images or {}
    hdr_ref = '<w:headerReference w:type="default" r:id="rIdH"/>' if header is not None else ""
    top, right, bottom, left = margins
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}'
                f'<w:sectPr>{hdr_ref}<w:pgSz w:w="{page[0]}" w:h="{page[1]}"/><w:pgMar w:top="{top}" '
                f'w:right="{right}" w:bottom="{bottom}" w:left="{left}" w:header="720" w:footer="720" w:gutter="0"/>'
                "</w:sectPr></w:body></w:document>")
    settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="{W}">'
                '<w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" '
                f'w:uri="http://schemas.microsoft.com/office/word" w:val="{compat}"/></w:compat></w:settings>')
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ct = "application/vnd.openxmlformats-officedocument"
    img_rels = "".join(f'<Relationship Id="{rid}" Type="{rel}/image" Target="media/{rid}.png"/>'
                       for rid in images)
    doc_rels = (f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>'
                f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/>' + img_rels)
    overrides = (f'<Override PartName="/word/document.xml" ContentType="{ct}.wordprocessingml.document.main+xml"/>'
                 f'<Override PartName="/word/styles.xml" ContentType="{ct}.wordprocessingml.styles+xml"/>'
                 f'<Override PartName="/word/settings.xml" ContentType="{ct}.wordprocessingml.settings+xml"/>')
    parts = {}
    if header is not None:
        doc_rels += f'<Relationship Id="rIdH" Type="{rel}/header" Target="header1.xml"/>'
        overrides += f'<Override PartName="/word/header1.xml" ContentType="{ct}.wordprocessingml.header+xml"/>'
        parts["word/header1.xml"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                                     f'<w:hdr {NS}>{header}</w:hdr>')
        parts["word/_rels/header1.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{img_rels}'
            '</Relationships>')
    if theme_xml is not None:
        doc_rels += f'<Relationship Id="rIdT" Type="{rel}/theme" Target="theme/theme1.xml"/>'
        overrides += f'<Override PartName="/word/theme/theme1.xml" ContentType="{ct}.theme+xml"/>'
        parts["word/theme/theme1.xml"] = theme_xml
    for rid, data in images.items():
        parts[f"word/media/{rid}.png"] = data
    parts.update({
        "[Content_Types].xml": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Default Extension="png" ContentType="image/png"/>' + overrides + "</Types>"),
        "_rels/.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>'),
        "word/_rels/document.xml.rels": (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{doc_rels}'
            '</Relationships>'),
        "word/document.xml": document,
        "word/styles.xml": STYLES,
        "word/settings.xml": settings,
    })
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in ["[Content_Types].xml"] + [n for n in parts if n != "[Content_Types].xml"]:
            z.writestr(name, parts[name])
    print(f"wrote {path}")


OUT = pathlib.Path(__file__).parent / "input.docx"


def style(idx):
    return (f'<wps:style><a:lnRef idx="{idx}"><a:schemeClr val="accent1"/></a:lnRef>'
            '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
            '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
            '<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></wps:style>')


def story(label):
    return f'<w:p><w:pPr><w:jc w:val="right"/></w:pPr>{run(label)}</w:p>'


OUT = pathlib.Path(__file__).parent / "input.docx"


def style(idx):
    return (f'<wps:style><a:lnRef idx="{idx}"><a:schemeClr val="accent1"/></a:lnRef>'
            '<a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>'
            '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>'
            '<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></wps:style>')


def solid(w=None):
    wa = f' w="{w}"' if w is not None else ""
    return f'<a:ln{wa}><a:solidFill><a:srgbClr val="CC0000"/></a:solidFill></a:ln>'


def story(label, jc):
    return f'<w:p><w:pPr><w:jc w:val="{jc}"/></w:pPr>{run(label)}</w:p>'


# Keep the retained boxes at their original page positions.
CASES = [
    (140, "no w lnRef 0", solid(), style(0)),
    (350, "no w lnRef 3", solid(), style(3)),
    (420, "w 0 lnRef 3", solid(0), style(3)),
    (490, "w 2 lnRef 3", solid(25400), style(3)),
]
M = ("margin", "margin")
boxes = []
for y, label, ln, st in CASES:
    boxes.append(box(0, y, 150, 50, story(label, "left"), rel=M, fill="FFFFFF", ln_xml=ln, style=st))
write(OUT, para("", extra="".join(boxes)), theme_xml=theme([25400, 76200, 152400]))
