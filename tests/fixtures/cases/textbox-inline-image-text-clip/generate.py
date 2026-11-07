"""Generate textbox-inline-image-text-clip: text in a box's insets beside inline pictures

A standalone fixed-size text box clips an inline picture to its text area,
inside the left and right insets. Its text is clipped only to the box's
edges: a hanging first line, a list label or a negative indent prints in the
inset, also when the paragraph or the box holds a picture (Word for Mac
examples, online export).

Columns: text and a picture in one paragraph / text and a picture
in separate paragraphs. Rows: first line hanging 18pt (past the 7.2pt inset,
cut at the box edge), dash list label hanging 7pt, left indent -10pt,
right indent -10pt with right alignment. The pictures are 120pt wide, wider
than the 92.2pt text area.
"""

import pathlib
import struct
import zipfile
import zlib

EMU = 12700
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"')
WPS_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
RPR = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="20"/></w:rPr>'
BOX_W, BOX_H = 106.6, 141.5
COLS, ROWS = (240, 420), (70, 260, 450, 640)
PIC = (120, 50)


def placeholder_png(n=60):
    """A framed checkerboard, a different colour per quadrant, so a clip shows."""
    colours = ((200, 0, 0), (0, 0, 200), (0, 140, 0), (0, 0, 0))
    rows = []
    for y in range(n):
        row = bytearray([0])
        for x in range(n):
            if min(x, y, n - 1 - x, n - 1 - y) < 3:
                row += bytes((0, 0, 0))
            elif (x // 4 + y // 4) % 2:
                row += bytes(colours[(y >= n // 2) * 2 + (x >= n // 2)])
            else:
                row += bytes((255, 255, 255))
        rows.append(bytes(row))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", n, n, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"".join(rows))) + chunk(b"IEND", b""))


ids = iter(range(100, 10000))


def emu(v):
    return int(round(v * EMU))


def inline_picture(w, h):
    i = next(ids)
    ext = f'cx="{emu(w)}" cy="{emu(h)}"'
    return ('<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
            f'<wp:extent {ext}/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:docPr id="{i}" name="Picture {i}"/>'
            '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="code.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext {ext}/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
            '</a:graphicData></a:graphic></wp:inline></w:drawing>')


INDENTS = {
    "hanging": ('<w:ind w:left="0" w:hanging="360"/>', "left"),
    "label": ('<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>'
              '<w:ind w:left="0" w:hanging="140"/>', "left"),
    "left": ('<w:ind w:left="-200"/>', "left"),
    "right": ('<w:ind w:right="-200"/>', "right"),
}


def paragraph(kind, text, picture=False):
    ind, jc = INDENTS[kind]
    ppr = f'<w:pPr>{ind}<w:jc w:val="{jc}"/>{RPR}</w:pPr>'
    pic = f"<w:r>{RPR}<w:br/>{inline_picture(*PIC)}</w:r>" if picture else ""
    return f'<w:p>{ppr}<w:r>{RPR}<w:t xml:space="preserve">{text}</w:t></w:r>{pic}</w:p>'


def picture_paragraph():
    return f'<w:p><w:pPr>{RPR}</w:pPr><w:r>{RPR}{inline_picture(*PIC)}</w:r></w:p>'


def story(kind, layout):
    text = "MMMMMMMMMMMMMMMMMMMMMMMM" if kind in ("left", "right") else "Hanging line of text that wraps"
    if layout == "same":
        return paragraph(kind, text, picture=True)
    return paragraph(kind, text) + picture_paragraph()


def shape(w, h, content):
    return ('<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr>'
            f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="EEEEEE"/></a:solidFill>'
            '<a:ln w="3175"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></wps:spPr>'
            f"<wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx>"
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" '
            'rIns="91440" bIns="45720" anchor="t" anchorCtr="0" upright="1"><a:noAutofit/>'
            "</wps:bodyPr></wps:wsp>")


def text_box(x, y, content):
    i = next(ids)
    return ('<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" '
            f'relativeHeight="{i}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
            '<wp:simplePos x="0" y="0"/>'
            f'<wp:positionH relativeFrom="page"><wp:posOffset>{emu(x)}</wp:posOffset></wp:positionH>'
            f'<wp:positionV relativeFrom="page"><wp:posOffset>{emu(y)}</wp:posOffset></wp:positionV>'
            f'<wp:extent cx="{emu(BOX_W)}" cy="{emu(BOX_H)}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>'
            f'<wp:wrapNone/><wp:docPr id="{i}" name="Box {i}"/><wp:cNvGraphicFramePr/>'
            f'<a:graphic><a:graphicData uri="{WPS_URI}">{shape(BOX_W, BOX_H, content)}</a:graphicData></a:graphic>'
            "</wp:anchor></w:drawing></w:r>")


boxes = "".join(text_box(COLS[c], ROWS[r], story(kind, layout))
                for r, kind in enumerate(INDENTS) for c, layout in enumerate(("same", "apart")))
title = ("Rows: hanging first line, list label, left indent -10pt, right indent -10pt; "
         "columns: text and picture, text then picture paragraph")
body = f'<w:p><w:r>{RPR}<w:t xml:space="preserve">{title}</w:t></w:r>{boxes}</w:p>'

document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="567" w:right="1134" w:bottom="567" w:left="1134" w:header="0" w:footer="0" w:gutter="0"/>'
            "</w:sectPr></w:body></w:document>")

styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
          '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          "</w:docDefaults></w:styles>")

numbering = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering {NS}>'
             '<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="singleLevel"/>'
             '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="–"/>'
             '<w:lvlJc w:val="left"/><w:pPr><w:ind w:left="0" w:hanging="140"/></w:pPr>'
             f'{RPR}</w:lvl></w:abstractNum>'
             '<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>')

content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
    "</Types>")

rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
package_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="word/document.xml"/></Relationships>')
document_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{rel}/styles" Target="styles.xml"/>'
    f'<Relationship Id="rId2" Type="{rel}/numbering" Target="numbering.xml"/>'
    f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/code.png"/></Relationships>')

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/numbering.xml", numbering)
    z.writestr("word/media/code.png", placeholder_png())
print(f"wrote {out}")
