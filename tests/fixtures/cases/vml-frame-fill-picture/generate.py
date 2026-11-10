"""Generate vml-frame-fill-picture: VML shapes filled with a stretched picture

VML rectangles whose fill is a picture, `<v:fill type="frame" r:id="..."/>`,
without `v:imagedata` and without a DrawingML alternative. Word stretches the
picture over the shape.

Tests:
- Page 1: a 595x842pt rectangle at the page origin, z-index -1, white stroke.
  It covers the header, which Word draws below the body's shapes.
- Page 1: inline 120x60pt and 60x60pt rectangles, stroke off, filled with the logo
- Page 2: the header's 141.75x70.85pt logo rectangle, white 2pt stroke
"""

import pathlib
import struct
import zipfile
import zlib

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = (
    f'{W} xmlns:r="{R_NS}" xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:o="urn:schemas-microsoft-com:office:office"'
)


def png(width, height, pixel):
    rows = []
    for y in range(height):
        row = bytearray([0])
        for x in range(width):
            row += bytes(pixel(x, y))
        rows.append(bytes(row))

    def chunk(tag, data):
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data))

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr)
            + chunk(b"IDAT", zlib.compress(b"".join(rows), 9)) + chunk(b"IEND", b""))


def page_picture(x, y):
    """A4 at 100 dpi: a dark band at the top, a thin rule at the bottom, a side stripe."""
    if y < 60:
        return (30, 60, 110)
    if 1120 <= y < 1126 or x < 18:
        return (30, 60, 110)
    return (255, 255, 255)


def logo(x, y):
    """Two coloured halves and a dark diagonal: shows the stretch."""
    if abs(x * 50 - y * 100) < 300:
        return (25, 25, 25)
    return (200, 60, 90) if x < 50 else (60, 170, 90)


def rect(style, fill, rid, extra=""):
    return (
        f'<w:r><w:pict><v:rect style="{style}" {extra}>'
        f'<v:fill r:id="{rid}" o:title="" recolor="t" type="frame"{fill}/></v:rect></w:pict></w:r>'
    )


def para(content="", text=""):
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{content}{run}</w:p>"


background = rect(
    "position:absolute;margin-left:0;margin-top:0;width:595.3pt;height:841.9pt;z-index:-1;"
    "mso-position-horizontal-relative:page;mso-position-vertical-relative:page",
    "", "rIdBg", 'strokecolor="white"',
)
inline = "width:120pt;height:60pt"
body = [
    para(background, "The page background is a VML rectangle filled with a picture."),
    para(text="Word stretches the picture over the whole rectangle, behind this text."),
    para(),
    para(text="Inline rectangles filled with the logo, stroke off:"),
    para(rect(inline, "", "rIdLogo", 'stroked="f"') + '<w:r><w:t xml:space="preserve">  </w:t></w:r>'
         + rect("width:60pt;height:60pt", "", "rIdLogo", 'stroked="f"')),
    para(text="End of the first page."),
    para('<w:r><w:br w:type="page"/></w:r>', "The second page has no background: the header logo shows."),
]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:headerReference w:type="default" r:id="rIdHdr"/><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="2268" w:right="1134" w:bottom="1134" w:left="1418" w:header="709" w:footer="567" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)
header_logo = rect(
    "position:absolute;margin-left:356pt;margin-top:-5.4pt;width:141.75pt;height:70.85pt;z-index:1;"
    "mso-position-horizontal-relative:text;mso-position-vertical-relative:text",
    ' rotate="t"', "rIdLogo", 'strokecolor="white" strokeweight="2pt"',
)
header = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {NS}>'
    + para(header_logo, "Header text beside the logo rectangle") + "</w:hdr>"
)
styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)

ct = "application/vnd.openxmlformats-officedocument.wordprocessingml"
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    f'<Override PartName="/word/document.xml" ContentType="{ct}.document.main+xml"/>'
    f'<Override PartName="/word/styles.xml" ContentType="{ct}.styles+xml"/>'
    f'<Override PartName="/word/header1.xml" ContentType="{ct}.header+xml"/>'
    "</Types>"
)


def rels(*entries):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="{i}" Type="{R_NS}/{t}" Target="{target}"/>' for i, t, target in entries)
        + "</Relationships>"
    )


out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels(("rId1", "officeDocument", "word/document.xml")))
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", rels(
        ("rId1", "styles", "styles.xml"), ("rIdHdr", "header", "header1.xml"),
        ("rIdBg", "image", "media/background.png"), ("rIdLogo", "image", "media/logo.png")))
    z.writestr("word/header1.xml", header)
    z.writestr("word/_rels/header1.xml.rels", rels(("rIdLogo", "image", "media/logo.png")))
    z.writestr("word/styles.xml", styles)
    z.writestr("word/media/background.png", png(827, 1169, page_picture))
    z.writestr("word/media/logo.png", png(100, 50, logo))
print(f"wrote {out}")
