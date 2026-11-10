"""Generate vml-frame-fill-textbox-group: VML picture fills under text and in groups

A `<v:fill type="frame" r:id="..."/>` stretches a picture over its shape.
vml-frame-fill-picture covers bare shapes; this covers the other two places.

Tests:
- A floating 200x100pt VML text box filled with the logo, its text on top
- A v:group (coordsize 1000x400 over 300x120pt) with a frame-filled rect, a
  rect with v:imagedata beside it
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


def logo(x, y):
    """Two coloured halves and a dark diagonal: shows the stretch."""
    if abs(x * 50 - y * 100) < 300:
        return (25, 25, 25)
    return (200, 60, 90) if x < 50 else (60, 170, 90)


def para(content="", text=""):
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f"<w:p>{content}{run}</w:p>"


FILL = '<v:fill r:id="rIdLogo" o:title="" recolor="t" type="frame"/>'
textbox = (
    '<w:r><w:pict><v:rect style="position:absolute;margin-left:0;margin-top:30pt;width:200pt;height:100pt;'
    'z-index:1;mso-position-horizontal-relative:text;mso-position-vertical-relative:text" '
    f'strokecolor="#1e3c6e" strokeweight="1pt">{FILL}<v:textbox>'
    '<w:txbxContent><w:p><w:r><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="28"/></w:rPr>'
    "<w:t>Text over a picture fill</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect></w:pict></w:r>"
)
group = (
    '<w:r><w:pict><v:group style="position:absolute;margin-left:0;margin-top:30pt;width:300pt;height:120pt;'
    'z-index:2;mso-position-horizontal-relative:text;mso-position-vertical-relative:text" '
    'coordorigin="0,0" coordsize="1000,400">'
    f'<v:rect style="position:absolute;left:0;top:0;width:600;height:300" stroked="f">{FILL}</v:rect>'
    '<v:rect style="position:absolute;left:650;top:0;width:300;height:300" stroked="f">'
    '<v:imagedata r:id="rIdLogo" o:title=""/></v:rect>'
    "</v:group></w:pict></w:r>"
)
body = [
    para(text="A VML text box filled with a picture, text on top:"),
    para(textbox),
    *[para() for _ in range(8)],
    para(text="A VML group: a frame-filled rectangle, an imagedata rectangle:"),
    para(group),
]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1418" w:right="1134" w:bottom="1134" w:left="1418" w:header="709" w:footer="567" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
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
        ("rId1", "styles", "styles.xml"), ("rIdLogo", "image", "media/logo.png")))
    z.writestr("word/styles.xml", styles)
    z.writestr("word/media/logo.png", png(100, 50, logo))
print(f"wrote {out}")
