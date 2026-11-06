"""Generate truncated-png: PNG pictures whose image data ends early

Some generators embed PNGs cut off in the middle of the IDAT chunk (the
chunk length claims more bytes than the file has; no CRC, no IEND). Word
still draws the rows it can decode.

Tests:
- 1-bit grayscale, 8-bit grayscale, 8-bit RGB and 8-bit RGBA pictures
- each cut off after half of its compressed image data
"""

import pathlib
import struct
import zipfile
import zlib

W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
     'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
     'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
     'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SIZE = 120


def chunk(kind, data):
    return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", zlib.crc32(kind + data))


def png(color_type, bit_depth, pixel):
    rows = b""
    for y in range(SIZE):
        if bit_depth == 1:
            bits = [pixel(x, y) for x in range(SIZE)]
            row = bytes(sum(b << (7 - i) for i, b in enumerate(bits[k:k + 8])) for k in range(0, SIZE, 8))
        else:
            row = b"".join(bytes(pixel(x, y)) for x in range(SIZE))
        rows += b"\0" + row
    ihdr = struct.pack(">IIBBBBB", SIZE, SIZE, bit_depth, color_type, 0, 0, 0)
    return b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", ihdr) + chunk(b"IDAT", zlib.compress(rows, 9)) + chunk(b"IEND", b"")


def truncated(data):
    idat = data.index(b"IDAT") - 4
    length = struct.unpack(">I", data[idat:idat + 4])[0]
    return data[:idat + 8 + length // 2]


def checker(x, y):
    """Pseudo-random 4-px modules, like a QR code, so the data spreads over the stream."""
    n = ((x // 4) * 7919 + (y // 4) * 104729) % 65521
    return (n * n // 7) % 2


pictures = {
    "gray1": png(0, 1, checker),
    "gray8": png(0, 8, lambda x, y: [60 + 150 * checker(x, y)]),
    "rgb": png(2, 8, lambda x, y: [200, 60, 60] if checker(x, y) else [60, 60, 200]),
    "rgba": png(6, 8, lambda x, y: [40, 160, 40, 255] if checker(x, y) else [200, 200, 40, 128]),
}
images = [(f"{name} truncated", truncated(data)) for name, data in pictures.items()]


def picture(i):
    emu = 1080000
    return (f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{emu}" cy="{emu}"/>'
            f'<wp:docPr id="{i + 1}" name="Picture {i + 1}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{i + 1}" name="image{i}.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            f'<pic:blipFill><a:blip r:embed="rIdImg{i}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu}" cy="{emu}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>")


body = "".join(
    f'<w:p><w:r><w:t xml:space="preserve">{label}</w:t></w:r></w:p><w:p>{picture(i)}</w:p>'
    for i, (label, _) in enumerate(images))
document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1134" w:right="1417" w:bottom="1134" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>'
            "</w:sectPr></w:body></w:document>")
styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Default Extension="png" ContentType="image/png"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    "</Types>"
)
package_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>'
)
document_rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    f'<Relationship Id="rId1" Type="{REL}/styles" Target="styles.xml"/>'
    + "".join(f'<Relationship Id="rIdImg{i}" Type="{REL}/image" Target="media/image{i}.png"/>' for i in range(len(images)))
    + "</Relationships>"
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    for i, (_, data) in enumerate(images):
        z.writestr(f"word/media/image{i}.png", data)
print(f"wrote {out}")
