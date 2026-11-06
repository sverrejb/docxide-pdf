"""Generate tab-picture-wrap: an inline picture that does not fit after a tab

Letterhead paragraphs often start with a tab before a page-wide banner picture.
When the picture does not fit between the tab stop and the right margin, Word
starts a new line with it: the tab stays behind on a line of its own (sized by
the tab's font) and the picture sits at the start of the next line. A tab after
the picture that no longer fits takes a third line, again sized by the tab's
font, not by the larger font of the picture's run. That last line also holds
the paragraph mark, so a larger mark makes it taller.

Cases, each between "Before" and "After" lines (tab stop at 83pt, 468pt column):
1. 420pt picture: next line, at the margin
2. 480pt picture in a 20pt run: the trailing tab line is a 10.5pt line
3. 480pt picture, 14pt paragraph mark: the trailing tab line is a 14pt line
"""

import pathlib
import struct
import zipfile
import zlib

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'
)


def placeholder_png(w=200, h=20):
    """A grey checkerboard, drawn here."""
    rows = b""
    for y in range(h):
        rows += b"\x00" + b"".join(
            bytes((v, v, v)) for v in (60 if (x // 10 + y // 10) % 2 else 150 for x in range(w))
        )

    def chunk(tag, data):
        body = tag + data
        return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(rows)) + chunk(b"IEND", b""))


def picture(w_pt, h_pt, pid, size=21):
    cx, cy = int(w_pt * 12700), int(h_pt * 12700)
    return (
        f'<w:r><w:rPr><w:sz w:val="{size}"/></w:rPr><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{pid}" name="Banner {pid}"/>'
        '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{pid}" name="banner.png"/><pic:cNvPicPr/></pic:nvPicPr>'
        '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
        "</wp:inline></w:drawing></w:r>"
    )


TAB = '<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:tab/></w:r>'
SPACE = '<w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>'


def marker(text):
    return f'<w:p><w:pPr><w:spacing w:line="260" w:lineRule="exact"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>'


def banner(w_pt, pid, size=21, mark=None):
    mark_xml = f'<w:rPr><w:sz w:val="{mark}"/></w:rPr>' if mark else ""
    return (
        '<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1660"/><w:tab w:val="center" w:pos="4025"/></w:tabs>'
        f"{mark_xml}</w:pPr>{TAB}{picture(w_pt, 30, pid, size)}{TAB}{SPACE}</w:p>"
    )


body = []
cases = ((420, 21, None), (480, 40, None), (480, 21, 28))
for n, (w, size, mark) in enumerate(cases, 1):
    body += [marker(f"Before {n}"), banner(w, n, size, mark), marker(f"After {n}")]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)

settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:defaultTabStop w:val="708"/>'
    '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
    "</w:settings>"
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
    f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/banner.png"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/media/banner.png", placeholder_png())
print(f"wrote {out}")
