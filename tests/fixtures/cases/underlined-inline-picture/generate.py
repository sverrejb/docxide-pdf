"""Generate underlined-inline-picture: an underlined picture alone on its line

Word underlines an inline picture whose run is underlined: the underline runs
under the picture at the run font's underline position, and the line keeps
that font's descent below the picture, as a line of text would. Without the
underline the next line starts right below the picture.

Each case: a label line, then a paragraph holding only a 100 x 40pt picture
whose run is in the given font and size, then the next label (Arial 10pt):
1. Tahoma 16pt underlined (+3.3pt).
2. Arial 24pt underlined (+5.2pt).
3. Arial 24pt, underline "words" (+5.2pt).
"""
import pathlib, zipfile, struct, zlib
out = pathlib.Path(__file__).parent
W = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
     'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
     'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
     'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
     'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def png():
    """A grey square, drawn here."""
    w = h = 40
    rows = b"".join(b"\x00" + b"".join(bytes((90, 90, 90)) for _ in range(w)) for _ in range(h))
    c = lambda t, d: struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)
    return b"\x89PNG\r\n\x1a\n" + c(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0)) + c(b"IDAT", zlib.compress(rows)) + c(b"IEND", b"")


ids = iter(range(1, 999))


def pic_run(font, size, u, h_pt=40):
    n = next(ids)
    cx, cy = int(100 * 12700), int(h_pt * 12700)
    upr = f'<w:u w:val="{u}"/>' if u else ""
    return (f'<w:r><w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/><w:sz w:val="{size*2}"/>{upr}</w:rPr><w:drawing>'
            f'<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{cx}" cy="{cy}"/><wp:docPr id="{n}" name="P{n}"/>'
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
            f'<pic:nvPicPr><pic:cNvPr id="{n}" name="p.png"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
            '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>')


def para(runs, mark_font="Arial", mark_size=10, mark_u=None):
    upr = f'<w:u w:val="{mark_u}"/>' if mark_u else ""
    return (f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:rPr><w:rFonts w:ascii="{mark_font}" w:hAnsi="{mark_font}"/>'
            f'<w:sz w:val="{mark_size*2}"/>{upr}</w:rPr></w:pPr>{runs}</w:p>')


def text(t):
    return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'


cases = [("Tahoma", 16, "single"), ("Arial", 24, "single"), ("Arial", 24, "words")]
body = ""
for n, (font, size, u) in enumerate(cases, 1):
    body += para(text(f"Case {n}: {font} {size}pt, underline {u}")) + para(pic_run(font, size, u), font, size, u)
body += para(text("End"))
doc = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>{body}'
       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1418" w:bottom="1134" w:left="1418" w:header="567" w:footer="567" w:gutter="0"/></w:sectPr></w:body></w:document>')
styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>')
settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:defaultTabStop w:val="708"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>')
ov = lambda p, t: f'<Override PartName="/word/{p}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{t}+xml"/>'
ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
      '<Default Extension="png" ContentType="image/png"/>' + ov("document.xml", "document.main") + ov("styles.xml", "styles") + ov("settings.xml", "settings") + "</Types>")
rels = lambda items: ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                      + "".join(f'<Relationship Id="{i}" Type="{REL}/{t}" Target="{g}"/>' for i, t, g in items) + "</Relationships>")
with zipfile.ZipFile(out / "input.docx", "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", ct)
    z.writestr("_rels/.rels", rels([("rId1", "officeDocument", "word/document.xml")]))
    z.writestr("word/document.xml", doc)
    z.writestr("word/_rels/document.xml.rels", rels([("rId1", "styles", "styles.xml"), ("rId2", "settings", "settings.xml"), ("rIdImg", "image", "media/p.png")]))
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
    z.writestr("word/media/p.png", png())
print(f"wrote {out / 'input.docx'}")
