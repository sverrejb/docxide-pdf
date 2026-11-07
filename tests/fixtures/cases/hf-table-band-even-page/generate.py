"""Generate hf-table-band-even-page: header table bands follow the shown page number

evenAndOddHeaders, page numbering from 2, a full-width floating table placed
against the page only in the even header (250 to 300pt down), Arial 10pt on
exact 12pt lines. The first sheet shows page 2 and its even header: its lines
skip the table's band; the second sheet (page 3, default header) runs through.

Then a two-column section on a new sheet, numbering on (page 6), whose headers
have no table, with an empty even footer and a default footer holding a
full-width table 605 to 735pt down, crossing the bottom margin; a continuous
section after it balances its columns. Page 6 holds 54 lines a column, odd
page 7 only the 44 above the table, and page 8 the last 8, balanced 4 and 4.
"""

import pathlib
import zipfile

NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
)
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
OUT = pathlib.Path(__file__).parent


def para(text=""):
    run = f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>' if text else ""
    return f'<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>{run}</w:p>'


def page_field():
    return (
        '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>'
        '<w:r><w:t xml:space="preserve">Page </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        '<w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
    )


def table(x, y, w, h, text="T", dist=""):
    borders = "".join(
        f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>' for s in ("top", "left", "bottom", "right")
    )
    return (
        f'<w:tbl><w:tblPr><w:tblpPr {dist} w:vertAnchor="page" w:horzAnchor="page" w:tblpX="{x}" w:tblpY="{y}"/>'
        f'<w:tblW w:w="{w}" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>{borders}</w:tblBorders></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{w}"/></w:tblGrid>'
        f'<w:tr><w:trPr><w:trHeight w:val="{h}" w:hRule="exact"/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>{para(text)}</w:tc></w:tr></w:tbl>'
    )


def write(name, sections, even_odd=False):
    """`sections`: (body paragraphs, sectPr options) each; every section but the
    last ends in its last paragraph."""
    parts = {}
    rels = [("rIdStyles", "styles", "styles.xml"), ("rIdSettings", "settings", "settings.xml")]

    def sect_pr(n, header=None, footer=None, even_header=None, even_footer=None, cols="", pgnum="", kind="nextPage"):
        refs = ""
        for part_kind, typ, content in (
            ("header", "default", header),
            ("footer", "default", footer),
            ("header", "even", even_header),
            ("footer", "even", even_footer),
        ):
            if content is None:
                continue
            part = f"{part_kind}_{typ}_{n}"
            tag = "hdr" if part_kind == "header" else "ftr"
            parts[f"word/{part}.xml"] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:{tag} {NS}>{content}</w:{tag}>', part_kind)
            rels.append((f"rId{part}", part_kind, f"{part}.xml"))
            refs += f'<w:{part_kind}Reference w:type="{typ}" r:id="rId{part}"/>'
        return (
            f'<w:sectPr>{refs}<w:type w:val="{kind}"/><w:pgSz w:w="12240" w:h="15840"/>'
            '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>'
            f"{pgnum}{cols}</w:sectPr>"
        )

    body = ""
    for n, (paras, opts) in enumerate(sections):
        sect = sect_pr(n, **opts)
        if n + 1 < len(sections):
            last = paras[-1].replace("</w:pPr>", f"{sect}</w:pPr>", 1)
            body += "".join(paras[:-1]) + last
        else:
            body += "".join(paras) + sect
    document = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{body}</w:body></w:document>'
    styles = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        "</w:docDefaults></w:styles>"
    )
    eo = "<w:evenAndOddHeaders/>" if even_odd else ""
    settings = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}>{eo}<w:defaultTabStop w:val="708"/>'
        '<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
        "</w:settings>"
    )
    ov = lambda part, kind: f'<Override PartName="/{part}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>'
    ct = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        + ov("word/document.xml", "document.main")
        + ov("word/styles.xml", "styles")
        + ov("word/settings.xml", "settings")
        + "".join(ov(p, k) for p, (_, k) in parts.items())
        + "</Types>"
    )
    rl = lambda items: (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(f'<Relationship Id="{i}" Type="{REL}/{k}" Target="{t}"/>' for i, k, t in items)
        + "</Relationships>"
    )
    with zipfile.ZipFile(OUT / name, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rl([("rId1", "officeDocument", "word/document.xml")]))
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", rl(rels))
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
        for p, (xml, _) in parts.items():
            z.writestr(p, xml)


write(
    "input.docx",
    [
        (
            [para(f"Line {i}") for i in range(1, 201)],
            dict(
                header=page_field(),
                even_header=table(1440, 5000, 9360, 1000) + page_field(),
                pgnum='<w:pgNumType w:start="2"/>',
            ),
        ),
        (
            [para(f"Column line {i}") for i in range(1, 205)],
            dict(
                header=page_field(),
                even_header=page_field(),
                footer=table(1440, 12100, 9360, 2600) + para(),
                even_footer=para(),
                cols='<w:cols w:num="2" w:space="720"/>',
            ),
        ),
        ([para()], dict(kind="continuous")),
    ],
    even_odd=True,
)
print(f"wrote {OUT / 'input.docx'}")
