"""Generate textbox-wrap-beside: text wraps beside square/tight text boxes and shapes

Word wraps a paragraph beside a floating text box or shape (wrapSquare, wrapTight,
wrapThrough) the way it wraps beside a picture, wherever a strip is left on a side
its wrapText allows. The line starts at the box's wrap edge (extent + effectExtent +
distL/distR) or at the paragraph's left indent, whichever is further right; a
first-line indent adds to that start, a hanging indent never takes the first line
left of the edge. Under compatibility mode 14 and older a hanging indent's first
line counts its indent from the edge instead.

One case per page, Arial 10pt on an exact 13pt line, A4 with 2cm margins:
1. The document's first page in two 200pt columns: an anchor after a 24pt space after,
   40pt above the first column's foot with a 180pt box: box and text start at the
   second column's top
2. Text box at the left (0.85pt into the column, effectExtent r=0.8pt, distR 9pt);
   paragraphs indented 129pt, 0pt, 200pt, 36pt + 18pt first line, 150pt - 40pt
   hanging, each beside the box and again below it
3. The same with a picture
4. Text box at the right margin; paragraphs with a 129pt right indent and none
5. Centred text box, wrapText bothSides: text on both sides of it
6. Text box 200pt in, wrapText largest: text on the wider left side only
7. Centred text box, wrapText right: text on its right only
8. Text box anchored in a paragraph with text: the anchor's own lines wrap
9. Two boxes in one anchor, at the left and right margins: text between them
10. Page-anchored box anchored in the first paragraph, reached by later ones
11. Tight text box with a rectangular wrap polygon (no effectExtent in the edge)
12. Group shape wrapping square as one box
13. Box with a 50pt strip, wrapText largest: the anchor's lines take the strip
14. Box with a 17pt strip, wrapText largest: the anchor's lines start under it
15. "Internationalization" (bold "I") beside a 40pt strip: the whole word goes below
16. A wrapText=right box and a bothSides box: no text left of the first
17. An empty clearing-break paragraph after boxes 250pt and 80pt tall: text below both
18. An empty anchor 40pt above the page foot with a 180pt box: both go to the next page
19. A 17pt strip beside a wrapText=largest picture in a paragraph that runs past the
    page: its first lines start below the picture, the rest on the next page
20. Hebrew w:bidi paragraphs between a wrapText=right box and a bothSides one: each
    line fills the strip right of the second box first, then the one between them
21. An anchor opening with "Internationalization" beside the 40pt strip of its own
    wrapText=right box: the whole word starts below the box
22. "Hi" + no-break space + "Internationalization" across runs beside a 40pt strip:
    the joined words go below the box together
23. Two 200pt columns, a widow-controlled two-line paragraph beside an earlier box at the
    first column's foot: it moves to the second column's full measure
24. The same in one column: it moves to the next page's full measure
25. A bothSides box with a 17pt strip on its left: the text takes the right side only
26. Courier "A" + no-break space + space + 30pt "B" beside a 40pt strip: the spaces
    measure 10pt, the joined word fits the strip

The compat-14 variant (textbox-wrap-beside-compat14) has cases 2 and 3 only.

usage: python3 generate.py   (writes this fixture's and the compat-14 variant's input.docx)
"""

import pathlib
import struct
import zipfile
import zlib

EMU = 12700
NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="wp14"'
)
SECT = (
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr>"
)
WORDS = ("alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november "
         "oscar papa quebec romeo sierra tango uniform victor whiskey xray yankee zulu").split()
COL_W = (11906 - 2 * 1134) / 20


def grey_png(w=8, h=8, grey=150):
    raw = b"".join(b"\x00" + bytes([grey]) * w for _ in range(h))

    def chunk(t, d):
        return struct.pack(">I", len(d)) + t + d + struct.pack(">I", zlib.crc32(t + d) & 0xFFFFFFFF)

    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 0, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def emu(pt):
    return int(round(pt * EMU))


def wrap_xml(wrap, wrap_text):
    if wrap == "square":
        return f'<wp:wrapSquare wrapText="{wrap_text}"/>'
    if wrap == "tight":
        pts = [(0, 0), (0, 21600), (21600, 21600), (21600, 0), (0, 0)]
        poly = f'<wp:start x="{pts[0][0]}" y="{pts[0][1]}"/>' + "".join(
            f'<wp:lineTo x="{x}" y="{y}"/>' for x, y in pts[1:])
        return f'<wp:wrapTight wrapText="{wrap_text}"><wp:wrapPolygon edited="0">{poly}</wp:wrapPolygon></wp:wrapTight>'
    raise ValueError(wrap)


SHAPE_ID = [100]


def shape(kind="tb", w=128.4, h=100, hrel="column", hoff=0.85, halign=None, vrel="paragraph", voff=10,
          ee=(0.55, 0.85, 0.8, 0.6), wrap="square", wt="bothSides", text="Box"):
    SHAPE_ID[0] += 1
    sid = SHAPE_ID[0]
    hpos = f"<wp:align>{halign}</wp:align>" if halign else f"<wp:posOffset>{emu(hoff)}</wp:posOffset>"
    el, et, er, eb = ee
    if kind == "pic":
        graphic = (
            '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{sid}" name="Picture {sid}"/><pic:cNvPicPr/></pic:nvPicPr>'
            '<pic:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
            f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>')
    elif kind == "group":
        half = w / 2

        def child(i, x):
            return (f'<wps:wsp><wps:cNvPr id="{sid * 10 + i}" name="Part {i}"/><wps:cNvSpPr/><wps:spPr>'
                    f'<a:xfrm><a:off x="{emu(x)}" y="0"/><a:ext cx="{emu(half - 4)}" cy="{emu(h)}"/></a:xfrm>'
                    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="A0A0A0"/></a:solidFill>'
                    '<a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp>')

        graphic = (
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">'
            f'<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{emu(w)}" cy="{emu(h)}"/>'
            f'<a:chOff x="0" y="0"/><a:chExt cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm></wpg:grpSpPr>'
            + child(1, 0) + child(2, half + 4) + "</wpg:wgp></a:graphicData></a:graphic>")
    else:
        graphic = (
            '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">'
            f'<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/>'
            f'<a:ext cx="{emu(w)}" cy="{emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="A0A0A0"/></a:solidFill><a:ln><a:noFill/></a:ln></wps:spPr>'
            f'<wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>'
            '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t">'
            "<a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic>")
    return (
        f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="{emu(9)}" distR="{emu(9)}" simplePos="0" '
        f'relativeHeight="{sid}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">'
        f'<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="{hrel}">{hpos}</wp:positionH>'
        f'<wp:positionV relativeFrom="{vrel}"><wp:posOffset>{emu(voff)}</wp:posOffset></wp:positionV>'
        f'<wp:extent cx="{emu(w)}" cy="{emu(h)}"/>'
        f'<wp:effectExtent l="{emu(el)}" t="{emu(et)}" r="{emu(er)}" b="{emu(eb)}"/>'
        + wrap_xml(wrap, wt)
        + f'<wp:docPr id="{sid}" name="Shape {sid}"/><wp:cNvGraphicFramePr/>' + graphic
        + "</wp:anchor></w:drawing></w:r>")


def run(text):
    return f'<w:r><w:t xml:space="preserve">{text}</w:t></w:r>'


def para(runs, ind=None):
    ppr = '<w:spacing w:after="0" w:line="260" w:lineRule="exact"/>'
    if ind:
        ppr += "<w:ind " + " ".join(f'w:{k}="{int(v * 20)}"' for k, v in ind.items()) + "/>"
    return f"<w:p><w:pPr>{ppr}</w:pPr>{runs}</w:p>"


def words(case, n, count=18):
    # Preserve each retained case's word sequence when renumbering after cuts.
    seed = case + (case >= 10) + (case >= 19)
    return f"C{case}P{n} " + " ".join(WORDS[(seed * 7 + n * 5 + i) % len(WORDS)] for i in range(count))


INDENTS = [{"left": 129}, None, {"left": 200}, {"left": 36, "firstLine": 18}, {"left": 150, "hanging": 40}]


def case(n, title, shapes, indents=(None,) * 10, anchor_text=None, before=0, pre=(), follow=(), opts=None):
    opts = opts or {}
    out = [para(run(f"Case {n}: {title}"))]
    out += [para(run(words(n, 90 + i))) for i in range(before)]
    out += list(pre)
    out.append(para("".join(shapes) + (run(anchor_text) if anchor_text else "") + opts.get("anchor_runs", "")))
    out += list(follow)
    out += [para(run(words(n, i)), ind) for i, ind in enumerate(indents)]
    return out


def courier(text, bold=False, size=None):
    rpr = '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>' + ("<w:b/>" if bold else "")
    rpr += f'<w:sz w:val="{size * 2}"/>' if size else ""
    return f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>'


HEBREW = "שלום עולם זה טקסט בעברית לבדיקה של עטיפה ליד תיבות טקסט צפות בעמוד".split()


def hebrew_para(tag, n):
    text = tag + " " + " ".join(HEBREW[i % len(HEBREW)] for i in range(n))
    return para(f'<w:r><w:rPr><w:rtl/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>').replace(
        "<w:pPr>", "<w:pPr><w:bidi/>", 1)


def widow_case(cols):
    width = 100 if cols else 300
    text = "Two lines alpha bravo charlie" + ("" if cols else " delta echo foxtrot golf hotel")
    return ([shape(hoff=0, voff=0, w=width, h=60, ee=(0, 0, 0, 0))], (None,) * 3, None, 0,
            [para(run(f"filler {i}")) for i in range(50)],
            [para(run(f"short {i}")) for i in range(3)]
            + [para(run(text)).replace("<w:pPr>", "<w:pPr><w:widowControl/>", 1)], {"cols": cols})


def cases_compat15():
    tall = dict(h=150)
    flat = (0, 0, 0, 0)
    hi = "Hi " + " ".join(WORDS[i % len(WORDS)] for i in range(60))
    letters = " ".join(["a bb cc dd ee"] * 40)
    cases = [
        ("indents beside a text box", [shape(**tall)], INDENTS * 2),
        ("indents beside a picture", [shape(kind="pic", **tall)], INDENTS * 2),
        ("text box at the right", [shape(halign="right", hoff=None, **tall)],
         [{"right": 129}, None, {"right": 129}, None, None, None]),
        ("centred, bothSides", [shape(halign="center", hoff=None, w=100)], (None,) * 8),
        ("200pt in, largest", [shape(hoff=200, wt="largest")], (None,) * 8),
        ("centred, right", [shape(halign="center", hoff=None, w=100, wt="right")], (None,) * 8),
        ("anchor with text", [shape(voff=0, ee=(0, 0, 0, 0))], (None,) * 6, " ".join(["anchor text wraps too"] * 12)),
        ("boxes at both margins", [shape(hoff=0, ee=(0, 0, 0, 0)), shape(halign="right", hoff=None, w=100)],
         (None,) * 8),
        ("page-anchored box", [shape(vrel="page", voff=245, h=105, hrel="margin", hoff=0, ee=(0, 0, 0, 0))],
         (None,) * 16),
        ("tight text box", [shape(wrap="tight")], [{"left": 36}] * 8),
        ("group", [shape(kind="group", ee=(0, 0, 0, 0))], (None,) * 8),
        ("largest, 50pt strip", [shape(hoff=0, voff=0, w=COL_W - 59, h=150, ee=flat, wt="largest")], (None,) * 4,
         hi),
        ("largest, 17pt strip", [shape(hoff=0, voff=0, w=COL_W - 26, h=150, ee=flat, wt="largest")], (None,) * 4,
         hi),
        ("first word across runs", [shape(hoff=0, voff=0, w=COL_W - 49, h=150, ee=flat, wt="right")], (None,) * 2,
         None, 0, (), [para(courier("I", True) + courier("nternationalization " + hi))]),
        ("right-only box and bothSides box", [shape(hoff=200, voff=0, w=100, h=150, ee=flat, wt="right"),
                                              shape(hoff=400, voff=0, w=40, h=150, ee=flat)], (None,) * 6),
        ("clearing break below both boxes", [shape(hoff=0, voff=0, w=100, h=250, ee=flat),
                                             shape(hoff=380, voff=0, w=100, h=80, ee=flat)], (None,) * 2,
         None, 0, (), ['<w:p><w:r><w:br w:clear="all"/></w:r></w:p>']),
        ("box past the page foot", [shape(hoff=0, voff=0, w=300, h=180, ee=flat)], (None,) * 3,
         None, 0, [para(run(f"filler {i}")) for i in range(52)]),
        ("split paragraph below a picture", [shape(kind="pic", hoff=0, voff=0, w=COL_W - 17, h=150, ee=flat,
                                                   wt="largest")], (), "Top " + " ".join(
            WORDS[i % len(WORDS)] for i in range(1100))),
        ("right-to-left between two boxes", [shape(hoff=200, voff=0, w=100, h=150, ee=flat, wt="right"),
                                             shape(hoff=400, voff=0, w=40, h=150, ee=flat)], (), None, 0, (),
         [hebrew_para("1", 60), hebrew_para("2", 60)]),
        ("oversized opening word", [shape(hoff=0, voff=0, w=COL_W - 40, h=150, ee=flat, wt="right")], (None,),
         None, 0, (), (), {"anchor_runs": courier("Internationalization " + " ".join(WORDS[:40]))}),
        ("no-break space across runs", [shape(hoff=0, voff=0, w=COL_W - 40, h=150, ee=flat, wt="right")],
         (None,), None, 0, (), [para(courier("Hi\u00a0", True) + courier("Internationalization " + hi))]),
        ("widow, next column", *widow_case(2)),
        ("widow, next page", *widow_case(None)),
        ("bothSides, 17pt strip on the left", [shape(hoff=26, voff=0, w=200, h=100, ee=flat)], (), None, 0, (),
         [para(run(f"{i} {letters}")) for i in range(2)]),
        ("spaces before a no-break join in their own size",
         [shape(hoff=0, voff=0, w=COL_W - 49, h=100, ee=flat, wt="right")], (None,), None, 0, (),
         [para(courier("A\u00a0 ") + courier("B", size=30) + courier(" ab cd ef gh ij kl mn op qr st uv wx yz" * 12))]),
    ]
    first = ("24pt after, next column",
             [shape(hoff=0, voff=0, w=100, h=180, ee=flat)], (None,) * 3, "short anchor", 0,
             [para(run(f"filler {i}")) for i in range(51)]
             + [para(run("filler 51")).replace('w:after="0"', 'w:after="480"', 1)], (), {"cols": 2})
    return [first] + cases


def build(path, cases, compat):
    body = []
    for n, spec in enumerate(cases, 1):
        title, shapes, indents, *rest = spec
        paras = case(n, title, shapes, indents, *rest)
        cols = rest[-1].get("cols") if rest and isinstance(rest[-1], dict) else None
        sect = SECT.replace("</w:sectPr>", f'<w:cols w:num="{cols}" w:space="{int((COL_W - 400) * 20)}"/></w:sectPr>') \
            if cols else SECT
        if n < len(cases):
            paras[-1] = paras[-1].replace("</w:pPr>", sect + "</w:pPr>", 1)
        body += paras
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>'
                + "".join(body) + SECT + "</w:body></w:document>")
    styles = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
        '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
        '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
        '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
        "</w:docDefaults></w:styles>")
    settings = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:defaultTabStop w:val="708"/>'
        f'<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" '
        f'w:val="{compat}"/></w:compat></w:settings>')
    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
        '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
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
        f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/>'
        f'<Relationship Id="rIdImg" Type="{rel}/image" Target="media/grey.png"/></Relationships>')
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", package_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", document_rels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/settings.xml", settings)
        z.writestr("word/media/grey.png", grey_png())
    print(f"wrote {path}")


here = pathlib.Path(__file__).parent
all_cases = cases_compat15()
build(here / "input.docx", all_cases, 15)
compat14 = here.parent / "textbox-wrap-beside-compat14"
compat14.mkdir(exist_ok=True)
build(compat14 / "input.docx", all_cases[1:3], 14)
