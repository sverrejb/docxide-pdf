"""Generate field-result-format-source: which run formats a regenerated field result

Word lays a PAGE, NUMPAGES or STYLEREF result it regenerates out in the
formatting of the first character of the field's instruction, not in that of
the cached result or of the field's begin or end marker; with \\* MERGEFORMAT
in that of the cached result's first character. A hidden instruction hides the
result.

Tests:
- Body: a wide instruction or cached result, a split instruction and a
  hidden instruction
- Body: \\* MERGEFORMAT with a wide result, a red first result run,
  a red result and wide instruction, and a hidden instruction
- Body: STYLEREF with a wide instruction
- Body: IF { PAGE } with \\* MERGEFORMAT inside quoted true text
  (no switch), and inside false text followed by a real switch
- Header: a small-caps STYLEREF in a text box, in a font no other text uses
- Header: PAGE \\* alphabetic in all caps, in a font no other text uses:
  the font applies, the caps do not (a, b)
- Header: NUMPAGES \\* alphabetic in all caps also drops the caps (b);
  SECTIONPAGES \\* alphabetic keeps them (B)
- Header: PAGE \\* alphabetic in small caps draws full-size lower case (a, b)
- Body: PAGE \\* alphabetic in all caps keeps the caps (A, B)
- The alphabetic results cache a wrong value, so Word must regenerate them
- Page two's heading changes the PAGE and STYLEREF header results
"""

import pathlib
import zipfile

W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
WP = 'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"'
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

WIDE = '<w:w w:val="200"/><w:u w:val="single"/>'
RED = '<w:color w:val="C00000"/>'
HIDDEN = "<w:vanish/>"
MERGE = "\\* MERGEFORMAT"


def run(rpr, inner):
    return f"<w:r><w:rPr>{rpr}</w:rPr>{inner}</w:r>"


def text(t, rpr=""):
    return run(rpr, f'<w:t xml:space="preserve">{t}</w:t>')


def field(instr, results, begin="", instr_rpr="", sep="", end=""):
    return (
        run(begin, '<w:fldChar w:fldCharType="begin"/>')
        + run(instr_rpr, f'<w:instrText xml:space="preserve"> {instr} </w:instrText>')
        + run(sep, '<w:fldChar w:fldCharType="separate"/>')
        + "".join(run(r, f"<w:t>{t}</w:t>") for r, t in results)
        + run(end, '<w:fldChar w:fldCharType="end"/>')
    )


def para(runs, ppr=""):
    return f'<w:p><w:pPr>{ppr}</w:pPr>{"".join(runs)}</w:p>'


def line(label, fld):
    return para([text(f"{label} ["), fld, text("]")])


def if_page(tail, cached):
    return (
        run("", '<w:fldChar w:fldCharType="begin"/>')
        + run("", '<w:instrText xml:space="preserve"> IF </w:instrText>')
        + field("PAGE", [("", "1")])
        + run("", f'<w:instrText xml:space="preserve"> = 1 {tail} </w:instrText>')
        + run("", '<w:fldChar w:fldCharType="separate"/>')
        + run(WIDE, f'<w:t xml:space="preserve">{cached}</w:t>')
        + run("", '<w:fldChar w:fldCharType="end"/>')
    )


QUOTED_MERGE = f'"literal {MERGE} text" "other"'
heading = lambda t: para([text(t)], '<w:pStyle w:val="Heading1"/>')
split_instr = (
    run(RED, '<w:fldChar w:fldCharType="begin"/>')
    + run(RED, '<w:instrText xml:space="preserve"> PA</w:instrText>')
    + run(WIDE, '<w:instrText xml:space="preserve">GE </w:instrText>')
    + run("", '<w:fldChar w:fldCharType="separate"/>')
    + run(WIDE, "<w:t>1</w:t>")
    + run(WIDE, '<w:fldChar w:fldCharType="end"/>')
)
lines = [
    line("Result wide", field("PAGE", [(WIDE, "1")])),
    line("Instruction wide", field("PAGE", [("", "1")], instr_rpr=WIDE)),
    line("Instruction red then wide", split_instr),
    line("Instruction hidden", field("PAGE", [("", "1")], instr_rpr=HIDDEN)),
    line("Merge, result wide", field("PAGE " + MERGE, [(WIDE, "1")])),
    line("Merge, first result run red", field("NUMPAGES " + MERGE, [(RED + WIDE, "1"), ("", "0")])),
    line("Merge, result red, instruction wide", field("PAGE " + MERGE, [(RED, "1")], instr_rpr=WIDE)),
    line("Merge, instruction hidden", field("PAGE " + MERGE, [("", "1")], instr_rpr=HIDDEN)),
    line("Style reference, instruction wide", field('STYLEREF "Heading 1"', [("", "x")], instr_rpr=WIDE)),
    line("If, merge quoted in the true text", if_page(QUOTED_MERGE, f"literal {MERGE} text")),
    line("If, merge quoted and after", if_page(f'"yes" "no {MERGE}" {MERGE}', "yes")),
]
CAPS = "<w:caps/>"
body_caps = line("Body, alphabetic in caps", field("PAGE \\* alphabetic", [(CAPS, "z")], instr_rpr=CAPS))
body = (
    [heading("quiet zone overview")]
    + lines
    + [body_caps, '<w:p><w:r><w:br w:type="page"/></w:r></w:p>', heading("quiet zone details"), body_caps]
)

COURIER = '<w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Courier New"/>'
GEORGIA = '<w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Georgia"/>'
box_w, box_h = 3200400, 360000
textbox = (
    f'<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" '
    f'behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" {WP}><wp:simplePos x="0" y="0"/>'
    '<wp:positionH relativeFrom="margin"><wp:posOffset>2880000</wp:posOffset></wp:positionH>'
    '<wp:positionV relativeFrom="page"><wp:posOffset>360000</wp:posOffset></wp:positionV>'
    f'<wp:extent cx="{box_w}" cy="{box_h}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>'
    '<wp:docPr id="1" name="Running head"/><wp:cNvGraphicFramePr/>'
    f'<a:graphic xmlns:a="{A}"><a:graphicData uri="{WPS}"><wps:wsp xmlns:wps="{WPS}"><wps:cNvSpPr txBox="1"/>'
    f'<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{box_w}" cy="{box_h}"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>'
    '<a:ln><a:noFill/></a:ln></wps:spPr>'
    f'<wps:txbx><w:txbxContent {W}>'
    + para([text("Section: ", COURIER), field('STYLEREF "Heading 1"', [(COURIER + "<w:smallCaps/>", "x")],
                                              instr_rpr=COURIER + "<w:smallCaps/>")])
    + "</w:txbxContent></wps:txbx>"
    '<wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t"/>'
    "</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>"
)
header_paras = [
    para([textbox, text("Sheet "), field("PAGE \\* alphabetic", [(GEORGIA + "<w:caps/>", "z")],
                                         instr_rpr=GEORGIA + "<w:caps/>"),
          text(" Pages "), field("NUMPAGES \\* alphabetic", [(GEORGIA + CAPS, "z")], instr_rpr=GEORGIA + CAPS),
          text(" Section pages "),
          field("SECTIONPAGES \\* alphabetic", [(GEORGIA + CAPS, "z")], instr_rpr=GEORGIA + CAPS)]),
    para([text("Small "), field("PAGE \\* alphabetic", [(GEORGIA + "<w:smallCaps/>", "z")],
                                instr_rpr=GEORGIA + "<w:smallCaps/>")]),
]

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W} {R}><w:body>'
    + "".join(body)
    + '<w:sectPr><w:headerReference w:type="default" r:id="rId3"/>'
    '<w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="1701" w:right="1134" w:bottom="1134" w:left="1418" w:header="708" w:footer="708" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)
header = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr {W}>{"".join(header_paras)}</w:hdr>'

styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults>'
    '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
    '<w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="100" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults>"
    '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>'
    '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>'
    '<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>'
    "</w:styles>"
)

ct = "application/vnd.openxmlformats-officedocument.wordprocessingml"
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    f'<Override PartName="/word/document.xml" ContentType="{ct}.document.main+xml"/>'
    f'<Override PartName="/word/styles.xml" ContentType="{ct}.styles+xml"/>'
    f'<Override PartName="/word/header1.xml" ContentType="{ct}.header+xml"/>'
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
    f'<Relationship Id="rId3" Type="{rel}/header" Target="header1.xml"/></Relationships>'
)

out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/header1.xml", header)
print(f"wrote {out}")
