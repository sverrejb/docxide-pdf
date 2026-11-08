"""Generate soft-hyphen-contexts: where a w:softHyphen may break and show

Word for Mac behavior, online export, compat 15, Arial 11pt:

- in a paragraph with a tab a soft hyphen breaks and shows as in one
  without: the hyphen counts when fitting ("Diagram-" stays once the line
  holds the prefix and the hyphen, justified once it holds the prefix);
- before a tab it shows only when the tab wraps to the next line;
- one in a w:caps run right after capitals is drawn as a plain hyphen,
  everywhere; one between plain text breaks as usual;
- a left tab wraps only when the text up to it and its hyphen does not fit,
  else the line breaks there ("Diagram-" stays after the tab);
- in a 21pt column "Ver-" fits and the word breaks there ("Ver-" | "ixx"),
  also after "ab";
- under condensed spacing (-4pt, the hyphen's advance below zero) a 40pt
  column breaks at it and shows "MMMM-", also after a tab;
- a word that goes on in a run of another colour or size breaks as one in a
  single run would ("Ver-" | "ixx", the "ixx" red or 8pt; "Veri-" | "xx" at
  24pt with the soft hyphen in the red run);
- no pair is kerned across a soft hyphen: "A", soft hyphen, "V" under
  w:kern is as wide as "AV" unkerned, also where a tab aligns it.

Each sweep keeps the two widths around Word's switch: a right indent ends
the first line k * 0.25pt after the prefix "...Diagram" (k * 0.5pt past
189.933105pt for the left tab before a word).
"""

import pathlib
import zipfile

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/></w:rPr>'
CAPS = RPR.replace("<w:sz", "<w:caps/><w:sz")
PREFIX = "Lorem ipsum dolor sit amet, Diagram"
PREFIX_W = 180.339844  # Arial 11pt advance widths of PREFIX
TEXT_W = 10466  # twips of the column at right indent 0
TAB_W = 36.0  # the tab stop before the text
TAB_STOP = '<w:tabs><w:tab w:val="left" w:pos="720"/></w:tabs>'
SHY = f'<w:r>{RPR}<w:softHyphen/></w:r>'
TAB = f'<w:r>{RPR}<w:tab/></w:r>'


def run(t, rpr=RPR):
    return f'<w:r>{rpr}<w:t xml:space="preserve">{t}</w:t></w:r>'


def in_run(*parts, rpr=RPR):
    """One run whose w:t elements are joined by soft hyphens."""
    texts = "<w:softHyphen/>".join(f'<w:t xml:space="preserve">{p}</w:t>' for p in parts)
    return f"<w:r>{rpr}{texts}</w:r>"


def own_runs(*parts, shy=SHY):
    """Runs of text with each soft hyphen in a run of its own."""
    return shy.join(run(p) for p in parts)


def para(content, jc="left", right=0, tabs=""):
    return (f'<w:p><w:pPr>{tabs}<w:spacing w:after="60"/><w:ind w:right="{right}"/><w:jc w:val="{jc}"/>{RPR}</w:pPr>'
            f'{content}</w:p>')


def label(t):
    return f'<w:p><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr>{run(t)}</w:p>'


def right_for(k, lead=0.0):
    return int(round(TEXT_W - (lead + PREFIX_W + k * 0.25) * 20))


WORD = (PREFIX, "matic", "ally and several words follow.")
caps_word = ("Dia", "gram", "matic", "ally ")
body = [label("In w:caps runs after capitals and between plain text")]
body.append(para("".join(in_run(*caps_word, rpr=CAPS) for _ in range(10)), "left", 5300))
body.append(para("".join(run("Dia", CAPS) + SHY.replace(RPR, CAPS) + run("gram matic ally ") for _ in range(10)),
                 "left", 5900))
body.append(para("".join(own_runs(*caps_word, shy=SHY.replace(RPR, CAPS)) for _ in range(10)), "left", 5300))
for jc, form, sweep in (("left", in_run, (14, 15)), ("both", own_runs, (-1, 0))):
    body.append(label(f"jc={jc}, after a tab"))
    for k in sweep:
        body.append(para(TAB + form(*WORD), jc, right_for(k, TAB_W), TAB_STOP))
body.append(label("Before a tab"))
for k in (14, 15):
    body.append(para(own_runs(PREFIX, "") + TAB + run("matically and several words follow."), "left", right_for(k)))
body.append(label("A left tab before a word the line is too short for"))
for k in (4, 5):
    body.append(para(run("Lorem ipsum dolor") + TAB + in_run("Diagram", "matically"), "left",
                     int(round(TEXT_W - (189.933105 + k * 0.5) * 20))))
body.append(label("A 21pt column"))
for content in (in_run("Ver", "ixx"), run("ab ") + in_run("Ver", "ixx")):
    body.append(para(content, "left", int(round(TEXT_W - 21 * 20))))
body.append(label("Condensed by 4pt"))
CONDENSED = RPR.replace("<w:sz", '<w:spacing w:val="-80"/><w:sz')
body.append(para(in_run("MMMM", "MMMM", rpr=CONDENSED), "left", int(round(TEXT_W - 40 * 20))))
body.append(para(TAB.replace(RPR, CONDENSED) + in_run("MMMM", "MMMM", rpr=CONDENSED), "left",
                 int(round(TEXT_W - (40 + TAB_W) * 20)), TAB_STOP))
body.append(label("The word goes on in runs of another colour or size"))
RED = RPR.replace("<w:sz", '<w:color w:val="FF0000"/><w:sz')
SMALL = RPR.replace('w:val="22"', 'w:val="16"')
for w in (21.0, 24.0):
    right = int(round(TEXT_W - w * 20))
    for content in (in_run("Ver", "") + run("ixx", RED), in_run("Ver", "") + in_run("i", "", rpr=RED) + run("xx"),
                    in_run("Ver", "") + run("ixx", SMALL)):
        body.append(para(content, "left", right))
body.append(label("Kerned Arial 24 after right, centre and left tabs at 150pt"))
KERN = RPR.replace("<w:sz", '<w:kern w:val="2"/><w:sz').replace('w:val="22"', 'w:val="48"')
for kind in ("right", "center", "left"):
    stop = f'<w:tabs><w:tab w:val="{kind}" w:pos="3000"/></w:tabs>'
    for parts in (("A", "V"), ("AVA", "VAV")):
        body.append(para(TAB.replace(RPR, KERN) + in_run(*parts, rpr=KERN) + run(" |", KERN), "left", 0, stop))

document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(body)}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="567" w:right="720" w:bottom="567" w:left="720" w:header="0" w:footer="0" w:gutter="0"/>'
            "</w:sectPr></w:body></w:document>")
styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="de-DE"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          "</w:docDefaults></w:styles>")
settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}>'
            '<w:defaultTabStop w:val="720"/><w:compat>'
            '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
            "</w:compat></w:settings>")
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
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
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/></Relationships>')
out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
print(f"wrote {out}")
