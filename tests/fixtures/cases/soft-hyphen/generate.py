"""Generate soft-hyphen: w:softHyphen as Word breaks and draws it

A w:softHyphen is an optional hyphen (Word for Mac behavior, online export,
compat 15):

- where the line does not break at it, it draws nothing and takes no room;
- a line may break after it, and then shows a hyphen, which must fit the
  line like the rest of the word: left-aligned, "Diagram-" stays on the line
  once the column holds the prefix and the hyphen (3.75pt past the prefix,
  not 3.5pt);
- a justified line may also keep it by narrowing its spaces, as for a hard
  hyphen ("Diagram-" stays once the column holds the prefix, not 0.25pt
  less);
- the same whether the soft hyphen sits between w:t elements of one run or
  in a run of its own;
- a long word breaks at its last soft hyphen that fits.

Each sweep keeps the two widths around Word's switch: a right indent ends
the first line k * 0.25pt after the prefix "...Diagram" (Arial 11pt).
"""

import pathlib
import zipfile

NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
RPR = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/></w:rPr>'
PREFIX = "Lorem ipsum dolor sit amet, Diagram"
PREFIX_W = 180.339844  # Arial 11pt advance widths of PREFIX
TEXT_W = 10466  # twips of the column at right indent 0


def run(t):
    return f'<w:r>{RPR}<w:t xml:space="preserve">{t}</w:t></w:r>'


def shy_run():
    return f'<w:r>{RPR}<w:softHyphen/></w:r>'


def para(content, jc, right):
    return (f'<w:p><w:pPr><w:spacing w:after="60"/><w:ind w:right="{right}"/><w:jc w:val="{jc}"/>{RPR}</w:pPr>'
            f'{content}</w:p>')


def label(t):
    return f'<w:p><w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr>{run(t)}</w:p>'




def in_run(*parts):
    """One run whose w:t elements are joined by soft hyphens."""
    shy = "<w:softHyphen/>"
    texts = shy.join(f'<w:t xml:space="preserve">{p}</w:t>' for p in parts)
    return f"<w:r>{RPR}{texts}</w:r>"


def own_runs(*parts):
    """Runs of text with each soft hyphen in a run of its own."""
    return shy_run().join(run(p) for p in parts)


WORD = (PREFIX, "matic", "ally and several words follow.")
body = []
for jc, form, name, sweep in (("left", in_run, "between the w:t of one run", (14, 15)),
                              ("both", own_runs, "in runs of their own", (-1, 0))):
    body.append(label(f"jc={jc}, soft hyphens {name}"))
    for k in sweep:
        right = int(round(TEXT_W - (PREFIX_W + k * 0.25) * 20))
        body.append(para(form(*WORD), jc, right))
    body.append('<w:p><w:r><w:br w:type="page"/></w:r></w:p>')
body.append(label("A long word breaks at its last soft hyphen that fits:"))
body.append(para(run("The supercal") + own_runs("ifrag", "ilist", "ice", "xpia", "lidocious works."), "left", 7685))

document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {NS}><w:body>{"".join(body)}'
            '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="567" w:right="720" w:bottom="567" w:left="720" w:header="0" w:footer="0" w:gutter="0"/>'
            "</w:sectPr></w:body></w:document>")
styles = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {NS}><w:docDefaults>'
          '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>'
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="de-DE"/></w:rPr></w:rPrDefault>'
          '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
          "</w:docDefaults></w:styles>")
settings = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {NS}><w:compat>'
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
