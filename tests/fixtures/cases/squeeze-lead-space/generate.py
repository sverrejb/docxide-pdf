"""Generate squeeze-lead-space

A justified line in Word 2013+ (compatibilityMode 15) keeps a word that
overflows the margin by narrowing the line's spaces, when at most a third of
the space before the word and the word sticks out, with that space already
narrowed by its share of the overflow (the spaces give it up in proportion to
their widths). So the same word may overflow further on a line with fewer
spaces. Two paragraphs' right indents put their line's last word 0.025 pt
inside and outside that threshold; the line must not wrap anywhere else, so a
second line of words follows.

Tests:
- Arial 11 pt "xx" and "der" after 10 spaces, "der" and "Haus" after 20,
  "xx", "der" and "Haus" after 30
- The same three words in Arial 16 pt after 14 spaces
- "der" after 10 spaces split into runs "de" and a coloured "r", plain and
  underlined (the underline bridges the spaces): judged as the whole word
"""

import pathlib
import sys
import zipfile

# Arial advance widths (units per 2048) of the characters used.
ADV = {"a": 1139, "b": 1139, "x": 1024, "d": 1139, "e": 1139, "r": 682, "H": 1479,
       "u": 1139, "s": 1024, " ": 569}


def width(text, size):
    return sum(ADV[c] for c in text) / 2048 * size


W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
TEXT_W = (11906 - 1418 - 1134) / 20

paras = []
for size, count, words in [
    (11, 10, ["xx", "der"]),
    (11, 20, ["der", "Haus"]),
    (11, 30, ["xx", "der", "Haus"]),
    (16, 14, ["xx", "der", "Haus"]),
]:
    space = width(" ", size)
    prefix = " ".join(["ab"] * count)
    share = 1 / count
    tail = " and more words after it to fill a second line."
    for last in words:
        word = width(last, size)
        threshold = (word + space) / (3 - 2 * share)
        natural = width(prefix, size) + space + word
        variants = [("", [(f"{prefix} {last}{tail}", "")])]
        if (size, count, last) == (11, 10, "der"):
            split = [(f"{prefix} de", ""), ("r", '<w:color w:val="C00000"/>'), (tail, "")]
            variants += [("", split), ('<w:u w:val="single"/>', split)]
        for k in (-1, 0):
            overflow = threshold + (k + 0.5) * 0.05
            right = round((TEXT_W - natural + overflow) * 20)
            for underline, runs in variants:
                body = "".join(
                    f'<w:r><w:rPr>{color}{underline}<w:sz w:val="{size * 2}"/></w:rPr>'
                    f'<w:t xml:space="preserve">{text}</w:t></w:r>'
                    for text, color in runs)
                paras.append(
                    f'<w:p><w:pPr><w:spacing w:after="40"/><w:ind w:right="{right}"/><w:jc w:val="both"/></w:pPr>'
                    f'{body}</w:p>')

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>' + "".join(paras)
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    '<w:pgMar w:top="600" w:right="1134" w:bottom="600" w:left="1418" w:header="0" w:footer="0" w:gutter="0"/>'
    "</w:sectPr></w:body></w:document>"
)
styles = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles {W}><w:docDefaults><w:rPrDefault><w:rPr>'
    '<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>'
    '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>'
    "</w:docDefaults></w:styles>"
)
settings = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings {W}><w:compat>'
    '<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>'
    "</w:compat></w:settings>"
)
content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
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
    f'<Relationship Id="rId2" Type="{rel}/settings" Target="settings.xml"/></Relationships>'
)
out = pathlib.Path(__file__).parent / "input.docx"
with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", package_rels)
    z.writestr("word/document.xml", document)
    z.writestr("word/_rels/document.xml.rels", document_rels)
    z.writestr("word/styles.xml", styles)
    z.writestr("word/settings.xml", settings)
print(f"wrote {out} ({len(paras)} paragraphs)", file=sys.stderr)
