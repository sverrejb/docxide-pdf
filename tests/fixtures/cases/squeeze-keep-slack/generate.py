"""Generate squeeze-keep-slack

A justified line in Word 2013+ (compatibilityMode 15) keeps a last word that
overflows the margin while at most a third of it and the space before it,
that space already narrowed by its share of the overflow, sticks out. Word
keeps the word a little longer than that: 0.0007 to 0.022pt, depending on the
number of spaces on the line. Each pair of paragraphs puts its line's last
word 0.00025pt past the threshold (Word keeps it) and just past the most Word
allowed for that line (Word wraps it). The first word is two letters at odd
half-point sizes and the right indent is in twips; together they set the
overflow to a ten-thousandth of a point.

Tests:
- Arial 11pt "der" after 10 spaces
- Times New Roman 12pt "Dále" after 16 spaces and "te" after 13
"""

import itertools
import pathlib
import sys
import zipfile

# Advance widths (units per 2048 em) of the characters used.
ADV = {
    "Arial": {"a": 1139, "b": 1139, "c": 1024, "d": 1139, "e": 1139, "i": 455, "n": 1139, "o": 1139,
              "r": 682, "s": 1024, "t": 569, "k": 1024, "x": 1024, " ": 569, "l": 455, "h": 1139,
              "m": 1706, "f": 569, "u": 1139, "w": 1479, "p": 1139, "y": 1024, "g": 1139, "D": 1479,
              "á": 1139, ".": 569, "v": 1024},
    "Times New Roman": {"a": 909, "b": 1024, "c": 909, "d": 1024, "e": 909, "i": 569, "n": 1024,
                        "o": 1024, "r": 682, "s": 797, "t": 569, "k": 1024, "x": 1024, " ": 512,
                        "l": 569, "h": 1024, "m": 1593, "f": 682, "u": 1024, "w": 1479, "p": 1024,
                        "y": 1024, "g": 1024, "D": 1479, "á": 909, ".": 512, "v": 1024},
}
TUNE = "abeinorst"
# (font, half-points, words before the last, last word, past the threshold in
# pt Word keeps, past it Word wraps); Word examples.
CASES = [
    ("Arial", 22, 7, "der", 0.00025, 0.0016),
    ("Times New Roman", 24, 13, "Dále", 0.00025, 0.0065),
    ("Times New Roman", 24, 10, "te", 0.00025, 0.0040),
]
FILL = " and then more words follow here to make a second line."


def width(font, text, hp):
    return sum(ADV[font][c] for c in text) * hp / 2 / 2048


W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
LEFT, RIGHT = 1418, 1134
TEXT_W = (11906 - LEFT - RIGHT) / 20

paras = []
for font, hp, count, last, keep, wrap in CASES:
    rest = "kx " + " ".join(["ab"] * count) + " cd"
    space = width(font, " ", hp)
    spaces = (rest.count(" ") + 2) * space
    word = width(font, last, hp)
    threshold = (word + space) / (3 - 2 * space / spaces)
    tail = 2 * space + width(font, rest, hp) + word
    tunes = [(width(font, c1, z1) + width(font, c2, z2), (c1, z1, c2, z2))
             for c1, c2 in itertools.product(TUNE, TUNE) for z1 in range(16, 33) for z2 in range(16, 33)]
    for past in (keep, wrap):
        # overflow = tune + tail - (TEXT_W - right / 20) = threshold + past
        target = threshold + past + TEXT_W - tail
        err, (c1, z1, c2, z2), right = min(
            (abs(w + round((target - w) * 20) / 20 - target), key, round((target - w) * 20)) for w, key in tunes)
        overflow = width(font, c1, z1) + width(font, c2, z2) + tail - (TEXT_W - right / 20)
        print(f"{font} {last}: past the threshold by {overflow - threshold:.6f}", file=sys.stderr)
        rpr = lambda z: f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/><w:sz w:val="{z}"/>'
        runs = [(rpr(z1), c1), (rpr(z2), c2), (rpr(hp), f" {rest} {last}{FILL}")]
        body = "".join(f'<w:r><w:rPr>{r}</w:rPr><w:t xml:space="preserve">{t}</w:t></w:r>' for r, t in runs)
        paras.append(f'<w:p><w:pPr><w:spacing w:after="120"/><w:ind w:right="{right}"/><w:jc w:val="both"/></w:pPr>{body}</w:p>')

document = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document {W}><w:body>' + "".join(paras)
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>'
    f'<w:pgMar w:top="600" w:right="{RIGHT}" w:bottom="600" w:left="{LEFT}" w:header="0" w:footer="0" w:gutter="0"/>'
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
