"""Generate char-scale-widths

Word measures a run with character scaling (w:w) by scaling the face with its
average character width rounded down to whole font units,
floor(scale * xAvgCharWidth) / xAvgCharWidth, and rounding each scaled advance
to whole font units; plain scale * advance is off by up to about a unit per
glyph. Each pair of left-aligned paragraphs puts the right margin 0.01-0.06pt
past and 0.06-0.11pt short of Word's width of the line "Probe <scaled word>",
so the word stays on the first line in one and wraps in the other.

Tests:
- Arial at 66, 80, 85, 90, 95, 120 and 150%, 11 and 14pt
- Arial bold, Times New Roman regular, bold and italic
- Arial bold at 95%, where 95% of xAvgCharWidth 980 is exactly 931 units
- Auto-width tables whose first column holds 20 Arial "n" at 50 and 150%:
  autofit sizes the column to the scaled width
- w:w 0 and 601, which Word draws at 100%
- A shadow at 150% and an emboss at 85% with letter spacing: Word draws the
  copies glyph by glyph at the face's positions, only offset
"""

import math
import pathlib
import sys
import zipfile

# Advance widths (units per 2048 em) and OS/2 xAvgCharWidth of the faces used.
FACES = {
    ("Arial", ""): (904, {"P": 1366, "r": 682, "o": 1139, "b": 1139, "e": 1139, " ": 569, "n": 1139, "m": 1706}),
    ("Arial", "b"): (980, {"P": 1366, "r": 797, "o": 1251, "b": 1251, "e": 1139, " ": 569, "n": 1251, "m": 1821}),
    ("Times New Roman", ""): (821, {"P": 1139, "r": 682, "o": 1024, "b": 1024, "e": 909, " ": 512, "n": 1024, "m": 1593}),
    ("Times New Roman", "b"): (874, {"P": 1251, "r": 909, "o": 1024, "b": 1139, "e": 909, " ": 512, "n": 1139, "m": 1706}),
    ("Times New Roman", "i"): (823, {"P": 1251, "r": 797, "o": 1024, "b": 1024, "e": 909, " ": 512, "n": 1024, "m": 1479}),
}
CASES = [
    ("Arial", "", 11, 66), ("Arial", "", 11, 80), ("Arial", "", 11, 85), ("Arial", "", 11, 90),
    ("Arial", "", 11, 95), ("Arial", "", 11, 120), ("Arial", "", 11, 150), ("Arial", "", 14, 85),
    ("Arial", "b", 14, 85),
    ("Times New Roman", "", 12, 85), ("Times New Roman", "", 12, 90),
    ("Times New Roman", "b", 12, 85), ("Times New Roman", "i", 12, 80),
    ("Arial", "b", 11, 95),
]


def scaled_units(adv, avg, scale):
    return math.floor(2 * adv * math.floor(scale * avg / 100 + 1e-9) / avg + 1) // 2


def word_for(avg, adv, scale, size):
    """The glyph Word's rounding moves most, repeated to about 250pt."""
    def error(c):
        return abs(scaled_units(adv[c], avg, scale) - adv[c] * scale / 100)
    glyph = max("nmeob", key=error)
    count = int(250 / (adv[glyph] * scale / 100 * size / 2048))
    return glyph * count


W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
LEFT, RIGHT = 1418, 1134
TEXT_W = (11906 - LEFT - RIGHT) / 20

paras = []
for font, style, size, scale in CASES:
    avg, adv = FACES[(font, style)]
    word = word_for(avg, adv, scale, size)
    prefix = sum(adv[c] for c in "Probe ") * size / 2048
    width = prefix + sum(scaled_units(adv[c], avg, scale) for c in word) * size / 2048
    ideal = prefix + sum(adv[c] for c in word) * scale / 100 * size / 2048
    face = f'<w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}"/>' + {"": "", "b": "<w:b/>", "i": "<w:i/>"}[style]
    rpr = f'{face}<w:sz w:val="{size * 2}"/>'
    for sign in (1, -1):
        right = math.floor((TEXT_W - width - 0.01) * 20) if sign > 0 else math.ceil((TEXT_W - width + 0.06) * 20)
        margin = TEXT_W - right / 20
        print(f"{font} {style or '-'} {size} {scale}% {word[0]}x{len(word)}: word {width:.3f} ideal {ideal:.3f} margin {margin:.3f}", file=sys.stderr)
        paras.append(
            f'<w:p><w:pPr><w:spacing w:after="60"/><w:ind w:right="{right}"/></w:pPr>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve">Probe </w:t></w:r>'
            f'<w:r><w:rPr>{rpr}<w:w w:val="{scale}"/></w:rPr><w:t>{word}</w:t></w:r>'
            f'<w:r><w:rPr>{rpr}</w:rPr><w:t xml:space="preserve"> to fill one more line.</w:t></w:r></w:p>')

BORDERS = "".join(f'<w:{side} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                  for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
for scale in (50, 150):
    cells = [f'<w:r><w:rPr><w:w w:val="{scale}"/></w:rPr><w:t>{"n" * 20}</w:t></w:r>', "<w:r><w:t>x</w:t></w:r>"]
    paras.append(
        f'<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>{BORDERS}</w:tblBorders>'
        '<w:tblLayout w:type="autofit"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid><w:tr>'
        + "".join(f'<w:tc><w:tcPr><w:tcW w:w="0" w:type="auto"/></w:tcPr><w:p>{c}</w:p></w:tc>' for c in cells)
        + "</w:tr></w:tbl><w:p/>")

for scale in (0, 601):
    paras.append(
        f'<w:p><w:r><w:t xml:space="preserve">Scale {scale}: </w:t></w:r>'
        f'<w:r><w:rPr><w:w w:val="{scale}"/></w:rPr><w:t>abc nnnnnnnnnn</w:t></w:r>'
        '<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>')
for effect, scale, spacing in (("shadow", 150, 20), ("emboss", 85, -10)):
    paras.append(
        f'<w:p><w:pPr><w:spacing w:after="60"/></w:pPr><w:r><w:t xml:space="preserve">{effect} </w:t></w:r>'
        f'<w:r><w:rPr><w:spacing w:val="{spacing}"/><w:w w:val="{scale}"/><w:{effect}/></w:rPr><w:t>{"n" * 30}</w:t></w:r>'
        '<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p>')

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
