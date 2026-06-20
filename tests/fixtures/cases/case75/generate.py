"""Generate case75: Section-wide footnote/endnote property bag (17.11.11 / 17.11.5)

Isolates the SECTION-WIDE w:footnotePr / w:endnotePr that live in w:sectPr
(word/document.xml), which docxside does not yet parse — src/docx/sections.rs
reads no footnote/endnote element. This is distinct from case74, which sets the
DOCUMENT-WIDE bag in the Settings part (word/settings.xml). To keep the section
bag the ONLY source of the note format, this fixture ships NO settings.xml
property bag at all — so a renderer that ignores the sectPr child falls back to
the plain-decimal default, while a renderer that honors it picks up:

  - sectPr/footnotePr/numFmt = upperRoman  -> footnote marks render  I, II, III
  - sectPr/endnotePr/numFmt  = lowerLetter -> endnote  marks render  a, b

A renderer ignoring the section property bag (today's docxside) prints both as
plain decimal (1, 2, 3 / 1, 2). The reference PDF reveals whether the sectPr-level
numFmt is honored.

Note: footnotes/endnotes cannot be authored via python-docx, so this follows the
raw-XML ZIP pattern of case74 / case18. The single change over case74 is moving
the two note property bags from word/settings.xml into the body's w:sectPr (where
CT_SectPr places footnotePr?/endnotePr? FIRST, before pgSz). pos="sectEnd" and the
multi-section override are deliberately left for the 17.11.21/.22 placement fixture.
"""

import zipfile
import io
import pathlib

WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def make_docx():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>""")

        z.writestr("_rels/.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL}">
  <Relationship Id="rId1" Type="{DOC_REL}/officeDocument" Target="word/document.xml"/>
</Relationships>""")

        z.writestr("word/_rels/document.xml.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL}">
  <Relationship Id="rId1" Type="{DOC_REL}/styles" Target="styles.xml"/>
  <Relationship Id="rId3" Type="{DOC_REL}/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId4" Type="{DOC_REL}/endnotes" Target="endnotes.xml"/>
</Relationships>""")

        z.writestr("word/styles.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="{WML}">
  <w:docDefaults>
    <w:rPrDefault><w:rPr>
      <w:sz w:val="24"/>
      <w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/>
    </w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr>
      <w:spacing w:after="160" w:line="278" w:lineRule="auto"/>
    </w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:pPr><w:spacing w:before="360" w:after="80"/><w:keepNext/></w:pPr>
    <w:rPr><w:sz w:val="40"/><w:b/><w:color w:val="0F4761"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="FootnoteReference">
    <w:name w:val="footnote reference"/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="FootnoteText">
    <w:name w:val="footnote text"/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:sz w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="EndnoteReference">
    <w:name w:val="endnote reference"/>
    <w:rPr><w:vertAlign w:val="superscript"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="EndnoteText">
    <w:name w:val="endnote text"/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
    <w:rPr><w:sz w:val="20"/></w:rPr>
  </w:style>
</w:styles>""")

        def fn_ref(fid):
            return (f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                    f'<w:footnoteReference w:id="{fid}"/></w:r>')

        def en_ref(eid):
            return (f'<w:r><w:rPr><w:rStyle w:val="EndnoteReference"/></w:rPr>'
                    f'<w:endnoteReference w:id="{eid}"/></w:r>')

        def run(t):
            return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'

        body = "".join([
            f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>{run("Section-Wide Note Formats")}</w:p>',
            f'<w:p>{run("First claim needing a footnote")}{fn_ref(2)}'
            f'{run(" and a second one")}{fn_ref(3)}{run(".")}</w:p>',
            f'<w:p>{run("A third footnote sits here")}{fn_ref(4)}'
            f'{run(", while this sentence carries an endnote")}{en_ref(2)}{run(".")}</w:p>',
            f'<w:p>{run("A closing endnote follows here")}{en_ref(3)}{run(".")}</w:p>',
        ])

        # The gap under test: the SECTION-WIDE note property bags. CT_SectPr places
        # footnotePr?/endnotePr? as its first children (before pgSz), each carrying
        # a NON-DEFAULT numFmt. No word/settings.xml bag exists, so the sectPr is
        # the only place these formats are declared.
        z.writestr("word/document.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WML}" xmlns:r="{DOC_REL}">
  <w:body>
    {body}
    <w:sectPr>
      <w:footnotePr>
        <w:numFmt w:val="upperRoman"/>
      </w:footnotePr>
      <w:endnotePr>
        <w:numFmt w:val="lowerLetter"/>
      </w:endnotePr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>""")

        def note_para(tag, ref_tag, nid, text):
            return (f'<w:{tag} w:id="{nid}"><w:p>'
                    f'<w:pPr><w:pStyle w:val="{"FootnoteText" if tag=="footnote" else "EndnoteText"}"/></w:pPr>'
                    f'<w:r><w:rPr><w:rStyle w:val="{"FootnoteReference" if tag=="footnote" else "EndnoteReference"}"/></w:rPr><w:{ref_tag}/></w:r>'
                    f'<w:r><w:t xml:space="preserve"> {text}</w:t></w:r>'
                    f'</w:p></w:{tag}>')

        z.writestr("word/footnotes.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{WML}">
  <w:footnote w:type="separator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  {note_para("footnote", "footnoteRef", 2, "First footnote text.")}
  {note_para("footnote", "footnoteRef", 3, "Second footnote text.")}
  {note_para("footnote", "footnoteRef", 4, "Third footnote text.")}
</w:footnotes>""")

        z.writestr("word/endnotes.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="{WML}">
  <w:endnote w:type="separator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:endnote>
  <w:endnote w:type="continuationSeparator" w:id="1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:endnote>
  {note_para("endnote", "endnoteRef", 2, "First endnote text.")}
  {note_para("endnote", "endnoteRef", 3, "Second endnote text.")}
</w:endnotes>""")

    return buf.getvalue()


if __name__ == "__main__":
    out = pathlib.Path(__file__).parent / "input.docx"
    out.write_bytes(make_docx())
    print(f"Wrote {out} ({out.stat().st_size} bytes)")
