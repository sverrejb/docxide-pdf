"""Generate case76: Footnote placement w:pos="beneathText" (17.11.21)

Isolates the footnote w:pos placement value, which docxside does not yet parse
(footnotes are always painted at the page-bottom margin — render_page_footnotes,
src/pdf/footnotes.rs). Sets the doc-wide footnote bag to:

  - footnotePr/pos = beneathText  -> footnotes sit directly UNDER the last line
                                     of body text, not at the bottom margin

The body is deliberately SHORT and near the top of a tall page, so "beneathText"
(notes high up, right under the text) is unmistakably different from today's
"pageBottom" (notes pinned to the bottom margin) in the reference PDF. Numbering
format is left at the default (decimal) so PLACEMENT is the only variable.

Note: footnotes cannot be authored via python-docx, so this follows the raw-XML
ZIP pattern of case74/case18. Endnote pos="sectEnd" (needs a multi-section doc to
differ from the docEnd default) is deliberately left for a later fixture.
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
  <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>""")

        z.writestr("_rels/.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL}">
  <Relationship Id="rId1" Type="{DOC_REL}/officeDocument" Target="word/document.xml"/>
</Relationships>""")

        z.writestr("word/_rels/document.xml.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL}">
  <Relationship Id="rId1" Type="{DOC_REL}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="{DOC_REL}/settings" Target="settings.xml"/>
  <Relationship Id="rId3" Type="{DOC_REL}/footnotes" Target="footnotes.xml"/>
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
</w:styles>""")

        # The gap under test: document-wide footnote bag with a NON-DEFAULT pos.
        # CT_FtnProps child order is pos?, numFmt?, ... — pos comes first.
        z.writestr("word/settings.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="{WML}">
  <w:footnotePr>
    <w:pos w:val="beneathText"/>
  </w:footnotePr>
</w:settings>""")

        def fn_ref(fid):
            return (f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>'
                    f'<w:footnoteReference w:id="{fid}"/></w:r>')

        def run(t):
            return f'<w:r><w:t xml:space="preserve">{t}</w:t></w:r>'

        # Short body near the top of a tall page so "beneathText" (notes hugging
        # the text) is visually distinct from "pageBottom" (notes at the margin).
        body = "".join([
            f'<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>{run("Footnote Placement")}</w:p>',
            f'<w:p>{run("The first sentence carries a footnote")}{fn_ref(2)}'
            f'{run(" and the second sentence carries another")}{fn_ref(3)}{run(".")}</w:p>',
        ])

        z.writestr("word/document.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WML}" xmlns:r="{DOC_REL}">
  <w:body>
    {body}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:docGrid w:linePitch="360"/>
    </w:sectPr>
  </w:body>
</w:document>""")

        def fn_para(nid, text):
            return (f'<w:footnote w:id="{nid}"><w:p>'
                    f'<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
                    f'<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
                    f'<w:r><w:t xml:space="preserve"> {text}</w:t></w:r>'
                    f'</w:p></w:footnote>')

        z.writestr("word/footnotes.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{WML}">
  <w:footnote w:type="separator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  {fn_para(2, "First footnote, placed beneath the text.")}
  {fn_para(3, "Second footnote, also beneath the text.")}
</w:footnotes>""")

    return buf.getvalue()


if __name__ == "__main__":
    out = pathlib.Path(__file__).parent / "input.docx"
    out.write_bytes(make_docx())
    print(f"Wrote {out} ({out.stat().st_size} bytes)")
