"""Generate case18: Footnotes

Tests:
- Footnote reference marks in body text (superscript numbers)
- Footnote content rendered at page bottom
- Separator line between body text and footnotes
- Multiple footnotes per page
- Footnote with bold/italic formatting
- Multiple footnotes in one paragraph
- Enough body text to verify footnotes reduce available body space
"""

import zipfile
import io

WML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

def make_docx():
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        # [Content_Types].xml
        z.writestr("[Content_Types].xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>""")

        # _rels/.rels
        z.writestr("_rels/.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL}">
  <Relationship Id="rId1" Type="{DOC_REL}/officeDocument" Target="word/document.xml"/>
</Relationships>""")

        # word/_rels/document.xml.rels
        z.writestr("word/_rels/document.xml.rels", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{REL}">
  <Relationship Id="rId1" Type="{DOC_REL}/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="{DOC_REL}/footnotes" Target="footnotes.xml"/>
</Relationships>""")

        # word/styles.xml — Aptos 12pt defaults + FootnoteReference + FootnoteText styles
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

        # Helper: footnote reference in body text
        def fn_ref(fn_id):
            return f"""<w:r>
        <w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
        <w:footnoteReference w:id="{fn_id}"/>
      </w:r>"""

        # Helper: simple text run
        def text_run(text, bold=False, italic=False):
            rpr = ""
            parts = []
            if bold:
                parts.append("<w:b/>")
            if italic:
                parts.append("<w:i/>")
            if parts:
                rpr = f"<w:rPr>{''.join(parts)}</w:rPr>"
            return f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'

        # Helper: paragraph
        def para(runs_xml, style=None):
            ppr = ""
            if style:
                ppr = f"<w:pPr><w:pStyle w:val=\"{style}\"/></w:pPr>"
            return f"<w:p>{ppr}{runs_xml}</w:p>"

        # Build document.xml
        body_paragraphs = []

        # Title
        body_paragraphs.append(para(text_run("Footnotes in Documents"), "Heading1"))

        # Para 1: single footnote
        body_paragraphs.append(para(
            text_run("Footnotes are a standard feature of academic and professional writing")
            + fn_ref(2)
            + text_run(". They provide additional context without interrupting the main text flow.")
        ))

        # Para 2: two footnotes in same paragraph
        body_paragraphs.append(para(
            text_run("The history of footnotes dates back to the invention of the printing press")
            + fn_ref(3)
            + text_run(", and they remain essential in modern publishing")
            + fn_ref(4)
            + text_run(". Different style guides have varying rules for their usage.")
        ))

        # Para 3: regular text (no footnotes)
        body_paragraphs.append(para(
            text_run("This paragraph has no footnotes. It exists to add body text and verify that normal paragraphs render correctly between paragraphs that contain footnote references.")
        ))

        # Para 4: footnote with longer reference text
        body_paragraphs.append(para(
            text_run("In scientific writing, footnotes serve a different purpose than in humanities")
            + fn_ref(5)
            + text_run(". Scientists typically prefer endnotes or inline citations, while historians and literary scholars often use extensive footnotes to discuss sources and provide commentary.")
        ))

        # Para 5: another footnote
        body_paragraphs.append(para(
            text_run("Legal documents frequently use footnotes for case citations and statutory references")
            + fn_ref(6)
            + text_run(". The footnote numbering restarts in some styles and continues in others.")
        ))

        # Para 6: closing paragraph with footnote
        body_paragraphs.append(para(
            text_run("This final paragraph tests that footnote rendering works correctly when multiple footnotes accumulate at the bottom of the page")
            + fn_ref(7)
            + text_run(".")
        ))

        doc_body = "\n".join(body_paragraphs)
        z.writestr("word/document.xml", f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{WML}" xmlns:r="{DOC_REL}">
  <w:body>
    {doc_body}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
      <w:docGrid w:linePitch="360"/>
      <w:footnotePr/>
    </w:sectPr>
  </w:body>
</w:document>""")

        # word/footnotes.xml
        def footnote_para(fn_id, text_parts):
            """text_parts is a list of (text, bold, italic) tuples"""
            runs = '<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>'
            runs += f'<w:r><w:t xml:space="preserve"> </w:t></w:r>'
            for text, bold, italic in text_parts:
                rpr_parts = ['<w:sz w:val="20"/>']
                if bold:
                    rpr_parts.append("<w:b/>")
                if italic:
                    rpr_parts.append("<w:i/>")
                rpr = f"<w:rPr>{''.join(rpr_parts)}</w:rPr>"
                runs += f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'
            return f"""<w:footnote w:id="{fn_id}">
      <w:p>
        <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
        {runs}
      </w:p>
    </w:footnote>"""

        footnotes_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="{WML}">
    <w:footnote w:type="separator" w:id="0">
      <w:p>
        <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
        <w:r><w:separator/></w:r>
      </w:p>
    </w:footnote>
    <w:footnote w:type="continuationSeparator" w:id="1">
      <w:p>
        <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>
        <w:r><w:continuationSeparator/></w:r>
      </w:p>
    </w:footnote>
    {footnote_para(2, [("This is a simple footnote providing additional context about the statement above.", False, False)])}
    {footnote_para(3, [("Gutenberg's movable type press, invented around 1440, revolutionized the dissemination of knowledge.", False, False)])}
    {footnote_para(4, [("See ", False, False), ("The Chicago Manual of Style", False, True), (", 17th edition, for comprehensive footnote formatting guidelines.", False, False)])}
    {footnote_para(5, [("Notable exceptions include the ", False, False), ("Nature", False, True), (" journal family, which uses a numbered reference system that functions similarly to footnotes.", False, False)])}
    {footnote_para(6, [("For example, ", False, False), ("Marbury v. Madison", False, True), (", 5 U.S. 137 (1803), established the principle of judicial review.", False, False)])}
    {footnote_para(7, [("Final footnote. When many footnotes appear on one page, Word allocates space at the bottom and reduces the body text area accordingly.", False, False)])}
</w:footnotes>"""

        z.writestr("word/footnotes.xml", footnotes_xml)

    return buf.getvalue()


if __name__ == "__main__":
    import pathlib
    out = pathlib.Path(__file__).parent / "input.docx"
    out.write_bytes(make_docx())
    print(f"Wrote {out} ({out.stat().st_size} bytes)")
