#!/usr/bin/env python3
"""case50: Deep style inheritance with run vs style vs paragraph conflicts.

Tests 3+ level basedOn chains where properties are set, overridden, and
inherited at different levels. Exercises style resolution logic that must
walk the full chain and merge run-level overrides on top.

Scenarios:
1. 3-level chain: Normal → BaseCustom → MidCustom → LeafCustom
   - BaseCustom: 14pt, bold, blue, 12pt space-after
   - MidCustom: overrides to 12pt (inherits bold+blue from BaseCustom), 6pt space-after
   - LeafCustom: adds italic (inherits 12pt+bold from MidCustom, blue from BaseCustom)
2. Run-level override on top of deep style: LeafCustom paragraph with
   run-level red color (should override blue from BaseCustom)
3. Paragraph-level spacing override: LeafCustom paragraph with explicit
   space-before that doesn't exist in any ancestor
4. 4-level chain: Normal → Level1 → Level2 → Level3 → Level4
   with font family changes at each level
5. Character style basedOn chain: BaseChar → MidChar → LeafChar
   applied as run styles on Normal paragraphs

Usage:
    uv run tests/fixtures/cases/case50/generate.py
"""

import os
import re
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Inches, Pt

OUT = Path("tests/fixtures/cases/case50/input.docx")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def generate():
    doc = Document()
    for section in doc.sections:
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)

    # Placeholders — replaced via ZIP post-processing
    doc.add_paragraph("STYLE_TEST_CONTENT_PLACEHOLDER")

    tmp = tempfile.mktemp(suffix=".docx")
    doc.save(tmp)

    # Custom styles to inject into styles.xml
    custom_styles = """
  <!-- 3-level paragraph style chain: Normal → BaseCustom → MidCustom → LeafCustom -->
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="BaseCustom">
    <w:name w:val="BaseCustom"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:after="240"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>
      <w:b/>
      <w:color w:val="2E75B6"/>
      <w:sz w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="MidCustom">
    <w:name w:val="MidCustom"/>
    <w:basedOn w:val="BaseCustom"/>
    <w:pPr>
      <w:spacing w:after="120"/>
    </w:pPr>
    <w:rPr>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="LeafCustom">
    <w:name w:val="LeafCustom"/>
    <w:basedOn w:val="MidCustom"/>
    <w:rPr>
      <w:i/>
    </w:rPr>
  </w:style>

  <!-- 4-level chain with font family changes: Normal → Level1 → Level2 → Level3 → Level4 -->
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="Level1">
    <w:name w:val="Level1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="120" w:after="120"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>
      <w:sz w:val="28"/>
      <w:color w:val="333333"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="Level2">
    <w:name w:val="Level2"/>
    <w:basedOn w:val="Level1"/>
    <w:rPr>
      <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/>
      <w:b/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="Level3">
    <w:name w:val="Level3"/>
    <w:basedOn w:val="Level2"/>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
      <w:i/>
      <w:sz w:val="22"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:customStyle="1" w:styleId="Level4">
    <w:name w:val="Level4"/>
    <w:basedOn w:val="Level3"/>
    <w:pPr>
      <w:spacing w:before="240" w:after="60"/>
      <w:ind w:left="720"/>
    </w:pPr>
    <w:rPr>
      <w:color w:val="C00000"/>
      <w:u w:val="single"/>
    </w:rPr>
  </w:style>

  <!-- Character style chain: BaseChar → MidChar → LeafChar -->
  <w:style w:type="character" w:customStyle="1" w:styleId="BaseChar">
    <w:name w:val="BaseChar"/>
    <w:rPr>
      <w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/>
      <w:b/>
      <w:sz w:val="28"/>
      <w:color w:val="006600"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:customStyle="1" w:styleId="MidChar">
    <w:name w:val="MidChar"/>
    <w:basedOn w:val="BaseChar"/>
    <w:rPr>
      <w:i/>
      <w:sz w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:customStyle="1" w:styleId="LeafChar">
    <w:name w:val="LeafChar"/>
    <w:basedOn w:val="MidChar"/>
    <w:rPr>
      <w:u w:val="single"/>
      <w:color w:val="660066"/>
    </w:rPr>
  </w:style>"""

    # Body content XML
    body_xml = """
<!-- Section 1: 3-level paragraph style chain -->
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Section 1: Three-Level Style Chain</w:t></w:r>
</w:p>

<w:p>
  <w:r>
    <w:rPr><w:sz w:val="22"/></w:rPr>
    <w:t xml:space="preserve">BaseCustom: 14pt bold blue, 12pt space-after. MidCustom inherits bold+blue, overrides to 12pt, 6pt space-after. LeafCustom adds italic, inherits everything else.</w:t>
  </w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="BaseCustom"/></w:pPr>
  <w:r><w:t>BaseCustom style: 14pt, bold, blue, Times New Roman</w:t></w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="MidCustom"/></w:pPr>
  <w:r><w:t>MidCustom style: 12pt (overridden), bold (inherited), blue (inherited)</w:t></w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="LeafCustom"/></w:pPr>
  <w:r><w:t>LeafCustom style: 12pt (from Mid), bold (from Base), blue (from Base), italic (own)</w:t></w:r>
</w:p>

<!-- Section 2: Run-level override on deep style -->
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Section 2: Run-Level Override on Deep Style</w:t></w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="LeafCustom"/></w:pPr>
  <w:r><w:t xml:space="preserve">LeafCustom with default blue color, then </w:t></w:r>
  <w:r>
    <w:rPr><w:color w:val="C00000"/></w:rPr>
    <w:t xml:space="preserve">this run overrides to red </w:t>
  </w:r>
  <w:r><w:t xml:space="preserve">and back to inherited blue.</w:t></w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="LeafCustom"/></w:pPr>
  <w:r>
    <w:rPr>
      <w:b w:val="0"/>
      <w:sz w:val="36"/>
    </w:rPr>
    <w:t>Run overrides: unbold (cancels inherited bold), 18pt (overrides 12pt from Mid)</w:t>
  </w:r>
</w:p>

<!-- Section 3: Paragraph-level spacing override on deep style -->
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Section 3: Paragraph-Level Spacing Override</w:t></w:r>
</w:p>

<w:p>
  <w:pPr>
    <w:pStyle w:val="LeafCustom"/>
    <w:spacing w:before="480"/>
  </w:pPr>
  <w:r><w:t>LeafCustom with explicit 24pt space-before (not in any ancestor)</w:t></w:r>
</w:p>

<w:p>
  <w:pPr>
    <w:pStyle w:val="MidCustom"/>
    <w:spacing w:after="0"/>
  </w:pPr>
  <w:r><w:t>MidCustom with space-after overridden to 0 (was 6pt from style, 12pt from Base)</w:t></w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="BaseCustom"/></w:pPr>
  <w:r><w:t>BaseCustom (12pt space-after) for visual comparison</w:t></w:r>
</w:p>

<!-- Section 4: 4-level font family chain -->
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Section 4: Four-Level Font Family Chain</w:t></w:r>
</w:p>

<w:p>
  <w:r>
    <w:rPr><w:sz w:val="22"/></w:rPr>
    <w:t xml:space="preserve">Level1: Georgia 14pt gray. Level2: Courier New (overrides font), adds bold. Level3: Arial (overrides font), adds italic, 11pt. Level4: inherits Arial+bold+italic, adds red+underline+indent.</w:t>
  </w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="Level1"/></w:pPr>
  <w:r><w:t>Level1: Georgia 14pt, dark gray color</w:t></w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="Level2"/></w:pPr>
  <w:r><w:t>Level2: Courier New (override), bold (own), 14pt+gray (from Level1)</w:t></w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="Level3"/></w:pPr>
  <w:r><w:t>Level3: Arial (override), italic (own), 11pt (override), bold+gray (inherited)</w:t></w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="Level4"/></w:pPr>
  <w:r><w:t>Level4: Arial+bold+italic (inherited), red+underline (own), 11pt (from L3), indented</w:t></w:r>
</w:p>

<!-- Section 5: Character style chain -->
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Section 5: Character Style Chain</w:t></w:r>
</w:p>

<w:p>
  <w:r>
    <w:rPr><w:sz w:val="22"/></w:rPr>
    <w:t xml:space="preserve">Normal paragraph with character styles applied to runs. BaseChar: Georgia bold 14pt green. MidChar: adds italic, overrides to 12pt. LeafChar: adds underline, overrides to purple.</w:t>
  </w:r>
</w:p>

<w:p>
  <w:r><w:t xml:space="preserve">Normal text, then </w:t></w:r>
  <w:r>
    <w:rPr><w:rStyle w:val="BaseChar"/></w:rPr>
    <w:t xml:space="preserve">BaseChar: Georgia bold 14pt green</w:t>
  </w:r>
  <w:r><w:t xml:space="preserve">, then normal again.</w:t></w:r>
</w:p>

<w:p>
  <w:r><w:t xml:space="preserve">Normal text, then </w:t></w:r>
  <w:r>
    <w:rPr><w:rStyle w:val="MidChar"/></w:rPr>
    <w:t xml:space="preserve">MidChar: italic+bold (inherited), 12pt (override), green (inherited)</w:t>
  </w:r>
  <w:r><w:t xml:space="preserve">, then normal again.</w:t></w:r>
</w:p>

<w:p>
  <w:r><w:t xml:space="preserve">Normal text, then </w:t></w:r>
  <w:r>
    <w:rPr><w:rStyle w:val="LeafChar"/></w:rPr>
    <w:t xml:space="preserve">LeafChar: underline+italic+bold (inherited), purple (override), 12pt (from Mid)</w:t>
  </w:r>
  <w:r><w:t xml:space="preserve">, then normal again.</w:t></w:r>
</w:p>

<!-- Section 6: Conflict — run props vs character style vs paragraph style -->
<w:p>
  <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
  <w:r><w:t>Section 6: Three-Way Conflict</w:t></w:r>
</w:p>

<w:p>
  <w:r>
    <w:rPr><w:sz w:val="22"/></w:rPr>
    <w:t xml:space="preserve">Paragraph style (BaseCustom) sets blue+bold. Character style (LeafChar) sets purple+underline+bold+italic. Run-level sets 20pt and unbold. The expected result is: purple (LeafChar overrides BaseCustom), unbold (run overrides both), italic (LeafChar), underline (LeafChar), 20pt (run).</w:t>
  </w:r>
</w:p>

<w:p>
  <w:pPr><w:pStyle w:val="BaseCustom"/></w:pPr>
  <w:r>
    <w:rPr>
      <w:rStyle w:val="LeafChar"/>
      <w:b w:val="0"/>
      <w:sz w:val="40"/>
    </w:rPr>
    <w:t>Three-way conflict: para=blue+bold, char=purple+italic+underline, run=unbold+20pt</w:t>
  </w:r>
</w:p>
"""

    # Post-process the ZIP
    with zipfile.ZipFile(tmp, "r") as zin:
        with zipfile.ZipFile(str(OUT), "w", zipfile.ZIP_DEFLATED) as zout:
            doc_xml = zin.read("word/document.xml").decode()
            styles_xml = zin.read("word/styles.xml").decode()

            # Replace placeholder
            doc_xml = re.sub(
                r'<w:p[^>]*><w:r><w:t>STYLE_TEST_CONTENT_PLACEHOLDER</w:t></w:r></w:p>',
                lambda m: body_xml,
                doc_xml,
                count=1,
            )

            # Inject custom styles
            styles_xml = styles_xml.replace(
                "</w:styles>", custom_styles + "\n</w:styles>"
            )

            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, doc_xml)
                elif item.filename == "word/styles.xml":
                    zout.writestr(item, styles_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))

    os.unlink(tmp)
    print(f"Generated {OUT}")


if __name__ == "__main__":
    generate()
