#!/usr/bin/env python3
"""case44: Extended WordArt — more warp presets, glow, combined effects, varied sizes.

Complements case43 by testing additional warp presets and effect combinations:
Page 1:
  1. textDeflate — pinched middle
  2. textInflate — bulging middle
  3. textChevron — V-shape
  4. textTriangle — wide top, narrow bottom
  5. textCircle — circular text path

Page 2:
  6. textCascadeDown — descending cascade
  7. textDoubleWave1 — double wave
  8. textFadeRight — perspective fade
  9. textCanDown — barrel bottom
  10. textRingOutside — ring outside

Page 3:
  11. textDeflateBottom — deflate bottom only
  12. textCurveUp — upward curve
  13. Glow effect (textPlain + glow, no warp distortion)
  14. Combined: outline + shadow + textWave2
  15. Small text with textStop (octagon shape)
"""

import os
import re
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

OUT = Path("tests/fixtures/cases/case44/input.docx")

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
WP14_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
GRAPHIC_DATA_URI = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

EMU_PER_PT = 12700
EMU_PER_INCH = 914400


def wordart_xml(
    text,
    warp_preset,
    y_offset_emu,
    color="4472C4",
    font_size_hps=72,
    outline_color=None,
    outline_width_emu=9525,
    shadow=False,
    glow=False,
    glow_color="FFC000",
    glow_radius_emu=63500,
    adj_values=None,
    width_emu=int(5.5 * EMU_PER_INCH),
    height_emu=int(1.0 * EMU_PER_INCH),
):
    """Generate mc:AlternateContent XML for a WordArt textbox."""

    # w14:textOutline
    outline_xml = ""
    if outline_color:
        outline_xml = (
            f'<w14:textOutline w14:w="{outline_width_emu}" w14:cap="flat">'
            f'<w14:solidFill><w14:srgbClr w14:val="{outline_color}"/></w14:solidFill>'
            f"</w14:textOutline>"
        )

    # w14:shadow
    shadow_xml = ""
    if shadow:
        shadow_xml = (
            f'<w14:shadow w14:blurRad="38100" w14:dist="25400" w14:dir="3600000">'
            f'<w14:srgbClr w14:val="808080">'
            f'<w14:alpha w14:val="60000"/>'
            f"</w14:srgbClr>"
            f"</w14:shadow>"
        )

    # w14:glow
    glow_xml = ""
    if glow:
        glow_xml = (
            f'<w14:glow w14:rad="{glow_radius_emu}">'
            f'<w14:srgbClr w14:val="{glow_color}">'
            f'<w14:alpha w14:val="60000"/>'
            f"</w14:srgbClr>"
            f"</w14:glow>"
        )

    # Adjustment values
    av_xml = "<a:avLst/>"
    if adj_values:
        gds = "".join(
            f'<a:gd name="{n}" fmla="val {v}"/>' for n, v in adj_values
        )
        av_xml = f"<a:avLst>{gds}</a:avLst>"

    doc_pr_id = abs(hash(text + warp_preset)) % 100000

    return f"""<mc:AlternateContent xmlns:mc="{MC_NS}">
  <mc:Choice Requires="wps">
    <w:drawing xmlns:w="{W_NS}">
      <wp:anchor distT="0" distB="0" distL="0" distR="0"
          simplePos="0" relativeHeight="251658240" behindDoc="0"
          locked="0" layoutInCell="1" allowOverlap="1"
          xmlns:wp="{WP_NS}" xmlns:wp14="{WP14_NS}">
        <wp:simplePos x="0" y="0"/>
        <wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>
        <wp:positionV relativeFrom="paragraph"><wp:posOffset>{y_offset_emu}</wp:posOffset></wp:positionV>
        <wp:extent cx="{width_emu}" cy="{height_emu}"/>
        <wp:effectExtent l="0" t="0" r="0" b="0"/>
        <wp:wrapNone/>
        <wp:docPr id="{doc_pr_id}" name="WordArt {text[:10]}"/>
        <a:graphic xmlns:a="{A_NS}">
          <a:graphicData uri="{GRAPHIC_DATA_URI}">
            <wps:wsp xmlns:wps="{WPS_NS}">
              <wps:cNvSpPr txBox="1"/>
              <wps:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="{width_emu}" cy="{height_emu}"/></a:xfrm>
                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                <a:noFill/>
                <a:ln><a:noFill/></a:ln>
              </wps:spPr>
              <wps:txbx>
                <w:txbxContent xmlns:w="{W_NS}">
                  <w:p>
                    <w:pPr><w:jc w:val="center"/></w:pPr>
                    <w:r>
                      <w:rPr>
                        <w:b/>
                        <w:color w:val="{color}"/>
                        <w:sz w:val="{font_size_hps}"/>
                        <w:szCs w:val="{font_size_hps}"/>
                        {outline_xml}
                        {shadow_xml}
                        {glow_xml}
                      </w:rPr>
                      <w:t>{text}</w:t>
                    </w:r>
                  </w:p>
                </w:txbxContent>
              </wps:txbx>
              <wps:bodyPr fromWordArt="1" wrap="none" lIns="0" tIns="0" rIns="0" bIns="0">
                <a:prstTxWarp prst="{warp_preset}">{av_xml}</a:prstTxWarp>
                <a:spAutoFit/>
              </wps:bodyPr>
            </wps:wsp>
          </a:graphicData>
        </a:graphic>
      </wp:anchor>
    </w:drawing>
  </mc:Choice>
</mc:AlternateContent>"""


def main():
    doc = Document()

    # Page 1 title
    title = doc.add_paragraph()
    run = title.add_run("Extended WordArt Tests \u2014 Page 1")
    run.bold = True
    run.font.size = Pt(14)

    # Spacers for page 1
    for _ in range(20):
        doc.add_paragraph("")

    # Page break
    p_break = doc.add_paragraph()
    run_br = p_break.add_run()
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    run_br._element.append(br)

    # Page 2 title
    title2 = doc.add_paragraph()
    run2 = title2.add_run("Extended WordArt Tests \u2014 Page 2")
    run2.bold = True
    run2.font.size = Pt(14)

    for _ in range(20):
        doc.add_paragraph("")

    # Page break
    p_break2 = doc.add_paragraph()
    run_br2 = p_break2.add_run()
    br2 = OxmlElement("w:br")
    br2.set(qn("w:type"), "page")
    run_br2._element.append(br2)

    # Page 3 title
    title3 = doc.add_paragraph()
    run3 = title3.add_run("Extended WordArt Tests \u2014 Page 3")
    run3.bold = True
    run3.font.size = Pt(14)

    for _ in range(20):
        doc.add_paragraph("")

    # Save base document
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        doc.save(tmp.name)
        tmp_path = tmp.name

    # Read document XML
    with zipfile.ZipFile(tmp_path, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        other_files = {}
        for name in zin.namelist():
            if name != "word/document.xml":
                other_files[name] = zin.read(name)

    # Add w14 namespace if not present
    if f'xmlns:w14="{W14_NS}"' not in doc_xml:
        doc_xml = doc_xml.replace(
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
            'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
            f' xmlns:w14="{W14_NS}"',
        )
    if "mc:Ignorable" in doc_xml and "w14" not in re.search(r'mc:Ignorable="([^"]*)"', doc_xml).group(1):
        doc_xml = re.sub(
            r'mc:Ignorable="([^"]*)"',
            r'mc:Ignorable="\1 w14"',
            doc_xml,
        )

    # Define the WordArt specs for each page
    # y offsets in inches from anchor paragraph
    page1_specs = [
        ("DEFLATE", "textDeflate", 0.7, "2E75B6", 72, "1F4E79", False, False, None, None),
        ("INFLATE", "textInflate", 2.2, "ED7D31", 72, "C44F1A", False, False, None, None),
        ("CHEVRON", "textChevron", 3.8, "70AD47", 72, None, False, False, None, None),
        ("TRIANGLE", "textTriangle", 5.5, "C00000", 72, "800000", False, False, None, None),
        ("CIRCLE", "textCircle", 7.2, "7030A0", 72, None, False, False,
         [("adj", "10800000")], int(1.5 * EMU_PER_INCH)),
    ]

    page2_specs = [
        ("CASCADE", "textCascadeDown", 0.7, "4472C4", 72, "2E5FA1", False, False, None, None),
        ("DOUBLE WAVE", "textDoubleWave1", 2.2, "ED7D31", 72, None, False, False, None, None),
        ("FADE RIGHT", "textFadeRight", 3.8, "70AD47", 72, "4A7A2E", False, False, None, None),
        ("BARREL", "textCanDown", 5.5, "C00000", 72, None, False, False, None, None),
        ("RING", "textRingOutside", 7.2, "7030A0", 72, None, False, False,
         None, int(1.5 * EMU_PER_INCH)),
    ]

    page3_specs = [
        ("DEFLATE BTM", "textDeflateBottom", 0.7, "2E75B6", 72, "1F4E79", False, False, None, None),
        ("CURVE UP", "textCurveUp", 2.2, "ED7D31", 72, None, False, False, None, None),
        ("GLOW TEXT", "textPlain", 3.8, "4472C4", 64, None, False, True, None, None),
        ("WAVE COMBO", "textWave2", 5.5, "C00000", 72, "800000", True, False, None, None),
        ("STOP", "textStop", 7.2, "70AD47", 48, "4A7A2E", False, False, None, None),
    ]

    def build_arts(specs):
        parts = []
        for text, preset, y_in, color, hps, outline, shadow, glow, adj, h_emu in specs:
            y_emu = int(y_in * EMU_PER_INCH)
            kwargs = {"color": color, "font_size_hps": hps}
            if outline:
                kwargs["outline_color"] = outline
            if shadow:
                kwargs["shadow"] = True
            if glow:
                kwargs["glow"] = True
                kwargs["glow_color"] = "FFC000"
            if adj:
                kwargs["adj_values"] = adj
            if h_emu:
                kwargs["height_emu"] = h_emu
            parts.append(wordart_xml(text=text, warp_preset=preset, y_offset_emu=y_emu, **kwargs))
        return "".join(parts)

    page1_xml = build_arts(page1_specs)
    page2_xml = build_arts(page2_specs)
    page3_xml = build_arts(page3_specs)

    # Inject WordArt by finding title paragraphs via their text content.
    # Each title has unique text we can locate.
    markers = [
        ("Extended WordArt Tests \u2014 Page 1", page1_xml),
        ("Extended WordArt Tests \u2014 Page 2", page2_xml),
        ("Extended WordArt Tests \u2014 Page 3", page3_xml),
    ]

    for marker_text, arts_xml in markers:
        # Find the closing </w:t> after the marker text, then the next </w:r>
        # and inject WordArt after the </w:r> (inside the same <w:p>)
        marker_pos = doc_xml.find(marker_text)
        if marker_pos == -1:
            print(f"WARNING: Could not find marker '{marker_text}'")
            continue
        # Find the next </w:r> after the marker
        r_close = doc_xml.find("</w:r>", marker_pos)
        if r_close == -1:
            print(f"WARNING: Could not find </w:r> after marker '{marker_text}'")
            continue
        insert_pos = r_close + len("</w:r>")
        doc_xml = doc_xml[:insert_pos] + arts_xml + doc_xml[insert_pos:]

    # Write output
    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(str(OUT), "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in other_files.items():
            zout.writestr(name, data)
        zout.writestr("word/document.xml", doc_xml.encode("utf-8"))

    os.unlink(tmp_path)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
