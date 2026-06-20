#!/usr/bin/env python3
"""case70: w:lnNumType (Line Numbering Settings, OOXML 17.6.8) — margin line numbers.

Isolates the section-level line-numbering feature: a sectPr child
`<w:lnNumType w:countBy="1" w:start="1" w:restart="continuous"/>` that prints a
line number in the left margin for every line of body text (classic legal /
contract / code-listing layout). Several short single-line paragraphs so the
1,2,3,... margin sequence is unmistakable in the reference.
"""

import os
import tempfile
import zipfile
from pathlib import Path

from docx import Document

OUT = Path("tests/fixtures/cases/case70/input.docx")

# 17.6.8: @countBy=number every Nth line, @start=first number,
# @restart=newPage|newSection|continuous, @distance=twips text->number.
# In CT_SectPr child order, w:lnNumType precedes w:cols (and w:pgNumType).
LN_NUM_TYPE = '<w:lnNumType w:countBy="1" w:start="1" w:restart="continuous"/>'


def main():
    doc = Document()
    for i in range(1, 13):
        doc.add_paragraph(f"Numbered line number {i} of the section body text.")

    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        doc.save(tmp.name)
        tmp_path = tmp.name

    with zipfile.ZipFile(tmp_path, "r") as zin:
        doc_xml = zin.read("word/document.xml").decode("utf-8")
        other = {n: zin.read(n) for n in zin.namelist() if n != "word/document.xml"}

    # lnNumType precedes w:cols in CT_SectPr; fall back to docGrid, then end.
    if "<w:cols" in doc_xml:
        doc_xml = doc_xml.replace("<w:cols", LN_NUM_TYPE + "<w:cols", 1)
    elif "<w:docGrid" in doc_xml:
        doc_xml = doc_xml.replace("<w:docGrid", LN_NUM_TYPE + "<w:docGrid", 1)
    else:
        doc_xml = doc_xml.replace("</w:sectPr>", LN_NUM_TYPE + "</w:sectPr>", 1)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(str(OUT), "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in other.items():
            zout.writestr(name, data)
        zout.writestr("word/document.xml", doc_xml.encode("utf-8"))

    os.unlink(tmp_path)
    print(f"Wrote {OUT}")


if __name__ == "__main__":
    main()
