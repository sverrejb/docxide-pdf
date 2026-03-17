"""Generate a bigram sheet DOCX for mining Word's inter-glyph adjustments.

Usage: uv run tools/generate_bigram_sheet.py [font_name] [font_size_pt] [output_path]

Sets the font via the Normal style's docDefaults to ensure Word exports
with direct Tf (no matrix scaling), matching handcrafted test cases.
"""
import sys
import string
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn

FONT = sys.argv[1] if len(sys.argv) > 1 else "Aptos"
SIZE = int(sys.argv[2]) if len(sys.argv) > 2 else 12
OUTPUT = sys.argv[3] if len(sys.argv) > 3 else f"tests/fixtures/kern_mining/{FONT.lower()}/input.docx"

CHARS = list(string.ascii_lowercase + string.ascii_uppercase + string.digits + '.,;:!?-\'"()/@ ')

import os
os.makedirs(os.path.dirname(OUTPUT), exist_ok=True)

doc = Document()

# Set font on the Normal style so Word uses it as the default
style = doc.styles['Normal']
style.font.name = FONT
style.font.size = Pt(SIZE)

# Generate bigrams: one paragraph per left character
for left in CHARS:
    words = []
    for right in CHARS:
        if left == ' ' and right == ' ':
            continue
        words.append(f"{left}{right}")
    text = " ".join(words)
    doc.add_paragraph(text)

doc.save(OUTPUT)
print(f"Generated {OUTPUT}: {len(CHARS)} chars, {len(CHARS)*len(CHARS)} pairs, font={FONT} {SIZE}pt via Normal style")
