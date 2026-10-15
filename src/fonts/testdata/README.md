# Synthetic non-breaking hyphen fonts

These tiny fonts contain only generated geometric outlines; no system or third-party font data is used. They are test fixtures under the library license.

Regenerate with `scripts/make_nb_hyphen_fonts.py` in stamp2docx (fontTools 4.67.0). In the upstream checkout the generator is `tools/make_nb_hyphen_fonts.py`.

- `ascii`: only U+002D, advance 400.
- `unicode`: U+002D and U+2010, advance 500 for U+2010.
- `native`: also U+2011, advance 700, to detect accidental replacement.
- `none`: none of these hyphens, an unsupported-character negative control.

All widths use 1000 units/em. Regeneration uses fixed font timestamps.
