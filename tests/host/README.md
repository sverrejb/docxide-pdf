# Host smoke fixture

`input.docx` is a synthetic, uncompressed ZIP containing only
`word/document.xml`. It has one body paragraph and one single-cell table, each
containing a hyphen in `NB Hyphen Fixture native`. The paragraphs carry the
bookmarks `wasm_body` and `wasm_cell`, respectively. It names the geometric font
fixture in `src/fonts/testdata/nb-hyphen-native.ttf`; no third-party document or
font data is included. This input needs no Word reference PDF.

`tools/check-wasm-host.sh` embeds this input and the font in its wasm module,
then verifies actual conversion with the fallback and with the registered font.
