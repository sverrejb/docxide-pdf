# Inline table after a floating table

Synthetic Word 2013-compatibility layout: a two-line floating subject table,
an empty paragraph with 12 pt before/after, and a two-row inline table.
The inline table must start below the floating table when their rectangles
intersect. Paragraph wrapping can use the remaining side strip; this inline
table cannot, because it retains its alignment and width.

Word's `Following paragraph` top is 151.2855 pt from the page top. Before the
fix it is 146.9396 pt; after the fix it is 151.1088 pt (error 0.177 pt).
`reference.pdf` was exported by Microsoft Word 16.0.17932.20190.

Controls: `inline-clear-control` has no float; `inline-beside-float` right-aligns
a narrow inline table outside the float's exclusion bounds. Both retain their
pre-fix output. The latter already has an unrelated vertical discrepancy
(Word 135.0855 pt, converter 146.9396 pt); this fix does not address it.
No real document content, signatures, images, or metadata are included.
