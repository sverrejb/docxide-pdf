# Annotations

One JSON file per annotation, `<group>/<case>/<timestamp>-<github login>.json`, opened as a PR by
the comparison page on the docxide site (docxide-demo repo). Notes sent without signing in are
`<timestamp>-anonymous-<4 hex>.json` with `"author": "anonymous"`. Coordinates are PDF points like the case browser's
`tests/output/annotations.json`: `x_pt` from the left, `y_pt` from the bottom, `page` 0-based,
`source` `Reference` (Word's PDF) or `Generated` (docxide-pdf). `generated` is the build the note was made
against. Set `"fixed": true` when resolved; the site then greys the marker out.

```json
{
  "case": "cases/case41",
  "page": 0,
  "source": "Generated",
  "x_pt": 123.4,
  "y_pt": 567.8,
  "note": "The second bullet sits one line too low.",
  "author": "octocat",
  "created": "2026-10-31T14:22:33Z",
  "generated": "0.20.1 @be95d87b",
  "fixed": false
}
```
