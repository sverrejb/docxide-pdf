# numeric-slash-suffix

Synthetic DOCX inputs with immutable Microsoft Word PDF references. No private documents.

The group/case layout is required by `tests/common::discover_fixtures`; case names are at most 16 characters for readable, stable case identifiers.

| Case | Control parameters |
|---|---|
| `s0-c14-b-sm` | `numeric-slash-case00-mode14-both-over1.8` |
| `s0-c15-b-sm` | `numeric-slash-case04-mode15-both-over1.8` |
| `s0-c15-b-lg` | `numeric-slash-case05-mode15-both-over6.0` |
| `s0-c15-l-sm` | `numeric-slash-case06-mode15-left-over1.8` |
| `s1-c15-b-sm` | `numeric-slash-case12-mode15-both-over1.8` |
| `s2-c15-b-sm` | `numeric-slash-case20-mode15-both-over1.8` |

Run from the repository root with the Word-compatible fonts and `mutool` available:

```sh
DOCXIDE_GROUP=numeric-slash-suffix cargo test --test visual_comparison --test page_geometry --test text_boundary -- --nocapture
```

Check that all cases appear in the output and `tests/output/latest_scores.json`; an empty test run is not validation. Do not regenerate references with the renderer under test.

## Validation

6 cases were actually discovered and measured with mutool 1.23.0. All reference/generated page counts agree; no case was skipped. Minimum SSIM: 0.8166. 272 library tests passed; the separate system-font-dependent `missing_fonts_resolve_as_word_does` test was excluded.

The new cases have explicit visual baselines in `tests/baselines.json`; existing baseline entries are unchanged. The original main text-boundary parser does not recognize MuPDF 1.23 character nodes; its zero-line scores are not accepted as baselines. Text lines are separately audited with the character-node parser fix. Fonts and renderer versions affect visual scores, so use the same Word-compatible fonts when reproducing them.
