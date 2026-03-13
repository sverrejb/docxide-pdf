# Fix: "November 2021" on wrong page in USEP handbook

## Context

In the `usep_handbook` fixture, "November 2021" renders on page 1 in our output but should be on page 2 (matching Word's reference). The text positions for everything UP TO "Consultant's Handbook" are nearly identical between generated and reference (within 1pt), confirming the content heights are correct. The issue is specifically about the page-break decision for "November 2021".

## Root Cause

The `keepNext` chain evaluation in the render loop (`src/pdf/mod.rs:1633-1659`) doesn't account for `page_break_after` on paragraphs that continue the chain.

**Document structure:**
- Para 22: "November 2021" — Heading1 style → `keepNext=true`
- Para 23: page break paragraph — Heading1 style → `keepNext=true` + `page_break_after=true`
- Para 24: empty Normal → `keepNext=false` (chain terminates)

**Current behavior:** When evaluating keepNext for para 22:
1. Walks to para 23, checks `page_break_before` (false) — not triggered
2. Para 23 has `keep_next = true`, chain continues
3. Walks to para 24 (no keepNext), calculates ~43pt of extra height
4. Total (needed + extra) ≈ 63pt, plenty of room on page → stays on page 1

**Word's behavior:** Word recognizes that a `page_break_after` in the middle of a keepNext chain means the chain can't be satisfied on one page (content after the break goes to the next page). So it moves the chain start ("November 2021") to a new page, giving:
- Page 1: content through empty paragraphs after "Consultant's Handbook"
- Page 2: only "November 2021" (+ the page-break paragraph)
- Page 3: "Welcome to the USEP Team!" and subsequent content

## Fix

**File:** `src/pdf/mod.rs` — keepNext chain evaluation (~line 1651)

After the `!next.keep_next` check (which terminates the chain normally), add a check: if the chain-continuing paragraph also has `page_break_after`, set `extra = f32::MAX` to force a page break.

```rust
// Current code at ~line 1646:
if !next.keep_next {
    let next_ls = next.line_spacing.unwrap_or(ctx.doc_line_spacing);
    let next_line_h = resolve_line_h(next_ls, nfs, nlhr);
    extra += next_inter + next_first_line_h + next_line_h;
    break;
}
// ADD after the above block, before `extra += next_inter + next_first_line_h;`:
if next.page_break_after {
    extra = f32::MAX;
    break;
}
extra += next_inter + next_first_line_h;
```

The check is placed AFTER `!next.keep_next` so that `page_break_after` on the chain **terminator** (no keepNext) is NOT affected — that case is handled normally after rendering. It only fires when a chain-continuing paragraph (keepNext=true) also breaks the page, making it impossible to keep the chain together.

## Verification

1. `cargo run -- tests/fixtures/scraped/usep_handbook/input.docx /tmp/usep_test.pdf`
2. `mutool draw -F text -o - /tmp/usep_test.pdf 1-3` — confirm:
   - Page 1: title page content, NO "November 2021"
   - Page 2: "November 2021" only
   - Page 3: "Welcome to the USEP Team!" and body content
3. `cargo test -- --nocapture` — check for REGRESSION lines (no regressions expected since this fixture is currently skipped)
4. Remove `usep_handbook` from `tests/fixtures/SKIPLIST` if scores improve sufficiently
