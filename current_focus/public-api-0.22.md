# Public API 0.22: one `convert()`, one `Conversion` (handover)

Planned 2026-11-19 with the user. Status: **planning done, nothing coded.**
Main was at `475c1e95f` when this was written. Update the status table at the
bottom as each step lands.

## Goal

The public API grew one function per input × output × options combination,
and contributor PRs add per-thread side channels next to the conversion call:

| Today (0.21.x, main) | Problem |
|---|---|
| `convert_docx_to_pdf[_with]` (path → path) | 7 functions for one operation |
| `convert_docx_bytes_to_pdf[_with]` (bytes → path) | path → bytes is missing, and every new output adds another column |
| `convert_docx_bytes_to_pdf_bytes[_with]` (bytes → bytes) | |
| `layout_docx_bytes_with` (bytes → `display::Document`) | renders the PDF too, then throws it away |
| `take_bookmark_positions()` (#169, thread-local) | hidden state read after the call; reset in every entry point |
| `take_font_requests()` (#172, thread-local, not merged yet) | same |
| `Error::Zip(zip::ZipError)`, `Error::Xml(roxmltree::Error)` | a `zip` or `roxmltree` major bump is a breaking release for us; `Error` is not `#[non_exhaustive]` |
| `reject_old_env_vars()` inside the `_with` functions | "explicit options never read the environment" is not quite true; #170 has to skip it on wasm |
| `RevisionView` not `#[non_exhaustive]` | Word also has Simple Markup and Original |
| `ConvertOptions::debug_wrap`, `Error::RenamedEnvVar` | a developer knob and a leftover from the `DOCXSIDE_` rename, both public |

Target for 0.22 (pre-1.0, so a breaking release is allowed):

```rust
pub fn convert(docx: &[u8], options: &ConvertOptions) -> Result<Conversion, Error>;

#[non_exhaustive]
pub struct Conversion {
    pub pdf: Vec<u8>,
    pub bookmarks: Vec<BookmarkPosition>,      // from #169
    pub font_requests: Vec<FontRequest>,       // from #172
    pub display: Option<display::Document>,    // from #168/#175, only when the options ask for it
}
```

The public surface afterwards: `convert`, `ConvertOptions`, `Conversion`,
`Error`, `RevisionView`, `BookmarkPosition`, `FontRequest` and the `display`
module. The caller does file I/O with `std::fs`:

```rust
let out = convert(&std::fs::read("in.docx")?, &ConvertOptions::from_env())?;
std::fs::write("out.pdf", &out.pdf)?;
```

## Decisions (user, 2026-11-19)

1. **Conversion struct route.** One render returns everything. Not a field
   on `display::Document` alone, because a caller wanting PDF + positions
   would then convert twice.
2. **The 7 old functions are removed in 0.22.** They are not deprecated
   first: pre-1.0 clean break.
3. **These go into the same 0.22 break:**
   - `Error`: hide the dependency error types and add `#[non_exhaustive]`.
   - No environment read in `convert()`: the old-`DOCXSIDE_` check moves
     into `ConvertOptions::from_env()`.
   - `RevisionView` becomes `#[non_exhaustive]`.
   - Drop `debug_wrap` from the public options, and drop `Error::RenamedEnvVar`.
4. **Contributors get credit: merge their PRs first, as they are, then
   refactor in our own commits.** Their commits stay in history under their
   names, GitHub shows the PRs as merged, and the release notes thank them.
5. #169 is already merged (`a23175b1b`). `take_bookmark_positions` is **not
   in any release**: crates.io has 0.21.4, bumped 2026-11-16, before #169.
   `BookmarkPosition` stays without `#[non_exhaustive]` (user's choice
   during the #169 review); revisit it only if the user asks.

## Order of operations

Do these one at a time, in this order. The PR review protocol is in the
memory file `pr-review-policy.md`. In short:
- run a security audit first; anything outside `src/` and `tests/` needs the
  user's approval; no new dependencies;
- use a worktree `.worktrees/pr-N` on branch `review/pr-N` with the `fonts`
  symlink;
- for refactor-type PRs, run the byte-identical check (all fixtures, 4
  variants), then `./tools/run-tests.sh --test visual_comparison`;
- run the `simplify` skill on the source diff, in a separate commit;
- push with `git push origin review/pr-N:main` from the worktree, only after
  the user's OK.

**During steps 1–4, skip API-shape polish.** Order changed 2026-11-19 (user): the wasm stack first, then #175. Step 5 replaces those APIs, so
judge each PR on security, correctness and byte-identical output only.

### Step 1: #170 (banduk), "Convert on bare wasm32-unknown-unknown"

- The base of the wasm stack (#170 → #171 → #172).
- Reviewed 2026-11-19 on `review/pr-170`. User decisions:
  - **No stand-in clock.** The PR's `src/clock.rs` (an `Instant` that always
    says zero) is gone; the timing statements and their log lines carry
    `#[cfg(not(all(target_arch = "wasm32", target_os = "unknown")))]`.
  - Kept: the CI workflow `.github/workflows/host.yml`, the smoke test
    (`tools/check-wasm-host.sh`, `.mjs`, `tools/wasm_host_smoke.rs`,
    `tests/host/`), the README section and the CLAUDE.md command line.
- The smoke script now builds the host as a Cargo `[[example]]` (cdylib):
  the PR's hand-written `rustc` call broke on Rust 1.100's new target-dir
  layout. Checked on Rust 1.99 and 1.100.
- Measured on bare wasm: `std::env::var` returns nothing, but
  `std::env::vars_os` **panics**, and so does `Instant::now`. File access
  returns errors, no panic.

### Step 2: #171 (banduk), "Register fonts in memory per thread"

- Stacked on #170: review only the top commit, `1d70662eb`.
- Public: `register_font(Vec<u8>) -> usize`, `clear_fonts()`,
  `use_system_fonts(bool)`, all thread-local; new `src/fonts/memory.rs`;
  `tests/suite/host.rs`.
- This is an *input* side channel. See the open question below. Merge it as
  is.
- Built on #170's original commit, so expect conflicts: it still imports
  `crate::clock::Instant` (in context lines) and its CLAUDE.md line mentions
  `clock.rs`. Both are gone after the #170 review; use `std::time::Instant`
  under the bare-wasm `#[cfg(not(...))]` wherever it adds timing. #172 has
  the same (`fonts/resolve.rs` `t0`).

### Step 3: #172 (banduk), "Report the font lookups a conversion missed"

- Stacked on #171: review only the top commit, `45ca9312a`.
- Public: `FontRequest`, `take_font_requests()` (thread-local); new
  `src/fonts/requests.rs`.
- Known conflict: `reset_conversion_metadata()` in `src/lib.rs` must call
  both `take_bookmark_positions()` and `take_font_requests()`. The README
  and CLAUDE.md lists also conflict trivially.
- The PR says 1243/1245 fixtures are byte-identical. case30 and charts_ex
  differ in the last float digit from release codegen partitioning;
  `CARGO_PROFILE_RELEASE_CODEGEN_UNITS=1` makes them identical. Verify that
  claim; don't take it on trust.

### Step 4: #175 (sboily), "Display list: marked content, structure tree and outline"

- Independent of the rest; adds `Item::BeginMarked`/`EndMarked`,
  `Document::structure`, `Document::outline`, `Info::pdf_ua` and `StructType`.
- Files: `src/` only (display.rs, pdf/{assembly, canvas, display_list,
  footnotes, link_tagger, mod, page_builder, table_layout, tagging,
  textbox_render}.rs).
- Already checked 2026-11-19: it merges cleanly onto main with #169. On a
  scratch branch, #169's 7 tests and #175's 5 display tests pass. #169's
  marker links are stripped before links are tagged, so #175's link indices
  are unaffected.
- It touches `tagging.rs`, so run `--test accessibility` too (policy rule 9).
- Known side issue, not #175's fault: `OutlineEntry.top` (from
  `heading_entries`' `y_position`) ignores section vertical alignment, while
  `BookmarkPosition.top` includes it. Our PDF's outline destinations have the
  same offset. It is a separate fix: apply `valign_offsets` to
  `heading_entries` the way `bookmarks::store` does.

### Step 5: our 0.22 refactor (separate commits, one concern each)

Work in a worktree. Suggested commits:

1. **`convert()` + `Conversion`.**
   - Internally, keep the thread-locals the contributors wrote. `convert()`
     clears them, renders, then drains bookmarks and font requests into
     `Conversion` before returning, so no state survives the call.
   - Ceiling: no nested conversion on the same thread (nothing does that
     today).
   - The display list: an option, name to pick (e.g.
     `ConvertOptions::display(true)`), that starts `canvas::Recording`
     around the same render. One render then gives both PDF and display list.
2. **Remove the 7 old functions, `take_bookmark_positions()`,
   `take_font_requests()` and `reset_conversion_metadata()`.**
   - Also delete `docx::parse(path)` (the path-reading parser) and
     `render_and_write` if nothing else uses them.
   - Make `pdf::bookmarks` and `fonts::requests` crate-private again.
3. **Environment.**
   - `convert()` reads no environment.
   - The old-`DOCXSIDE_` check moves into `ConvertOptions::from_env()`. It
     can no longer return an error from there, so pick one: `from_env`
     returns `Result`, or it panics, or it logs. **Ask the user.**
   - The check uses `std::env::vars_os`, which panics on bare wasm: keep
     #170's `cfg!` guard when it moves into `from_env`.
4. **`Error`.**
   - Add `#[non_exhaustive]`.
   - Replace `Zip(ZipError)` and `Xml(roxmltree::Error)` with variants that
     don't expose the crate types (e.g. fold them into `InvalidDocx` with the
     message, or keep the source as a boxed `dyn Error` behind `source()`).
   - Drop `RenamedEnvVar`.
   - `Io` is only produced by file reading and writing (`docx/mod.rs` opens
     the input path; `render_and_write` writes). With no library file I/O
     left, check whether `Io` is still reachable and drop it if not.
5. **`RevisionView`: `#[non_exhaustive]`.**
6. **`debug_wrap`.**
   - Remove it from the public `ConvertOptions` builder/fields.
   - The CLI flag `--debug-wrap` and `DOCXIDE_DEBUG_WRAP` must keep working,
     because the case browser and wrap debugging use them. `src/main.rs` is
     a separate crate and cannot see `pub(crate)`.
   - Options: a `#[doc(hidden)] pub` field/method, or only `from_env()` sets
     it and the CLI exports the variable. Pick the smaller one.
7. **Callers.**
   - In this repo:
     - `src/main.rs` (CLI);
     - `tests/common/mod.rs` (the harness's conversion);
     - `tests/suite/bookmark_positions.rs` and `tests/suite/host.rs`: rewrite
       `take_*()` into `out.bookmarks` / `out.font_requests` and keep every
       assertion;
     - the README "Usage" and "Bookmark positions" sections, and #171/#172's
       README sections;
     - the CLAUDE.md architecture line for `lib.rs`. It describes the three
       `convert_*` functions and the `_with` convention; rewrite it, and
       state the new rule: a new output becomes a `Conversion` field, never
       a new function or side channel.
   - Other repos (ask the user before touching them): docxide-template's
     README example and the docxide-demo site's conversion call.
8. **Version bump to 0.22.0, with release notes.** The release notes list
   the API change with a before/after snippet and thank banduk (#169–#172)
   and sboily (#168, #175) by PR.

**Checks for step 5:**
- The byte-identical check against the main commit before the refactor, on
  all fixtures in the default, markup, debug-wrap and no-resample variants.
  The memory file `byte-identical-refactor-check.md` has the recipe; use the
  release CLI and `DOCXIDE_NO_FONT_CACHE=1`.
- The visual suite.
- `cargo test --lib`.
- `cargo clippy --all-targets`.
- The wasm smoke check from #170, if the user kept it.

## Rules while the steps are in flight

- **No crate release between step 1 and the end of step 5.** Main briefly
  carries three thread-local APIs. docxide-demo auto-deploys on a new crate
  version.
- **Tell the contributors early.** One short comment on each PR: "merging
  as is; the API becomes `convert() -> Conversion` in 0.22". That keeps them
  from building more on the `take_*` functions. Comments are outward-facing:
  **draft them and get the user's OK before posting.** The user chose not to
  comment on #169.
- **Other sessions may be running.** Check `git status` and
  `pgrep -fl "claude|cargo"` first, and never switch branches in the main
  checkout.

## Open questions (decide before step 5; none blocks steps 1–4)

1. **Font input (#171).** Keep the per-thread `register_font()` registry, or
   pass fonts in `ConvertOptions`?
   - Options would fit "explicit options never read hidden state".
   - But wasm hosts convert many documents with the same fonts, so options
     would need a shared, pre-parsed `FontSet` (e.g. `Arc`) to avoid
     re-parsing faces on every call.
   - Ask the user. The lowest-effort version: `ConvertOptions::fonts(&FontSet)`
     installs the existing thread-local registry for the duration of
     `convert()`.
2. **`from_env()` and the old-variable error** (see step 5.3).
3. **The name of the display option** and whether `Conversion::display`
   always carries the outline/structure from #175 (it does, once the
   recording runs).

## Status

| Step | State |
|---|---|
| #169 bookmark positions | merged `a23175b1b` (2026-11-19) |
| 1. #170 bare wasm | reviewed on `review/pr-170`, awaiting push |
| 2. #171 memory fonts | not started |
| 3. #172 font requests | not started |
| 4. #175 display tags | not started |
| 5. 0.22 refactor | not started |
| Release 0.22.0 | not started |
