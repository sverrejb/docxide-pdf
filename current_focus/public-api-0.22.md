# Public API 0.22: one `convert()`, one `Conversion` (handover)

Planned 2026-11-19 with the user; public API reviewed with the user
2026-11-20 (decisions 6–15, none open). Status: **step 5.1–5.8 coded
2026-11-21 on branch `api-0.22` (not pushed); 5.9 and the release open.**
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
| `take_font_requests()` (#172, thread-local) | same |
| `Error::Zip(zip::ZipError)`, `Error::Xml(roxmltree::Error)` | a `zip` or `roxmltree` major bump is a breaking release for us; `Error` is not `#[non_exhaustive]` |
| `reject_old_env_vars()` inside the `_with` functions | "explicit options never read the environment" is not quite true; #170 has to skip it on wasm |
| `RevisionView` not `#[non_exhaustive]` | Word also has Simple Markup and Original |
| `ConvertOptions::debug_wrap`, `Error::RenamedEnvVar` | a developer knob and a leftover from the `DOCXSIDE_` rename, both public |

Target for 0.22 (pre-1.0, so a breaking release is allowed). **After 0.22
we try to avoid breaking changes**, so every public item below is shaped so
that later additions are additive (user, 2026-11-20).

```rust
// The only function.
pub fn convert(docx: &[u8], options: &ConvertOptions) -> Result<Conversion, Error>;

#[non_exhaustive]
#[derive(Debug, Clone)]
pub struct Conversion {
    pub pdf: Vec<u8>,
    pub bookmarks: Vec<BookmarkPosition>,      // from #169
    pub missing_fonts: Vec<MissingFont>,       // from #172 (its take_font_requests); decision 13
    pub display: Option<display::Document>,    // from #168/#175; Some only with .display_list(true)
}

#[non_exhaustive]
#[derive(Clone, Debug, Default)]
pub struct ConvertOptions { /* all fields private */ }

impl ConvertOptions {
    pub fn new() -> Self;                       // Word's online export
    pub fn revisions(self, RevisionView) -> Self;
    pub fn resample_pictures(self, bool) -> Self;
    pub fn font_dirs(self, impl IntoIterator<Item = impl Into<PathBuf>>) -> Self;
    pub fn font(self, data: Vec<u8>) -> Self;   // replaces register_font/clear_fonts (#171)
    pub fn system_fonts(self, bool) -> Self;    // replaces use_system_fonts (#171)
    pub fn font_cache(self, bool) -> Self;
    pub fn display_list(self, bool) -> Self;
    #[doc(hidden)] pub fn debug_wrap(self, bool) -> Self;
}

#[non_exhaustive]
pub enum RevisionView { Final, Markup }

#[non_exhaustive]
pub enum Error { InvalidDocx(String) }          // Zip, Xml, Pdf, Io, RenamedEnvVar gone

#[non_exhaustive]
pub struct BookmarkPosition { .. }
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MissingFont { pub name: String, pub bold: bool, pub italic: bool }
pub mod display;                                // + #175's structure/outline types
```

Gone: the 7 `convert_*`/`layout_*` functions, `ConvertOptions::from_env`,
the public option fields, `take_bookmark_positions`, `take_font_requests`,
`register_font`, `clear_fonts`, `use_system_fonts`. **The library reads no
environment variables**; the `DOCXIDE_*` variables become a CLI feature. The
caller does file I/O with `std::fs`:

```rust
let out = convert(&std::fs::read("in.docx")?, &ConvertOptions::new())?;
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
   - No environment read in `convert()` (superseded by decision 6: no
     environment read in the library at all).
   - `RevisionView` becomes `#[non_exhaustive]`.
   - Drop `debug_wrap` from the public options, and drop `Error::RenamedEnvVar`.
4. **Contributors get credit: merge their PRs first, as they are, then
   refactor in our own commits.** Their commits stay in history under their
   names, GitHub shows the PRs as merged, and the release notes thank them.
5. #169 is already merged (`a23175b1b`). `take_bookmark_positions` is **not
   in any release**: crates.io has 0.21.4, bumped 2026-11-16, before #169.
   `BookmarkPosition` stays without `#[non_exhaustive]` (user's choice
   during the #169 review); reversed 2026-11-20, see decision 11.

## Decisions (user, 2026-11-20, public API review)

6. **Drop `ConvertOptions::from_env()`.** The library reads no environment.
   The CLI reads the `DOCXIDE_*` variables through clap's `env` attribute;
   the test harness reads `DOCXIDE_FONTS` with `std::env::split_paths`. The
   old-`DOCXSIDE_` check is deleted, not moved (the rename shipped
   2026-10-15). This removes old open question 2.
7. **Option fields become private; builders only.** A public field locks its
   representation (e.g. `font_dirs: Vec<PathBuf>` could never become a font
   source enum). Nothing in this repo writes the fields directly.
8. **Fonts in the options: `ConvertOptions::font(Vec<u8>)`** and
   `ConvertOptions::system_fonts(bool)`. Faces are parsed once when the
   option is built and kept behind an `Arc`, so reusing one
   `&ConvertOptions` for many conversions re-parses nothing and `clone` is
   cheap. No `FontSet` type, no per-thread state: today a native server
   that registers on one thread and converts on another (tokio
   `spawn_blocking`) silently gets no fonts. Unusable data is skipped with a
   log line; a `fonts_loaded()` getter can come later, additively. This
   removes old open question 1.
9. **`Conversion::missing_fonts`**, not `font_requests` (it is an output).
10. **`display_list(bool)`** names the display option. Removes old open
    question 3.
11. **`#[non_exhaustive]` on `BookmarkPosition` and `MissingFont`.** After
    0.22 a new field (an x position, a font weight) must not break callers.
12. **`cli` stays a default feature.** The README's usage section documents
    the opt-out for library users:
    `docxide-pdf = { version = "0.22", default-features = false }`, which
    skips clap and env_logger.
13. **`missing_fonts` is feedback for the end user**, e.g. a web app that
    says "this document uses Aptos Bold, which isn't available: upload it
    or change the font". Each entry is a font the document asks for, with
    the requested style; type renamed `FontRequest` → `MissingFont`
    (`name`, `bold`, `italic`). Rules:
    - reported: a family and style the document asks for that no embedded,
      `.font()` or system face provides; once per (name, bold, italic);
    - reported: a missing style even when synthetic bold/italic covers it
      (uploading the real face changes the output);
    - not reported: our fallback chain, the symbol/CJK/math rescue probes,
      fonts the DOCX embeds.
    - The face used instead (`replaced_by`) can be added later as a field
      without a break; not now.
    - Considered and rejected for 0.22: a parse-only "fonts this document
      uses" call before converting (saves the second render for a host that
      downloads fonts on demand). No stated user; it would be additive.
14. **`display` stays an ordinary public module** (user, 2026-11-21: "this
    crate is a moving target"). No `unstable-display` feature.
15. **No `Conversion::save(path)` in 0.22** (recommended 2026-11-20; the
    user did not object). `std::fs::write(path, &out.pdf)?` is the same
    length; a save would bring file I/O back into the library and could only
    fail on bare wasm. Adding it later is non-breaking.

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
- This is an *input* side channel; decision 8 replaces it. Merge it as
  is.
- Reviewed 2026-11-20 on `review/pr-171`. User decisions: keep the CI step
  (`./tools/run-tests.sh --test host`) and the smoke-host font check; leave
  out its README and CLAUDE.md edits.
- **Fake paths (user: replace in step 5, not before).** A registered font
  travels as a made-up `PathBuf` `memory:N`. Every place that turns a font
  path into bytes asks the registry first: `probe_face` (discovery.rs) and
  `memory::read_font` in `fonts/mod.rs` `font_file`, `fonts/resolve.rs`
  `try_font` and `pdf/wordart.rs` `load_font_data`, plus
  `pdf/display_list.rs` (where the fake path leaked into the public
  `display::FontSource::File { path: "memory:0" }`; fixed in the review with
  a test, `host_fonts_reach_the_display_list_as_bytes`).
- Also from the review: `use_system_fonts(false)` disables the caller's own
  `ConvertOptions::font_dirs` too, silently.
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
- Reviewed 2026-11-20 on `review/pr-172`: README and CLAUDE.md edits left
  out (user); all 1245 fixtures byte-identical here in 4 variants, case30
  and charts_ex included.
- **Noisy list (user: decide in step 5, merged as is).** `note_miss` sits in
  `FontIndex::find_font_file`, so every lookup that misses is recorded:
  - each request lists the document's name and then every fallback that
    missed; on a host with no fonts every request ends with
    `LAST_RESORT_FONTS` (Arial, Liberation Sans, Arimo, Helvetica, DejaVu
    Sans);
  - lookups outside `register_font` (the symbol/CJK rescue probes via
    `covered_chars`, the math font via `font_file`) each become a one-name
    request;
  - a font the DOCX embeds is still requested (the index is checked first).

  Registering any of these would change output, so none is wrong, but a host
  cannot tell the font the document asked for from a fallback.

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
- Reviewed and merged 2026-11-20 (`bdf5fd103` merge, `c336d95f2` simplify).
  PDFs byte-identical on all fixtures, 4 variants. The display list drops
  empty artifacts that the PDF still writes (empty header `BDC…EMC`); they
  draw nothing. Left for step 5: the `tagged_links` position scan in
  `assembly.rs` (index could be recorded when the annotation is pushed) and
  `display_list::links` repeating the annotation loop's missing-bookmark
  filter.

### Step 5: our 0.22 refactor (separate commits, one concern each)

Work in a worktree. Every commit must build and pass the byte-identical
check, so callers move in the same commit as the API they use (reviewed
2026-11-20: the earlier order deleted functions six commits before their
callers moved). Suggested commits:

1. **`convert()` + `Conversion`, and move every in-repo caller to it.**
   - Internally, keep the thread-locals the contributors wrote. `convert()`
     clears them, renders, then drains bookmarks and font requests into
     `Conversion` before returning, so no state survives the call.
   - `// ponytail:` ceiling: no nested conversion on the same thread
     (nothing does that today).
   - `display_list(true)` starts `canvas::Recording` around the same render,
     so one render gives both PDF and display list.
   - Callers in this commit: `src/main.rs` (CLI), `tests/common/mod.rs`
     (`:276` path → path, `:552` bytes → bytes),
     `tests/suite/bookmark_positions.rs` and `tests/suite/host.rs` (rewrite
     `take_*()` into `out.bookmarks` / `out.missing_fonts`, keep every
     assertion), and `tools/wasm_host_smoke.rs`.
   - **What the library's file I/O did moves to the CLI and the harness:**
     the forced `.pdf` extension (`render_and_write`'s
     `with_extension("pdf")`), the input path in NotFound/PermissionDenied
     errors (`docx/mod.rs:1041`), and the timing log lines (keep the
     bare-wasm `cfg` on them).
   - Add a test: one fixture with `display_list(true)` and without gives
     identical `pdf` bytes ("one render gives both" depends on it).
   - Add a compile-time check that `ConvertOptions` and `Conversion` are
     `Send + Sync` (`fn _assert<T: Send + Sync>()`), so a later `Rc` cannot
     silently break server users.
   - `missing_fonts` reports what decision 13 says (see step 3, "Noisy
     list", for what #172 records today). Record per resolution in
     `fonts/resolve.rs` `register_font`, where the document's name and the
     fallback chain are known (about +15 lines there, about −35 in
     `requests.rs`, `note_miss` out of `find_font_file`); push only the
     document's name with the requested style, after the embedded fonts
     were checked. Rename `FontRequest` to `MissingFont`. Rewrite
     `tests/suite/host.rs`'s request tests to the new rules, and add one
     each: a fallback is not reported, an embedded font is not reported.
2. **Remove the 7 old functions, `take_bookmark_positions()`,
   `take_font_requests()` and `reset_conversion_metadata()`.**
   - Pure deletion now that commit 1 moved the callers.
   - Also delete `docx::parse(path)` (the path-reading parser) and
     `render_and_write`.
   - Make `pdf::bookmarks` and `fonts::requests` crate-private again.
3. **Options: private fields, no environment (decisions 6, 7).**
   - Make every `ConvertOptions` field private; delete `from_env`,
     `from_vars`, `reject_old_env_vars`, `old_env_var` and their unit tests
     in `lib.rs` (`:285–318`).
   - `debug_wrap`: `#[doc(hidden)] pub fn debug_wrap(self, bool)`. The
     alternative (the CLI sets the variable) needs `std::env::set_var`,
     which is `unsafe` in edition 2024.
   - CLI: enable clap's `env` feature (part of clap, not a new crate) and put
     `env = "DOCXIDE_…"` on each flag. Flags then override variables, which
     fixes the README's "a flag can't undo a variable" quirk. Two details:
     - `DOCXIDE_REVISIONS=markup` is not a boolean: either the flag becomes
       `--revisions <final|markup>` with `env`, or the CLI reads that one
       variable by hand. Pick the smaller one.
     - `DOCXIDE_FONTS` is `:`-separated (`;` on Windows); split it with
       `std::env::split_paths`, not a fixed `value_delimiter`.
     - clap reads `""`, `0`, `false`, `no`, `off` as off for a flag: a
       superset of today's rule, so no existing setting changes meaning.
   - Harness (`tests/common/mod.rs:274`): `ConvertOptions::new()` plus
     `font_dirs(std::env::split_paths(..))` from `DOCXIDE_FONTS`, so
     `.cargo/config.toml` keeps working. Check whether the harness relies on
     any other `DOCXIDE_*` option variable before deleting `from_env`.
   - The byte-identical check runs through the CLI, so its four variants
     (default, markup, debug-wrap, no-resample) prove the CLI still honours
     the variables.
4. **Fonts in the options (decision 8), then real bytes instead of fake
   paths.** Two commits.
   - 4a, API: `font(Vec<u8>)` parses into an `Arc`'d index kept in the
     options; `convert()` installs it as the thread's registry for the
     call and removes it afterwards. `system_fonts(bool)` replaces
     `use_system_fonts`, so it can no longer silently override `font_dirs`.
     Delete `register_font`, `clear_fonts`, `use_system_fonts`. Update
     `tests/suite/host.rs` and the wasm smoke host.
   - 4b, internal: registered fonts carry their bytes the way DOCX-embedded
     fonts already do (`FontEntry::embedded_data`, `Arc<Vec<u8>>`): the font
     lookup returns a source (path or bytes) instead of a path, and the
     `memory:N` paths, `memory::read_font`, the `probe_face` registry check
     and the `display_list.rs` special case go.
   - The measurements that take a path must take the source instead:
     `hhea_box` (called from `pdf/fonts.rs`) and `face_is_script_design`
     (`fonts/resolve.rs`). Today they work for registered fonts, because
     both go through `probe_face`, which asks the registry first (checked
     2026-11-21); once 4b removes the fake paths they would stop working.
     Check for others: grep `font_path` and `probe_face`.
   - 4b changes no public API. If it grows, it can ship in 0.22.x instead of
     holding up the release (ask the user then).
   - Proof: the host tests, the wasm smoke test, and the byte-identical
     check (no fixture uses registered fonts, so native output must not
     move). Add a test that a registered font and the same font in a
     directory give the same PDF.
5. **`Error`.**
   - Add `#[non_exhaustive]`.
   - Delete the `From<ZipError>`, `From<roxmltree::Error>` and
     `From<io::Error>` impls and let the compiler list the sites; each
     becomes `InvalidDocx(message)`. No boxed source: `#[non_exhaustive]`
     lets one be added later.
   - Delete `Pdf` (nothing constructs it, checked 2026-11-20), `Io` (only
     `docx::parse(path)` produced it; gone in commit 2) and `RenamedEnvVar`.
   - Expected result: `Error` has one variant, `InvalidDocx(String)`.
6. **`#[non_exhaustive]` on `RevisionView`, `BookmarkPosition` and
   `MissingFont`** (decision 11). `tests/suite/bookmark_positions.rs` only
   reads fields, so it is unaffected (checked 2026-11-20).
7. **Docs and other repos.**
   - README: "Usage", "Bookmark positions", #171/#172's sections, and the
     options table (the environment column becomes CLI-only), and the
     `default-features = false` opt-out for library users (decision 12).
   - CLAUDE.md: rewrite the `lib.rs` architecture line (three `convert_*`
     functions, `_with` convention, "three spellings" incl. `from_vars`).
     New rules: a new output becomes a `Conversion` field, never a new
     function or side channel; a new option is a builder method plus a CLI
     flag with `env = "DOCXIDE_…"`; the library never reads the environment.
     Also the "Environment variables (library)" heading becomes CLI.
   - Other repos (ask the user before touching them): docxide-template's
     README example and the docxide-demo site's conversion call
     (`convert_docx_bytes_to_pdf_bytes` in its `src/main.rs`; its
     `fly.toml`/`Dockerfile` set no `DOCXIDE_` variable, so dropping the
     environment read changes nothing there).
8. **Version bump to 0.22.0, with release notes.** The release notes list
   the API change with a before/after snippet, say that the library no
   longer reads `DOCXIDE_*` (pass options; the CLI still reads them), and
   thank banduk (#169–#172) and sboily (#168, #175) by PR.
9. **Whole-repo `cargo fmt`, at the very end (user, 2026-11-20).** The tree
   has drifted from rustfmt (`cargo fmt --check` flags files nobody
   touched, e.g. `pdf/comments.rs`, `pdf/layout.rs`, `pdf/paragraph.rs`).
   - The last commit before the release, after every other step: a
     formatting-only commit with nothing else in it, so review and
     `git blame --ignore-rev` can skip it.
   - Check that no other session has open branches first: they would
     conflict on every reformatted line.
   - Proof: `cargo fmt --check` clean, then the byte-identical check
     (formatting must not move a single byte of output).

**Checks for step 5:**
- The byte-identical check against the main commit before the refactor, on
  all fixtures in the default, markup, debug-wrap and no-resample variants.
  The memory file `byte-identical-refactor-check.md` has the recipe; use the
  release CLI and `DOCXIDE_NO_FONT_CACHE=1`.
- The visual suite.
- `./tools/run-tests.sh --test host`.
- `cargo test --lib`.
- `cargo clippy --all-targets`.
- The wasm smoke check, `./tools/check-wasm-host.sh` (kept in step 1).
- The display on/off byte test and the `Send + Sync` check from commit 1.

## Rules while the steps are in flight

- **No whole-tree `cargo fmt` until step 5.9.** Format only the lines you
  change (`rustfmt --edition 2024 --check` and read the diff for your files;
  revert the reformatting of anything else).

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

## Open questions

None. Every question from 2026-11-19 and 2026-11-20 is decided
(decisions 6–15). Ask the user about anything new that changes the public
surface.

## Step 5 notes (2026-11-21, for the next session)

Base `83d235da2`, rebased onto `4d72ce49b` (tools and docs only). Every commit: release CLI, all 1246 fixtures × 4
variants byte-identical to the base; clippy warnings identical to the base.
At the end: visual 1244 unchanged, accessibility 1242 unchanged, host 11,
bookmark_positions 7, `cargo test --lib` 376 (no failures, as at base),
wasm smoke passes.

- 5.1: `note_if_missing` is called from the two document loops of
  `collect_and_register_fonts`, not inside `register_font`, because the
  math font, the rescue fonts and the empty-document Helvetica go through
  `register_font` too. It checks the name `register_font` looks up
  (`word_substitute`: Helvetica → Arial) and skips Bahnschrift
  (`868a7c3bc`, its own commit so the user can drop it). Chart title font (`charts::TITLE_FONT`), theme minor font, line
  number and comment-pane fonts are reported when missing: Word draws
  them, so supplying them changes output.
- 5.1: file I/O moved out: the CLI forces `.pdf` before picking a free
  name (simplify pass; before, `-o out` checked `out` and wrote
  `out.pdf`, which could overwrite). The timing log line now covers parse
  + render (the library no longer times).
- 5.2: `reject_old_env_vars`/`old_env_var` went here, not in 5.3: with the
  old functions gone they were dead code.
- 5.3: clap's `env` on set-true flags needs `FalseyValueParser` (the
  default `BoolValueParser` rejects `1`). Off values now include `no`,
  `off`. A flag still can't switch a variable off (set-true flags); the
  README says so. `DOCXIDE_REVISIONS` and `DOCXIDE_FONTS` are read by hand
  in `main.rs`. The harness and `test_index` read `DOCXIDE_FONTS` and
  `DOCXIDE_NO_FONT_CACHE` (`""`, `0`, `false` = off).
- 5.4a: `ConvertOptions` keeps `Arc<MemoryFonts>`; `convert()` installs it
  as the thread's registry (`memory::install`, a drop guard).
  `register_font(..) -> usize` counted faces; `font()` has no count (a
  warning is logged), so the host test asserts "unusable bytes embed no
  font" instead. `FontSetup.system` replaces the early return of an empty
  index; a scan without system dirs does not save the disk cache (else it
  dropped the system entries).
- 5.4b: `FontEntry::font_path` → `font_source: Option<FontSource>`
  (`File(PathBuf)` | `Data(Arc<Vec<u8>>)`); `is_some()` still means
  "installed or supplied, not embedded" (hhea line box, Bahnschrift,
  WordArt, script check). Test: supplied and the same font in a directory
  give identical PDFs (regular and bold).
- 5.5: one site needed a message (`roxmltree::Document::parse` of
  `word/document.xml`).
- 5.8: no CHANGELOG or GitHub releases exist; the release notes are the
  body of `4fb6d14c8`.

Open after step 5:
- Other repos (need the user's OK): docxide-template README example,
  docxide-demo `src/main.rs` (`convert_docx_bytes_to_pdf_bytes`).
- `missing_fonts` edge: a Mac-only family with an altName is skipped by
  `register_font` but not reported.
- Supplied fonts still reach `FontIndex` through a thread-local; passing
  `&options.fonts` with the index (about 6 signatures) would delete
  `install`. `pdf::render`/`render_pass` return `Result` but never fail.

## Status

| Step | State |
|---|---|
| #169 bookmark positions | merged `a23175b1b` (2026-11-19) |
| 1. #170 bare wasm | merged 2026-11-20 |
| 2. #171 memory fonts | merged 2026-11-20 |
| 3. #172 font requests | merged 2026-11-20 |
| 4. #175 display tags | merged `c336d95f2` (2026-11-20) |
| 5.1 `convert()` + `Conversion` | `d2fa9a9d8` on `api-0.22` (2026-11-21) |
| 5.2 old functions removed | `f4aefc62a` |
| 5.3 private options, no environment | `87bd7a4a7` |
| 5.4a `font()`, `system_fonts()` | `618d24c98` |
| 5.4b bytes, no `memory:N` paths | `46f753360` |
| 5.5 `Error` | `2c069140c` |
| 5.6 `#[non_exhaustive]` | `89d6a2a99` |
| 5.7 README, CLAUDE.md (other repos not touched) | `99448c852` |
| 5.8 version 0.22.0, release notes in the commit message | `4fb6d14c8` |
| simplify pass | `9ae109c39` |
| `missing_fonts` names the substitute (Arial for Helvetica) | `868a7c3bc`, **user to keep or drop** (decision 13 said the document's name) |
| `display.rs` doc links | `d7d5c0eeb` |
| second simplify pass: no scan without dirs, test setup via options | `0ff9cf4e6` |
| push `api-0.22` to main | waiting for the user |
| 5.9 whole-repo `cargo fmt` | not started |
| Release 0.22.0 | not started |
