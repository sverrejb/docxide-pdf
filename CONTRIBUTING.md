# Contributing

Pull requests are welcome. This file is about the shape of a pull request;
`CLAUDE.md` has the commands, architecture, testing rules and Word layout
notes. Agents read both.

Most rules below come from what the maintainer changed on the way in: the
`Simplify PR #N` commits that follow a merge, and the merge comments. A PR
that already has that shape needs no follow-up, and gets merged sooner.

## Scope

- One concern per PR. A series is a stack of PRs that each build and pass
  on their own (#169 → #170 → #171 → #172); a stacked PR says which commit
  to review.
- `src/` and `tests/` only, where possible. Anything else (CI workflows,
  scripts in `tools/`, a new language, README and `CLAUDE.md` edits) is the
  maintainer's call and may be left out (#171, #172). Keep those edits small
  and in their own commit, so they can be dropped.
- No new dependencies, no version bumps, no release changes.
- Base on current `main` and rebase; don't merge `main` in.

## Code: the simplest form that writes the same bytes

The follow-up commits remove lines (#168: 75 net, #178: 60 net) without
moving output. Look for these before you open the PR:

- **No machinery for a case that never happens.** An `Option`, not a stack,
  when calls never nest (#172); no save and restore of state that nothing
  re-enters (#175); `entry().or_default()`, not a guard for an unreachable
  branch (#178).
- **No stand-ins for one platform.** #170's always-zero clock became
  `#[cfg(...)]` on the timing lines; an empty value comes from
  `Default::default()`, not a second static (#170, #171).
- **One helper, not copies.** One function for one concept, called at every
  site (#169's `bookmark_names`, #175's heading table). A `_noting`/`_with`
  variant left with one caller folds into the function (#178). Test helpers
  live once in `tests/common` (`convert_bytes_with`, `pdf_contains`,
  `pdf_streams`), not copied into each suite (#171, #177).
- **An optimization must pay for its code.** #178 counted letters per
  stretch of one language; a `FastMap` per letter is as fast and plain. A
  buffer declared once above the loop replaces take, reset and put back.
- **Arguments, not ambient state.** In #177 the picture policy (fast mode,
  resample threshold) is passed through the embedder like `resample`; only
  `compression::zlib` reads the per-thread setting, scoped once in
  `convert()` next to the font registry.
- **Idioms.** `LocalKey::set` and `with_borrow_mut` over
  `with(|c| c.borrow_mut())`; `Default::default()` where the type is known;
  imports in rustfmt order; `cargo fmt --check` clean (main is formatted
  since 0.22).
- **Comments say why**, with the number when there is one (#177: "a page at
  200% zoom on a 2x display shows ~288 ppi"). `// ponytail:` marks a known
  limit.

## Public API and CLI

The rules in `CLAUDE.md` (Architecture, `lib.rs`) hold; in practice:

- A new output is a `Conversion` field; a new option is a private field and
  a builder method. No thread-local side channel in the public API (0.22
  folded #169, #171 and #172's `take_*` functions into `convert()`). New
  public types are `#[non_exhaustive]`.
- Every option is a CLI flag whose environment variable is read by clap:
  `#[arg(long, env = "DOCXIDE_…")]`, with `FalseyValueParser` for booleans
  and `value_delimiter` for lists. It then shows in `--help` and invalid
  values are rejected. Never `std::env::var` in `main.rs` (#177's
  hand-read `DOCXIDE_COMPRESSION=fast` became `DOCXIDE_FAST_COMPRESSION`).
  The variable is named after the flag.
- An option is documented as one row of the README's Configuration table
  (#177), not a new section.

## Finish the feature

An option should do everything its name promises, and keep only what
measures. #177's `Fast` grew from a zlib level into stored JPEGs, cheaper
resampling above 384 ppi and no trial JPEG encode; skipping object streams
was tried and dropped (no measurable gain, files 21% larger). Before opening:
what else would a user of this option expect, and is each part measured?

## Proof in the description

The maintainer re-checks every claim, so give each one with the command that
reproduces it:

- **Refactors and performance work:** PDFs byte-identical to `main` on all
  fixtures, in the variants default, markup (`--markup`), debug-wrap,
  no-resample and fast compression. Convert every
  `tests/fixtures/*/*/input.docx` with the release CLI from both commits
  (`DOCXIDE_NO_FONT_CACHE=1`) and compare; if floats differ in the last
  digit, retry with `CARGO_PROFILE_RELEASE_CODEGEN_UNITS=1`. Also the
  display list, bookmark positions and missing fonts, on a debug build (no
  overflow panic).
- **Layout changes:** the visual suite, with improvements accepted and any
  drop shown (see Testing in `CLAUDE.md`).
- **Performance:** CPU time over all fixtures, runs interleaved, as a table
  (main vs the PR, with total size), and the documents that gain most.
- **Features:** tests that need no installed fonts (a test font from
  `src/fonts/testdata`, the DOCX built in the test). Touching hosts or wasm:
  `--test host` and `./tools/check-wasm-host.sh`; touching `tagging.rs`:
  `--test accessibility`.
- `cargo clippy --all-targets`: no warning the base does not have.

## Commits and description

- Subject: an imperative sentence, capitalized, no `type(scope):` prefix
  and no PR number ("Add an opt-in fast PDF compression mode"). Body: what
  changed, as bullets, then the proof ("PDFs byte-identical on all 1,246
  fixtures in 5 variants").
- Every commit builds and passes; callers move in the same commit as the
  API they use.
- The description says what and why, the public surface it adds, its
  limits, and the validation with numbers.
- The maintainer merges a PR as it is and refactors in follow-up commits,
  so authorship stays in history. Read those commits after a merge: the next PR
  should not need them.
