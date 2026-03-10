# Codebase Structure Improvement Plan

## Phase 1: Quantitative Survey

### Module size & complexity scan
- Count lines per file to find oversized modules (candidates for splitting)
- Count `pub`/`pub(super)`/`pub(crate)` items per module to assess API surface area
- Identify the largest functions (line count) — long functions often signal mixed responsibilities

### Dependency graph
- Map `use`/`mod` relationships between modules to see coupling
- Identify circular or surprising dependencies (e.g., does `pdf/` ever reach back into `docx/`?)
- Check which types from `model.rs` each module touches — this reveals the IR's role as a coupling point

## Phase 2: Qualitative Code Reading

### Read the hot paths — focus on the modules the quantitative scan flagged:
- `src/docx/mod.rs` — likely the largest file given it's the "parse orchestrator" plus table/paragraph parsing
- `src/pdf/layout.rs` — text layout is inherently complex; check if it mixes concerns
- `src/model.rs` — the IR is the bridge between parsing and rendering; assess if it's the right abstraction level
- `src/fonts.rs` — font handling touches everything; check boundaries

### Read the boundaries:
- `src/lib.rs` — the public API surface. Is it minimal and clean?
- How errors flow through the system (`error.rs` -> callers)

## Phase 3: Pattern Detection

### Look for recurring code smells:
- Grep for `TODO`, `FIXME`, `HACK`, `workaround` — the codebase's own self-assessment
- Look for long match/if-else chains that could be trait-dispatched
- Find duplicated logic between modules (e.g., XML attribute parsing patterns repeated across `docx/` submodules)
- Check for "god structs" — types with many fields that get threaded through everything

### Assess the IR (`model.rs`):
- Is it a clean intermediate representation, or does it leak DOCX-specific or PDF-specific concerns?
- Are there types that exist only to shuttle data between two specific modules?

## Phase 4: Test & Tooling Review

- Read `tests/visual_comparison.rs` and `tests/text_boundary.rs` — are tests well-organized or monolithic?
- Check if the `tools/` workspace duplicates any logic from the main crate
- Look at fixture organization — any structural issues there?

## Phase 5: Synthesis

With the above data collected, categorize findings into:
1. **Split candidates** — modules doing too many things
2. **Merge candidates** — tiny modules that are always used together
3. **Abstraction opportunities** — repeated patterns that could be unified
4. **API boundary issues** — leaky abstractions, over-exposed internals
5. **IR improvements** — model.rs changes that would simplify both sides
