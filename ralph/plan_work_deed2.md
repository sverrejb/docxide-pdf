# Plan: Implement startOverride for transition_to_work_deed

## Context

The `transition_to_work_deed` fixture scores 25.6% Jaccard / 37.6% SSIM (target: 70% SSIM).
After the numStyleLink fix (Fix 1), numbering labels now appear but **never restart** between clauses.

**Root cause**: The document has **585 `w:startOverride`** elements across **69 of 151 numId instances**,
all completely ignored by our parser. Each legal clause uses a different numId with startOverride
to restart sub-level counters at 1. Without this, counters accumulate across the entire document:

- Page 10: labels show `(ag)`, `(ah)` instead of `(a)`, `(b)`
- Page 50: labels show `(ct)`, `(cu)` instead of `(a)`, `(b)`

This causes: wrong label text, wider labels (2-3 chars vs 1), different line wrapping,
cascading page break drift. By page 40 the content is visually desynchronized.
By page 80 it's completely different content on the same page number.
Generated: 148 pages vs reference: 153 pages.

## Fix: Parse and apply startOverride

**File**: `src/docx/numbering.rs`

### Step 1: Add `start_overrides` to NumberingInfo (line 16-19) ✅ COMPLETED

Add a new field to `NumberingInfo`:
```rust
pub(super) struct NumberingInfo {
    pub(super) abstract_nums: HashMap<String, HashMap<u8, LevelDef>>,
    pub(super) num_to_abstract: HashMap<String, String>,
    pub(super) start_overrides: HashMap<String, HashMap<u8, u32>>,  // NEW: numId → (ilvl → start)
}
```

Update the `#[derive(Default)]` — HashMap derives Default, so this just works.

### Step 2: Parse lvlOverride/startOverride in `parse_numbering()` (line 94-102) ✅ COMPLETED

In the `"num"` match arm, after the existing `num_to_abstract.insert(...)`, add:
```rust
let mut overrides: HashMap<u8, u32> = HashMap::new();
for ovr in node.children() {
    if ovr.tag_name().name() == "lvlOverride" && ovr.tag_name().namespace() == Some(WML_NS) {
        if let Some(ilvl) = ovr.attribute((WML_NS, "ilvl")).and_then(|v| v.parse::<u8>().ok()) {
            if let Some(start_ovr) = wml(ovr, "startOverride") {
                if let Some(val) = start_ovr.attribute((WML_NS, "val")).and_then(|v| v.parse::<u32>().ok()) {
                    overrides.insert(ilvl, val);
                }
            }
        }
    }
}
if !overrides.is_empty() {
    start_overrides.insert(num_id.to_string(), overrides);
}
```

Note: `start_overrides` is a new local variable initialized alongside `num_style_link` etc. at line 32-33.

### Step 3: Return start_overrides from parse_numbering() (line 121-124) ✅ COMPLETED

```rust
NumberingInfo {
    abstract_nums,
    num_to_abstract,
    start_overrides,
}
```

### Step 4: Use start_overrides in `parse_list_info()` (line 254-259) ✅ COMPLETED

Change the counter initialization to use the override value when available:
```rust
let start = numbering.start_overrides
    .get(num_id)
    .and_then(|m| m.get(&ilvl))
    .copied()
    .unwrap_or(def.start);
let current_counter = *counters
    .entry((num_id.to_string(), ilvl))
    .and_modify(|c| *c += 1)
    .or_insert(start);
```

The existing code at line 255-259:
```rust
let start = def.start;
let current_counter = *counters
    .entry((num_id.to_string(), ilvl))
    .and_modify(|c| *c += 1)
    .or_insert(start);
```

### Why this works

The XML structure is:
```xml
<w:num w:numId="24">
  <w:abstractNumId w:val="29"/>
  <w:lvlOverride w:ilvl="0"/>            <!-- no override for level 0 -->
  <w:lvlOverride w:ilvl="1">
    <w:startOverride w:val="1"/>          <!-- restart level 1 at 1 -->
  </w:lvlOverride>
  <!-- ... levels 2-8 also restart at 1 -->
</w:num>
```

Each clause gets a different numId (69 of 151 have overrides). Since counters are keyed
by `(num_id, ilvl)`, different numIds have independent counters. The startOverride just sets
the INITIAL counter value for each (numId, ilvl) pair.

## Verification

1. `cargo check` — compiles
2. `DOCXIDE_CASE=transition_to_work_deed cargo test -- --nocapture` — check improved scores
3. Text verification:
   - `mutool draw -F text` page 10: should show `(a)`, `(b)` instead of `(ag)`, `(ah)`
   - `mutool draw -F text` page 50: should show `(a)`, `(b)` instead of `(ct)`, `(cu)`
4. Page count: should be closer to 153 (currently 148)
5. `cargo test -- --nocapture` — full suite, check for "REGRESSION in:" lines
6. `./tools/target/debug/analyze-fixtures` — score overview
