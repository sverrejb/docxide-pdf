#!/bin/bash
set -e

if [ -z "$1" ]; then
  echo "Usage: $0 <iterations>"
  exit 1
fi

iterations="$1"
progress="./ralph/marge_progress.md"

if [ ! -f "$progress" ]; then
  cat > "$progress" <<'EOF'
# Progress for Marge — OOXML spec coverage audit

Marge walks the OOXML spec (ECMA-376 Part 1 / ISO/IEC 29500-1) one chapter per
iteration, checking whether docxside-pdf covers each chapter's features.

## Scope & master checklist

Walk top to bottom; do the first `[ ]` item each run; mark `[x]` when audited.
When every box is `[x]`, the audit is complete.

**Clause 17 — WordprocessingML** (core: `src/docx/*`, `src/pdf/*`)
- [ ] 17.2 Main Document Story
- [ ] 17.3 Paragraphs and Rich Formatting
- [ ] 17.4 Tables
- [ ] 17.5 Custom Markup
- [ ] 17.6 Sections
- [ ] 17.7 Styles
- [ ] 17.8 Fonts
- [ ] 17.9 Numbering
- [ ] 17.10 Headers and Footers
- [ ] 17.11 Footnotes and Endnotes
- [ ] 17.12 Glossary Document
- [ ] 17.13 Annotations (comments, revisions / tracked changes)
- [ ] 17.14 Mail Merge
- [ ] 17.15 Settings
- [ ] 17.16 Fields and Hyperlinks
- [ ] 17.18 Simple Types (ST_* value domains — reference only, low audit value)

**Clause 20 — DrawingML Framework** (`src/geometry/*`, `docx/textbox.rs`, `docx/images.rs`)
- [ ] 20.1 DrawingML Main (shapes, presetGeometry, fills, effects, text body — `a:`)
- [ ] 20.2 DrawingML Picture (`pic:` picture container)
- [ ] 20.4 DrawingML WordprocessingML Drawing (`wp:` inline/anchor positioning, wrap)

**Clause 21 — DrawingML Components** (`docx/charts.rs`, `docx/smartart.rs`, `pdf/charts*.rs`, `pdf/smartart.rs`)
- [ ] 21.2 DrawingML Charts (`c:` chartSpace)
- [ ] 21.4 DrawingML Diagrams (`dgm:`/`dsp:` SmartArt)

**Clause 22 — Shared MLs**
- [ ] 22.1 Math (OMML `m:` namespace)
- [ ] 22.6 Bibliography (`b:` citation sources)

**Out of scope** (no print impact / not in a DOCX): clause 18 SpreadsheetML,
clause 19 PresentationML, 20.3 Locked Canvas, 20.5 SpreadsheetML Drawing,
21.1/21.3 chart-internal drawings, 22.2–22.5 / 22.7–22.9 (doc properties,
variant types, custom-XML props, additional characteristics, simple-type
domains). VML (legacy `v:` drawings) is ECMA-376 Part 4, not this spec PDF.

## Chapters processed
(none yet)

## Coverage gaps found
(none yet)
EOF
fi

for ((i=1; i<=iterations; i++)); do
    echo "====================="
    echo "Iteration $i / $iterations — $(date '+%Y-%m-%d %H:%M:%S')"
    echo "====================="

    result=$(claude --model claude-opus-4-8 --permission-mode acceptEdits -p "@${progress} \
  You are auditing docxside-pdf for coverage of the OOXML spec (ECMA-376 Part 1 / ISO/IEC 29500-1), ONE chapter per run. Steps:
1. Read ${progress}. The spec is ingested in the local-rag MCP — query it via mcp__local-rag__query_documents (\`/Users/sverrejb/specs/office.pdf\`). \
2. In the 'Scope & master checklist' section, find the FIRST item still marked '[ ]' (top to bottom). That is your chapter for this run — work on ONE and ONLY ONE. If every item is already '[x]', the audit is finished: output <promise>COMPLETE</promise> and do nothing else. \
3. Enumerate that chapter's features/elements from the spec. For each, check whether docxside-pdf handles it — search @src for the relevant XML element/attribute names. Use the analysis tools (docx-inspect, analyze-fixtures --grep) rather than ad-hoc bash. \
4. Record findings in ${progress}: change that checklist item from '[ ]' to '[x]', append the chapter under 'Chapters processed' with a timestamp, and list any unhandled or partially-handled features under 'Coverage gaps found'. This file is the ONLY deliverable — write each gap so a fresh agent with no prior context can pick it up and implement it: spec element/attribute name, spec section reference, what is missing vs. what exists today, and which src/ file would likely own the fix. \
5. DO NOT implement, fix, or change any code, and DO NOT commit anything — this is an audit only. The progress file is the work product; another agent will act on it later. \
6. If and ONLY IF every checklist item is now '[x]', output <promise>COMPLETE</promise>. Otherwise end the run without that marker.")

  echo "$result"

  if [[ "$result" == *"<promise>COMPLETE</promise>"* ]]; then
    echo "Marge audit complete — all chapters covered after $i iterations."
    exit 0
  fi
done
