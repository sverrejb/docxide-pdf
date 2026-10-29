//! Accessibility metrics, always relative to the Word reference: veraPDF
//! PDF/UA-1 rule parity plus a comparison of the structure trees as Poppler
//! reads them (`pdfinfo -struct-text`). `coverage` alone needs no reference:
//! it checks the DOCX's own text against our structure tree.

use rayon::prelude::*;
use serde::{Deserialize, Serialize};
use std::collections::{BTreeMap, HashMap};
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::Mutex;

/// Pseudo element kind for a content item's text, kept in the element list so
/// text stays interleaved with inline children in reading order.
pub const TEXT: &str = "#text";

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct Elem {
    pub depth: usize,
    pub kind: String,
    pub inline: bool,
    pub alt: bool,
    pub text: String,
}

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct Rule {
    pub failed: u64,
    pub passed: u64,
    pub description: String,
}

#[derive(Clone, Debug, Default, Serialize, Deserialize)]
pub struct Analysis {
    /// Structure tree in pre-order; empty for an untagged PDF.
    pub elems: Vec<Elem>,
    /// PDF/UA-1 rule id ("7.1-3") → check counts.
    pub rules: BTreeMap<String, Rule>,
}

pub fn tools_available() -> bool {
    Command::new("pdfinfo").arg("-v").output().is_ok()
        && Command::new("verapdf").arg("--version").output().is_ok()
}

/// How a PDF fares against Word's reference: what tests/accessibility.rs
/// scores and the engine comparison shows.
pub struct Scores {
    /// PDF/UA-1 rules the PDF fails on its own.
    pub ua_fail: usize,
    /// The PDF claims PDF/UA-1 (rule 5-1 passes).
    pub claims_ua: bool,
    /// None for an untagged (macOS print-path) reference: no Word bar to measure against.
    pub vs_word: Option<VsWord>,
}

pub struct VsWord {
    pub deficit: Vec<String>,
    pub struct_score: f64,
    pub text_score: f64,
}

pub fn scores(reference: &Analysis, generated: &Analysis) -> Scores {
    Scores {
        ua_fail: generated.rules.values().filter(|r| r.failed > 0).count(),
        claims_ua: !generated.rules.get("5-1").is_some_and(|r| r.failed > 0),
        vs_word: (!reference.elems.is_empty()).then(|| VsWord {
            deficit: ua_deficit(&reference.rules, &generated.rules)
                .into_iter()
                .map(String::from)
                .collect(),
            struct_score: struct_score(&reference.elems, &generated.elems),
            text_score: text_score(&reference.elems, &generated.elems),
        }),
    }
}

/// The `analyze_cached` cache for `<dir>/<name>.pdf`, given `<dir>/<name>`.
pub fn cache_path(stem: &Path) -> PathBuf {
    stem.with_extension("a11y.json")
}

/// True while `cache` is newer than `pdf`: what `analyze_cached` reuses.
pub fn cache_fresh(pdf: &Path, cache: &Path) -> bool {
    matches!((super::metrics::mtime(cache), super::metrics::mtime(pdf)), (Some(c), Some(p)) if c >= p)
}

/// Analyse `pdf`, reusing `cache` while it is newer than the PDF.
// ponytail: cache ignores the veraPDF/Poppler version; delete tests/output/**/*.a11y.json after upgrading them
pub fn analyze_cached(pdf: &Path, cache: &Path) -> Result<Analysis, String> {
    if cache_fresh(pdf, cache)
        && let Some(a) = fs::read_to_string(cache)
            .ok()
            .and_then(|s| serde_json::from_str(&s).ok())
    {
        return Ok(a);
    }
    // veraPDF first, so a machine without it fails before the pdfinfo work.
    let rules = ua_rules(pdf)?;
    let analysis = Analysis {
        elems: struct_elems(pdf)?,
        rules,
    };
    if let Ok(json) = serde_json::to_string(&analysis) {
        fs::write(cache, json).ok();
    }
    Ok(analysis)
}

fn struct_elems(pdf: &Path) -> Result<Vec<Elem>, String> {
    let out = Command::new("pdfinfo")
        .arg("-struct-text")
        .arg(pdf)
        .output()
        .map_err(|e| format!("pdfinfo: {e}"))?;
    if !out.status.success() {
        return Err(format!(
            "pdfinfo: {}",
            String::from_utf8_lossy(&out.stderr).trim()
        ));
    }
    Ok(parse_struct_text(&String::from_utf8_lossy(&out.stdout)))
}

/// Parse `pdfinfo -struct-text`: two spaces of indent per level, element lines
/// like `Note <Note 1> (inline)` or `Figure ["alt"]:`, attribute lines starting
/// with `/`, annotation references `Object 11 0`, and quoted text lines.
pub fn parse_struct_text(dump: &str) -> Vec<Elem> {
    let mut elems = Vec::new();
    for line in dump.lines() {
        let body = line.trim_start();
        let depth = (line.len() - body.len()) / 2;
        if body.is_empty() || body.starts_with('/') || body.starts_with("Object ") {
            continue;
        }
        if let Some(quoted) = body.strip_prefix('"') {
            elems.push(Elem {
                depth,
                kind: TEXT.into(),
                inline: true,
                alt: false,
                text: quoted.strip_suffix('"').unwrap_or(quoted).to_string(),
            });
            continue;
        }
        let kind: String = body
            .chars()
            .take_while(|c| !matches!(c, ' ' | ':' | '[' | '<' | '('))
            .collect();
        let alt = match (body.find("[\""), body.rfind("\"]")) {
            (Some(a), Some(b)) if b > a + 1 => !body[a + 2..b].trim().is_empty(),
            _ => false,
        };
        elems.push(Elem {
            depth,
            kind,
            inline: body.contains("(inline)"),
            alt,
            text: String::new(),
        });
    }
    elems
}

/// PDFs per verapdf process. Starting its JVM costs about six files' worth of
/// validation, and with `--success` the report runs to megabytes per file, so
/// batches stay small enough to parse comfortably.
const BATCH: usize = 24;

type Rules = BTreeMap<String, Rule>;
/// One verapdf job: the file as verapdf names it (absolute), and its rules or the error that stopped it.
type Job = (PathBuf, Result<Rules, String>);

/// Results of `prefetch_ua_rules` by canonical path, each taken once by `ua_rules`.
static PREFETCHED: Mutex<BTreeMap<PathBuf, Result<Rules, String>>> = Mutex::new(BTreeMap::new());

/// Validate `pdfs` a batch per verapdf process, for `analyze_cached` to pick up
/// later. One process per PDF would make the JVM start the suite's main cost.
pub fn prefetch_ua_rules(pdfs: &[PathBuf]) {
    let pdfs: Vec<PathBuf> = pdfs
        .iter()
        .filter_map(|p| fs::canonicalize(p).ok())
        .collect();
    let results: Vec<_> = pdfs
        .par_chunks(BATCH)
        .filter_map(|batch| run_verapdf(batch).ok())
        .flatten()
        .collect();
    PREFETCHED.lock().unwrap().extend(results);
}

fn ua_rules(pdf: &Path) -> Result<Rules, String> {
    if let Some(prefetched) = fs::canonicalize(pdf)
        .ok()
        .and_then(|p| PREFETCHED.lock().unwrap().remove(&p))
    {
        return prefetched;
    }
    run_verapdf(&[pdf.to_path_buf()])?
        .pop()
        .ok_or("verapdf: no job in report")?
        .1
}

/// One verapdf run over `pdfs`: each file's rule map, or why it has none.
fn run_verapdf(pdfs: &[PathBuf]) -> Result<Vec<Job>, String> {
    // Forced to ua1: Word writes no pdfuaid, so auto-detection would validate
    // PDF/A-1b. --success lists passing rules too, with per-rule check counts.
    let out = Command::new("verapdf")
        .args([
            "-f",
            "ua1",
            "--format",
            "json",
            "--success",
            "--maxfailuresdisplayed",
            "0",
        ])
        .args(pdfs)
        .output()
        .map_err(|e| format!("verapdf: {e}"))?;
    // The exit code only summarises the jobs (1: non-compliant, 7: one failed to
    // parse); the report says which file, so it is what gets checked.
    parse_verapdf_report(&String::from_utf8_lossy(&out.stdout))
        .map_err(|e| format!("{e} (exit {:?})", out.status.code()))
}

/// Every job of a verapdf JSON report; rule ids are "7.1-3" style.
pub fn parse_verapdf_report(json: &str) -> Result<Vec<Job>, String> {
    let v: serde_json::Value =
        serde_json::from_str(json).map_err(|e| format!("verapdf json: {e}"))?;
    let jobs = v["report"]["jobs"]
        .as_array()
        .ok_or("verapdf json: no jobs")?;
    Ok(jobs
        .iter()
        .map(|job| {
            let path = PathBuf::from(job["itemDetails"]["name"].as_str().unwrap_or(""));
            let rules = match job["validationResult"][0]["details"]["ruleSummaries"].as_array() {
                Some(summaries) => Ok(summaries.iter().map(rule_summary).collect()),
                None => Err(format!(
                    "verapdf: {}",
                    job["taskResult"]["exceptionMessage"]
                        .as_str()
                        .unwrap_or("no validation result")
                )),
            };
            (path, rules)
        })
        .collect())
}

fn rule_summary(r: &serde_json::Value) -> (String, Rule) {
    let id = format!(
        "{}-{}",
        r["clause"].as_str().unwrap_or("?"),
        r["testNumber"]
    );
    let rule = Rule {
        failed: r["failedChecks"].as_u64().unwrap_or(0),
        passed: r["passedChecks"].as_u64().unwrap_or(0),
        description: r["description"].as_str().unwrap_or("").to_string(),
    };
    (id, rule)
}

/// Rules where we do worse than Word: failed where Word passes, or failing a
/// larger share of the rule's checks than Word does. The share test matters for
/// rules like 7.1-3 (untagged content) that Word fails on one stray item while
/// an untagged PDF fails them everywhere.
pub fn ua_deficit<'a>(
    reference: &BTreeMap<String, Rule>,
    generated: &'a BTreeMap<String, Rule>,
) -> Vec<&'a str> {
    let share = |r: &Rule| r.failed as f64 / (r.failed + r.passed).max(1) as f64;
    generated
        .iter()
        .filter(|(id, g)| {
            g.failed > 0
                && reference
                    .get(*id)
                    .is_none_or(|r| r.failed == 0 || share(g) > share(r) + 0.01)
        })
        .map(|(id, _)| id.as_str())
        .collect()
}

/// Element types in reading order, Span dropped (Word splits them per
/// formatting run) and a Figure's alt text presence folded into its token.
fn struct_tokens(elems: &[Elem]) -> Vec<&str> {
    elems
        .iter()
        .filter(|e| e.kind != TEXT && e.kind != "Span")
        .map(|e| {
            if e.kind == "Figure" && e.alt {
                "Figure+alt"
            } else {
                e.kind.as_str()
            }
        })
        .collect()
}

fn levenshtein<T: PartialEq>(a: &[T], b: &[T]) -> usize {
    let mut prev: Vec<usize> = (0..=b.len()).collect();
    for x in a {
        let mut cur = vec![prev[0] + 1];
        for (j, y) in b.iter().enumerate() {
            cur.push(
                (prev[j] + usize::from(x != y))
                    .min(prev[j + 1] + 1)
                    .min(cur[j] + 1),
            );
        }
        prev = cur;
    }
    prev[b.len()]
}

/// 1 − edit distance between the two element-type sequences, relative to the longer.
pub fn struct_score(reference: &[Elem], generated: &[Elem]) -> f64 {
    let (a, b) = (struct_tokens(reference), struct_tokens(generated));
    let n = a.len().max(b.len());
    if n == 0 {
        return 1.0;
    }
    1.0 - levenshtein(&a, &b) as f64 / n as f64
}

/// Symbol glyphs (checkboxes, arrows, dingbats, symbol-font private-use codes)
/// all count as one character. Word extracts the same Wingdings checkbox as raw
/// U+F0A8 in one document and as Unicode in the next, and which symbol it is
/// isn't what this score measures. Math operators (U+2200–22FF) and • stay
/// distinct.
fn fold_symbol(c: char) -> char {
    match c as u32 {
        0xF000..=0xF0FF // symbol-font private use
        | 0x2190..=0x21FF // arrows
        | 0x2300..=0x23FF // misc technical
        | 0x25A0..=0x27FF // geometric shapes, misc symbols, dingbats, misc math symbols-A, arrows-A
        | 0x2900..=0x297F // arrows-B
        | 0x2B00..=0x2BFF // misc symbols and arrows
        | 0x1F300..=0x1FAFF => '\u{FFFC}', // pictographs, emoji, arrows-C
        _ => c,
    }
}

/// Whitespace- and symbol-normalised (`fold_symbol`) text per block element in
/// reading order; inline content (Span, Link, Note, text) folds into its
/// nearest block ancestor.
pub fn block_texts(elems: &[Elem]) -> Vec<String> {
    let mut blocks: Vec<String> = Vec::new();
    let mut open: Vec<(usize, usize)> = Vec::new(); // (depth, block index) of enclosing blocks
    for e in elems {
        while open.last().is_some_and(|&(d, _)| d >= e.depth) {
            open.pop();
        }
        if e.kind == TEXT {
            match open.last() {
                Some(&(_, i)) => blocks[i].push_str(&e.text),
                None => blocks.push(e.text.clone()),
            }
        } else if !e.inline {
            open.push((e.depth, blocks.len()));
            blocks.push(String::new());
        }
    }
    blocks
        .into_iter()
        .map(|b| b.split_whitespace().collect::<Vec<_>>().join(" "))
        .map(|b| b.chars().map(fold_symbol).collect::<String>())
        .filter(|b| !b.is_empty())
        .collect()
}

/// Characters in blocks that match exactly and in the same order (LCS over
/// blocks, weighted by length), relative to the longer text. A reordered,
/// missing, extra or word-merged paragraph costs its whole length.
pub fn text_score(reference: &[Elem], generated: &[Elem]) -> f64 {
    let (a, b) = (block_texts(reference), block_texts(generated));
    let total = |v: &[String]| v.iter().map(|s| s.chars().count()).sum::<usize>();
    let n = total(&a).max(total(&b));
    if n == 0 {
        return 1.0;
    }
    let mut prev = vec![0usize; b.len() + 1];
    for x in &a {
        let w = x.chars().count();
        let mut cur = vec![0usize; b.len() + 1];
        for (j, y) in b.iter().enumerate() {
            cur[j + 1] = if x == y {
                prev[j] + w
            } else {
                prev[j + 1].max(cur[j])
            };
        }
        prev = cur;
    }
    prev[b.len()] as f64 / n as f64
}

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const MC_NS: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/// How much of the DOCX's text (body, footnotes, endnotes) reaches the
/// structure tree, which is what a screen reader gets.
pub struct Coverage {
    /// Letters and digits that never do (case-folded, counted with repeats).
    pub missing: usize,
    /// Paragraphs whose text isn't there whole, when anything is missing.
    pub lost: Vec<String>,
}

/// Text hidden by a style and textbox overflow Word clips too still count as
/// missing; the baselines absorb them, as `ua_fail` does source-limited rules.
pub fn coverage(docx: &Path, generated: &[Elem]) -> Result<Coverage, String> {
    let file = fs::File::open(docx).map_err(|e| format!("{}: {e}", docx.display()))?;
    let mut zip = zip::ZipArchive::new(file).map_err(|e| format!("zip: {e}"))?;
    let mut paras = Vec::new();
    for part in [
        "word/document.xml",
        "word/footnotes.xml",
        "word/endnotes.xml",
    ] {
        let Ok(entry) = zip.by_name(part) else {
            continue;
        };
        let xml = std::io::read_to_string(entry).map_err(|e| format!("{part}: {e}"))?;
        paras.extend(docx_paragraphs(&xml)?);
    }
    let got: String = generated
        .iter()
        .filter(|e| e.kind == TEXT)
        .map(|e| letters(&e.text))
        .collect();
    let count = |s: &str| {
        let mut n: HashMap<char, usize> = HashMap::new();
        s.chars().for_each(|c| *n.entry(c).or_default() += 1);
        n
    };
    let have = count(&got);
    let missing = count(&letters(&paras.concat()))
        .iter()
        .map(|(c, n)| n.saturating_sub(have.get(c).copied().unwrap_or(0)))
        .sum();
    // Paragraphs come in reading order: look after the last one found first.
    let mut from = 0;
    let lost = paras
        .into_iter()
        .filter(|p| {
            let own = letters(p);
            if missing == 0 || own.chars().count() <= 3 {
                return false;
            }
            match got[from..]
                .find(&own)
                .map(|i| from + i)
                .or_else(|| got.find(&own))
            {
                Some(at) => {
                    from = at + own.len();
                    false
                }
                None => true,
            }
        })
        .collect();
    Ok(Coverage { missing, lost })
}

/// Letters and digits only, case-folded: what coverage compares.
fn letters(text: &str) -> String {
    text.chars()
        .filter(|c| c.is_alphanumeric())
        .flat_map(char::to_lowercase)
        .collect()
}

/// A WordprocessingML part's paragraph texts as a reader should get them: no
/// deleted, hidden or field-code text, no cached result of a field the
/// renderer recomputes, no `mc:Fallback` copies, no note separators; a
/// textbox's paragraphs follow the one that anchors it.
pub fn docx_paragraphs(xml: &str) -> Result<Vec<String>, String> {
    let doc = roxmltree::Document::parse(xml).map_err(|e| format!("xml: {e}"))?;
    let mut out = Vec::new();
    collect_paragraphs(doc.root_element(), None, &mut out, &mut Vec::new());
    Ok(out)
}

/// Fields whose result the renderer computes at layout time (the keywords
/// `parse_field_code` in src/docx/runs.rs knows), so the result cached in the
/// DOCX is not what the page shows nor what the reader should get: Word
/// updates them on export too, and a stale body STYLEREF still reads
/// "Heading pending" in the XML while the page says the real heading.
fn recomputed_field(instr: &str) -> bool {
    instr.split_whitespace().next().is_some_and(|kw| {
        ["PAGE", "NUMPAGES", "STYLEREF", "PAGEREF"]
            .iter()
            .any(|f| kw.eq_ignore_ascii_case(f))
    })
}

/// `fields` holds the open complex fields in document order: each one's
/// instruction and whether its separator has passed, after which its runs
/// are the cached result. The walk is in document order, so a field that
/// spans paragraphs stays open across them.
fn collect_paragraphs(
    node: roxmltree::Node,
    mut para: Option<usize>,
    out: &mut Vec<String>,
    fields: &mut Vec<(String, bool)>,
) {
    let separator = || {
        matches!(
            node.attribute((W_NS, "type")),
            Some("separator" | "continuationSeparator" | "continuationNotice")
        )
    };
    match (node.tag_name().namespace(), node.tag_name().name()) {
        (Some(MC_NS), "Fallback") | (Some(W_NS), "del") => return,
        (Some(W_NS), "footnote" | "endnote") if separator() => return,
        (Some(W_NS), "r") if hidden_run(node) => return,
        (Some(W_NS), "fldSimple")
            if recomputed_field(node.attribute((W_NS, "instr")).unwrap_or("")) =>
        {
            return;
        }
        (Some(W_NS), "fldChar") => match node.attribute((W_NS, "fldCharType")) {
            Some("begin") => fields.push((String::new(), false)),
            Some("separate") => {
                if let Some(field) = fields.last_mut() {
                    field.1 = true;
                }
            }
            Some("end") => {
                fields.pop();
            }
            _ => {}
        },
        (Some(W_NS), "instrText") => {
            if let (Some(field), Some(t)) = (fields.last_mut(), node.text()) {
                field.0.push_str(t);
            }
        }
        (Some(W_NS), "p") => {
            out.push(String::new());
            para = Some(out.len() - 1);
        }
        (Some(W_NS), "t") => {
            let cached_result = fields
                .iter()
                .any(|(instr, in_result)| *in_result && recomputed_field(instr));
            if let (Some(i), Some(t)) = (para, node.text())
                && !cached_result
            {
                out[i].push_str(t);
            }
        }
        _ => {}
    }
    for child in node.children().filter(|n| n.is_element()) {
        collect_paragraphs(child, para, out, fields);
    }
}

fn hidden_run(run: roxmltree::Node) -> bool {
    run.children()
        .find(|n| n.has_tag_name((W_NS, "rPr")))
        .and_then(|p| p.children().find(|n| n.has_tag_name((W_NS, "vanish"))))
        .is_some_and(|v| !matches!(v.attribute((W_NS, "val")), Some("0" | "false" | "off")))
}

#[cfg(test)]
mod tests {
    use super::docx_paragraphs;

    fn wml(body: &str) -> String {
        format!(
            r#"<w:document xmlns:w="{}"><w:body>{body}</w:body></w:document>"#,
            super::W_NS
        )
    }

    fn field(instr: &str, result: &str) -> String {
        format!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> {instr} </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{result}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
        )
    }

    #[test]
    fn recomputed_field_results_are_not_docx_text() {
        // The fixture cases/styleref-body-field: a stale STYLEREF result.
        let stale = field(r#"STYLEREF "Example Heading""#, "Heading pending");
        let xml = wml(&format!(
            "<w:p><w:r><w:t>Body heading: </w:t></w:r>{stale}<w:r><w:t> after</w:t></w:r></w:p>"
        ));
        assert_eq!(docx_paragraphs(&xml).unwrap(), ["Body heading:  after"]);

        // A field the renderer never evaluates keeps its cached result; nesting
        // and the simple form follow the same rule.
        let date = field("DATE", "29 Oct 2026");
        let page = field("PAGE", "7");
        let xml = wml(&format!(
            r#"<w:p>{date}<w:r><w:t> p. </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>9</w:t></w:r></w:fldSimple></w:p><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> REF x </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>see </w:t></w:r>{page}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>"#
        ));
        assert_eq!(docx_paragraphs(&xml).unwrap(), ["29 Oct 2026 p. ", "see "]);
    }
}
