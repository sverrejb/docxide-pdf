use crate::common;

use crate::common::a11y::{self, Analysis, VsWord};
use rayon::prelude::*;
use std::collections::{BTreeMap, HashMap};
use std::fs;
use std::path::{Path, PathBuf};

struct Scored {
    name: String,
    /// `ua_fail` counts the rules the source can't satisfy too (no title,
    /// pictures without descr); the baseline absorbs them.
    scores: a11y::Scores,
    /// Letters and digits of the DOCX that never reach the structure tree.
    missing: usize,
}

/// Which PDF in tests/output/<group>/<case>/ to score; e.g. `libreoffice.pdf`
/// scores LibreOffice's export with the same yardstick.
fn gen_pdf_name() -> String {
    std::env::var("DOCXIDE_A11Y_GEN").unwrap_or_else(|_| "generated.pdf".into())
}

fn detail_json(
    reference: &Analysis,
    generated: &Analysis,
    v: &VsWord,
    failing: &BTreeMap<&str, u64>,
) -> serde_json::Value {
    let rules: Vec<_> = v
        .deficit
        .iter()
        .map(|id| {
            let counts = |a: &Analysis| a.rules.get(id).map(|r| [r.failed, r.passed]);
            serde_json::json!({
                "rule": id,
                "description": generated.rules[id].description,
                "reference_failed_passed": counts(reference),
                "generated_failed_passed": counts(generated),
            })
        })
        .collect();
    serde_json::json!({
        "ua_fail": failing,
        "ua_deficit": rules,
        "a11y_struct": v.struct_score,
        "a11y_text": v.text_score,
    })
}

fn analyze_fixture(fixture: &Path, gen_name: &str) -> Option<Result<Scored, String>> {
    let reference = fixture.join("reference.pdf");
    if !fixture.join("input.docx").exists() || !reference.exists() {
        return None;
    }
    let name = common::case_key(fixture);
    let out = common::output_dir(fixture);
    fs::create_dir_all(&out).ok();
    let fail = |e: String| Some(Err(format!("{name}: {e}")));

    let ref_a = match a11y::analyze_cached(&reference, &a11y::cache_path(&out.join("reference"))) {
        Ok(a) => a,
        Err(e) => return fail(e),
    };

    let gen_pdf = if gen_name == "generated.pdf" {
        match common::ensure_generated_pdf(fixture) {
            Ok(p) => p,
            Err(e) => return fail(e),
        }
    } else {
        out.join(gen_name)
    };
    if !gen_pdf.exists() {
        return None;
    }
    let stem = gen_pdf.file_stem().unwrap().to_string_lossy().to_string();
    let gen_a = match a11y::analyze_cached(&gen_pdf, &a11y::cache_path(&out.join(&stem))) {
        Ok(a) => a,
        Err(e) => return fail(e),
    };
    let failing: BTreeMap<&str, u64> = gen_a
        .rules
        .iter()
        .filter(|(_, r)| r.failed > 0)
        .map(|(id, r)| (id.as_str(), r.failed))
        .collect();
    let scores = a11y::scores(&ref_a, &gen_a);
    let coverage = match a11y::coverage(&fixture.join("input.docx"), &gen_a.elems) {
        Ok(c) => c,
        Err(e) => return fail(e),
    };
    let mut detail = match &scores.vs_word {
        Some(v) => detail_json(&ref_a, &gen_a, v, &failing),
        None => serde_json::json!({ "ua_fail": failing }),
    };
    detail["a11y_missing"] = coverage.missing.into();
    detail["lost_paragraphs"] = coverage.lost.into();
    fs::write(
        out.join(format!("{stem}.deficit.json")),
        serde_json::to_string_pretty(&detail).unwrap(),
    )
    .ok();
    Some(Ok(Scored {
        name,
        scores,
        missing: coverage.missing,
    }))
}

/// The reference and generated PDFs of `fixture` whose analysis cache is stale.
fn pdfs_to_validate(fixture: &Path, gen_name: &str) -> Vec<PathBuf> {
    let reference = fixture.join("reference.pdf");
    if !fixture.join("input.docx").exists() || !reference.exists() {
        return Vec::new();
    }
    let out = common::output_dir(fixture);
    let generated = if gen_name == "generated.pdf" {
        common::ensure_generated_pdf(fixture).ok()
    } else {
        Some(out.join(gen_name))
    };
    [Some(reference), generated]
        .into_iter()
        .flatten()
        .filter(|pdf| {
            let cache = a11y::cache_path(&out.join(pdf.file_stem().unwrap()));
            pdf.exists() && !a11y::cache_fresh(pdf, &cache)
        })
        .collect()
}

#[test]
fn accessibility_vs_reference() {
    if !a11y::tools_available() {
        println!("SKIP accessibility: needs verapdf and pdfinfo (brew install verapdf poppler)");
        return;
    }
    let gen_name = gen_pdf_name();
    let fixtures = common::discover_fixtures().expect("Failed to read tests/fixtures");
    // One verapdf process per batch of PDFs, not per PDF: its JVM start would
    // otherwise be most of this test's time.
    let stale: Vec<PathBuf> = fixtures
        .par_iter()
        .flat_map_iter(|f| pdfs_to_validate(f, &gen_name))
        .collect();
    a11y::prefetch_ua_rules(&stale);
    let outcomes: Vec<Result<Scored, String>> = fixtures
        .par_iter()
        .filter_map(|f| analyze_fixture(f, &gen_name))
        .collect();

    let mut results = Vec::new();
    let mut errors = Vec::new();
    for o in outcomes {
        match o {
            Ok(s) => results.push(s),
            Err(e) => errors.push(e),
        }
    }
    results.sort_by(|a, b| a.name.cmp(&b.name));
    let compared: Vec<(&Scored, &VsWord)> = results
        .iter()
        .filter_map(|r| Some((r, r.scores.vs_word.as_ref()?)))
        .collect();

    let name_w = common::name_width(results.iter().map(|r| r.name.as_str()), 4);
    println!(
        "\n  {:<name_w$}  UaFail  UaDef  Struct    Text  Deficit rules ({gen_name})",
        "Case"
    );
    for (r, v) in &compared {
        let shown: Vec<&str> = v.deficit.iter().take(6).map(String::as_str).collect();
        let more = v.deficit.len().saturating_sub(shown.len());
        println!(
            "  {:<name_w$}  {:>6}  {:>5}  {:>5.1}%  {:>5.1}%  {}{}",
            r.name,
            r.scores.ua_fail,
            v.deficit.len(),
            v.struct_score * 100.0,
            v.text_score * 100.0,
            shown.join(" "),
            if more > 0 {
                format!(" +{more}")
            } else {
                String::new()
            }
        );
    }
    for r in results.iter().filter(|r| r.scores.vs_word.is_none()) {
        println!(
            "  {:<name_w$}  {:>6}      -       -       -  (untagged reference)",
            r.name, r.scores.ua_fail
        );
    }
    let n = compared.len().max(1) as f64;
    println!(
        "\n  a11y ≥ Word (no UA-1 deficit): {}/{} · struct {:.1}% · text {:.1}% (means) · N/A untagged reference: {}",
        compared
            .iter()
            .filter(|(_, v)| v.deficit.is_empty())
            .count(),
        compared.len(),
        compared.iter().map(|(_, v)| v.struct_score).sum::<f64>() / n * 100.0,
        compared.iter().map(|(_, v)| v.text_score).sum::<f64>() / n * 100.0,
        results.len() - compared.len(),
    );
    let claims: Vec<&Scored> = results.iter().filter(|r| r.scores.claims_ua).collect();
    println!(
        "  PDF/UA-1 rules failed on our own: {} total over {} PDFs · {} fail only one · {} claim PDF/UA-1",
        results.iter().map(|r| r.scores.ua_fail).sum::<usize>(),
        results.len(),
        results.iter().filter(|r| r.scores.ua_fail == 1).count(),
        claims.len(),
    );
    println!(
        "  DOCX text missing from the structure tree: {} letters and digits in {} PDFs (lost paragraphs in *.deficit.json)",
        results.iter().map(|r| r.missing).sum::<usize>(),
        results.iter().filter(|r| r.missing > 0).count(),
    );
    for e in &errors {
        println!("  ERROR {e}");
    }
    assert!(
        errors.is_empty(),
        "{} fixtures could not be analysed",
        errors.len()
    );

    // Other PDFs are scored for comparison only; baselines track our own output.
    if gen_name != "generated.pdf" {
        return;
    }
    let updates: HashMap<String, common::Baselines> = results
        .iter()
        .map(|r| {
            let v = r.scores.vs_word.as_ref();
            let b = common::Baselines {
                ua_fail: Some(r.scores.ua_fail),
                a11y_missing: Some(r.missing),
                ua_deficit: v.map(|v| v.deficit.len()),
                a11y_struct: v.map(|v| v.struct_score),
                a11y_text: v.map(|v| v.text_score),
                ..Default::default()
            };
            (r.name.clone(), b)
        })
        .collect();
    common::write_latest_scores(&updates);

    // A PDF that declares PDF/UA conformance must pass every machine check.
    let false_claims: Vec<String> = claims
        .iter()
        .filter(|r| r.scores.ua_fail > 0)
        .map(|r| format!("{} ({} rules)", r.name, r.scores.ua_fail))
        .collect();
    assert!(
        false_claims.is_empty(),
        "PDF/UA claimed but failing: {}",
        false_claims.join(", ")
    );

    let baselines = common::read_baselines();
    let mut regressions: Vec<&str> = updates
        .iter()
        .filter(|(name, r)| {
            baselines.get(*name).is_some_and(|b| {
                let worse = |old: Option<usize>, new: Option<usize>| matches!((old, new), (Some(o), Some(n)) if n > o);
                let dropped = |old: Option<f64>, new: Option<f64>| {
                    matches!((old, new), (Some(o), Some(n)) if n < o - common::REGRESSION_SLACK)
                };
                worse(b.ua_fail, r.ua_fail)
                    || worse(b.a11y_missing, r.a11y_missing)
                    || worse(b.ua_deficit, r.ua_deficit)
                    || dropped(b.a11y_struct, r.a11y_struct)
                    || dropped(b.a11y_text, r.a11y_text)
            })
        })
        .map(|(name, _)| name.as_str())
        .collect();
    regressions.sort();
    assert!(
        regressions.is_empty(),
        "a11y regression in: {}",
        regressions.join(", ")
    );
}

const LIST_DUMP: &str = r#"Document
  H1 (block)
    Span (inline)
      "Test"
      " "
  L (block):
     /ListNumbering /Disc
    LI (block)
      Lbl (block)
        "•"
      LBody (block)
        Span (inline)
          " "
          "This"
          " "
  P (block)
    Link (inline)
      Object 11 0
      Span (inline)
        "1"
      Note <Note 1> (inline)
        P (block)
          "A note."
    "after"
  Figure ["A close up of a sign??Description automatically generated "]:
     /Placement /Block
    ""
  Figure [""]:
    ""
"#;

#[test]
fn verapdf_report_keeps_every_job() {
    let json = r#"{"report":{"jobs":[
        {"itemDetails":{"name":"/a.pdf"},"validationResult":[{"details":{"ruleSummaries":[
            {"clause":"7.1","testNumber":3,"failedChecks":2,"passedChecks":5,"description":"tagged"}]}}]},
        {"itemDetails":{"name":"/b.pdf"},"taskResult":{"exceptionMessage":"Couldn't parse stream"}}
    ]}}"#;
    let jobs = a11y::parse_verapdf_report(json).unwrap();
    assert_eq!(jobs[0].0, Path::new("/a.pdf"));
    let rule = &jobs[0].1.as_ref().unwrap()["7.1-3"];
    assert_eq!(
        (rule.failed, rule.passed, rule.description.as_str()),
        (2, 5, "tagged")
    );
    assert_eq!(
        jobs[1].1.as_ref().unwrap_err(),
        "verapdf: Couldn't parse stream"
    );
}

#[test]
fn parses_struct_dump() {
    let elems = a11y::parse_struct_text(LIST_DUMP);
    let kinds: Vec<&str> = elems
        .iter()
        .filter(|e| e.kind != a11y::TEXT)
        .map(|e| e.kind.as_str())
        .collect();
    assert_eq!(
        kinds,
        [
            "Document", "H1", "Span", "L", "LI", "Lbl", "LBody", "Span", "P", "Link", "Span",
            "Note", "P", "Figure", "Figure"
        ]
    );
    let figures: Vec<bool> = elems
        .iter()
        .filter(|e| e.kind == "Figure")
        .map(|e| e.alt)
        .collect();
    assert_eq!(figures, [true, false]);
    assert_eq!(
        a11y::block_texts(&elems),
        ["Test", "•", "This", "1after", "A note."]
    );
}

#[test]
fn identical_trees_score_perfectly() {
    let elems = a11y::parse_struct_text(LIST_DUMP);
    assert_eq!(a11y::struct_score(&elems, &elems), 1.0);
    assert_eq!(a11y::text_score(&elems, &elems), 1.0);
    assert_eq!(a11y::struct_score(&elems, &[]), 0.0);
    assert_eq!(a11y::text_score(&elems, &[]), 0.0);
}

#[test]
fn text_score_penalises_merged_words_and_reordering() {
    let doc = |texts: &[&str]| {
        let mut dump = String::from("Document\n");
        for t in texts {
            dump += &format!("  P (block)\n    \"{t}\"\n");
        }
        a11y::parse_struct_text(&dump)
    };
    let reference = doc(&["Hello world", "Second"]);
    assert_eq!(
        a11y::text_score(&reference, &doc(&["Helloworld", "Second"])),
        6.0 / 17.0
    );
    assert_eq!(
        a11y::text_score(&reference, &doc(&["Second", "Hello world"])),
        11.0 / 17.0
    );
}

#[test]
fn text_score_ignores_which_symbol_a_glyph_maps_to() {
    let doc =
        |text: &str| a11y::parse_struct_text(&format!("Document\n  P (block)\n    \"{text}\"\n"));
    assert_eq!(
        a11y::text_score(&doc("Yes \u{F0A8} No \u{F0A8}"), &doc("Yes ◻ No ☐")),
        1.0
    );
    assert!(a11y::text_score(&doc("a ≤ b"), &doc("a ≥ b")) < 1.0);
}

#[test]
fn deficit_counts_only_rules_worse_than_word() {
    let json = |rules: &[(&str, u32, u64, u64)]| {
        let summaries: Vec<_> = rules
            .iter()
            .map(|(clause, test, failed, passed)| {
                serde_json::json!({"clause": clause, "testNumber": test, "failedChecks": failed, "passedChecks": passed, "description": "d"})
            })
            .collect();
        let report = serde_json::json!({"report": {"jobs": [{"validationResult": [{"details": {"ruleSummaries": summaries}}]}]}});
        let mut jobs = a11y::parse_verapdf_report(&report.to_string()).unwrap();
        jobs.pop().unwrap().1.unwrap()
    };
    // Word fails 5-1 (no pdfuaid) and one stray 7.1-3 check.
    let word = json(&[("5", 1, 1, 0), ("7.1", 3, 1, 999), ("6.2", 1, 0, 1)]);
    assert!(a11y::ua_deficit(&word, &word).is_empty());
    let ours = json(&[("5", 1, 1, 0), ("7.1", 3, 50, 50), ("6.2", 1, 1, 0)]);
    assert_eq!(a11y::ua_deficit(&word, &ours), ["6.2-1", "7.1-3"]);
}

#[test]
fn docx_paragraphs_keep_what_a_reader_gets() {
    let w = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#;
    let xml = format!(
        r#"<w:document {w}><w:body>
          <w:p><w:r><w:t>Kept </w:t></w:r><w:del><w:r><w:t>deleted</w:t></w:r></w:del>
            <w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden</w:t></w:r>
            <w:r><w:rPr><w:vanish w:val="0"/></w:rPr><w:t>shown</w:t></w:r>
            <mc:AlternateContent><mc:Choice><w:txbxContent><w:p><w:r><w:t>box</w:t></w:r></w:p></w:txbxContent></mc:Choice>
            <mc:Fallback><w:txbxContent><w:p><w:r><w:t>copy</w:t></w:r></w:p></w:txbxContent></mc:Fallback></mc:AlternateContent></w:p>
          <w:endnote w:type="separator"><w:p><w:r><w:t>separator</w:t></w:r></w:p></w:endnote>
        </w:body></w:document>"#
    );
    assert_eq!(a11y::docx_paragraphs(&xml).unwrap(), ["Kept shown", "box"]);
}
