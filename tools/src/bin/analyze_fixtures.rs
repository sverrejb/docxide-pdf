//! Analyze scraped DOCX fixtures to categorize which features they use.
//!
//! Usage:
//!   analyze-fixtures [fixtures_dir]
//!   analyze-fixtures --failing     only show fixtures below Jaccard threshold
//!   analyze-fixtures --fonts       show font list per fixture
//!
//! Defaults to tests/fixtures/scraped/ relative to the working directory.
//! Scans each fixture's input.docx for unsupported features, extracts fonts,
//! checks reference PDF page count, reads test scores from CSV, and prints
//! a per-fixture summary table plus aggregate tally.

use std::collections::HashMap;
use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::Command;
use zip::ZipArchive;

struct FixtureAnalysis {
    name: String,
    textboxes: usize,
    anchored_images: usize,
    floating_tables: usize,
    multi_column: bool,
    shapes_vml: usize,
    ole_objects: usize,
    smartart: bool,
    footnotes: bool,
    endnotes: bool,
    math: usize,
    sdt_content: usize,
    alternate_content: usize,
    total_paragraphs: usize,
    total_tables: usize,
    total_images_inline: usize,
    fonts: Vec<String>,
    ref_pages: Option<usize>,
    jaccard: Option<f64>,
    ssim: Option<f64>,
    skipped: bool,
    dominant_issue: String,
}

fn count_pattern(xml: &str, pattern: &str) -> usize {
    xml.matches(pattern).count()
}

fn extract_fonts(archive: &mut ZipArchive<fs::File>) -> Vec<String> {
    let Some(xml) = read_entry(archive, "word/fontTable.xml") else {
        return vec![];
    };
    let Ok(doc) = roxmltree::Document::parse(&xml) else {
        return vec![];
    };
    let wml = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    let mut fonts: Vec<String> = doc
        .root_element()
        .children()
        .filter(|n| n.tag_name().name() == "font" && n.tag_name().namespace() == Some(wml))
        .filter_map(|n| n.attribute((wml, "name")).map(String::from))
        .collect();
    fonts.sort();
    fonts.dedup();
    fonts
}

fn ref_page_count(fixture_path: &Path) -> Option<usize> {
    let ref_pdf = fixture_path.join("reference.pdf");
    if !ref_pdf.exists() {
        return None;
    }
    let output = Command::new("mutool")
        .args(["info", ref_pdf.to_str()?])
        .output()
        .ok()?;
    let stdout = String::from_utf8_lossy(&output.stdout);
    for line in stdout.lines() {
        if line.starts_with("Pages:") {
            return line.split(':').nth(1)?.trim().parse().ok();
        }
    }
    None
}

fn load_scores(csv_name: &str, score_col: usize) -> HashMap<String, f64> {
    let csv_path = PathBuf::from("tests/output").join(csv_name);
    let mut scores = HashMap::new();
    let Ok(content) = fs::read_to_string(&csv_path) else {
        return scores;
    };
    for line in content.lines().skip(1) {
        let parts: Vec<&str> = line.split(',').collect();
        if parts.len() > score_col {
            if let (Some(name), Ok(val)) = (parts.get(1), parts[score_col].parse::<f64>()) {
                scores.insert(name.to_string(), val);
            }
        }
    }
    scores
}

fn analyze_docx(path: &Path) -> Option<FixtureAnalysis> {
    let name = path.file_name()?.to_str()?.to_string();

    let docx_path = path.join("input.docx");
    if !docx_path.exists() {
        return None;
    }

    let file = fs::File::open(&docx_path).ok()?;
    let mut archive = ZipArchive::new(file).ok()?;

    let doc_xml = read_entry(&mut archive, "word/document.xml").unwrap_or_default();

    let textboxes = count_pattern(&doc_xml, "txbxContent")
        + count_pattern(&doc_xml, "v:textbox")
        + count_pattern(&doc_xml, "<wps:txbx");
    let anchored_images = count_pattern(&doc_xml, "wp:anchor");
    let floating_tables = count_pattern(&doc_xml, "tblpPr");
    let shapes_vml = count_pattern(&doc_xml, "v:shape") + count_pattern(&doc_xml, "v:rect");
    let ole_objects =
        count_pattern(&doc_xml, "o:OLEObject") + count_pattern(&doc_xml, "w:object");
    let math = count_pattern(&doc_xml, "m:oMath");
    let sdt_content = count_pattern(&doc_xml, "w:sdtContent");
    let alternate_content = count_pattern(&doc_xml, "mc:AlternateContent");
    let total_paragraphs = count_pattern(&doc_xml, "<w:p ") + count_pattern(&doc_xml, "<w:p>");
    let total_tables = count_pattern(&doc_xml, "<w:tbl>") + count_pattern(&doc_xml, "<w:tbl ");
    let total_images_inline = count_pattern(&doc_xml, "wp:inline");

    let multi_column = doc_xml.contains("w:cols ") && {
        if let Some(pos) = doc_xml.find("w:cols ") {
            let snippet = &doc_xml[pos..doc_xml.len().min(pos + 200)];
            snippet.contains("w:num=\"2")
                || snippet.contains("w:num=\"3")
                || snippet.contains("w:num=\"4")
        } else {
            false
        }
    };

    let smartart = archive.by_name("word/diagrams/data1.xml").is_ok();
    let footnotes = {
        let fn_xml = read_entry(&mut archive, "word/footnotes.xml").unwrap_or_default();
        count_pattern(&fn_xml, "<w:footnote ") > 2
    };
    let endnotes = {
        let en_xml = read_entry(&mut archive, "word/endnotes.xml").unwrap_or_default();
        count_pattern(&en_xml, "<w:endnote ") > 2
    };

    let fonts = extract_fonts(&mut archive);
    let ref_pages = ref_page_count(path);

    let dominant_issue = if textboxes > 5 {
        format!("textboxes ({})", textboxes)
    } else if anchored_images > 3 {
        format!("anchored images ({})", anchored_images)
    } else if floating_tables > 0 {
        format!("floating tables ({})", floating_tables)
    } else if multi_column {
        "multi-column layout".to_string()
    } else if shapes_vml > 5 {
        format!("VML shapes ({})", shapes_vml)
    } else if ole_objects > 0 {
        format!("OLE objects ({})", ole_objects)
    } else if math > 0 {
        format!("math equations ({})", math)
    } else if smartart {
        "SmartArt".to_string()
    } else if textboxes > 0 {
        format!("textboxes ({})", textboxes)
    } else if anchored_images > 0 {
        format!("anchored images ({})", anchored_images)
    } else if sdt_content > 3 {
        format!("structured doc tags ({})", sdt_content)
    } else if alternate_content > 0 {
        format!("mc:AlternateContent ({})", alternate_content)
    } else {
        "text/layout only".to_string()
    };

    Some(FixtureAnalysis {
        name,
        textboxes,
        anchored_images,
        floating_tables,
        multi_column,
        shapes_vml,
        ole_objects,
        smartart,
        footnotes,
        endnotes,
        math,
        sdt_content,
        alternate_content,
        total_paragraphs,
        total_tables,
        total_images_inline,
        fonts,
        ref_pages,
        jaccard: None,
        ssim: None,
        skipped: false,
        dominant_issue,
    })
}

fn read_entry(archive: &mut ZipArchive<fs::File>, name: &str) -> Option<String> {
    let mut entry = archive.by_name(name).ok()?;
    let mut content = String::new();
    entry.read_to_string(&mut content).ok()?;
    Some(content)
}

fn audit_fixtures(fixtures_dir: &Path) {
    let features: &[(&str, &[&str])] = &[
        // Run properties that may not be implemented
        ("w:caps", &["w:caps"]),
        ("w:smallCaps", &["w:smallCaps"]),
        ("w:dstrike (double-strike)", &["w:dstrike"]),
        ("w:vanish (hidden text)", &["w:vanish"]),
        ("w:spacing (char spacing)", &["w:spacing"]),
        ("w:kern", &["w:kern"]),
        ("w:emboss/imprint/shadow", &["w:emboss", "w:imprint", "w:shadow"]),
        // Paragraph properties
        ("w:ind w:right (right indent)", &["w:right=\""]),
        ("w:mirrorIndents", &["w:mirrorIndents"]),
        ("w:numPr (lists)", &["w:numPr"]),
        ("w:sectPr in pPr (mid-doc sections)", &["<w:sectPr>"]),
        // Document-level
        ("w:sdtContent (struct doc tags)", &["w:sdtContent"]),
        ("mc:AlternateContent", &["mc:AlternateContent"]),
        ("w:txbxContent (textboxes)", &["txbxContent"]),
        ("wp:anchor (anchored drawings)", &["wp:anchor"]),
        ("w:tblpPr (floating tables)", &["tblpPr"]),
        ("w:cols multi-col", &["w:num=\"2", "w:num=\"3", "w:num=\"4"]),
        ("w:drawing (any drawing)", &["w:drawing"]),
        ("m:oMath (math)", &["m:oMath"]),
        ("v:shape (VML)", &["v:shape"]),
        ("w:object (OLE)", &["w:object"]),
    ];

    let mut entries: Vec<_> = fs::read_dir(fixtures_dir)
        .unwrap()
        .filter_map(|e| e.ok())
        .filter(|e| e.path().is_dir())
        .collect();
    entries.sort_by_key(|e| e.file_name());

    let jaccard_scores = load_scores("results.csv", 3);
    let skip_fixtures = load_skip_list();

    // Collect per-feature counts: feature_name -> (failing_fixtures, passing_fixtures, total_count)
    let mut feature_stats: Vec<(&str, usize, usize, usize, usize)> = Vec::new();

    for &(name, patterns) in features {
        let mut failing = 0usize;
        let mut passing = 0usize;
        let mut skipped_count = 0usize;
        let mut total_hits = 0usize;

        for entry in &entries {
            let fixture_name = entry.file_name().to_string_lossy().to_string();
            let docx_path = entry.path().join("input.docx");
            let Ok(file) = fs::File::open(&docx_path) else { continue };
            let Ok(mut archive) = ZipArchive::new(file) else { continue };

            let doc_xml = read_entry(&mut archive, "word/document.xml").unwrap_or_default();
            let style_xml = read_entry(&mut archive, "word/styles.xml").unwrap_or_default();
            let all_xml = format!("{}{}", doc_xml, style_xml);

            let hits: usize = patterns.iter().map(|p| count_pattern(&all_xml, p)).sum();
            if hits > 0 {
                total_hits += hits;
                let jaccard = jaccard_scores.get(&fixture_name).copied().unwrap_or(0.0);
                let is_skipped = skip_fixtures.contains(&fixture_name);
                if is_skipped {
                    skipped_count += 1;
                } else if jaccard < 0.20 {
                    failing += 1;
                } else {
                    passing += 1;
                }
            }
        }

        if total_hits > 0 {
            feature_stats.push((name, failing, passing, skipped_count, total_hits));
        }
    }

    // Sort by failing count descending, then total hits
    feature_stats.sort_by(|a, b| b.1.cmp(&a.1).then(b.4.cmp(&a.4)));

    println!("{:<40} {:>6} {:>6} {:>6} {:>8}", "Feature", "Fail", "Pass", "Skip", "Hits");
    println!("{}", "─".repeat(72));
    for (name, failing, passing, skipped, hits) in &feature_stats {
        println!(
            "{:<40} {:>6} {:>6} {:>6} {:>8}",
            name, failing, passing, skipped, hits
        );
    }
}

fn grep_fixtures(fixtures_dir: &Path, pattern: &str) {
    let mut entries: Vec<_> = fs::read_dir(fixtures_dir)
        .unwrap()
        .filter_map(|e| e.ok())
        .filter(|e| e.path().is_dir())
        .collect();
    entries.sort_by_key(|e| e.file_name());

    let jaccard_scores = load_scores("results.csv", 3);

    println!("{:<66} {:>6} {:>6} {:>6}  {}", "Fixture", "Jaccd", "doc", "style", "Total");
    println!("{}", "─".repeat(100));
    let mut total_fixtures = 0;
    for entry in &entries {
        let name = entry.file_name().to_string_lossy().to_string();
        let docx_path = entry.path().join("input.docx");
        let Ok(file) = fs::File::open(&docx_path) else { continue };
        let Ok(mut archive) = ZipArchive::new(file) else { continue };

        let doc_count = read_entry(&mut archive, "word/document.xml")
            .map(|x| count_pattern(&x, pattern))
            .unwrap_or(0);
        let style_count = read_entry(&mut archive, "word/styles.xml")
            .map(|x| count_pattern(&x, pattern))
            .unwrap_or(0);
        let total = doc_count + style_count;

        if total > 0 {
            let jaccard = jaccard_scores.get(&name).copied().unwrap_or(0.0);
            let short_name = if name.len() > 64 {
                format!("{}…", &name[..63])
            } else {
                name
            };
            println!(
                "{:<66} {:>5.1}% {:>6} {:>6}  {}",
                short_name,
                jaccard * 100.0,
                doc_count,
                style_count,
                total,
            );
            total_fixtures += 1;
        }
    }
    println!("\n{} fixtures contain '{}'", total_fixtures, pattern);
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let only_failing = args.iter().any(|a| a == "--failing");
    let show_fonts = args.iter().any(|a| a == "--fonts");
    let do_audit = args.iter().any(|a| a == "--audit");
    let grep_pattern = args.iter().position(|a| a == "--grep").and_then(|i| args.get(i + 1));
    let fixtures_dir = args
        .iter()
        .skip(1)
        .find(|a| !a.starts_with('-') && args.iter().position(|x| x == "--grep").is_none_or(|gi| args.get(gi + 1) != Some(a)))
        .map(PathBuf::from)
        .unwrap_or_else(|| PathBuf::from("tests/fixtures/scraped"));

    if !fixtures_dir.is_dir() {
        eprintln!("Directory not found: {}", fixtures_dir.display());
        std::process::exit(1);
    }

    if do_audit {
        audit_fixtures(&fixtures_dir);
        return;
    }

    if let Some(pattern) = grep_pattern {
        grep_fixtures(&fixtures_dir, pattern);
        return;
    }

    // Load skip list
    let skip_fixtures = load_skip_list();

    // Load test scores: results.csv = timestamp,case,pages,avg_jaccard,pass
    //                    ssim_results.csv = timestamp,case,pages,avg_ssim
    let jaccard_scores = load_scores("results.csv", 3);
    let ssim_scores = load_scores("ssim_results.csv", 3);

    let mut entries: Vec<_> = fs::read_dir(&fixtures_dir)
        .unwrap()
        .filter_map(|e| e.ok())
        .filter(|e| e.path().is_dir())
        .collect();
    entries.sort_by_key(|e| e.file_name());

    let mut analyses = Vec::new();
    for entry in &entries {
        let Some(mut a) = analyze_docx(&entry.path()) else {
            continue;
        };
        a.jaccard = jaccard_scores.get(&a.name).copied();
        a.ssim = ssim_scores.get(&a.name).copied();
        a.skipped = skip_fixtures.contains(&a.name);
        analyses.push(a);
    }

    if only_failing {
        analyses.retain(|a| {
            a.jaccard.is_none_or(|j| j < 0.20) && !a.skipped
        });
    }

    // Print per-fixture table
    println!(
        "{:<66} {:>6} {:>6} {:>3} {:>4} {:>4} {:>4} {:>3} {:>4}  {}",
        "Fixture",
        "Jaccd",
        "SSIM",
        "Pg",
        "TxBx",
        "Anch",
        "FTbl",
        "Col",
        "AltC",
        "Dominant Issue"
    );
    println!("{}", "─".repeat(140));
    for a in &analyses {
        let short_name = if a.name.len() > 64 {
            format!("{}…", &a.name[..63])
        } else {
            a.name.clone()
        };
        let skip_marker = if a.skipped { " [SKIP]" } else { "" };
        println!(
            "{:<66} {:>5.1}% {:>5.1}% {:>3} {:>4} {:>4} {:>4} {:>3} {:>4}  {}{}",
            short_name,
            a.jaccard.unwrap_or(0.0) * 100.0,
            a.ssim.unwrap_or(0.0) * 100.0,
            a.ref_pages.map(|p| p.to_string()).unwrap_or("-".into()),
            a.textboxes,
            a.anchored_images,
            a.floating_tables,
            if a.multi_column { "Y" } else { "-" },
            a.alternate_content,
            a.dominant_issue,
            skip_marker,
        );
        if show_fonts {
            println!("    fonts: {}", a.fonts.join(", "));
        }
    }

    // Aggregate tally
    let counted = if only_failing { &analyses } else { &analyses };
    let mut tally: HashMap<&str, usize> = HashMap::new();
    for a in counted {
        if a.textboxes > 0 {
            *tally.entry("textboxes").or_default() += 1;
        }
        if a.anchored_images > 0 {
            *tally.entry("anchored_images").or_default() += 1;
        }
        if a.floating_tables > 0 {
            *tally.entry("floating_tables").or_default() += 1;
        }
        if a.multi_column {
            *tally.entry("multi_column").or_default() += 1;
        }
        if a.shapes_vml > 0 {
            *tally.entry("vml_shapes").or_default() += 1;
        }
        if a.ole_objects > 0 {
            *tally.entry("ole_objects").or_default() += 1;
        }
        if a.math > 0 {
            *tally.entry("math").or_default() += 1;
        }
        if a.sdt_content > 0 {
            *tally.entry("structured_doc_tags").or_default() += 1;
        }
        if a.alternate_content > 0 {
            *tally.entry("alternate_content").or_default() += 1;
        }
        if a.footnotes {
            *tally.entry("footnotes").or_default() += 1;
        }
        if a.endnotes {
            *tally.entry("endnotes").or_default() += 1;
        }
        if a.smartart {
            *tally.entry("smartart").or_default() += 1;
        }
        if a.dominant_issue == "text/layout only" {
            *tally.entry("text/layout only").or_default() += 1;
        }
    }

    println!(
        "\n\nFeature Tally (across {} fixtures):",
        counted.len()
    );
    println!("{}", "─".repeat(40));
    let mut sorted: Vec<_> = tally.into_iter().collect();
    sorted.sort_by(|a, b| b.1.cmp(&a.1));
    for (feature, count) in &sorted {
        println!("  {:<25} {:>3} fixtures", feature, count);
    }

    // Summary stats
    let total_paras: usize = analyses.iter().map(|a| a.total_paragraphs).sum();
    let total_tbls: usize = analyses.iter().map(|a| a.total_tables).sum();
    let total_imgs: usize = analyses.iter().map(|a| a.total_images_inline).sum();
    let passing = analyses
        .iter()
        .filter(|a| a.jaccard.is_some_and(|j| j >= 0.20))
        .count();
    let failing = analyses
        .iter()
        .filter(|a| a.jaccard.is_some_and(|j| j < 0.20) && !a.skipped)
        .count();
    let skipped = analyses.iter().filter(|a| a.skipped).count();
    println!(
        "\nContent: {} paragraphs, {} tables, {} inline images across {} fixtures",
        total_paras,
        total_tbls,
        total_imgs,
        analyses.len()
    );
    println!(
        "Scores: {} passing, {} failing, {} skipped (font issues)",
        passing, failing, skipped
    );
}

fn load_skip_list() -> Vec<String> {
    // Parse SKIP_FIXTURES from tests/common/mod.rs
    let Ok(content) = fs::read_to_string("tests/common/mod.rs") else {
        return vec![];
    };
    let mut skips = Vec::new();
    let mut in_skip = false;
    for line in content.lines() {
        if line.contains("SKIP_FIXTURES") && line.contains("&[") {
            in_skip = true;
        }
        if in_skip {
            if let Some(start) = line.find('"') {
                if let Some(end) = line[start + 1..].find('"') {
                    skips.push(line[start + 1..start + 1 + end].to_string());
                }
            }
            if line.contains("];") {
                break;
            }
        }
    }
    skips
}
