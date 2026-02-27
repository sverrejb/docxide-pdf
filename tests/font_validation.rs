mod common;

use rayon::prelude::*;
use std::collections::{BTreeSet, HashMap};
use std::io::Read;
use std::path::Path;
use std::process::Command;
use std::fs;

/// Extract unique font family names from a PDF using `mutool info`.
fn extract_pdf_fonts(pdf: &Path) -> Result<BTreeSet<String>, String> {
    let output = Command::new("mutool")
        .args(["info", pdf.to_str().unwrap()])
        .output()
        .map_err(|e| format!("Failed to run mutool info: {e}"))?;
    let text = String::from_utf8_lossy(&output.stdout);

    let mut families = BTreeSet::new();
    for line in text.lines() {
        if let Some(start) = line.find('\'') {
            if let Some(end) = line[start + 1..].find('\'') {
                let raw_name = &line[start + 1..start + 1 + end];
                let family = normalize_pdf_font_name(raw_name);
                if !family.is_empty() {
                    families.insert(family);
                }
            }
        }
    }
    Ok(families)
}

/// Normalize a PDF font name to its base family, removing subset prefixes and style suffixes.
fn normalize_pdf_font_name(name: &str) -> String {
    let name = if name.len() > 7
        && name.as_bytes()[6] == b'+'
        && name[..6].chars().all(|c| c.is_ascii_uppercase())
    {
        &name[7..]
    } else {
        name
    };

    let name = name
        .replace("-BoldItalic", "")
        .replace("-BoldMT", "")
        .replace("-ItalicMT", "")
        .replace("-Bold", "")
        .replace("-Italic", "");

    if let Some(stripped) = name.strip_suffix("PSMT") {
        stripped.to_string()
    } else if let Some(stripped) = name.strip_suffix("PS") {
        stripped.to_string()
    } else if let Some(stripped) = name.strip_suffix("MT") {
        stripped.to_string()
    } else {
        name
    }
}

/// Normalize a display font name (from DOCX) to match PDF PostScript naming.
/// "Times New Roman" → "TimesNewRoman", "Courier New" → "CourierNew"
fn normalize_docx_font_name(name: &str) -> String {
    name.replace(' ', "")
}

/// Extract font family names the DOCX actually uses by parsing its XML.
fn extract_docx_fonts(docx_path: &Path) -> Result<BTreeSet<String>, String> {
    let file = fs::File::open(docx_path).map_err(|e| format!("open: {e}"))?;
    let mut archive = zip::ZipArchive::new(file).map_err(|e| format!("zip: {e}"))?;

    // Parse theme fonts (major=heading, minor=body)
    let (theme_major, theme_minor) = parse_theme_fonts(&mut archive);

    // Resolve the default body font: Normal style > docDefaults > theme minor
    let default_font = parse_default_font(&mut archive, &theme_major, &theme_minor);

    // Build style→font map from styles.xml (resolves theme refs + basedOn inheritance)
    let style_fonts = parse_style_fonts(&mut archive, &theme_major, &theme_minor);

    // Collect explicit fonts and used styles from document.xml, headers, footers
    let xml_files = collect_xml_names(&mut archive);
    let mut fonts = BTreeSet::new();
    let mut used_styles: BTreeSet<String> = BTreeSet::new();
    let mut has_unstyled_runs = false;

    for xml_name in &xml_files {
        let Ok(mut entry) = archive.by_name(xml_name) else {
            continue;
        };
        let mut content = String::new();
        if entry.read_to_string(&mut content).is_err() {
            continue;
        }
        let Ok(doc) = roxmltree::Document::parse(&content) else {
            continue;
        };
        collect_fonts_from_xml(&doc, &theme_major, &theme_minor, &mut fonts);
        collect_used_styles(&doc, &mut used_styles);
        if !has_unstyled_runs {
            has_unstyled_runs = has_runs_without_font(&doc);
        }
    }

    // Add fonts from styles actually used in the document
    for style_id in &used_styles {
        if let Some(font) = style_fonts.get(style_id.as_str()) {
            fonts.insert(normalize_docx_font_name(font));
        }
    }

    // Include default body font if any runs rely on style/default inheritance
    if has_unstyled_runs {
        if let Some(name) = &default_font {
            fonts.insert(normalize_docx_font_name(name));
        }
    }
    if fonts.is_empty() {
        if let Some(name) = &default_font {
            fonts.insert(normalize_docx_font_name(name));
        }
    }

    Ok(fonts)
}

fn collect_xml_names(archive: &mut zip::ZipArchive<fs::File>) -> Vec<String> {
    let mut names = Vec::new();
    for i in 0..archive.len() {
        if let Ok(entry) = archive.by_index(i) {
            let name = entry.name().to_string();
            if name == "word/document.xml"
                || name.starts_with("word/header")
                || name.starts_with("word/footer")
            {
                names.push(name);
            }
        }
    }
    names
}

/// Resolve the default body font from styles.xml.
/// Priority: Normal style w:ascii > docDefaults w:ascii > docDefaults theme ref > theme minor
fn parse_default_font(
    archive: &mut zip::ZipArchive<fs::File>,
    theme_major: &Option<String>,
    theme_minor: &Option<String>,
) -> Option<String> {
    let Ok(mut entry) = archive.by_name("word/styles.xml") else {
        return theme_minor.clone();
    };
    let mut content = String::new();
    if entry.read_to_string(&mut content).is_err() {
        return theme_minor.clone();
    }
    let Ok(doc) = roxmltree::Document::parse(&content) else {
        return theme_minor.clone();
    };

    let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // 1. docDefaults font
    let mut doc_default_font = None;
    for node in doc.descendants() {
        if node.tag_name().name() == "docDefaults" {
            for rpr_default in node.descendants() {
                if rpr_default.tag_name().name() == "rFonts" {
                    if let Some(name) = rpr_default.attribute((w, "ascii"))
                        .or_else(|| rpr_default.attribute("ascii"))
                    {
                        doc_default_font = Some(name.to_string());
                    } else {
                        let theme = rpr_default.attribute((w, "asciiTheme"))
                            .or_else(|| rpr_default.attribute("asciiTheme"));
                        if let Some(t) = theme {
                            doc_default_font = resolve_theme(t, theme_major, theme_minor);
                        }
                    }
                }
            }
        }
    }

    // 2. Normal style font (overrides docDefaults)
    for node in doc.descendants() {
        if node.tag_name().name() == "style" {
            let style_id = node.attribute((w, "styleId"))
                .or_else(|| node.attribute("styleId"));
            if style_id == Some("Normal") {
                for rfonts in node.descendants() {
                    if rfonts.tag_name().name() == "rFonts" {
                        if let Some(name) = rfonts.attribute((w, "ascii"))
                            .or_else(|| rfonts.attribute("ascii"))
                        {
                            return Some(name.to_string());
                        }
                        let theme = rfonts.attribute((w, "asciiTheme"))
                            .or_else(|| rfonts.attribute("asciiTheme"));
                        if let Some(t) = theme {
                            if let Some(resolved) = resolve_theme(t, theme_major, theme_minor) {
                                return Some(resolved);
                            }
                        }
                    }
                }
            }
        }
    }

    doc_default_font.or_else(|| theme_minor.clone())
}

fn resolve_theme(
    theme: &str,
    theme_major: &Option<String>,
    theme_minor: &Option<String>,
) -> Option<String> {
    match theme {
        "majorHAnsi" | "majorBidi" | "majorEastAsia" => theme_major.clone(),
        "minorHAnsi" | "minorBidi" | "minorEastAsia" => theme_minor.clone(),
        _ => None,
    }
}

fn parse_theme_fonts(
    archive: &mut zip::ZipArchive<fs::File>,
) -> (Option<String>, Option<String>) {
    let Ok(mut entry) = archive.by_name("word/theme/theme1.xml") else {
        return (None, None);
    };
    let mut content = String::new();
    if entry.read_to_string(&mut content).is_err() {
        return (None, None);
    }
    let Ok(doc) = roxmltree::Document::parse(&content) else {
        return (None, None);
    };

    let mut major = None;
    let mut minor = None;
    for node in doc.descendants() {
        if node.tag_name().name() == "majorFont" {
            for child in node.children() {
                if child.tag_name().name() == "latin" {
                    major = child.attribute("typeface").map(String::from);
                }
            }
        }
        if node.tag_name().name() == "minorFont" {
            for child in node.children() {
                if child.tag_name().name() == "latin" {
                    minor = child.attribute("typeface").map(String::from);
                }
            }
        }
    }
    (major, minor)
}

/// Check if any w:r element lacks an explicit w:rFonts (relies on style/default font).
fn has_runs_without_font(doc: &roxmltree::Document) -> bool {
    let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    for node in doc.descendants() {
        if node.tag_name().name() == "r"
            && (node.tag_name().namespace() == Some(w) || node.tag_name().namespace().is_none())
        {
            // Check if this run has w:rPr/w:rFonts
            let has_font = node.children().any(|child| {
                child.tag_name().name() == "rPr"
                    && child
                        .children()
                        .any(|n| n.tag_name().name() == "rFonts")
            });
            if !has_font {
                // Check that this run has actual text content (not just formatting)
                let has_text = node
                    .children()
                    .any(|c| c.tag_name().name() == "t" || c.tag_name().name() == "br");
                if has_text {
                    return true;
                }
            }
        }
    }
    false
}

/// Collect w:pStyle and w:rStyle values from document content XML.
fn collect_used_styles(doc: &roxmltree::Document, styles: &mut BTreeSet<String>) {
    let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    for node in doc.descendants() {
        let name = node.tag_name().name();
        if name == "pStyle" || name == "rStyle" {
            if let Some(val) = node.attribute((w, "val")).or_else(|| node.attribute("val")) {
                styles.insert(val.to_string());
            }
        }
    }
}

/// Build a map of styleId → resolved font name from styles.xml.
/// Handles theme references and basedOn inheritance.
fn parse_style_fonts(
    archive: &mut zip::ZipArchive<fs::File>,
    theme_major: &Option<String>,
    theme_minor: &Option<String>,
) -> HashMap<String, String> {
    let Ok(mut entry) = archive.by_name("word/styles.xml") else {
        return HashMap::new();
    };
    let mut content = String::new();
    if entry.read_to_string(&mut content).is_err() {
        return HashMap::new();
    }
    let Ok(doc) = roxmltree::Document::parse(&content) else {
        return HashMap::new();
    };

    let w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    // First pass: collect direct font and basedOn for each style
    let mut direct_font: HashMap<String, String> = HashMap::new();
    let mut based_on: HashMap<String, String> = HashMap::new();

    for node in doc.descendants() {
        if node.tag_name().name() != "style" {
            continue;
        }
        let Some(style_id) = node.attribute((w, "styleId")).or_else(|| node.attribute("styleId"))
        else {
            continue;
        };

        for child in node.descendants() {
            if child.tag_name().name() == "basedOn" {
                if let Some(val) = child.attribute((w, "val")).or_else(|| child.attribute("val")) {
                    based_on.insert(style_id.to_string(), val.to_string());
                }
            }
            if child.tag_name().name() == "rFonts" {
                if let Some(name) = child
                    .attribute((w, "ascii"))
                    .or_else(|| child.attribute("ascii"))
                {
                    direct_font.insert(style_id.to_string(), name.to_string());
                } else if let Some(theme) = child
                    .attribute((w, "asciiTheme"))
                    .or_else(|| child.attribute("asciiTheme"))
                {
                    if let Some(resolved) = resolve_theme(theme, theme_major, theme_minor) {
                        direct_font.insert(style_id.to_string(), resolved);
                    }
                }
            }
        }
    }

    // Second pass: resolve inheritance (walk basedOn chain)
    let style_ids: Vec<String> = direct_font
        .keys()
        .chain(based_on.keys())
        .cloned()
        .collect::<BTreeSet<_>>()
        .into_iter()
        .collect();
    let mut resolved: HashMap<String, String> = HashMap::new();
    for id in &style_ids {
        if let Some(font) = resolve_style_font(id, &direct_font, &based_on, &mut resolved, 10) {
            resolved.insert(id.clone(), font);
        }
    }

    resolved
}

fn resolve_style_font(
    id: &str,
    direct: &HashMap<String, String>,
    based_on: &HashMap<String, String>,
    cache: &mut HashMap<String, String>,
    depth: u8,
) -> Option<String> {
    if depth == 0 {
        return None;
    }
    if let Some(cached) = cache.get(id) {
        return Some(cached.clone());
    }
    if let Some(font) = direct.get(id) {
        return Some(font.clone());
    }
    if let Some(parent) = based_on.get(id) {
        return resolve_style_font(parent, direct, based_on, cache, depth - 1);
    }
    None
}

fn collect_fonts_from_xml(
    doc: &roxmltree::Document,
    theme_major: &Option<String>,
    theme_minor: &Option<String>,
    fonts: &mut BTreeSet<String>,
) {
    for node in doc.descendants() {
        if node.tag_name().name() == "rFonts" {
            // Direct font name
            if let Some(name) = node.attribute((
                "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                "ascii",
            )) {
                fonts.insert(normalize_docx_font_name(name));
            } else if let Some(name) = node.attribute("ascii") {
                fonts.insert(normalize_docx_font_name(name));
            }

            // Theme font reference → resolve to actual name
            let theme_attr = node
                .attribute((
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    "asciiTheme",
                ))
                .or_else(|| node.attribute("asciiTheme"));
            if let Some(theme) = theme_attr {
                let resolved = match theme {
                    "majorHAnsi" | "majorBidi" | "majorEastAsia" => theme_major.as_deref(),
                    "minorHAnsi" | "minorBidi" | "minorEastAsia" => theme_minor.as_deref(),
                    _ => None,
                };
                if let Some(name) = resolved {
                    fonts.insert(normalize_docx_font_name(name));
                }
            }
        }
    }

}

struct FixtureResult {
    name: String,
    group: String,
    docx_fonts: BTreeSet<String>,
    pdf_fonts: BTreeSet<String>,
    missing: BTreeSet<String>,
    unexpected_fallbacks: BTreeSet<String>,
    pass: bool,
}

const FALLBACK_FONTS: &[&str] = &["Helvetica"];

fn analyze_fixture(fixture_dir: &Path) -> Option<FixtureResult> {
    let name = common::display_name(fixture_dir);
    let input_docx = fixture_dir.join("input.docx");
    if !input_docx.exists() {
        return None;
    }

    let output_dir = common::output_dir(fixture_dir);
    fs::create_dir_all(&output_dir).ok();
    let generated_pdf = output_dir.join("generated.pdf");

    // Reuse existing generated.pdf if newer than input.docx
    let needs_convert = !generated_pdf.exists() || {
        let docx_mtime = fs::metadata(&input_docx)
            .and_then(|m| m.modified())
            .unwrap_or(std::time::SystemTime::UNIX_EPOCH);
        let pdf_mtime = fs::metadata(&generated_pdf)
            .and_then(|m| m.modified())
            .unwrap_or(std::time::SystemTime::UNIX_EPOCH);
        pdf_mtime < docx_mtime
    };
    if needs_convert {
        if let Err(e) = docxide_pdf::convert_docx_to_pdf(&input_docx, &generated_pdf) {
            println!("  [SKIP] {name}: {e}");
            return None;
        }
    }

    let docx_fonts = match extract_docx_fonts(&input_docx) {
        Ok(f) => f,
        Err(e) => {
            println!("  [SKIP] {name}: docx parse error: {e}");
            return None;
        }
    };

    let pdf_fonts = match extract_pdf_fonts(&generated_pdf) {
        Ok(f) => f,
        Err(e) => {
            println!("  [SKIP] {name}: pdf font extraction error: {e}");
            return None;
        }
    };

    // Fonts the DOCX expects but the PDF doesn't have
    let missing: BTreeSet<String> = docx_fonts
        .iter()
        .filter(|f| !pdf_fonts.contains(*f))
        .cloned()
        .collect();

    // Fallback fonts in PDF that the DOCX didn't ask for
    let unexpected_fallbacks: BTreeSet<String> = pdf_fonts
        .iter()
        .filter(|f| FALLBACK_FONTS.contains(&f.as_str()) && !docx_fonts.contains(*f))
        .cloned()
        .collect();

    let pass = missing.is_empty() && unexpected_fallbacks.is_empty();

    let group = common::group_name(fixture_dir);

    Some(FixtureResult {
        name,
        group,
        docx_fonts,
        pdf_fonts,
        missing,
        unexpected_fallbacks,
        pass,
    })
}

#[test]
fn font_families_match_docx() {
    let _ = env_logger::try_init();
    let fixtures = common::discover_fixtures().expect("Failed to read tests/fixtures");
    if fixtures.is_empty() {
        return;
    }

    let mut results: Vec<FixtureResult> = fixtures
        .par_iter()
        .filter_map(|f| analyze_fixture(f))
        .collect();
    results.sort_by(|a, b| a.name.cmp(&b.name));

    let ts = common::timestamp();
    let name_w = results
        .iter()
        .map(|r| r.name.len())
        .max()
        .unwrap_or(4)
        .max(4);

    struct RowDisplay {
        matched: String,
        diff: String,
    }

    let rows: Vec<RowDisplay> = results
        .iter()
        .map(|r| {
            let matched: Vec<&str> = r
                .docx_fonts
                .iter()
                .filter(|f| r.pdf_fonts.contains(*f))
                .map(|s| s.as_str())
                .collect();
            let mut diff_parts: Vec<String> = Vec::new();
            for f in &r.docx_fonts {
                if !r.pdf_fonts.contains(f) {
                    diff_parts.push(format!("-{f}"));
                }
            }
            for f in &r.pdf_fonts {
                if !r.docx_fonts.contains(f) {
                    diff_parts.push(format!("+{f}"));
                }
            }
            RowDisplay {
                matched: matched.join(", "),
                diff: diff_parts.join(", "),
            }
        })
        .collect();

    let match_w = rows
        .iter()
        .map(|r| r.matched.len())
        .max()
        .unwrap_or(7)
        .max(7);
    let diff_w = rows
        .iter()
        .map(|r| r.diff.len())
        .max()
        .unwrap_or(4)
        .max(4);

    let sep = format!(
        "+-{}-+------+-{}-+-{}-+",
        "-".repeat(name_w),
        "-".repeat(match_w),
        "-".repeat(diff_w)
    );
    let thin = format!(
        "+-{}-+------+-{}-+-{}-+",
        "-".repeat(name_w),
        "-".repeat(match_w),
        "-".repeat(diff_w)
    );

    println!("\n{sep}");
    println!(
        "| {:<name_w$} | Pass | {:<match_w$} | {:<diff_w$} |",
        "Case", "Matched", "Diff"
    );
    println!("{sep}");

    for (i, (r, row)) in results.iter().zip(&rows).enumerate() {
        if i > 0 {
            println!("{thin}");
        }
        let status = if r.pass { "Y" } else { "N" };
        println!(
            "| {:<name_w$} | {:<4} | {:<match_w$} | {:<diff_w$} |",
            r.name, status, row.matched, row.diff
        );

        common::log_csv(
            "font_validation_results.csv",
            "timestamp,case,pass,docx_fonts,pdf_fonts,missing,unexpected_fallbacks",
            &format!(
                "{},{},{},{},{},{},{}",
                ts,
                r.name,
                r.pass,
                r.docx_fonts
                    .iter()
                    .cloned()
                    .collect::<Vec<_>>()
                    .join(";"),
                r.pdf_fonts
                    .iter()
                    .cloned()
                    .collect::<Vec<_>>()
                    .join(";"),
                r.missing
                    .iter()
                    .cloned()
                    .collect::<Vec<_>>()
                    .join(";"),
                r.unexpected_fallbacks
                    .iter()
                    .cloned()
                    .collect::<Vec<_>>()
                    .join(";"),
            ),
        );
    }

    println!("{sep}");
    println!("  + font in PDF but not declared in DOCX | - declared in DOCX but missing from PDF");
    let case_failures: Vec<&str> = results
        .iter()
        .filter(|r| !r.pass && r.group == "cases")
        .map(|r| r.name.as_str())
        .collect();
    assert!(
        case_failures.is_empty(),
        "Font mismatches in cases/: {}",
        case_failures.join(", ")
    );
}
