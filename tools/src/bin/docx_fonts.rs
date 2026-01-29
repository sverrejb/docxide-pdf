//! Analyze font usage in a DOCX file.
//! Reads theme1.xml, styles.xml, and document.xml to summarise:
//!   - Theme major/minor font names
//!   - Document defaults (font, size, spacing)
//!   - Per named-style font settings
//!   - Top font+size combos found in body runs
//!
//! Usage: docx-fonts <file.docx>

use std::collections::HashMap;
use std::fs;
use std::io::Read;
use zip::ZipArchive;

const WML: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const DML: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

fn wml<'a>(node: &roxmltree::Node<'a, 'a>, local: &str) -> Option<&'a str> {
    node.attribute((WML, local))
}

fn read_zip_text(archive: &mut ZipArchive<fs::File>, path: &str) -> Option<String> {
    let mut entry = archive.by_name(path).ok()?;
    let mut s = String::new();
    entry.read_to_string(&mut s).ok()?;
    Some(s)
}

// ── Theme ─────────────────────────────────────────────────────────────────────

fn print_theme_fonts(archive: &mut ZipArchive<fs::File>) {
    let xml = match read_zip_text(archive, "word/theme/theme1.xml") {
        Some(s) => s,
        None => {
            println!("  (word/theme/theme1.xml not found)");
            return;
        }
    };

    let doc = roxmltree::Document::parse(&xml).expect("invalid theme XML");

    let major = doc
        .descendants()
        .find(|n| n.tag_name().namespace() == Some(DML) && n.tag_name().name() == "majorFont")
        .and_then(|n| {
            n.children()
                .find(|c| c.tag_name().name() == "latin")
                .and_then(|c| c.attribute("typeface"))
        });

    let minor = doc
        .descendants()
        .find(|n| n.tag_name().namespace() == Some(DML) && n.tag_name().name() == "minorFont")
        .and_then(|n| {
            n.children()
                .find(|c| c.tag_name().name() == "latin")
                .and_then(|c| c.attribute("typeface"))
        });

    println!("  major (heading): {}", major.unwrap_or("?"));
    println!("  minor (body):    {}", minor.unwrap_or("?"));
}

// ── styles.xml ────────────────────────────────────────────────────────────────

#[derive(Default, Debug)]
struct StyleInfo {
    name: String,
    font_ascii: Option<String>,
    font_theme: Option<String>,
    size_half_pt: Option<u32>, // w:sz val (half-points)
    bold: bool,
    italic: bool,
}

impl StyleInfo {
    fn size_pt(&self) -> Option<f32> {
        self.size_half_pt.map(|v| v as f32 / 2.0)
    }

    fn font_display(&self) -> String {
        match (&self.font_ascii, &self.font_theme) {
            (Some(f), _) => f.clone(),
            (_, Some(t)) => format!("<theme:{t}>"),
            _ => "inherited".to_string(),
        }
    }
}

fn parse_rpr_into(rpr: &roxmltree::Node, info: &mut StyleInfo) {
    for child in rpr.children() {
        match child.tag_name().name() {
            "rFonts" => {
                if let Some(v) = wml(&child, "ascii") {
                    info.font_ascii = Some(v.to_string());
                }
                if let Some(v) = wml(&child, "asciiTheme") {
                    info.font_theme = Some(v.to_string());
                }
                // hAnsi as fallback
                if info.font_ascii.is_none() && info.font_theme.is_none() {
                    if let Some(v) = wml(&child, "hAnsi") {
                        info.font_ascii = Some(v.to_string());
                    }
                    if let Some(v) = wml(&child, "hAnsiTheme") {
                        info.font_theme = Some(v.to_string());
                    }
                }
            }
            "sz" => {
                if let Some(v) = wml(&child, "val").and_then(|s| s.parse().ok()) {
                    info.size_half_pt = Some(v);
                }
            }
            "b" => {
                // presence of <w:b/> without w:val="0" means bold
                let val = wml(&child, "val").unwrap_or("true");
                info.bold = val != "0" && val != "false";
            }
            "i" => {
                let val = wml(&child, "val").unwrap_or("true");
                info.italic = val != "0" && val != "false";
            }
            _ => {}
        }
    }
}

fn print_styles(archive: &mut ZipArchive<fs::File>) {
    let xml = match read_zip_text(archive, "word/styles.xml") {
        Some(s) => s,
        None => {
            println!("  (word/styles.xml not found)");
            return;
        }
    };

    let doc = roxmltree::Document::parse(&xml).expect("invalid styles XML");

    // Document defaults
    if let Some(defaults) = doc
        .descendants()
        .find(|n| n.tag_name().namespace() == Some(WML) && n.tag_name().name() == "docDefaults")
    {
        let mut info = StyleInfo::default();

        if let Some(rpr) = defaults.descendants().find(|n| {
            n.tag_name().namespace() == Some(WML) && n.tag_name().name() == "rPr"
        }) {
            parse_rpr_into(&rpr, &mut info);
        }

        let mut spacing_after: Option<u32> = None;
        let mut line_spacing: Option<String> = None;
        if let Some(ppr) = defaults.descendants().find(|n| {
            n.tag_name().namespace() == Some(WML) && n.tag_name().name() == "pPr"
        }) {
            for child in ppr.children() {
                if child.tag_name().name() == "spacing" {
                    if let Some(v) = wml(&child, "after").and_then(|s| s.parse::<u32>().ok()) {
                        spacing_after = Some(v);
                    }
                    if let Some(line) = wml(&child, "line") {
                        let rule = wml(&child, "lineRule").unwrap_or("auto");
                        line_spacing = Some(format!("{line} ({rule})"));
                    }
                }
            }
        }

        println!("  font:         {}", info.font_display());
        if let Some(pt) = info.size_pt() {
            println!("  size:         {pt}pt");
        }
        if let Some(twips) = spacing_after {
            println!("  space after:  {}pt  ({twips} twips)", twips / 20);
        }
        if let Some(ls) = line_spacing {
            println!("  line spacing: {ls}");
        }
    }

    // Named paragraph styles
    println!();
    println!("Paragraph styles:");
    println!("  {:<20}  {:<28}  {}", "style name", "font", "size");
    println!("  {}", "─".repeat(60));

    let interesting = [
        "Normal",
        "Default Paragraph Font",
        "Heading 1",
        "Heading 2",
        "Heading 3",
        "Heading 4",
        "Title",
        "Subtitle",
        "Body Text",
        "Caption",
    ];

    for style_node in doc.descendants().filter(|n| {
        n.tag_name().namespace() == Some(WML) && n.tag_name().name() == "style"
    }) {
        let style_name = style_node
            .children()
            .find(|n| n.tag_name().name() == "name")
            .and_then(|n| wml(&n, "val"))
            .unwrap_or("");

        if !interesting.contains(&style_name) {
            continue;
        }

        let mut info = StyleInfo { name: style_name.to_string(), ..Default::default() };

        if let Some(rpr) = style_node.children().find(|n| {
            n.tag_name().namespace() == Some(WML) && n.tag_name().name() == "rPr"
        }) {
            parse_rpr_into(&rpr, &mut info);
        }

        let modifiers: Vec<&str> = [info.bold.then_some("bold"), info.italic.then_some("italic")]
            .into_iter()
            .flatten()
            .collect();
        let mods = if modifiers.is_empty() { String::new() } else { format!(" ({})", modifiers.join(", ")) };

        let size_str = info.size_pt().map(|p| format!("{p}pt")).unwrap_or_else(|| "inherited".to_string());

        println!(
            "  {:<20}  {:<28}  {}{}",
            info.name,
            info.font_display(),
            size_str,
            mods
        );
    }
}

// ── document.xml run analysis ─────────────────────────────────────────────────

fn print_run_fonts(archive: &mut ZipArchive<fs::File>) {
    let xml = match read_zip_text(archive, "word/document.xml") {
        Some(s) => s,
        None => {
            println!("  (word/document.xml not found)");
            return;
        }
    };

    let doc = roxmltree::Document::parse(&xml).expect("invalid document XML");

    let mut combos: HashMap<String, usize> = HashMap::new();
    let mut run_count = 0usize;
    let mut no_rpr = 0usize;

    for run in doc.descendants().filter(|n| {
        n.tag_name().namespace() == Some(WML) && n.tag_name().name() == "r"
    }) {
        run_count += 1;

        let rpr = run.children().find(|n| {
            n.tag_name().namespace() == Some(WML) && n.tag_name().name() == "rPr"
        });

        let Some(rpr) = rpr else {
            no_rpr += 1;
            combos
                .entry("(no rPr — inherits)".to_string())
                .and_modify(|c| *c += 1)
                .or_insert(1);
            continue;
        };

        let mut info = StyleInfo::default();
        parse_rpr_into(&rpr, &mut info);

        let font = info.font_display();
        let size = info.size_pt().map(|p| format!("{p}pt")).unwrap_or_else(|| "inherited".to_string());
        let key = format!("{font} {size}");
        combos.entry(key).and_modify(|c| *c += 1).or_insert(1);
    }

    let mut sorted: Vec<_> = combos.iter().collect();
    sorted.sort_by(|a, b| b.1.cmp(a.1));

    println!("  Total runs: {run_count}  ({no_rpr} with no explicit rPr)");
    println!();
    println!("  {:>5}  {}", "count", "font + size");
    println!("  {}", "─".repeat(50));
    for (key, count) in sorted.iter().take(20) {
        println!("  {:>5}  {}", count, key);
    }
}

// ── main ──────────────────────────────────────────────────────────────────────

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() != 2 {
        eprintln!("Usage: docx-fonts <file.docx>");
        std::process::exit(1);
    }

    let file = fs::File::open(&args[1]).unwrap_or_else(|e| {
        eprintln!("Cannot open '{}': {e}", args[1]);
        std::process::exit(1);
    });
    let mut archive = ZipArchive::new(file).unwrap_or_else(|e| {
        eprintln!("Not a valid ZIP/DOCX: {e}");
        std::process::exit(1);
    });

    println!("=== {} ===", args[1]);
    println!();
    println!("Theme fonts (word/theme/theme1.xml):");
    print_theme_fonts(&mut archive);

    println!();
    println!("Document defaults (word/styles.xml):");
    print_styles(&mut archive);

    println!();
    println!("Run font usage (word/document.xml):");
    print_run_fonts(&mut archive);
}
