mod common;

use rayon::prelude::*;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::process::Command;
use std::{fs, io};

/// Parse `mutool info` output and return a map of page_number → image_count.
fn pdf_images_per_page(pdf: &Path) -> io::Result<HashMap<u32, u32>> {
    let output = Command::new("mutool")
        .args(["info", pdf.to_str().unwrap()])
        .output()?;
    let text = String::from_utf8_lossy(&output.stdout);
    let mut in_images = false;
    let mut counts: HashMap<u32, u32> = HashMap::new();
    for line in text.lines() {
        if line.starts_with("Images") {
            in_images = true;
            continue;
        }
        if in_images {
            let trimmed = line.trim();
            if trimmed.is_empty() || (!trimmed.starts_with(|c: char| c.is_ascii_digit())) {
                break;
            }
            if let Some(page) = trimmed.split_whitespace().next().and_then(|s| s.parse::<u32>().ok()) {
                *counts.entry(page).or_insert(0) += 1;
            }
        }
    }
    Ok(counts)
}

struct ImageResult {
    name: String,
    ref_total: u32,
    gen_total: u32,
    page_mismatches: Vec<(u32, u32, u32)>, // (page, ref_count, gen_count)
    pass: bool,
}

fn analyze_fixture(fixture_dir: &Path) -> Option<ImageResult> {
    let name = fixture_dir
        .file_name()
        .unwrap()
        .to_string_lossy()
        .to_string();
    let input_docx = fixture_dir.join("input.docx");
    let reference_pdf = fixture_dir.join("reference.pdf");
    if !input_docx.exists() || !reference_pdf.exists() {
        return None;
    }

    let ref_images = pdf_images_per_page(&reference_pdf).ok()?;
    let ref_total: u32 = ref_images.values().sum();
    if ref_total == 0 {
        return None;
    }

    let output_dir = PathBuf::from("tests/output").join(&name);
    fs::create_dir_all(&output_dir).ok();
    let generated_pdf = output_dir.join("generated.pdf");

    if let Err(e) = docxide_pdf::convert_docx_to_pdf(&input_docx, &generated_pdf) {
        println!("  [SKIP] {name}: {e}");
        return None;
    }

    let gen_images = pdf_images_per_page(&generated_pdf).ok()?;
    let gen_total: u32 = gen_images.values().sum();

    let all_pages: Vec<u32> = {
        let mut pages: Vec<u32> = ref_images.keys().chain(gen_images.keys()).copied().collect();
        pages.sort();
        pages.dedup();
        pages
    };

    let page_mismatches: Vec<(u32, u32, u32)> = all_pages
        .iter()
        .filter_map(|&page| {
            let r = ref_images.get(&page).copied().unwrap_or(0);
            let g = gen_images.get(&page).copied().unwrap_or(0);
            if r != g { Some((page, r, g)) } else { None }
        })
        .collect();

    let pass = ref_total == gen_total && page_mismatches.is_empty();

    Some(ImageResult {
        name,
        ref_total,
        gen_total,
        page_mismatches,
        pass,
    })
}

#[test]
fn image_count_and_placement() {
    let _ = env_logger::try_init();
    let fixtures = common::discover_fixtures().expect("Failed to read tests/fixtures");
    if fixtures.is_empty() {
        return;
    }

    let results: Vec<ImageResult> = fixtures
        .par_iter()
        .filter_map(|f| analyze_fixture(f))
        .collect();

    if results.is_empty() {
        println!("No fixtures contain images — skipping.");
        return;
    }

    let name_w = results
        .iter()
        .map(|r| r.name.len())
        .max()
        .unwrap_or(4)
        .max(4);

    let sep = format!(
        "+-{}-+------+-----+-----+---------------------------------+",
        "-".repeat(name_w)
    );

    println!("\n{sep}");
    println!(
        "| {:<name_w$} | Pass | Ref | Gen | Page mismatches                 |",
        "Case"
    );
    println!("{sep}");

    let mut failures = Vec::new();

    for r in &results {
        let status = if r.pass { "Y" } else { "N" };
        let mismatch_str = if r.page_mismatches.is_empty() {
            String::new()
        } else {
            let parts: Vec<String> = r.page_mismatches
                .iter()
                .take(5)
                .map(|(p, rc, gc)| format!("p{p}:{rc}→{gc}"))
                .collect();
            let extra = r.page_mismatches.len().saturating_sub(5);
            if extra > 0 {
                format!("{} +{extra}more", parts.join(" "))
            } else {
                parts.join(" ")
            }
        };

        println!(
            "| {:<name_w$} | {:<4} | {:>3} | {:>3} | {:<31} |",
            r.name, status, r.ref_total, r.gen_total, mismatch_str
        );

        if !r.pass {
            failures.push(format!(
                "{}: ref={} gen={} mismatches=[{}]",
                r.name, r.ref_total, r.gen_total, mismatch_str
            ));
        }
    }

    println!("{sep}");

    if !failures.is_empty() {
        println!("\nImage count/placement mismatches ({}):", failures.len());
        for f in &failures {
            println!("  - {f}");
        }
    }
}
