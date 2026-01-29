//! Compare a test case's generated PDF against the Word reference.
//! Renders both with mutool, computes Jaccard per page, prints a table.
//!
//! Usage (run from project root):
//!   case-diff <case-name>          e.g. case-diff case1
//!   case-diff <case-name> --fresh  re-render even if PNGs already exist
//!   case-diff --all                compare every case in tests/fixtures/

use image::GenericImageView;
use std::path::{Path, PathBuf};
use std::process::{Command, Stdio};
use std::{fs, process};

const DPI: &str = "150";

// ── rendering ─────────────────────────────────────────────────────────────────

fn render_pdf(pdf: &Path, out_dir: &Path) -> Result<(), String> {
    fs::create_dir_all(out_dir).map_err(|e| e.to_string())?;
    let pattern = out_dir.join("page_%03d.png");
    let status = Command::new("mutool")
        .args([
            "draw",
            "-F", "png",
            "-r", DPI,
            "-o", pattern.to_str().unwrap(),
            pdf.to_str().unwrap(),
        ])
        .stdout(Stdio::null())
        .stderr(Stdio::null())
        .status()
        .map_err(|e| format!("mutool not found: {e}"))?;
    if status.success() {
        Ok(())
    } else {
        Err(format!("mutool exited {}", status.code().unwrap_or(-1)))
    }
}

fn pngs_in(dir: &Path) -> Vec<PathBuf> {
    let Ok(entries) = fs::read_dir(dir) else {
        return Vec::new();
    };
    let mut pages: Vec<_> = entries
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("png"))
        .collect();
    pages.sort();
    pages
}

// ── image comparison ──────────────────────────────────────────────────────────

fn is_ink(r: u8, g: u8, b: u8) -> bool {
    (0.299 * r as f32 + 0.587 * g as f32 + 0.114 * b as f32) < 200.0
}

fn jaccard(a: &Path, b: &Path) -> Result<f64, String> {
    let img_a = image::open(a).map_err(|e| format!("{e}"))?;
    let img_b = image::open(b).map_err(|e| format!("{e}"))?;

    let (w, h) = img_a.dimensions();
    if img_b.dimensions() != (w, h) {
        return Err(format!(
            "size mismatch {}×{} vs {}×{}",
            w, h, img_b.dimensions().0, img_b.dimensions().1
        ));
    }

    let (mut inter, mut union) = (0u64, 0u64);
    for y in 0..h {
        for x in 0..w {
            let [ra, ga, ba, _] = img_a.get_pixel(x, y).0;
            let [rb, gb, bb, _] = img_b.get_pixel(x, y).0;
            let ai = is_ink(ra, ga, ba);
            let bi = is_ink(rb, gb, bb);
            if ai || bi { union += 1; }
            if ai && bi { inter += 1; }
        }
    }
    Ok(if union == 0 { 1.0 } else { inter as f64 / union as f64 })
}

// ── case comparison ───────────────────────────────────────────────────────────

fn compare_case(fixture_dir: &Path, output_dir: &Path, fresh: bool) {
    let name = fixture_dir.file_name().unwrap().to_string_lossy();
    let reference_pdf = fixture_dir.join("reference.pdf");
    let generated_pdf = output_dir.join("generated.pdf");

    println!("\n=== {name} ===");

    if !reference_pdf.exists() {
        println!("  SKIP: no reference.pdf at {}", reference_pdf.display());
        return;
    }
    if !generated_pdf.exists() {
        println!("  SKIP: no generated.pdf (run `cargo test` first)");
        return;
    }

    let ref_render = output_dir.join("reference");
    let gen_render = output_dir.join("generated");

    let need_ref = fresh || pngs_in(&ref_render).is_empty();
    let need_gen = fresh || pngs_in(&gen_render).is_empty();

    if need_ref {
        print!("  Rendering reference... ");
        match render_pdf(&reference_pdf, &ref_render) {
            Ok(_) => println!("ok"),
            Err(e) => { println!("FAILED: {e}"); return; }
        }
    }

    if need_gen {
        print!("  Rendering generated... ");
        match render_pdf(&generated_pdf, &gen_render) {
            Ok(_) => println!("ok"),
            Err(e) => { println!("FAILED: {e}"); return; }
        }
    }

    let ref_pages = pngs_in(&ref_render);
    let gen_pages = pngs_in(&gen_render);

    if ref_pages.is_empty() {
        println!("  No reference pages after render");
        return;
    }

    let n = ref_pages.len().min(gen_pages.len());
    let mut scores = Vec::new();

    println!();
    println!("  {:<8}  {:>7}  status", "page", "jaccard");
    println!("  {}", "─".repeat(35));

    for i in 0..n {
        let page_name = ref_pages[i].file_stem().unwrap().to_string_lossy().to_string();
        match jaccard(&ref_pages[i], &gen_pages[i]) {
            Ok(score) => {
                let status = if score >= 0.40 { "PASS" } else { "FAIL" };
                println!("  {:<8}  {:>6.2}%  {}", page_name, score * 100.0, status);
                scores.push(score);
            }
            Err(e) => println!("  {:<8}  ERROR: {e}", page_name),
        }
    }

    if ref_pages.len() != gen_pages.len() {
        println!(
            "  WARN: page count mismatch — ref={} gen={}",
            ref_pages.len(), gen_pages.len()
        );
    }

    if !scores.is_empty() {
        let avg = scores.iter().sum::<f64>() / scores.len() as f64;
        let overall = if avg >= 0.40 { "PASS" } else { "FAIL" };
        println!("  {}", "─".repeat(35));
        println!("  {:<8}  {:>6.2}%  {}", "AVERAGE", avg * 100.0, overall);
    }

    println!();
    println!("  ref renders:  {}", ref_render.display());
    println!("  gen renders:  {}", gen_render.display());
}

// ── main ──────────────────────────────────────────────────────────────────────

fn find_project_root() -> PathBuf {
    let mut dir = std::env::current_dir().unwrap();
    loop {
        if dir.join("Cargo.toml").exists() && dir.join("tests").exists() {
            return dir;
        }
        if !dir.pop() {
            eprintln!("Cannot find project root (no Cargo.toml + tests/ found)");
            process::exit(1);
        }
    }
}

fn main() {
    let args: Vec<String> = std::env::args().collect();

    if args.len() < 2 {
        eprintln!("Usage:");
        eprintln!("  case-diff <case-name>           compare one case");
        eprintln!("  case-diff <case-name> --fresh   re-render before comparing");
        eprintln!("  case-diff --all                 compare every fixture");
        process::exit(1);
    }

    let fresh = args.contains(&"--fresh".to_string());
    let root = find_project_root();
    let fixtures_dir = root.join("tests/fixtures");
    let output_dir_base = root.join("tests/output");

    if args[1] == "--all" {
        let mut cases: Vec<_> = fs::read_dir(&fixtures_dir)
            .expect("Cannot read tests/fixtures")
            .filter_map(|e| e.ok())
            .map(|e| e.path())
            .filter(|p| p.is_dir())
            .collect();
        cases.sort();
        for case in cases {
            let name = case.file_name().unwrap().to_string_lossy().to_string();
            let out = output_dir_base.join(&name);
            compare_case(&case, &out, fresh);
        }
    } else {
        let case_name = &args[1];
        let fixture = fixtures_dir.join(case_name);
        if !fixture.exists() {
            eprintln!("Fixture not found: {}", fixture.display());
            process::exit(1);
        }
        let out = output_dir_base.join(case_name);
        compare_case(&fixture, &out, fresh);
    }
}
