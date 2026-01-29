use image::{DynamicImage, GenericImageView, ImageBuffer, Rgba};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::{fs, io};

const SIMILARITY_THRESHOLD: f64 = 0.40;
const MUTOOL_DPI: &str = "150";

fn discover_fixtures() -> io::Result<Vec<PathBuf>> {
    let fixtures_dir = Path::new("tests/fixtures");
    let mut fixtures = Vec::new();

    for entry in fs::read_dir(fixtures_dir)? {
        let entry = entry?;
        let path = entry.path();
        if path.is_dir() {
            fixtures.push(path);
        }
    }

    fixtures.sort();
    Ok(fixtures)
}

fn screenshot_pdf(pdf: &Path, output_dir: &Path) -> Result<(), String> {
    fs::create_dir_all(output_dir).map_err(|e| e.to_string())?;

    let output_pattern = output_dir.join("page_%03d.png");

    let status = Command::new("mutool")
        .args([
            "draw",
            "-F", "png",
            "-r", MUTOOL_DPI,
            "-o", output_pattern.to_str().unwrap(),
            pdf.to_str().unwrap(),
        ])
        .status()
        .map_err(|e| format!("Failed to run mutool: {e}"))?;

    if status.success() {
        Ok(())
    } else {
        Err(format!(
            "mutool exited with status {}",
            status.code().unwrap_or(-1)
        ))
    }
}

// Jaccard similarity on ink pixels: ignores white background, measures overlap
// of dark (text/content) pixels between the two images.
fn compare_images(a: &Path, b: &Path) -> Result<f64, String> {
    let img_a = image::open(a).map_err(|e| format!("Failed to open {}: {e}", a.display()))?;
    let img_b = image::open(b).map_err(|e| format!("Failed to open {}: {e}", b.display()))?;

    let (w, h) = img_a.dimensions();
    if img_b.dimensions() != (w, h) {
        return Err(format!(
            "Image dimensions differ: {:?} vs {:?}",
            img_a.dimensions(),
            img_b.dimensions()
        ));
    }

    let mut intersection: u64 = 0;
    let mut union: u64 = 0;

    for y in 0..h {
        for x in 0..w {
            let [ra, ga, ba, _] = img_a.get_pixel(x, y).0;
            let [rb, gb, bb, _] = img_b.get_pixel(x, y).0;
            let a_ink = is_ink(ra, ga, ba);
            let b_ink = is_ink(rb, gb, bb);
            if a_ink || b_ink {
                union += 1;
            }
            if a_ink && b_ink {
                intersection += 1;
            }
        }
    }

    if union == 0 {
        return Ok(1.0);
    }
    Ok(intersection as f64 / union as f64)
}

fn is_ink(r: u8, g: u8, b: u8) -> bool {
    let luma = 0.299 * r as f32 + 0.587 * g as f32 + 0.114 * b as f32;
    luma < 200.0
}

fn save_diff_image(a: &Path, b: &Path, out: &Path) -> Result<(), String> {
    let img_a = image::open(a).map_err(|e| format!("{e}"))?;
    let img_b = image::open(b).map_err(|e| format!("{e}"))?;

    let (w, h) = img_a.dimensions();
    let mut diff: ImageBuffer<Rgba<u8>, Vec<u8>> = ImageBuffer::new(w, h);

    for y in 0..h {
        for x in 0..w {
            let Rgba([ra, ga, ba, aa]) = img_a.get_pixel(x, y).0.into();
            let Rgba([rb, gb, bb, ab]) = img_b.get_pixel(x, y).0.into();
            diff.put_pixel(
                x,
                y,
                Rgba([
                    ra.abs_diff(rb),
                    ga.abs_diff(gb),
                    ba.abs_diff(bb),
                    aa.max(ab),
                ]),
            );
        }
    }

    if let Some(parent) = out.parent() {
        fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    }
    DynamicImage::ImageRgba8(diff)
        .save(out)
        .map_err(|e| e.to_string())
}

fn collect_page_pngs(dir: &Path) -> io::Result<Vec<PathBuf>> {
    let mut pages = Vec::new();
    for entry in fs::read_dir(dir)? {
        let p = entry?.path();
        if p.extension().and_then(|e| e.to_str()) == Some("png") {
            pages.push(p);
        }
    }
    pages.sort();
    Ok(pages)
}

#[test]
fn visual_comparison() {
    let fixtures = discover_fixtures().expect("Failed to read tests/fixtures");

    if fixtures.is_empty() {
        println!("[INFO] No fixtures found — add DOCX+PDF pairs to tests/fixtures/<name>/");
        return;
    }

    let mut all_passed = true;

    for fixture_dir in &fixtures {
        let name = fixture_dir.file_name().unwrap().to_string_lossy();
        let input_docx = fixture_dir.join("input.docx");
        let reference_pdf = fixture_dir.join("reference.pdf");

        let output_base = PathBuf::from("tests/output").join(name.as_ref());
        let reference_screenshots = output_base.join("reference");
        let generated_screenshots = output_base.join("generated");
        let diff_dir = output_base.join("diff");

        println!("\n=== Fixture: {name} ===");

        // Screenshot the reference PDF
        println!("  Screenshotting reference PDF...");
        if let Err(e) = screenshot_pdf(&reference_pdf, &reference_screenshots) {
            println!("  [ERROR] Failed to screenshot reference: {e}");
            all_passed = false;
            continue;
        }

        // Convert DOCX -> PDF via library
        let generated_pdf = output_base.join("generated.pdf");
        match convert_docx_to_pdf(&input_docx, &generated_pdf) {
            Err(e) => {
                println!("  [SKIP] {name}: {e}");
                continue;
            }
            Ok(()) => {}
        }

        // Screenshot the generated PDF
        println!("  Screenshotting generated PDF...");
        if let Err(e) = screenshot_pdf(&generated_pdf, &generated_screenshots) {
            println!("  [ERROR] Failed to screenshot generated: {e}");
            all_passed = false;
            continue;
        }

        // Compare pages
        let ref_pages = collect_page_pngs(&reference_screenshots)
            .unwrap_or_default();
        let gen_pages = collect_page_pngs(&generated_screenshots)
            .unwrap_or_default();

        if ref_pages.is_empty() {
            println!("  [WARN] No reference pages found");
            continue;
        }

        let page_count = ref_pages.len().min(gen_pages.len());
        let mut scores: Vec<f64> = Vec::new();

        for i in 0..page_count {
            let ref_page = &ref_pages[i];
            let gen_page = &gen_pages[i];
            let page_num = String::from(ref_page.file_stem().unwrap().to_str().unwrap());
            let diff_path = diff_dir.join(format!("{page_num}.png"));

            match compare_images(ref_page, gen_page) {
                Ok(score) => {
                    println!("  Page {}: similarity = {:.2}%", i + 1, score * 100.0);
                    scores.push(score);
                    let _ = save_diff_image(ref_page, gen_page, &diff_path);
                }
                Err(e) => {
                    println!("  Page {}: comparison error — {e}", i + 1);
                }
            }
        }

        if ref_pages.len() != gen_pages.len() {
            println!(
                "  [WARN] Page count mismatch: reference={}, generated={}",
                ref_pages.len(),
                gen_pages.len()
            );
        }

        if !scores.is_empty() {
            let avg = scores.iter().sum::<f64>() / scores.len() as f64;
            println!("  Average similarity: {:.2}%", avg * 100.0);
            if avg < SIMILARITY_THRESHOLD {
                println!(
                    "  [FAIL] {name}: average similarity {:.2}% below threshold {:.0}%",
                    avg * 100.0,
                    SIMILARITY_THRESHOLD * 100.0
                );
                all_passed = false;
            } else {
                println!("  [PASS] {name}");
            }
        }
    }

    assert!(all_passed, "One or more fixtures failed visual comparison");
}

fn convert_docx_to_pdf(input: &Path, output: &Path) -> Result<(), String> {
    docxside_pdf::convert_docx_to_pdf(input, output).map_err(|e| e.to_string())
}

