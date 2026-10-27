use crate::common;

use image::DynamicImage;
use rayon::prelude::*;
use std::fs;
use std::process::Command;
use std::time::Instant;

struct Row {
    name: String,
    pages: usize,
    ours_jaccard: f64,
    lo_jaccard: f64,
    ours_ssim: f64,
    lo_ssim: f64,
}

#[test]
fn libreoffice_comparison() {
    let _ = env_logger::try_init();

    // Opt-in only: spawns one soffice process per fixture, which is expensive
    // and irrelevant to the usual change-verification loop.
    if std::env::var_os("DOCXIDE_LO_COMPARE").is_none() {
        println!(
            "libreoffice_comparison skipped — opt in with DOCXIDE_LO_COMPARE=1 \
             or ./tools/run-tests.sh --libreoffice"
        );
        return;
    }

    let Some(soffice) = common::find_libreoffice() else {
        eprintln!(
            "LibreOffice not found — skipping libreoffice_comparison.\n\
             Install with: brew install libreoffice\n\
             Or set LIBREOFFICE_PATH=/path/to/soffice"
        );
        return;
    };

    let version = Command::new(&soffice)
        .arg("--version")
        .output()
        .ok()
        .and_then(|o| String::from_utf8(o.stdout).ok())
        .unwrap_or_else(|| "unknown".to_string());
    println!("\n  LibreOffice: {}", version.trim());
    println!("  soffice:     {}", soffice.display());

    let fixtures = common::discover_fixtures().expect("Failed to read tests/fixtures");
    if fixtures.is_empty() {
        println!("  (no fixtures discovered)");
        return;
    }

    let t0 = Instant::now();
    let rows: Vec<Row> = fixtures
        .par_iter()
        .filter_map(|fixture_dir| score_fixture(fixture_dir, &soffice))
        .collect();
    let elapsed = t0.elapsed().as_secs_f64();

    if rows.is_empty() {
        println!("  (no fixtures scored)");
        return;
    }

    let mut rows = rows;
    rows.sort_by(|a, b| a.name.cmp(&b.name));
    print_report(&rows, elapsed);
}

fn score_fixture(fixture_dir: &std::path::Path, soffice: &std::path::Path) -> Option<Row> {
    let name = common::case_key(fixture_dir);
    let reference_pdf = fixture_dir.join("reference.pdf");
    if !reference_pdf.exists() {
        return None;
    }

    let out_base = common::output_dir(fixture_dir);
    let ref_dir = out_base.join("reference");
    let gen_dir = out_base.join("generated");
    let lo_dir = out_base.join("libreoffice");

    let ours_pdf = common::generated_pdf_or_skip(fixture_dir, &name)?;
    let lo_pdf = match common::ensure_libreoffice_pdf(fixture_dir, soffice) {
        Ok(p) => p,
        Err(e) => {
            eprintln!("  [SKIP] {name}: LibreOffice conversion failed: {e}");
            return None;
        }
    };

    common::ensure_screenshots(&reference_pdf, &ref_dir).ok()?;
    common::ensure_screenshots(&ours_pdf, &gen_dir).ok()?;
    common::ensure_screenshots(&lo_pdf, &lo_dir).ok()?;

    let ref_pages = common::collect_page_pngs(&ref_dir).ok()?;
    let gen_pages = common::collect_page_pngs(&gen_dir).ok()?;
    let lo_pages = common::collect_page_pngs(&lo_dir).ok()?;
    if ref_pages.is_empty() {
        return None;
    }

    let lo_diff_dir = out_base.join("libreoffice_diff");
    let _ = fs::create_dir_all(&lo_diff_dir);

    let page_scores: Vec<(Option<f64>, Option<f64>, Option<f64>, Option<f64>)> = (0..ref_pages
        .len())
        .collect::<Vec<_>>()
        .par_iter()
        .map(|&i| {
            let Ok(r) = image::open(&ref_pages[i]) else {
                return (None, None, None, None);
            };
            let mut ours_j = None;
            let mut ours_s = None;
            let mut lo_j = None;
            let mut lo_s = None;
            if i < gen_pages.len()
                && let Ok(g) = image::open(&gen_pages[i])
            {
                if let Ok(pr) = common::compare_and_diff(&r, &g) {
                    ours_j = Some(pr.jaccard);
                }
                if let Ok(s) = common::ssim_score(&r, &g) {
                    ours_s = Some(s);
                }
            }
            if i < lo_pages.len()
                && let Ok(l) = image::open(&lo_pages[i])
            {
                if let Ok(pr) = common::compare_and_diff(&r, &l) {
                    lo_j = Some(pr.jaccard);
                    if let Some(stem) = ref_pages[i].file_stem().and_then(|s| s.to_str()) {
                        let _ = DynamicImage::ImageRgba8(pr.diff_img)
                            .save(lo_diff_dir.join(format!("{stem}.png")));
                    }
                }
                if let Ok(s) = common::ssim_score(&r, &l) {
                    lo_s = Some(s);
                }
            }
            (ours_j, ours_s, lo_j, lo_s)
        })
        .collect();

    let avg = |vals: Vec<f64>| -> Option<f64> {
        if vals.is_empty() {
            None
        } else {
            Some(vals.iter().sum::<f64>() / vals.len() as f64)
        }
    };
    let ours_jaccard = avg(page_scores.iter().filter_map(|p| p.0).collect())?;
    let ours_ssim = avg(page_scores.iter().filter_map(|p| p.1).collect())?;
    let lo_jaccard = avg(page_scores.iter().filter_map(|p| p.2).collect())?;
    let lo_ssim = avg(page_scores.iter().filter_map(|p| p.3).collect())?;

    Some(Row {
        name,
        pages: ref_pages.len(),
        ours_jaccard,
        lo_jaccard,
        ours_ssim,
        lo_ssim,
    })
}

fn print_report(rows: &[Row], elapsed_s: f64) {
    let name_w = common::name_width(rows.iter().map(|r| r.name.as_str()), 20);
    let pct = |v: f64| format!("{:>10}", format!("{:.1}%", v * 100.0));

    println!(
        "\n  docxide-pdf vs LibreOffice — accuracy against MS Word reference ({} fixtures, {:.1}s)",
        rows.len(),
        elapsed_s
    );
    println!("  (higher is better; Δ = docxide-pdf − LibreOffice, in percentage points)\n");

    println!(
        "  {:<name_w$}  {:>5}  {:>10}  {:>10}  {:>8}    {:>10}  {:>10}  {:>8}",
        "Case", "Pages", "Jaccard ours", "Jaccard LO", "Δ", "SSIM ours", "SSIM LO", "Δ"
    );
    println!(
        "  {}",
        "─".repeat(name_w + 5 + 10 + 10 + 8 + 10 + 10 + 8 + 17)
    );

    let mut jacc_wins = 0;
    let mut ssim_wins = 0;
    for r in rows {
        let dj = (r.ours_jaccard - r.lo_jaccard) * 100.0;
        let ds = (r.ours_ssim - r.lo_ssim) * 100.0;
        if dj > 0.0 {
            jacc_wins += 1;
        }
        if ds > 0.0 {
            ssim_wins += 1;
        }
        println!(
            "  {:<name_w$}  {:>5}  {:>10}  {:>10}  {}    {:>10}  {:>10}  {}",
            r.name,
            r.pages,
            common::color_score(r.ours_jaccard, &pct(r.ours_jaccard)),
            common::color_score(r.lo_jaccard, &pct(r.lo_jaccard)),
            color_delta(dj),
            common::color_score(r.ours_ssim, &pct(r.ours_ssim)),
            common::color_score(r.lo_ssim, &pct(r.lo_ssim)),
            color_delta(ds),
        );
    }

    println!(
        "  {}",
        "─".repeat(name_w + 5 + 10 + 10 + 8 + 10 + 10 + 8 + 17)
    );
    let n = rows.len() as f64;
    let mean_oj = rows.iter().map(|r| r.ours_jaccard).sum::<f64>() / n;
    let mean_lj = rows.iter().map(|r| r.lo_jaccard).sum::<f64>() / n;
    let mean_os = rows.iter().map(|r| r.ours_ssim).sum::<f64>() / n;
    let mean_ls = rows.iter().map(|r| r.lo_ssim).sum::<f64>() / n;
    println!(
        "  {:<name_w$}  {:>5}  {:>10}  {:>10}  {}    {:>10}  {:>10}  {}",
        "Mean",
        "",
        common::color_score(mean_oj, &pct(mean_oj)),
        common::color_score(mean_lj, &pct(mean_lj)),
        color_delta((mean_oj - mean_lj) * 100.0),
        common::color_score(mean_os, &pct(mean_os)),
        common::color_score(mean_ls, &pct(mean_ls)),
        color_delta((mean_os - mean_ls) * 100.0),
    );

    let total = rows.len();
    let jp = jacc_wins as f64 / total as f64 * 100.0;
    let sp = ssim_wins as f64 / total as f64 * 100.0;
    println!(
        "\n  docxide-pdf wins:  Jaccard {jacc_wins}/{total} ({jp:.0}%)   \
         SSIM {ssim_wins}/{total} ({sp:.0}%)"
    );
    println!("  Diff images: tests/output/<group>/<case>/libreoffice_diff/");
}

fn color_delta(pp: f64) -> String {
    let text = if pp.abs() < 0.05 {
        "  ~0".to_string()
    } else if pp > 0.0 {
        format!("+{pp:.1}pp")
    } else {
        format!("{pp:.1}pp")
    };
    let (r, g, b) = if pp > 0.05 {
        (60u8, 180, 60)
    } else if pp < -0.05 {
        (220u8, 60, 60)
    } else {
        (160u8, 160, 160)
    };
    format!("\x1b[38;2;{r};{g};{b}m{text:>8}\x1b[0m")
}
