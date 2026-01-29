//! Compute Jaccard similarity on ink pixels between two PNG images.
//! Matches the metric used in the visual_comparison test harness.
//!
//! Usage:
//!   jaccard <img_a.png> <img_b.png>
//!   jaccard <dir_a/> <dir_b/>      — compares page_NNN.png files in both dirs

use image::GenericImageView;
use std::path::Path;
use std::{fs, process};

fn is_ink(r: u8, g: u8, b: u8) -> bool {
    (0.299 * r as f32 + 0.587 * g as f32 + 0.114 * b as f32) < 200.0
}

fn jaccard(a: &Path, b: &Path) -> Result<f64, String> {
    let img_a = image::open(a).map_err(|e| format!("open {}: {e}", a.display()))?;
    let img_b = image::open(b).map_err(|e| format!("open {}: {e}", b.display()))?;

    let (w, h) = img_a.dimensions();
    if img_b.dimensions() != (w, h) {
        return Err(format!(
            "size mismatch: {}×{} vs {}×{}",
            w, h, img_b.dimensions().0, img_b.dimensions().1
        ));
    }

    let (mut intersection, mut union) = (0u64, 0u64);
    for y in 0..h {
        for x in 0..w {
            let [ra, ga, ba, _] = img_a.get_pixel(x, y).0;
            let [rb, gb, bb, _] = img_b.get_pixel(x, y).0;
            let ai = is_ink(ra, ga, ba);
            let bi = is_ink(rb, gb, bb);
            if ai || bi {
                union += 1;
            }
            if ai && bi {
                intersection += 1;
            }
        }
    }

    if union == 0 {
        Ok(1.0)
    } else {
        Ok(intersection as f64 / union as f64)
    }
}

fn collect_pngs(dir: &Path) -> Vec<std::path::PathBuf> {
    let mut pages: Vec<_> = fs::read_dir(dir)
        .unwrap_or_else(|e| {
            eprintln!("Cannot read dir {}: {e}", dir.display());
            process::exit(1);
        })
        .filter_map(|e| e.ok())
        .map(|e| e.path())
        .filter(|p| p.extension().and_then(|e| e.to_str()) == Some("png"))
        .collect();
    pages.sort();
    pages
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    if args.len() != 3 {
        eprintln!("Usage: jaccard <a.png> <b.png>");
        eprintln!("       jaccard <dir_a/> <dir_b/>");
        process::exit(1);
    }

    let a = Path::new(&args[1]);
    let b = Path::new(&args[2]);

    if a.is_dir() && b.is_dir() {
        let pages_a = collect_pngs(a);
        let pages_b = collect_pngs(b);
        let n = pages_a.len().min(pages_b.len());
        let mut scores = Vec::new();

        for i in 0..n {
            match jaccard(&pages_a[i], &pages_b[i]) {
                Ok(score) => {
                    let name = pages_a[i].file_name().unwrap().to_string_lossy();
                    println!("{name}  {:.2}%", score * 100.0);
                    scores.push(score);
                }
                Err(e) => eprintln!("page {i}: {e}"),
            }
        }

        if pages_a.len() != pages_b.len() {
            eprintln!(
                "WARN: page count differs ({} vs {})",
                pages_a.len(),
                pages_b.len()
            );
        }

        if !scores.is_empty() {
            let avg = scores.iter().sum::<f64>() / scores.len() as f64;
            println!("─────────────────────────────");
            println!("avg  {:.2}%  ({} pages)", avg * 100.0, scores.len());
        }
    } else {
        match jaccard(a, b) {
            Ok(score) => println!("{:.2}%", score * 100.0),
            Err(e) => {
                eprintln!("Error: {e}");
                process::exit(1);
            }
        }
    }
}
