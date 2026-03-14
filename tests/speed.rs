mod common;

use std::collections::HashMap;
use std::time::Instant;

use common::Baselines;

/// Allowed regression factor for per-fixture timing.
/// A fixture that previously took 100ms can take up to 150ms before flagging.
const SPEED_REGRESSION_FACTOR: f64 = 1.5;

/// Allowed regression for total conversion time (all fixtures summed).
const TOTAL_REGRESSION_FACTOR: f64 = 1.3;

struct TimingResult {
    name: String,
    convert_ms: f64,
    previous_ms: Option<f64>,
}

#[test]
fn conversion_speed() {
    let fixtures = common::discover_fixtures().expect("discover fixtures");
    let baselines = common::read_baselines();

    let mut results: Vec<TimingResult> = Vec::new();

    for fixture in &fixtures {
        let input_docx = fixture.join("input.docx");
        if !input_docx.exists() {
            continue;
        }

        let out_dir = common::output_dir(fixture);
        std::fs::create_dir_all(&out_dir).ok();
        let output_pdf = out_dir.join("speed_test.pdf");

        let name = common::display_name(fixture);
        let previous_ms = baselines.get(&name).and_then(|b| b.convert_ms);

        let t0 = Instant::now();
        let result =
            std::panic::catch_unwind(|| docxide_pdf::convert_docx_to_pdf(&input_docx, &output_pdf));
        let elapsed_ms = t0.elapsed().as_secs_f64() * 1000.0;

        // Clean up the temp PDF
        std::fs::remove_file(&output_pdf).ok();

        if result.is_err() || result.unwrap().is_err() {
            continue;
        }

        results.push(TimingResult {
            name,
            convert_ms: elapsed_ms,
            previous_ms,
        });
    }

    // Print results table
    let mut regressions: Vec<String> = Vec::new();
    let mut total_ms = 0.0;
    let mut total_previous_ms = 0.0;
    let mut has_previous_total = true;

    println!("\n{:<40} {:>10} {:>10} {:>8}", "Fixture", "Time(ms)", "Prev(ms)", "Delta");
    println!("{}", "-".repeat(72));

    for r in &results {
        total_ms += r.convert_ms;

        let prev_str = match r.previous_ms {
            Some(p) => {
                total_previous_ms += p;
                format!("{p:.1}")
            }
            None => {
                has_previous_total = false;
                "-".to_string()
            }
        };

        let delta_str = match r.previous_ms {
            Some(p) if p > 0.0 => {
                let ratio = r.convert_ms / p;
                if ratio > SPEED_REGRESSION_FACTOR {
                    regressions.push(format!(
                        "{}: {:.0}ms → {:.0}ms ({:.0}%)",
                        r.name,
                        p,
                        r.convert_ms,
                        (ratio - 1.0) * 100.0
                    ));
                    format!("{:+.0}%  ⚠", (ratio - 1.0) * 100.0)
                } else {
                    format!("{:+.0}%", (ratio - 1.0) * 100.0)
                }
            }
            _ => String::new(),
        };

        println!(
            "{:<40} {:>10.1} {:>10} {:>8}",
            r.name, r.convert_ms, prev_str, delta_str
        );
    }

    println!("{}", "-".repeat(72));
    let total_prev_str = if has_previous_total {
        format!("{total_previous_ms:.1}")
    } else {
        "-".to_string()
    };
    println!(
        "{:<40} {:>10.1} {:>10}",
        "TOTAL", total_ms, total_prev_str
    );

    if has_previous_total && total_previous_ms > 0.0 {
        let total_ratio = total_ms / total_previous_ms;
        println!(
            "Total time ratio: {:.2}x (threshold: {:.1}x)",
            total_ratio, TOTAL_REGRESSION_FACTOR
        );
        if total_ratio > TOTAL_REGRESSION_FACTOR {
            regressions.push(format!(
                "TOTAL: {:.0}ms → {:.0}ms ({:.0}% slower)",
                total_previous_ms,
                total_ms,
                (total_ratio - 1.0) * 100.0
            ));
        }
    }

    // Update baselines with new timings
    let mut updates: HashMap<String, Baselines> = HashMap::new();
    for r in &results {
        updates.insert(
            r.name.clone(),
            Baselines {
                convert_ms: Some((r.convert_ms * 10.0).round() / 10.0),
                ..Default::default()
            },
        );
    }
    common::update_baselines(&updates);

    if !regressions.is_empty() {
        println!("\n  REGRESSION in: {}", regressions.join(", "));
    }
}
