/// Compare hinted vs unhinted glyph advances against Word's TJ adjustments.
///
/// Tests whether a TrueType hinting interpreter (fontdue) can reproduce
/// the micro-adjustments Word applies in its PDF output.
///
/// Usage: hinting-compare <font-path> "<text>" <font-size-pt> [dpi1,dpi2,...]
use std::path::Path;

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    if args.len() < 3 {
        eprintln!(
            "Usage: hinting-compare <font-path> \"<text>\" <font-size-pt> [dpi1,dpi2,...]\n\
             Example: hinting-compare /path/to/Aptos.ttf \"Hello, world!\" 12 72,96,150,300"
        );
        std::process::exit(1);
    }

    let font_path = &args[0];
    let text = &args[1];
    let font_size_pt: f32 = args[2].parse().expect("invalid font size");
    let dpis: Vec<f32> = if args.len() > 3 {
        args[3].split(',').filter_map(|s| s.parse().ok()).collect()
    } else {
        vec![72.0, 96.0, 120.0, 144.0, 150.0, 160.0, 200.0, 288.0, 300.0, 600.0]
    };

    let font_data = std::fs::read(font_path).expect("can't read font");

    // ttf-parser: unhinted advances
    let face = ttf_parser::Face::parse(&font_data, 0).expect("can't parse font");
    let upm = face.units_per_em() as f32;

    let chars: Vec<char> = text.chars().collect();
    let gids: Vec<u16> = chars
        .iter()
        .filter_map(|&ch| face.glyph_index(ch).map(|g| g.0))
        .collect();
    let unhinted_advances: Vec<f32> = gids
        .iter()
        .filter_map(|&gid| {
            face.glyph_hor_advance(ttf_parser::GlyphId(gid))
                .map(|a| a as f32)
        })
        .collect();

    println!("Font: {}", Path::new(font_path).file_name().unwrap().to_string_lossy());
    println!("Text: \"{}\" ({} chars, {} glyphs)", text, chars.len(), gids.len());
    println!("UPM: {}", upm);
    println!("Font size: {}pt", font_size_pt);
    println!();

    // Show unhinted advances in 1000-units
    println!("Unhinted advances (1000-units):");
    for (i, (&ch, &adv)) in chars.iter().zip(unhinted_advances.iter()).enumerate() {
        let adv_1000 = adv / upm * 1000.0;
        print!("  {}={:.1}", ch, adv_1000);
        if (i + 1) % 10 == 0 {
            println!();
        }
    }
    println!("\n");

    // fontdue: hinted advances at each DPI
    let fontdue_font = fontdue::Font::from_bytes(
        font_data.as_slice(),
        fontdue::FontSettings::default(),
    )
    .expect("fontdue can't parse font");

    // Word's case1 TJ adjustments for reference (if this is "Hello, world!" at 12pt)
    let word_adj: Option<Vec<(usize, f32)>> = if text == "Hello, world!" && (font_size_pt - 12.0).abs() < 0.1 {
        // Adjustments between char[i] and char[i+1], in 1000-units at Tf=12
        // e→l: 5, l→l: 10, l→o: 10, o→,: 31, ,→sp: -5, sp→w: -4, w→o: 13, o→r: 10, l→d: 11
        Some(vec![
            (1, 5.0),   // H→e: 5  (between index 0 and 1... wait, adjustment is AFTER glyph)
            // Actually: adj between char[i-1] and char[i]:
            // index 1: e (after H): TJ says He...then 5...then l
            // So adj at position 1-2 boundary = e→l = 5
        ])
    } else {
        None
    };

    // For each DPI, compute hinted advances and the per-glyph correction
    println!("{:<6} {:>6}  corrections (hinted - unhinted, in 1000-units)", "DPI", "ppem");
    println!("{}", "-".repeat(80));

    for &dpi in &dpis {
        let ppem = font_size_pt * dpi / 72.0;
        let font_size_px = ppem; // fontdue uses pixels = ppem

        // Get hinted advance for each glyph
        let mut corrections = Vec::new();
        let mut correction_str = String::new();

        for (i, &gid) in gids.iter().enumerate() {
            let (metrics, _) = fontdue_font.rasterize_indexed(gid as u16, font_size_px);
            let hinted_advance_px = metrics.advance_width;

            // Convert to font units for comparison
            // hinted_advance in font_units = hinted_advance_px / ppem * upm
            let hinted_fu = hinted_advance_px / ppem * upm;
            let unhinted_fu = unhinted_advances[i];
            let correction_fu = hinted_fu - unhinted_fu;
            let correction_1000 = correction_fu / upm * 1000.0;

            corrections.push(correction_1000);

            if correction_1000.abs() > 0.1 {
                correction_str.push_str(&format!(
                    "{}→{}:{:+.0} ",
                    chars[i],
                    if i + 1 < chars.len() { chars[i + 1].to_string() } else { "·".to_string() },
                    correction_1000
                ));
            }
        }

        // Compute inter-glyph adjustments (cumulative position correction)
        // The TJ adjustment between glyph i and i+1 = correction on glyph i's advance
        let mut adj_str = String::new();
        for i in 0..corrections.len().saturating_sub(1) {
            let adj = corrections[i];
            if adj.abs() > 0.1 {
                adj_str.push_str(&format!(
                    "{}→{}:{:+.0} ",
                    chars[i],
                    chars[i + 1],
                    adj
                ));
            }
        }

        println!(
            "{:<6} {:>6.1}  {}",
            format!("{:.0}", dpi),
            ppem,
            if adj_str.is_empty() { "(all zero)".to_string() } else { adj_str }
        );
    }

    // Show Word's reference values
    if text == "Hello, world!" && (font_size_pt - 12.0).abs() < 0.1 {
        println!();
        println!(
            "Word:         e→l:+5 l→l:+10 l→o:+10 o→,:+31 ,→ :-5  →w:-4 w→o:+13 o→r:+10 l→d:+11"
        );
    }
}
