use std::collections::HashMap;
use std::path::{Path, PathBuf};

use lopdf::{Document, Object, ObjectId};
use memmap2::Mmap;

// ---------------------------------------------------------------------------
// CLI
// ---------------------------------------------------------------------------

fn main() {
    let args: Vec<String> = std::env::args().skip(1).collect();
    let summary_only = args.iter().any(|a| a == "--summary");
    let run_all = args.iter().any(|a| a == "--all");
    let verbose = args.iter().any(|a| a == "--verbose");

    let fixtures: Vec<PathBuf> = if run_all {
        collect_all_fixtures()
    } else {
        let path = args
            .iter()
            .find(|a| !a.starts_with("--"))
            .expect("Usage: shaping-compare <fixture-path> | --all [--summary]");
        vec![PathBuf::from(path)]
    };

    let font_index = build_font_index();
    let mut all_runs: Vec<RunComparison> = Vec::new();

    for fixture in &fixtures {
        let ref_pdf = fixture.join("reference.pdf");
        if !ref_pdf.exists() {
            continue;
        }
        match process_fixture(&ref_pdf, &font_index, verbose) {
            Ok(runs) => {
                if !summary_only && !runs.is_empty() {
                    print_fixture_report(fixture, &runs);
                }
                all_runs.extend(runs);
            }
            Err(e) => {
                eprintln!("  Error processing {}: {e}", fixture.display());
            }
        }
    }

    if all_runs.is_empty() {
        println!("No text runs extracted.");
        return;
    }

    print_aggregate(&all_runs);
}

fn collect_all_fixtures() -> Vec<PathBuf> {
    let base = Path::new("tests/fixtures");
    let mut fixtures = Vec::new();
    for subdir in &["cases", "scraped", "samples"] {
        let dir = base.join(subdir);
        if !dir.is_dir() {
            continue;
        }
        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };
        for entry in entries.flatten() {
            let p = entry.path();
            if p.is_dir() && p.join("reference.pdf").exists() {
                fixtures.push(p);
            }
        }
    }
    fixtures.sort();
    fixtures
}

// ---------------------------------------------------------------------------
// Data types
// ---------------------------------------------------------------------------

#[derive(Debug)]
struct RunComparison {
    text: String,
    font_name: String,
    font_size: f32,
    pdf_width: f32,
    rustybuzz_width: Option<f32>,
    per_char_width: Option<f32>,
}

impl RunComparison {
    /// True if text decoding likely succeeded: per-char width is within 20% of PDF width.
    /// When decoding fails (CJK, custom encodings), both rb and pc diverge wildly from PDF.
    fn text_reliable(&self) -> bool {
        if let Some(pc) = self.per_char_width {
            if self.pdf_width > 0.0 {
                let ratio = pc / self.pdf_width;
                return (0.8..=1.2).contains(&ratio);
            }
        }
        false
    }
}

struct PdfFont {
    base_font: String,
    glyph_widths: HashMap<u16, f32>,
    to_unicode: HashMap<u16, String>,
}

struct TextRun {
    text: String,
    font_name: String,
    font_size: f32,
    width: f32,
}

// ---------------------------------------------------------------------------
// PDF parsing
// ---------------------------------------------------------------------------

fn process_fixture(
    ref_pdf: &Path,
    font_index: &FontIndex,
    verbose: bool,
) -> Result<Vec<RunComparison>, String> {
    let doc = Document::load(ref_pdf).map_err(|e| format!("load: {e}"))?;
    let mut results = Vec::new();

    for page_id in doc.page_iter() {
        let fonts = extract_page_fonts(&doc, page_id)?;
        let text_runs = extract_text_runs(&doc, page_id, &fonts)?;

        for run in text_runs {
            if run.text.trim().is_empty() || run.text.len() > 200 {
                continue;
            }
            let clean_name = strip_subset_prefix(&run.font_name);
            let (rb_w, pc_w) =
                compute_comparison_widths(&clean_name, &run.text, run.font_size, font_index, verbose);
            results.push(RunComparison {
                text: run.text,
                font_name: clean_name.to_string(),
                font_size: run.font_size,
                pdf_width: run.width,
                rustybuzz_width: rb_w,
                per_char_width: pc_w,
            });
        }
    }
    Ok(results)
}

fn strip_subset_prefix(name: &str) -> &str {
    // Subset prefix is 6 uppercase letters followed by '+'
    if name.len() > 7
        && name.as_bytes()[6] == b'+'
        && name[..6].chars().all(|c| c.is_ascii_uppercase())
    {
        &name[7..]
    } else {
        name
    }
}

fn extract_page_fonts(
    doc: &Document,
    page_id: ObjectId,
) -> Result<HashMap<Vec<u8>, PdfFont>, String> {
    let mut fonts = HashMap::new();

    let page = doc
        .get_object(page_id)
        .map_err(|e| format!("page obj: {e}"))?;
    let page_dict = page.as_dict().map_err(|e| format!("page dict: {e}"))?;

    let resources = resolve_resources(doc, page_dict)?;
    let Some(resources_dict) = resources else {
        return Ok(fonts);
    };

    let font_dict = match resources_dict.get(b"Font") {
        Ok(obj) => {
            let resolved = resolve_object(doc, obj);
            match resolved.as_dict() {
                Ok(d) => Some(d.clone()),
                Err(_) => None,
            }
        }
        Err(_) => None,
    };
    let Some(font_dict) = font_dict else {
        return Ok(fonts);
    };

    for (name, obj) in font_dict.iter() {
        let font_obj = resolve_object(doc, obj);
        let Ok(fd) = font_obj.as_dict() else { continue };

        let base_font = fd
            .get(b"BaseFont")
            .ok()
            .and_then(|o| resolve_object(doc, o).as_name().ok().map(|n| n.to_vec()))
            .map(|b| String::from_utf8_lossy(&b).to_string())
            .unwrap_or_default();

        let glyph_widths = extract_cid_widths(doc, fd);
        let to_unicode = extract_to_unicode(doc, fd);

        fonts.insert(
            name.clone(),
            PdfFont {
                base_font,
                glyph_widths,
                to_unicode,
            },
        );
    }
    Ok(fonts)
}

fn resolve_resources<'a>(
    doc: &'a Document,
    page_dict: &'a lopdf::Dictionary,
) -> Result<Option<lopdf::Dictionary>, String> {
    match page_dict.get(b"Resources") {
        Ok(obj) => {
            let resolved = resolve_object(doc, obj);
            match resolved.as_dict() {
                Ok(d) => Ok(Some(d.clone())),
                Err(_) => Ok(None),
            }
        }
        Err(_) => Ok(None),
    }
}

fn resolve_object<'a>(doc: &'a Document, obj: &'a Object) -> &'a Object {
    match obj {
        Object::Reference(id) => doc.get_object(*id).unwrap_or(obj),
        _ => obj,
    }
}

fn extract_cid_widths(doc: &Document, font_dict: &lopdf::Dictionary) -> HashMap<u16, f32> {
    let mut widths = HashMap::new();

    // Try DescendantFonts for CIDFont
    if let Ok(desc_fonts) = font_dict.get(b"DescendantFonts") {
        let desc_fonts = resolve_object(doc, desc_fonts);
        if let Ok(arr) = desc_fonts.as_array() {
            for obj in arr {
                let cid_obj = resolve_object(doc, obj);
                if let Ok(cid_dict) = cid_obj.as_dict() {
                    if let Ok(w_obj) = cid_dict.get(b"W") {
                        let w_obj = resolve_object(doc, w_obj);
                        parse_w_array(w_obj, &mut widths);
                    }
                    // Also get DW (default width)
                    if let Ok(dw) = cid_dict.get(b"DW") {
                        let dw = resolve_object(doc, dw);
                        if let Ok(v) = dw.as_i64() {
                            widths.entry(0xFFFF).or_insert(v as f32);
                        }
                    }
                }
            }
        }
    }

    // Try simple Widths array + FirstChar for non-CID fonts
    if widths.is_empty() {
        if let (Ok(first_char), Ok(w_obj)) =
            (font_dict.get(b"FirstChar"), font_dict.get(b"Widths"))
        {
            let first_char = resolve_object(doc, first_char);
            let w_obj = resolve_object(doc, w_obj);
            if let (Ok(fc), Ok(arr)) = (first_char.as_i64(), w_obj.as_array()) {
                for (i, obj) in arr.iter().enumerate() {
                    let obj = resolve_object(doc, obj);
                    if let Ok(w) = obj_to_f32(obj) {
                        widths.insert((fc as u16) + i as u16, w);
                    }
                }
            }
        }
    }

    widths
}

fn parse_w_array(obj: &Object, widths: &mut HashMap<u16, f32>) {
    let Ok(arr) = obj.as_array() else { return };

    let mut i = 0;
    while i < arr.len() {
        let Some(start_gid) = obj_to_u16(&arr[i]) else {
            i += 1;
            continue;
        };
        i += 1;
        if i >= arr.len() {
            break;
        }
        match &arr[i] {
            Object::Array(inner) => {
                // [start [w1 w2 w3 ...]]
                for (j, w_obj) in inner.iter().enumerate() {
                    if let Ok(w) = obj_to_f32(w_obj) {
                        widths.insert(start_gid + j as u16, w);
                    }
                }
                i += 1;
            }
            _ => {
                // [start end width]
                if i + 1 < arr.len() {
                    let end_gid = obj_to_u16(&arr[i]).unwrap_or(start_gid);
                    i += 1;
                    if let Ok(w) = obj_to_f32(&arr[i]) {
                        for gid in start_gid..=end_gid {
                            widths.insert(gid, w);
                        }
                    }
                }
                i += 1;
            }
        }
    }
}

fn obj_to_u16(obj: &Object) -> Option<u16> {
    match obj {
        Object::Integer(n) => Some(*n as u16),
        Object::Real(n) => Some(*n as u16),
        _ => None,
    }
}

fn obj_to_f32(obj: &Object) -> Result<f32, ()> {
    match obj {
        Object::Integer(n) => Ok(*n as f32),
        Object::Real(n) => Ok(*n as f32),
        _ => Err(()),
    }
}

fn extract_to_unicode(doc: &Document, font_dict: &lopdf::Dictionary) -> HashMap<u16, String> {
    let mut map = HashMap::new();
    let Ok(tu_obj) = font_dict.get(b"ToUnicode") else {
        return map;
    };
    let tu_obj = resolve_object(doc, tu_obj);

    let cmap_data = match tu_obj {
        Object::Reference(id) => doc
            .get_object(*id)
            .ok()
            .and_then(|o| {
                if let Object::Stream(ref stream) = *o {
                    stream.decompressed_content().ok()
                } else {
                    None
                }
            }),
        Object::Stream(stream) => stream.decompressed_content().ok(),
        _ => None,
    };
    let Some(data) = cmap_data else { return map };
    let text = String::from_utf8_lossy(&data);
    parse_cmap(&text, &mut map);
    map
}

fn parse_cmap(text: &str, map: &mut HashMap<u16, String>) {
    let mut lines = text.lines();
    while let Some(line) = lines.next() {
        let line = line.trim();
        if line.ends_with("beginbfchar") {
            // Parse bfchar entries
            for line in lines.by_ref() {
                let line = line.trim();
                if line.contains("endbfchar") {
                    break;
                }
                let parts: Vec<&str> = line.split_whitespace().collect();
                if parts.len() >= 2 {
                    if let (Some(src), Some(dst)) =
                        (parse_hex_token(parts[0]), decode_unicode_hex(parts[1]))
                    {
                        map.insert(src as u16, dst);
                    }
                }
            }
        } else if line.ends_with("beginbfrange") {
            // Parse bfrange entries
            for line in lines.by_ref() {
                let line = line.trim();
                if line.contains("endbfrange") {
                    break;
                }
                let parts: Vec<&str> = line.split_whitespace().collect();
                if parts.len() >= 3 {
                    if let (Some(start), Some(end)) =
                        (parse_hex_token(parts[0]), parse_hex_token(parts[1]))
                    {
                        if parts[2].starts_with('[') {
                            // Array of individual mappings
                            let joined = parts[2..].join(" ");
                            let inner = joined
                                .trim_start_matches('[')
                                .trim_end_matches(']');
                            let tokens: Vec<&str> = inner.split_whitespace().collect();
                            for (i, tok) in tokens.iter().enumerate() {
                                if let Some(dst) = decode_unicode_hex(tok) {
                                    map.insert((start + i as u32) as u16, dst);
                                }
                            }
                        } else if let Some(dst_start) = parse_hex_token(parts[2]) {
                            for code in start..=end {
                                let offset = code - start;
                                let unicode_val = dst_start + offset;
                                if let Some(ch) = char::from_u32(unicode_val) {
                                    map.insert(code as u16, ch.to_string());
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}

fn parse_hex_token(s: &str) -> Option<u32> {
    let s = s.trim_start_matches('<').trim_end_matches('>');
    u32::from_str_radix(s, 16).ok()
}

fn decode_unicode_hex(s: &str) -> Option<String> {
    let s = s.trim_start_matches('<').trim_end_matches('>');
    if s.len() <= 4 {
        let val = u32::from_str_radix(s, 16).ok()?;
        char::from_u32(val).map(|c| c.to_string())
    } else {
        // Multi-char: pairs of 4 hex digits
        let mut result = String::new();
        let mut i = 0;
        while i + 4 <= s.len() {
            let val = u32::from_str_radix(&s[i..i + 4], 16).ok()?;
            result.push(char::from_u32(val)?);
            i += 4;
        }
        Some(result)
    }
}

// ---------------------------------------------------------------------------
// Content stream text extraction
// ---------------------------------------------------------------------------

fn extract_text_runs(
    doc: &Document,
    page_id: ObjectId,
    fonts: &HashMap<Vec<u8>, PdfFont>,
) -> Result<Vec<TextRun>, String> {
    let page = doc
        .get_object(page_id)
        .map_err(|e| format!("page: {e}"))?;
    let page_dict = page.as_dict().map_err(|e| format!("dict: {e}"))?;

    let content_data = get_page_content(doc, page_dict)?;
    let content =
        lopdf::content::Content::decode(&content_data).map_err(|e| format!("decode: {e}"))?;

    let mut runs = Vec::new();
    let mut current_font_key: Vec<u8> = Vec::new();
    let mut current_font_size: f32 = 12.0;
    let mut text_matrix_scale_y: f32 = 1.0;

    for op in &content.operations {
        match op.operator.as_str() {
            "Tf" => {
                if let (Some(name), Some(size)) = (op.operands.first(), op.operands.get(1)) {
                    if let Object::Name(ref n) = *name {
                        current_font_key = n.clone();
                    }
                    if let Ok(s) = obj_to_f32(size) {
                        current_font_size = s;
                    }
                }
            }
            "Tm" => {
                // Text matrix: [a b c d e f]
                if op.operands.len() >= 6 {
                    if let Ok(d) = obj_to_f32(&op.operands[3]) {
                        text_matrix_scale_y = d;
                    }
                }
            }
            "TJ" | "Tj" => {
                let font = fonts.get(&current_font_key);
                let effective_size = current_font_size * text_matrix_scale_y.abs();
                if effective_size < 0.5 {
                    continue;
                }

                let (text, width) = decode_text_op(&op.operator, &op.operands, font, effective_size);
                if !text.is_empty() {
                    let font_name = font
                        .map(|f| f.base_font.clone())
                        .unwrap_or_else(|| String::from_utf8_lossy(&current_font_key).to_string());
                    // Split on spaces for word-level comparison
                    for word in split_into_words(&text, width) {
                        if !word.0.is_empty() {
                            runs.push(TextRun {
                                text: word.0,
                                font_name: font_name.clone(),
                                font_size: effective_size,
                                width: word.1,
                            });
                        }
                    }
                }
            }
            _ => {}
        }
    }
    Ok(runs)
}

fn get_page_content(doc: &Document, page_dict: &lopdf::Dictionary) -> Result<Vec<u8>, String> {
    let contents = page_dict
        .get(b"Contents")
        .map_err(|e| format!("no Contents: {e}"))?;
    match contents {
        Object::Reference(id) => {
            let obj = doc.get_object(*id).map_err(|e| format!("ref: {e}"))?;
            match obj {
                Object::Stream(s) => {
                    s.decompressed_content().map_err(|e| format!("decompress: {e}"))
                }
                _ => Err("Contents not a stream".to_string()),
            }
        }
        Object::Array(arr) => {
            let mut data = Vec::new();
            for item in arr {
                if let Object::Reference(id) = item {
                    if let Ok(obj) = doc.get_object(*id) {
                        if let Object::Stream(ref s) = *obj {
                            if let Ok(d) = s.decompressed_content() {
                                if !data.is_empty() {
                                    data.push(b' ');
                                }
                                data.extend_from_slice(&d);
                            }
                        }
                    }
                }
            }
            Ok(data)
        }
        _ => Err("Unexpected Contents type".to_string()),
    }
}

fn decode_text_op(
    operator: &str,
    operands: &[Object],
    font: Option<&PdfFont>,
    font_size: f32,
) -> (String, f32) {
    let mut text = String::new();
    let mut total_width: f32 = 0.0;

    match operator {
        "TJ" => {
            if let Some(Object::Array(arr)) = operands.first() {
                for item in arr {
                    match item {
                        Object::String(bytes, _) => {
                            let (t, w) = decode_glyph_string(bytes, font, font_size);
                            text.push_str(&t);
                            total_width += w;
                        }
                        _ => {
                            // Numeric adjustment: negative = move right (increase width)
                            if let Ok(adj) = obj_to_f32(item) {
                                total_width -= adj / 1000.0 * font_size;
                            }
                        }
                    }
                }
            }
        }
        "Tj" => {
            if let Some(Object::String(bytes, _)) = operands.first() {
                let (t, w) = decode_glyph_string(bytes, font, font_size);
                text.push_str(&t);
                total_width += w;
            }
        }
        _ => {}
    }
    (text, total_width)
}

fn decode_glyph_string(bytes: &[u8], font: Option<&PdfFont>, font_size: f32) -> (String, f32) {
    let mut text = String::new();
    let mut width: f32 = 0.0;

    let Some(font) = font else {
        return (String::from_utf8_lossy(bytes).to_string(), 0.0);
    };

    let default_w = font.glyph_widths.get(&0xFFFF).copied().unwrap_or(1000.0);

    if !font.to_unicode.is_empty() {
        // CID font: bytes are big-endian glyph IDs (2 bytes each)
        let mut i = 0;
        while i + 1 < bytes.len() {
            let gid = u16::from_be_bytes([bytes[i], bytes[i + 1]]);
            if let Some(unicode) = font.to_unicode.get(&gid) {
                text.push_str(unicode);
            }
            let gw = font.glyph_widths.get(&gid).copied().unwrap_or(default_w);
            width += gw / 1000.0 * font_size;
            i += 2;
        }
    } else {
        // Simple font: bytes are character codes
        for &b in bytes {
            let ch = b as char;
            text.push(ch);
            let gw = font.glyph_widths.get(&(b as u16)).copied().unwrap_or(default_w);
            width += gw / 1000.0 * font_size;
        }
    }

    (text, width)
}

fn split_into_words(text: &str, total_width: f32) -> Vec<(String, f32)> {
    let trimmed = text.trim();
    if trimmed.is_empty() {
        return vec![];
    }
    // If short enough or no spaces, return as single run
    if !trimmed.contains(' ') || trimmed.len() <= 30 {
        return vec![(trimmed.to_string(), total_width)];
    }
    // For longer strings, we can't accurately split the width per word from TJ alone,
    // so return the whole run
    vec![(trimmed.to_string(), total_width)]
}

// ---------------------------------------------------------------------------
// Font discovery (simplified — no caching)
// ---------------------------------------------------------------------------

/// Maps lowercase lookup key → (path, face_index, is_bold, is_italic)
type FontIndex = HashMap<String, Vec<(PathBuf, u32, bool, bool)>>;

fn build_font_index() -> FontIndex {
    let t0 = std::time::Instant::now();
    let mut index = FontIndex::new();
    let dirs = font_directories();
    let mut count = 0u32;

    for dir in dirs {
        scan_font_dir(&dir, &mut index, &mut count);
    }
    eprintln!(
        "Font index: {} keys from {} files in {:.0}ms",
        index.len(),
        count,
        t0.elapsed().as_secs_f64() * 1000.0
    );
    index
}

fn scan_font_dir(dir: &Path, index: &mut FontIndex, count: &mut u32) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            scan_font_dir(&path, index, count);
            continue;
        }
        let ext = path
            .extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_ascii_lowercase());
        if !matches!(ext.as_deref(), Some("ttf" | "otf" | "ttc")) {
            continue;
        }
        *count += 1;
        let Ok(file) = std::fs::File::open(&path) else {
            continue;
        };
        let Ok(data) = (unsafe { Mmap::map(&file) }) else {
            continue;
        };
        let face_count = ttf_parser::fonts_in_collection(&data).unwrap_or(1);
        for face_idx in 0..face_count {
            let Ok(face) = ttf_parser::Face::parse(&data, face_idx) else {
                continue;
            };
            let bold = face.is_bold();
            let italic = face.is_italic();
            let entry_val = (path.clone(), face_idx, bold, italic);

            // Index by family name
            for name in face.names() {
                if name.is_unicode() {
                    if let Some(s) = name.to_string() {
                        if name.name_id == ttf_parser::name_id::FAMILY
                            || name.name_id == ttf_parser::name_id::POST_SCRIPT_NAME
                            || name.name_id == ttf_parser::name_id::FULL_NAME
                        {
                            let key = s.to_lowercase();
                            index.entry(key).or_default().push(entry_val.clone());
                        }
                    }
                }
            }
        }
    }
}

fn font_directories() -> Vec<PathBuf> {
    let mut dirs: Vec<PathBuf> = Vec::new();

    if let Ok(val) = std::env::var("DOCXSIDE_FONTS") {
        let sep = if cfg!(windows) { ';' } else { ':' };
        for part in val.split(sep) {
            let trimmed = part.trim();
            if !trimmed.is_empty() {
                dirs.push(PathBuf::from(trimmed));
            }
        }
    }

    #[cfg(target_os = "macos")]
    {
        dirs.extend([
            "/Library/Fonts".into(),
            "/System/Library/Fonts".into(),
            "/System/Library/Fonts/Supplemental".into(),
        ]);
        if let Ok(home) = std::env::var("HOME") {
            dirs.push(PathBuf::from(&home).join("Library/Fonts"));
        }
    }

    #[cfg(target_os = "linux")]
    {
        dirs.extend(["/usr/share/fonts".into(), "/usr/local/share/fonts".into()]);
        if let Ok(home) = std::env::var("HOME") {
            dirs.push(PathBuf::from(home).join(".local/share/fonts"));
        }
    }

    #[cfg(target_os = "windows")]
    {
        if let Ok(windir) = std::env::var("WINDIR") {
            dirs.push(PathBuf::from(windir).join("Fonts"));
        } else {
            dirs.push("C:\\Windows\\Fonts".into());
        }
    }

    dirs
}

// ---------------------------------------------------------------------------
// Width computation: rustybuzz vs per-char
// ---------------------------------------------------------------------------

fn ps_name_to_family(ps_name: &str) -> Option<&'static str> {
    // Common PostScript name → family name mappings from Word PDFs
    match ps_name.to_lowercase().as_str() {
        "timesnewromanpsmt" | "timesnewromanps-mt" => Some("times new roman"),
        "timesnewromanps-boldmt" => Some("times new roman"),
        "timesnewromanps-italicmt" => Some("times new roman"),
        "timesnewromanps-bolditalicmt" => Some("times new roman"),
        "arialmt" => Some("arial"),
        "arial-boldmt" => Some("arial"),
        "arial-italicmt" => Some("arial"),
        "arial-bolditalicmt" => Some("arial"),
        "couriernewpsmt" | "couriernewps-mt" => Some("courier new"),
        "couriernewps-italicmt" => Some("courier new"),
        "couriernewps-boldmt" => Some("courier new"),
        "symbolmt" => Some("symbol"),
        _ => None,
    }
}

fn parse_style_from_ps_name(name: &str) -> (bool, bool) {
    let lower = name.to_lowercase();
    let bold = lower.contains("bold");
    let italic = lower.contains("italic");
    (bold, italic)
}

fn find_font_in_index<'a>(
    font_name: &str,
    font_index: &'a FontIndex,
) -> Option<&'a (PathBuf, u32, bool, bool)> {
    let lower = font_name.to_lowercase();
    let (want_bold, want_italic) = parse_style_from_ps_name(font_name);

    // Build candidate lookup keys
    let mut keys_to_try: Vec<String> = vec![lower.clone()];

    // Try PS name alias
    if let Some(family) = ps_name_to_family(font_name) {
        keys_to_try.push(family.to_string());
    }

    // Strip style suffixes: "Aptos-Bold" → "aptos"
    let stripped = lower
        .replace("-bolditalic", "")
        .replace("-bold", "")
        .replace("-italic", "")
        .replace("-regular", "");
    if stripped != lower {
        keys_to_try.push(stripped);
    }

    for key in &keys_to_try {
        if let Some(entries) = font_index.get(key.as_str()) {
            // Try exact style match first
            if let Some(e) = entries.iter().find(|e| e.2 == want_bold && e.3 == want_italic) {
                return Some(e);
            }
            // Fall back to regular
            if let Some(e) = entries.iter().find(|e| !e.2 && !e.3) {
                return Some(e);
            }
            // Take first available
            return entries.first();
        }
    }
    None
}

fn compute_comparison_widths(
    font_name: &str,
    text: &str,
    font_size: f32,
    font_index: &FontIndex,
    verbose: bool,
) -> (Option<f32>, Option<f32>) {
    let Some(entry) = find_font_in_index(font_name, font_index) else {
        return (None, None);
    };

    if verbose {
        eprintln!("  Font match: '{}' → {} (idx={}, bold={}, italic={})",
            font_name, entry.0.display(), entry.1, entry.2, entry.3);
    }

    let Ok(file) = std::fs::File::open(&entry.0) else {
        return (None, None);
    };
    let Ok(data) = (unsafe { Mmap::map(&file) }) else {
        return (None, None);
    };

    let rb_w = rustybuzz_width(text, &data, entry.1, font_size);
    let pc_w = per_char_width(text, &data, entry.1, font_size);
    (rb_w, pc_w)
}

fn rustybuzz_width(text: &str, font_data: &[u8], face_index: u32, font_size: f32) -> Option<f32> {
    let face = rustybuzz::Face::from_slice(font_data, face_index)?;
    let mut buffer = rustybuzz::UnicodeBuffer::new();
    buffer.push_str(text);
    let output = rustybuzz::shape(&face, &[], buffer);
    let scale = font_size / face.units_per_em() as f32;
    Some(
        output
            .glyph_positions()
            .iter()
            .map(|p| p.x_advance as f32 * scale)
            .sum(),
    )
}

fn per_char_width(text: &str, font_data: &[u8], face_index: u32, font_size: f32) -> Option<f32> {
    let face = ttf_parser::Face::parse(font_data, face_index).ok()?;
    let upm = face.units_per_em() as f32;
    let scale = font_size / upm;
    let mut total: f32 = 0.0;
    for ch in text.chars() {
        if let Some(gid) = face.glyph_index(ch) {
            if let Some(adv) = face.glyph_hor_advance(gid) {
                total += adv as f32 * scale;
            }
        }
    }
    Some(total)
}

// ---------------------------------------------------------------------------
// Reporting
// ---------------------------------------------------------------------------

fn print_fixture_report(fixture: &Path, runs: &[RunComparison]) {
    let name = fixture
        .file_name()
        .unwrap_or_default()
        .to_string_lossy();
    println!("\n=== {} ({} runs) ===", name, runs.len());
    println!(
        "  {:<30} {:<20} {:>5} {:>8} {:>8} {:>8} {:>8} {:>8}",
        "text", "font", "size", "pdf_w", "rb_w", "pc_w", "d_rb", "d_pc"
    );

    let mut rb_closer = 0u32;
    let mut total_with_both = 0u32;

    for r in runs {
        let truncated: String = if r.text.len() > 28 {
            format!("{}...", &r.text[..25])
        } else {
            r.text.clone()
        };

        let rb_str = r
            .rustybuzz_width
            .map(|w| format!("{w:.2}"))
            .unwrap_or_else(|| "---".into());
        let pc_str = r
            .per_char_width
            .map(|w| format!("{w:.2}"))
            .unwrap_or_else(|| "---".into());
        let d_rb = r
            .rustybuzz_width
            .map(|w| format!("{:+.2}", w - r.pdf_width))
            .unwrap_or_else(|| "---".into());
        let d_pc = r
            .per_char_width
            .map(|w| format!("{:+.2}", w - r.pdf_width))
            .unwrap_or_else(|| "---".into());

        let font_short: String = if r.font_name.len() > 18 {
            format!("{}...", &r.font_name[..15])
        } else {
            r.font_name.clone()
        };

        println!(
            "  {:<30} {:<20} {:>5.1} {:>8.2} {:>8} {:>8} {:>8} {:>8}",
            truncated, font_short, r.font_size, r.pdf_width, rb_str, pc_str, d_rb, d_pc
        );

        if let (Some(rb), Some(pc)) = (r.rustybuzz_width, r.per_char_width) {
            total_with_both += 1;
            if (rb - r.pdf_width).abs() < (pc - r.pdf_width).abs() {
                rb_closer += 1;
            }
        }
    }

    if total_with_both > 0 {
        let mean_rb: f32 = runs
            .iter()
            .filter_map(|r| r.rustybuzz_width.map(|w| (w - r.pdf_width).abs()))
            .sum::<f32>()
            / total_with_both as f32;
        let mean_pc: f32 = runs
            .iter()
            .filter_map(|r| r.per_char_width.map(|w| (w - r.pdf_width).abs()))
            .sum::<f32>()
            / total_with_both as f32;
        println!(
            "  Summary: mean|d_rb|={mean_rb:.3}pt, mean|d_pc|={mean_pc:.3}pt, rustybuzz closer: {rb_closer}/{total_with_both} ({:.0}%)",
            rb_closer as f32 / total_with_both as f32 * 100.0
        );
    }
}

fn print_aggregate(runs: &[RunComparison]) {
    let with_both: Vec<&RunComparison> = runs
        .iter()
        .filter(|r| r.rustybuzz_width.is_some() && r.per_char_width.is_some())
        .collect();
    let no_font: usize = runs
        .iter()
        .filter(|r| r.rustybuzz_width.is_none())
        .count();

    let reliable: Vec<&RunComparison> = with_both
        .iter()
        .filter(|r| r.text_reliable())
        .copied()
        .collect();
    let unreliable = with_both.len() - reliable.len();

    println!("\n========== AGGREGATE ==========");
    println!(
        "Total runs: {}, with font match: {}, no font: {}, reliable text: {}, bad decode: {}",
        runs.len(),
        with_both.len(),
        no_font,
        reliable.len(),
        unreliable
    );

    if reliable.is_empty() {
        println!("No reliable runs for comparison.");
        if no_font > 0 {
            print_missing_fonts(runs, no_font);
        }
        return;
    }

    let mut rb_closer = 0u32;
    let mut rb_deltas: Vec<f32> = Vec::new();
    let mut pc_deltas: Vec<f32> = Vec::new();
    let mut rb_per_ch: Vec<f32> = Vec::new();
    let mut pc_per_ch: Vec<f32> = Vec::new();
    let mut font_rb_deltas: HashMap<String, Vec<f32>> = HashMap::new();
    let mut font_pc_deltas: HashMap<String, Vec<f32>> = HashMap::new();
    let mut font_rb_per_ch: HashMap<String, Vec<f32>> = HashMap::new();
    let mut font_pc_per_ch: HashMap<String, Vec<f32>> = HashMap::new();

    for r in &reliable {
        let rb = r.rustybuzz_width.unwrap();
        let pc = r.per_char_width.unwrap();
        let d_rb = (rb - r.pdf_width).abs();
        let d_pc = (pc - r.pdf_width).abs();
        let nchars = r.text.chars().count().max(1) as f32;
        rb_deltas.push(d_rb);
        pc_deltas.push(d_pc);
        rb_per_ch.push(d_rb / nchars);
        pc_per_ch.push(d_pc / nchars);
        if d_rb < d_pc {
            rb_closer += 1;
        }
        font_rb_deltas
            .entry(r.font_name.clone())
            .or_default()
            .push(d_rb);
        font_pc_deltas
            .entry(r.font_name.clone())
            .or_default()
            .push(d_pc);
        font_rb_per_ch
            .entry(r.font_name.clone())
            .or_default()
            .push(d_rb / nchars);
        font_pc_per_ch
            .entry(r.font_name.clone())
            .or_default()
            .push(d_pc / nchars);
    }

    let mean_rb: f32 = rb_deltas.iter().sum::<f32>() / rb_deltas.len() as f32;
    let mean_pc: f32 = pc_deltas.iter().sum::<f32>() / pc_deltas.len() as f32;
    let median_rb = median(&mut rb_deltas);
    let median_pc = median(&mut pc_deltas);
    let mean_rb_pch: f32 = rb_per_ch.iter().sum::<f32>() / rb_per_ch.len() as f32;
    let mean_pc_pch: f32 = pc_per_ch.iter().sum::<f32>() / pc_per_ch.len() as f32;

    println!(
        "\nReliable runs only (per-char width within ±20% of PDF width):"
    );
    println!(
        "rustybuzz closer: {rb_closer}/{} ({:.1}%)",
        reliable.len(),
        rb_closer as f32 / reliable.len() as f32 * 100.0
    );
    println!("Mean absolute delta:       rustybuzz {mean_rb:.3}pt, per-char {mean_pc:.3}pt");
    println!("Median absolute delta:     rustybuzz {median_rb:.3}pt, per-char {median_pc:.3}pt");
    println!("Mean per-character delta:   rustybuzz {mean_rb_pch:.4}pt/ch, per-char {mean_pc_pch:.4}pt/ch");

    // Per-font breakdown
    struct FontStat {
        name: String,
        count: usize,
        mean_rb: f32,
        mean_pc: f32,
        mean_rb_pch: f32,
        mean_pc_pch: f32,
        rb_closer_pct: f32,
    }

    let mut font_stats: Vec<FontStat> = font_rb_deltas
        .keys()
        .map(|font| {
            let rb = font_rb_deltas.get(font).unwrap();
            let pc = font_pc_deltas.get(font).unwrap();
            let rb_pch = font_rb_per_ch.get(font).unwrap();
            let pc_pch = font_pc_per_ch.get(font).unwrap();
            let n = rb.len();
            let rb_closer = rb.iter().zip(pc.iter()).filter(|(r, p)| r < p).count();
            FontStat {
                name: font.clone(),
                count: n,
                mean_rb: rb.iter().sum::<f32>() / n as f32,
                mean_pc: pc.iter().sum::<f32>() / n as f32,
                mean_rb_pch: rb_pch.iter().sum::<f32>() / n as f32,
                mean_pc_pch: pc_pch.iter().sum::<f32>() / n as f32,
                rb_closer_pct: rb_closer as f32 / n as f32 * 100.0,
            }
        })
        .collect();
    font_stats.sort_by(|a, b| b.count.cmp(&a.count));

    println!("\nPer-font breakdown (reliable runs only, sorted by run count):");
    println!(
        "  {:<25} {:>6} {:>8} {:>8} {:>10} {:>10} {:>7}",
        "font", "runs", "|d_rb|", "|d_pc|", "|d_rb|/ch", "|d_pc|/ch", "rb_win%"
    );
    for s in &font_stats {
        let font_short: String = if s.name.len() > 23 {
            format!("{}...", &s.name[..20])
        } else {
            s.name.clone()
        };
        println!(
            "  {:<25} {:>6} {:>8.3} {:>8.3} {:>10.4} {:>10.4} {:>6.1}%",
            font_short, s.count, s.mean_rb, s.mean_pc, s.mean_rb_pch, s.mean_pc_pch, s.rb_closer_pct
        );
    }

    // Missing fonts
    if no_font > 0 {
        print_missing_fonts(runs, no_font);
    }
}

fn print_missing_fonts(runs: &[RunComparison], no_font: usize) {
    let mut missing: HashMap<String, u32> = HashMap::new();
    for r in runs {
        if r.rustybuzz_width.is_none() {
            *missing
                .entry(strip_subset_prefix(&r.font_name).to_string())
                .or_default() += 1;
        }
    }
    let mut missing: Vec<_> = missing.into_iter().collect();
    missing.sort_by(|a, b| b.1.cmp(&a.1));
    println!("\nMissing fonts ({no_font} runs):");
    for (font, count) in missing.iter().take(20) {
        println!("  {font}: {count} runs");
    }
}

fn median(values: &mut [f32]) -> f32 {
    values.sort_by(|a, b| a.partial_cmp(b).unwrap_or(std::cmp::Ordering::Equal));
    let n = values.len();
    if n == 0 {
        return 0.0;
    }
    if n % 2 == 0 {
        (values[n / 2 - 1] + values[n / 2]) / 2.0
    } else {
        values[n / 2]
    }
}
