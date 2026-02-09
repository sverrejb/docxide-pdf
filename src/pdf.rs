use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::sync::OnceLock;

use pdf_writer::{Content, Filter, Name, Pdf, Rect, Ref, Str};
use ttf_parser::Face;

use crate::error::Error;
use crate::model::{Alignment, Block, Document, Run, Table};

struct FontEntry {
    pdf_name: String,
    font_ref: Ref,
    widths_1000: Vec<f32>,       // 224 entries for WinAnsi chars 32..=255
    line_h_ratio: Option<f32>,   // (ascender - descender + lineGap) / UPM; None for Type1 fallback
    ascender_ratio: Option<f32>, // ascender / UPM; used to place baseline within the slot
}

/// (lowercase family name, bold, italic) -> file path
type FontLookup = HashMap<(String, bool, bool), PathBuf>;

static FONT_INDEX: OnceLock<FontLookup> = OnceLock::new();

fn font_family_name(face: &Face) -> Option<String> {
    // Use ID 1 (Family) — matches what DOCX references and distinguishes
    // "Aptos Display" from "Aptos" from "Aptos Narrow".
    // ID 16 (Typographic Family) groups all these under one name, causing collisions.
    for name in face.names() {
        if name.name_id == ttf_parser::name_id::FAMILY
            && name.is_unicode()
            && let Some(s) = name.to_string()
        {
            return Some(s);
        }
    }
    None
}

fn read_font_style(path: &Path) -> Option<(String, bool, bool)> {
    let data = std::fs::read(path).ok()?;
    let face = Face::parse(&data, 0).ok()?;
    let family = font_family_name(&face)?;
    Some((family, face.is_bold(), face.is_italic()))
}

fn font_directories() -> Vec<PathBuf> {
    let mut dirs: Vec<PathBuf> = Vec::new();

    // 1. User-configured directories via DOCXSIDE_FONTS env var
    if let Ok(val) = std::env::var("DOCXSIDE_FONTS") {
        let sep = if cfg!(windows) { ';' } else { ':' };
        for part in val.split(sep) {
            let trimmed = part.trim();
            if !trimmed.is_empty() {
                dirs.push(PathBuf::from(trimmed));
            }
        }
    }

    // 2. Platform-specific system font directories
    #[cfg(target_os = "macos")]
    {
        dirs.extend([
            "/Applications/Microsoft Word.app/Contents/Resources/DFonts".into(),
            "/Library/Fonts".into(),
            "/Library/Fonts/Microsoft".into(),
            "/System/Library/Fonts".into(),
            "/System/Library/Fonts/Supplemental".into(),
        ]);
        if let Ok(home) = std::env::var("HOME") {
            let cloud = PathBuf::from(&home)
                .join("Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts");
            if let Ok(families) = std::fs::read_dir(&cloud) {
                for entry in families.flatten() {
                    if entry.path().is_dir() {
                        dirs.push(entry.path());
                    }
                }
            }
        }
    }

    #[cfg(target_os = "linux")]
    {
        dirs.extend([
            "/usr/share/fonts".into(),
            "/usr/local/share/fonts".into(),
        ]);
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

fn scan_font_dirs() -> FontLookup {
    let mut index = FontLookup::new();
    let dirs = font_directories();

    // Recursive walk using a stack
    let mut stack: Vec<PathBuf> = dirs;
    while let Some(dir) = stack.pop() {
        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                stack.push(path);
                continue;
            }
            match path.extension().and_then(|e| e.to_str()) {
                Some("ttf" | "otf" | "TTF" | "OTF") => {}
                _ => continue,
            }
            if let Some((family, bold, italic)) = read_font_style(&path) {
                index
                    .entry((family.to_lowercase(), bold, italic))
                    .or_insert(path);
            }
        }
    }
    index
}

fn get_font_index() -> &'static FontLookup {
    FONT_INDEX.get_or_init(scan_font_dirs)
}

/// Look up a font file by family name and style using the OS/2 table metadata index.
/// Falls back to the regular variant if the requested bold/italic is not available.
fn find_font_file(font_name: &str, bold: bool, italic: bool) -> Option<PathBuf> {
    let index = get_font_index();
    let key = font_name.to_lowercase();
    index
        .get(&(key.clone(), bold, italic))
        .or_else(|| {
            if bold || italic {
                index.get(&(key, false, false))
            } else {
                None
            }
        })
        .cloned()
}

/// Windows-1252 (WinAnsi) byte to Unicode char mapping.
/// Bytes 0x80-0x9F are remapped; all others map directly to their Unicode codepoint.
fn winansi_to_char(byte: u8) -> char {
    match byte {
        0x80 => '\u{20AC}',
        0x82 => '\u{201A}',
        0x83 => '\u{0192}',
        0x84 => '\u{201E}',
        0x85 => '\u{2026}',
        0x86 => '\u{2020}',
        0x87 => '\u{2021}',
        0x88 => '\u{02C6}',
        0x89 => '\u{2030}',
        0x8A => '\u{0160}',
        0x8B => '\u{2039}',
        0x8C => '\u{0152}',
        0x8E => '\u{017D}',
        0x91 => '\u{2018}',
        0x92 => '\u{2019}',
        0x93 => '\u{201C}',
        0x94 => '\u{201D}',
        0x95 => '\u{2022}', // bullet
        0x96 => '\u{2013}',
        0x97 => '\u{2014}',
        0x98 => '\u{02DC}',
        0x99 => '\u{2122}',
        0x9A => '\u{0161}',
        0x9B => '\u{203A}',
        0x9C => '\u{0153}',
        0x9E => '\u{017E}',
        0x9F => '\u{0178}',
        _ => byte as char,
    }
}

/// Convert a UTF-8 string to WinAnsi (Windows-1252) bytes for PDF Str encoding.
fn to_winansi_bytes(s: &str) -> Vec<u8> {
    s.chars()
        .filter_map(|c| match c as u32 {
            0x0000..=0x007F => Some(c as u8),
            0x00A0..=0x00FF => Some(c as u8), // Latin-1 supplement maps directly
            0x20AC => Some(0x80),
            0x201A => Some(0x82),
            0x0192 => Some(0x83),
            0x201E => Some(0x84),
            0x2026 => Some(0x85),
            0x2020 => Some(0x86),
            0x2021 => Some(0x87),
            0x02C6 => Some(0x88),
            0x2030 => Some(0x89),
            0x0160 => Some(0x8A),
            0x2039 => Some(0x8B),
            0x0152 => Some(0x8C),
            0x017D => Some(0x8E),
            0x2018 => Some(0x91),
            0x2019 => Some(0x92),
            0x201C => Some(0x93),
            0x201D => Some(0x94),
            0x2022 => Some(0x95), // bullet
            0x2013 => Some(0x96),
            0x2014 => Some(0x97),
            0x02DC => Some(0x98),
            0x2122 => Some(0x99),
            0x0161 => Some(0x9A),
            0x203A => Some(0x9B),
            0x0153 => Some(0x9C),
            0x017E => Some(0x9E),
            0x0178 => Some(0x9F),
            _ => None,
        })
        .collect()
}

/// Approximate Helvetica widths at 1000 units/em for WinAnsi chars 32..=255.
fn helvetica_widths() -> Vec<f32> {
    (32u8..=255u8)
        .map(|b| match b {
            32 => 278.0,                          // space
            33..=47 => 333.0,                     // punctuation
            48..=57 => 556.0,                     // digits
            58..=64 => 333.0,                     // more punctuation
            73 | 74 => 278.0,                     // I J (narrow uppercase)
            77 => 833.0,                          // M (wide)
            65..=90 => 667.0,                     // uppercase A-Z (average)
            91..=96 => 333.0,                     // brackets etc.
            102 | 105 | 106 | 108 | 116 => 278.0, // narrow lowercase: f i j l t
            109 | 119 => 833.0,                   // m w (wide)
            97..=122 => 556.0,                    // lowercase a-z (average)
            _ => 556.0,
        })
        .collect()
}

/// Embed a TrueType/OpenType font (raw bytes) into the PDF.
fn embed_truetype(
    pdf: &mut Pdf,
    font_ref: Ref,
    descriptor_ref: Ref,
    data_ref: Ref,
    font_name: &str,
    font_data: &[u8],
) -> Option<(Vec<f32>, f32, f32)> {
    let face = Face::parse(font_data, 0).ok()?;

    let units = face.units_per_em() as f32;
    let ascent = face.ascender() as f32 / units * 1000.0;
    let descent = face.descender() as f32 / units * 1000.0;
    let cap_height = face
        .capital_height()
        .map(|h| h as f32 / units * 1000.0)
        .unwrap_or(700.0);

    let bb = face.global_bounding_box();
    let bbox = Rect::new(
        bb.x_min as f32 / units * 1000.0,
        bb.y_min as f32 / units * 1000.0,
        bb.x_max as f32 / units * 1000.0,
        bb.y_max as f32 / units * 1000.0,
    );

    let widths: Vec<f32> = (32u8..=255u8)
        .map(|byte| {
            face.glyph_index(winansi_to_char(byte))
                .and_then(|gid| face.glyph_hor_advance(gid))
                .map(|adv| adv as f32 / units * 1000.0)
                .unwrap_or(0.0)
        })
        .collect();

    let data_len = i32::try_from(font_data.len()).ok()?;
    pdf.stream(data_ref, font_data)
        .pair(Name(b"Length1"), data_len);

    let ps_name = font_name.replace(' ', "");

    pdf.font_descriptor(descriptor_ref)
        .name(Name(ps_name.as_bytes()))
        .flags(pdf_writer::types::FontFlags::NON_SYMBOLIC)
        .bbox(bbox)
        .italic_angle(0.0)
        .ascent(ascent)
        .descent(descent)
        .cap_height(cap_height)
        .stem_v(80.0)
        .font_file2(data_ref);

    {
        let mut d = pdf.indirect(font_ref).dict();
        d.pair(Name(b"Type"), Name(b"Font"));
        d.pair(Name(b"Subtype"), Name(b"TrueType"));
        d.pair(Name(b"BaseFont"), Name(ps_name.as_bytes()));
        d.pair(Name(b"Encoding"), Name(b"WinAnsiEncoding"));
        d.pair(Name(b"FirstChar"), 32i32);
        d.pair(Name(b"LastChar"), 255i32);
        d.pair(Name(b"FontDescriptor"), descriptor_ref);
        d.insert(Name(b"Widths"))
            .array()
            .items(widths.iter().copied());
    }

    let line_gap = face.line_gap() as f32;
    let line_h_ratio = (face.ascender() as f32 - face.descender() as f32 + line_gap) / units;
    let ascender_ratio = face.ascender() as f32 / units;

    Some((widths, line_h_ratio, ascender_ratio))
}

fn primary_font_name(name: &str) -> &str {
    name.split(';').next().unwrap_or(name).trim()
}

fn font_key(run: &Run) -> String {
    let base = primary_font_name(&run.font_name);
    match (run.bold, run.italic) {
        (true, true) => format!("{}/BI", base),
        (true, false) => format!("{}/B", base),
        (false, true) => format!("{}/I", base),
        (false, false) => base.to_string(),
    }
}

type EmbeddedFonts = HashMap<(String, bool, bool), Vec<u8>>;

fn register_font(
    pdf: &mut Pdf,
    font_name: &str,
    bold: bool,
    italic: bool,
    pdf_name: String,
    alloc: &mut impl FnMut() -> Ref,
    embedded_fonts: &EmbeddedFonts,
) -> FontEntry {
    let font_ref = alloc();
    let descriptor_ref = alloc();
    let data_ref = alloc();

    let embedded_key = (font_name.to_lowercase(), bold, italic);
    let embedded_data = embedded_fonts.get(&embedded_key);

    let (widths, line_h_ratio, ascender_ratio) = embedded_data
        .and_then(|data| {
            embed_truetype(pdf, font_ref, descriptor_ref, data_ref, font_name, data)
        })
        .or_else(|| {
            find_font_file(font_name, bold, italic).and_then(|path| {
                let data = std::fs::read(&path).ok()?;
                embed_truetype(pdf, font_ref, descriptor_ref, data_ref, font_name, &data)
            })
        })
        .map(|(w, r, ar)| (w, Some(r), Some(ar)))
        .unwrap_or_else(|| {
            log::warn!("Font not found: {font_name} bold={bold} italic={italic} — using Helvetica");
            pdf.type1_font(font_ref)
                .base_font(Name(b"Helvetica"))
                .encoding_predefined(Name(b"WinAnsiEncoding"));
            (helvetica_widths(), None, None)
        });

    FontEntry {
        pdf_name,
        font_ref,
        widths_1000: widths,
        line_h_ratio,
        ascender_ratio,
    }
}

struct WordChunk {
    pdf_font: String,
    text: String,
    font_size: f32,
    color: Option<[u8; 3]>,
    x_offset: f32, // x relative to line start
    width: f32,
}

struct TextLine {
    chunks: Vec<WordChunk>,
    total_width: f32,
}

fn finish_line(chunks: &mut Vec<WordChunk>) -> TextLine {
    let total_width = chunks.last().map(|c| c.x_offset + c.width).unwrap_or(0.0);
    TextLine {
        chunks: std::mem::take(chunks),
        total_width,
    }
}

/// Layout runs into wrapped lines.
fn build_paragraph_lines(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    max_width: f32,
) -> Vec<TextLine> {
    let mut lines: Vec<TextLine> = Vec::new();
    let mut current_chunks: Vec<WordChunk> = Vec::new();
    let mut current_x: f32 = 0.0;

    for run in runs {
        let key = font_key(run);
        let entry = seen_fonts.get(&key).expect("font registered");
        let space_w = entry.widths_1000[0] * run.font_size / 1000.0;

        for word in run.text.split_whitespace() {
            let ww: f32 = to_winansi_bytes(word)
                .iter()
                .filter(|&&b| b >= 32)
                .map(|&b| entry.widths_1000[(b - 32) as usize] * run.font_size / 1000.0)
                .sum();

            if !current_chunks.is_empty() && current_x + ww > max_width {
                lines.push(finish_line(&mut current_chunks));
                current_x = 0.0;
            }

            current_chunks.push(WordChunk {
                pdf_font: entry.pdf_name.clone(),
                text: word.to_string(),
                font_size: run.font_size,
                color: run.color,
                x_offset: current_x,
                width: ww,
            });
            current_x += ww + space_w;
        }
    }

    if !current_chunks.is_empty() {
        lines.push(finish_line(&mut current_chunks));
    }

    if lines.is_empty() {
        lines.push(TextLine {
            chunks: vec![],
            total_width: 0.0,
        });
    }
    lines
}

/// Render pre-built lines applying the paragraph alignment.
/// `total_line_count` is the full paragraph line count (for justify: last line stays left-aligned).
fn render_paragraph_lines(
    content: &mut Content,
    lines: &[TextLine],
    alignment: &Alignment,
    margin_left: f32,
    text_width: f32,
    first_baseline_y: f32,
    line_pitch: f32,
    total_line_count: usize,
    first_line_index: usize,
) {
    let mut current_color: Option<[u8; 3]> = None;

    let last_line_idx = total_line_count.saturating_sub(1);
    for (line_num, line) in lines.iter().enumerate() {
        let y = first_baseline_y - line_num as f32 * line_pitch;
        let global_line_idx = first_line_index + line_num;

        let is_justified = *alignment == Alignment::Justify
            && global_line_idx != last_line_idx
            && line.chunks.len() > 1;

        let line_start_x = match alignment {
            Alignment::Center => margin_left + (text_width - line.total_width) / 2.0,
            Alignment::Right => margin_left + text_width - line.total_width,
            Alignment::Left | Alignment::Justify => margin_left,
        };

        let extra_per_gap = if is_justified {
            (text_width - line.total_width) / (line.chunks.len() - 1) as f32
        } else {
            0.0
        };

        for (chunk_idx, chunk) in line.chunks.iter().enumerate() {
            let x = line_start_x + chunk.x_offset + chunk_idx as f32 * extra_per_gap;
            if chunk.color != current_color {
                if let Some([r, g, b]) = chunk.color {
                    content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
                } else {
                    content.set_fill_gray(0.0);
                }
                current_color = chunk.color;
            }
            let text_bytes = to_winansi_bytes(&chunk.text);
            content
                .begin_text()
                .set_font(Name(chunk.pdf_font.as_bytes()), chunk.font_size)
                .next_line(x, y)
                .show(Str(&text_bytes))
                .end_text();
        }
    }
    if current_color.is_some() {
        content.set_fill_gray(0.0);
    }
}

fn font_metric(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
    get: impl Fn(&FontEntry) -> Option<f32>,
) -> Option<f32> {
    runs.first()
        .map(font_key)
        .and_then(|k| seen_fonts.get(&k))
        .and_then(get)
}

/// Compute the effective font_size, line_h_ratio, and ascender_ratio for a set of runs
/// by picking the run that produces the tallest visual ascent (font_size * ascender_ratio).
fn tallest_run_metrics(
    runs: &[Run],
    seen_fonts: &HashMap<String, FontEntry>,
) -> (f32, Option<f32>, Option<f32>) {
    let mut best_font_size = runs.first().map_or(12.0, |r| r.font_size);
    let mut best_ascent = 0.0f32;
    let mut best_line_h_ratio: Option<f32> = None;
    let mut best_ascender_ratio: Option<f32> = None;

    for run in runs {
        let key = font_key(run);
        let entry = seen_fonts.get(&key);
        let ar = entry.and_then(|e| e.ascender_ratio).unwrap_or(0.75);
        let ascent = run.font_size * ar;
        if ascent > best_ascent {
            best_ascent = ascent;
            best_font_size = run.font_size;
            best_ascender_ratio = entry.and_then(|e| e.ascender_ratio);
            best_line_h_ratio = entry.and_then(|e| e.line_h_ratio);
        }
    }
    (best_font_size, best_line_h_ratio, best_ascender_ratio)
}

const TABLE_CELL_PAD_LEFT: f32 = 5.4;
const TABLE_CELL_PAD_TOP: f32 = 0.0;
const TABLE_CELL_PAD_BOTTOM: f32 = 0.0;
const TABLE_BORDER_WIDTH: f32 = 0.5;

/// Auto-fit column widths so that the longest non-breakable word in each column
/// fits within the cell (including padding). Columns that need more space grow;
/// other columns shrink proportionally. Total width is preserved.
fn auto_fit_columns(
    table: &Table,
    seen_fonts: &HashMap<String, FontEntry>,
) -> Vec<f32> {
    let ncols = table.col_widths.len();
    if ncols == 0 {
        return table.col_widths.clone();
    }

    let mut min_widths = vec![0.0f32; ncols];

    for row in &table.rows {
        for (ci, cell) in row.cells.iter().enumerate() {
            if ci >= ncols {
                break;
            }
            for para in &cell.paragraphs {
                for run in &para.runs {
                    let key = font_key(run);
                    let Some(entry) = seen_fonts.get(&key) else {
                        continue;
                    };
                    for word in run.text.split_whitespace() {
                        let ww: f32 = to_winansi_bytes(word)
                            .iter()
                            .filter(|&&b| b >= 32)
                            .map(|&b| entry.widths_1000[(b - 32) as usize] * run.font_size / 1000.0)
                            .sum();
                        min_widths[ci] = min_widths[ci].max(ww);
                    }
                }
            }
        }
    }

    let total: f32 = table.col_widths.iter().sum();
    let mut widths = table.col_widths.clone();

    // Expand columns that need it, track how much extra space is needed
    let mut extra_needed: f32 = 0.0;
    let mut shrinkable: f32 = 0.0;
    for i in 0..ncols {
        if min_widths[i] > widths[i] {
            extra_needed += min_widths[i] - widths[i];
            widths[i] = min_widths[i];
        } else {
            shrinkable += widths[i] - min_widths[i];
        }
    }

    if extra_needed > 0.0 && shrinkable > 0.0 {
        let factor = extra_needed.min(shrinkable) / shrinkable;
        for i in 0..ncols {
            if widths[i] > min_widths[i] {
                let available = widths[i] - min_widths[i];
                widths[i] -= available * factor;
            }
        }
        // Normalize to preserve total
        let new_total: f32 = widths.iter().sum();
        if (new_total - total).abs() > 0.01 {
            let scale = total / new_total;
            for w in &mut widths {
                *w *= scale;
            }
        }
    }

    widths
}

struct RowLayout {
    height: f32,
    cell_lines: Vec<(Vec<TextLine>, f32, f32)>, // (lines, line_h, font_size) per cell
}

fn compute_row_layouts(
    table: &Table,
    col_widths: &[f32],
    doc: &Document,
    seen_fonts: &HashMap<String, FontEntry>,
) -> Vec<RowLayout> {
    table
        .rows
        .iter()
        .map(|row| {
            let mut max_h: f32 = 0.0;
            let cell_lines: Vec<(Vec<TextLine>, f32, f32)> = row
                .cells
                .iter()
                .enumerate()
                .map(|(ci, cell)| {
                    let col_w = col_widths.get(ci).copied().unwrap_or(cell.width);
                    let cell_text_w = col_w;
                    let mut total_h: f32 = TABLE_CELL_PAD_TOP + TABLE_CELL_PAD_BOTTOM;
                    let mut all_lines = Vec::new();
                    let mut first_font_size = 12.0f32;
                    let mut first_line_h = 14.4f32;

                    for para in &cell.paragraphs {
                        let font_size = para.runs.first().map_or(12.0, |r| r.font_size);
                        let effective_ls = para.line_spacing.unwrap_or(doc.line_spacing);
                        let line_h = font_metric(&para.runs, seen_fonts, |e| e.line_h_ratio)
                            .map(|ratio| font_size * ratio * effective_ls)
                            .unwrap_or(font_size * 1.2);

                        if all_lines.is_empty() {
                            first_font_size = font_size;
                            first_line_h = line_h;
                        }

                        if !para.runs.is_empty() {
                            let lines = build_paragraph_lines(&para.runs, seen_fonts, cell_text_w);
                            total_h += lines.len() as f32 * line_h;
                            all_lines.extend(lines);
                        }
                    }

                    max_h = max_h.max(total_h);
                    (all_lines, first_line_h, first_font_size)
                })
                .collect();

            RowLayout {
                height: max_h + TABLE_BORDER_WIDTH,
                cell_lines,
            }
        })
        .collect()
}

fn render_table(
    table: &Table,
    doc: &Document,
    seen_fonts: &HashMap<String, FontEntry>,
    content: &mut Content,
    all_contents: &mut Vec<Content>,
    slot_top: &mut f32,
    prev_space_after: f32,
) {
    let col_widths = auto_fit_columns(table, seen_fonts);
    let row_layouts = compute_row_layouts(table, &col_widths, doc, seen_fonts);

    *slot_top -= prev_space_after;

    for (ri, (row, layout)) in table.rows.iter().zip(row_layouts.iter()).enumerate() {
        let row_h = layout.height;
        log::debug!(
            "TABLE row={} row_h={:.2} cells={} slot_top={:.2}",
            ri,
            row_h,
            layout.cell_lines.len(),
            *slot_top
        );
        let at_page_top = (*slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;

        if !at_page_top && *slot_top - row_h < doc.margin_bottom {
            all_contents.push(std::mem::replace(content, Content::new()));
            *slot_top = doc.page_height - doc.margin_top;
        }

        let row_top = *slot_top;
        let row_bottom = row_top - row_h;

        // Render cell contents — text inset by cell padding
        let mut cell_x = doc.margin_left;
        for (ci, (cell, (lines, line_h, font_size))) in
            row.cells.iter().zip(layout.cell_lines.iter()).enumerate()
        {
            let col_w = col_widths.get(ci).copied().unwrap_or(cell.width);
            let text_x = cell_x + TABLE_CELL_PAD_LEFT;
            let text_w = col_w;

            if !lines.is_empty() && !lines.iter().all(|l| l.chunks.is_empty()) {
                let first_run = cell.paragraphs.first().and_then(|p| p.runs.first());
                let ascender_ratio = first_run
                    .map(font_key)
                    .and_then(|k| seen_fonts.get(&k))
                    .and_then(|e| e.ascender_ratio)
                    .unwrap_or(0.75);
                let baseline_y = row_top - TABLE_CELL_PAD_TOP - font_size * ascender_ratio;
                let alignment = cell
                    .paragraphs
                    .first()
                    .map(|p| p.alignment)
                    .unwrap_or(Alignment::Left);

                render_paragraph_lines(
                    content,
                    lines,
                    &alignment,
                    text_x,
                    text_w,
                    baseline_y,
                    *line_h,
                    lines.len(),
                    0,
                );
            }

            cell_x += col_w;
        }

        // Draw cell borders — first cell extends left by pad_left,
        // right border aligns with body text right edge.
        content.save_state();
        content.set_line_width(TABLE_BORDER_WIDTH);
        let mut bx = doc.margin_left - TABLE_CELL_PAD_LEFT;
        for (ci, cell) in row.cells.iter().enumerate() {
            let col_w = col_widths.get(ci).copied().unwrap_or(cell.width);
            let border_w = if ci == 0 {
                col_w + TABLE_CELL_PAD_LEFT
            } else {
                col_w
            };
            content.rect(bx, row_bottom, border_w, row_h).stroke();
            bx += border_w;
        }
        content.restore_state();

        *slot_top = row_bottom;
    }
}

pub fn render(doc: &Document) -> Result<Vec<u8>, Error> {
    let mut pdf = Pdf::new();
    let mut next_id = 1i32;
    let mut alloc = || {
        let r = Ref::new(next_id);
        next_id += 1;
        r
    };

    let catalog_id = alloc();
    let pages_id = alloc();

    // Phase 1: collect unique font names (with variant) and embed them
    let mut seen_fonts: HashMap<String, FontEntry> = HashMap::new();
    let mut font_order: Vec<String> = Vec::new();

    // Collect all runs from all blocks (paragraphs and table cells)
    let all_runs: Vec<&Run> = doc
        .blocks
        .iter()
        .flat_map(|block| -> Box<dyn Iterator<Item = &Run> + '_> {
            match block {
                Block::Paragraph(para) => Box::new(para.runs.iter()),
                Block::Table(table) => Box::new(
                    table
                        .rows
                        .iter()
                        .flat_map(|row| row.cells.iter())
                        .flat_map(|cell| cell.paragraphs.iter())
                        .flat_map(|para| para.runs.iter()),
                ),
            }
        })
        .collect();

    for run in &all_runs {
        let key = font_key(run);
        if !seen_fonts.contains_key(&key) {
            let base = primary_font_name(&run.font_name);
            let pdf_name = format!("F{}", font_order.len() + 1);
            let entry = register_font(
                &mut pdf,
                base,
                run.bold,
                run.italic,
                pdf_name,
                &mut alloc,
                &doc.embedded_fonts,
            );
            seen_fonts.insert(key.clone(), entry);
            font_order.push(key);
        }
    }

    if seen_fonts.is_empty() {
        let pdf_name = "F1".to_string();
        let entry = register_font(
            &mut pdf,
            "Helvetica",
            false,
            false,
            pdf_name,
            &mut alloc,
            &doc.embedded_fonts,
        );
        seen_fonts.insert("Helvetica".to_string(), entry);
        font_order.push("Helvetica".to_string());
    }

    let text_width = doc.page_width - doc.margin_left - doc.margin_right;

    // Phase 1b: embed images
    let mut image_pdf_names: HashMap<usize, String> = HashMap::new();
    let mut image_xobjects: Vec<(String, Ref)> = Vec::new();
    for (block_idx, block) in doc.blocks.iter().enumerate() {
        if let Block::Paragraph(para) = block
            && let Some(img) = &para.image
        {
            let xobj_ref = alloc();
            let pdf_name = format!("Im{}", image_xobjects.len() + 1);

            let mut xobj = pdf.image_xobject(xobj_ref, &img.data);
            xobj.filter(Filter::DctDecode);
            xobj.width(img.pixel_width as i32);
            xobj.height(img.pixel_height as i32);
            xobj.color_space().device_rgb();
            xobj.bits_per_component(8);

            image_xobjects.push((pdf_name.clone(), xobj_ref));
            image_pdf_names.insert(block_idx, pdf_name);
        }
    }

    // Phase 2: build multi-page content streams
    let mut all_contents: Vec<Content> = Vec::new();
    let mut current_content = Content::new();
    let mut slot_top = doc.page_height - doc.margin_top;
    let mut prev_space_after: f32 = 0.0;

    let adjacent_para = |idx: usize| -> Option<&crate::model::Paragraph> {
        match doc.blocks.get(idx)? {
            Block::Paragraph(p) => Some(p),
            Block::Table(_) => None,
        }
    };

    for (block_idx, block) in doc.blocks.iter().enumerate() {
        match block {
            Block::Paragraph(para) => {
                let next_para = adjacent_para(block_idx + 1);
                let prev_para = if block_idx > 0 {
                    adjacent_para(block_idx - 1)
                } else {
                    None
                };

                let effective_space_before =
                    if para.contextual_spacing && prev_para.is_some_and(|p| p.contextual_spacing) {
                        0.0
                    } else {
                        para.space_before
                    };
                let effective_space_after =
                    if para.contextual_spacing && next_para.is_some_and(|p| p.contextual_spacing) {
                        0.0
                    } else {
                        para.space_after
                    };

                let mut inter_gap = f32::max(prev_space_after, effective_space_before);

                let (font_size, tallest_lhr, tallest_ar) =
                    tallest_run_metrics(&para.runs, &seen_fonts);
                let effective_line_spacing = para.line_spacing.unwrap_or(doc.line_spacing);
                let line_h = tallest_lhr
                    .map(|ratio| font_size * ratio * effective_line_spacing)
                    .unwrap_or(font_size * 1.2);

                let para_text_x = doc.margin_left + para.indent_left;
                let para_text_width = (text_width - para.indent_left).max(1.0);
                let label_x = doc.margin_left + (para.indent_left - para.indent_hanging).max(0.0);

                let lines = if para.image.is_some() || para.runs.is_empty() {
                    vec![]
                } else {
                    build_paragraph_lines(&para.runs, &seen_fonts, para_text_width)
                };

                let content_h = if para.image.is_some() || para.runs.is_empty() {
                    para.content_height.max(doc.line_pitch)
                } else {
                    lines.len() as f32 * line_h
                };

                let needed = inter_gap + content_h + effective_space_after;
                let at_page_top = (slot_top - (doc.page_height - doc.margin_top)).abs() < 1.0;

                let keep_next_extra = if para.keep_next {
                    next_para.map_or(0.0, |next| {
                        let (nfs, nlhr, _) = tallest_run_metrics(&next.runs, &seen_fonts);
                        let next_inter = f32::max(effective_space_after, next.space_before);
                        let next_first_line_h = nlhr
                            .map(|ratio| nfs * ratio)
                            .unwrap_or(nfs * 1.2);
                        next_inter + next_first_line_h
                    })
                } else {
                    0.0
                };

                if !at_page_top && slot_top - needed - keep_next_extra < doc.margin_bottom {
                    let available = slot_top - inter_gap - doc.margin_bottom;
                    let first_line_h = tallest_lhr
                        .map(|ratio| font_size * ratio)
                        .unwrap_or(font_size);
                    let mut lines_that_fit = if line_h > 0.0 && available >= first_line_h {
                        1 + ((available - first_line_h) / line_h).floor() as usize
                    } else {
                        0
                    };

                    // Reduce to ensure at least 2 lines remain on next page (orphan control)
                    if lines_that_fit > 0 && lines.len().saturating_sub(lines_that_fit) < 2 {
                        lines_that_fit = lines.len().saturating_sub(2);
                    }

                    if lines_that_fit >= 2 && lines_that_fit < lines.len() {
                        let first_part = &lines[..lines_that_fit];
                        slot_top -= inter_gap;
                        let ascender_ratio = tallest_ar.unwrap_or(0.75);
                        let baseline_y = slot_top - font_size * ascender_ratio;

                        if !para.list_label.is_empty() {
                            let (label_font_name, label_bytes) =
                                label_for_run(&para.runs[0], &seen_fonts, &para.list_label);
                            current_content
                                .begin_text()
                                .set_font(Name(label_font_name.as_bytes()), font_size)
                                .next_line(label_x, baseline_y)
                                .show(Str(&label_bytes))
                                .end_text();
                        }

                        render_paragraph_lines(
                            &mut current_content,
                            first_part,
                            &para.alignment,
                            para_text_x,
                            para_text_width,
                            baseline_y,
                            line_h,
                            lines.len(),
                            0,
                        );

                        all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                        slot_top = doc.page_height - doc.margin_top;

                        let rest = &lines[lines_that_fit..];
                        let rest_content_h = rest.len() as f32 * line_h;
                        let baseline_y2 = slot_top - font_size * ascender_ratio;

                        render_paragraph_lines(
                            &mut current_content,
                            rest,
                            &para.alignment,
                            para_text_x,
                            para_text_width,
                            baseline_y2,
                            line_h,
                            lines.len(),
                            lines_that_fit,
                        );

                        slot_top -= rest_content_h;
                        prev_space_after = effective_space_after;
                        continue;
                    }

                    all_contents.push(std::mem::replace(&mut current_content, Content::new()));
                    slot_top = doc.page_height - doc.margin_top;
                    inter_gap = effective_space_before;
                }

                slot_top -= inter_gap;

                if (para.image.is_some() || para.runs.is_empty()) && para.content_height > 0.0 {
                    if let Some(pdf_name) = image_pdf_names.get(&block_idx) {
                        let img = para.image.as_ref().unwrap();
                        let y_bottom = slot_top - img.display_height;
                        let x = doc.margin_left + (text_width - img.display_width).max(0.0) / 2.0;
                        current_content.save_state();
                        current_content.transform([
                            img.display_width,
                            0.0,
                            0.0,
                            img.display_height,
                            x,
                            y_bottom,
                        ]);
                        current_content.x_object(Name(pdf_name.as_bytes()));
                        current_content.restore_state();
                    } else {
                        current_content
                            .set_fill_gray(0.5)
                            .rect(doc.margin_left, slot_top - content_h, text_width, content_h)
                            .fill_nonzero()
                            .set_fill_gray(0.0);
                    }
                } else if !lines.is_empty() {
                    let ascender_ratio = tallest_ar.unwrap_or(0.75);
                    let baseline_y = slot_top - font_size * ascender_ratio;

                    if !para.list_label.is_empty() {
                        let (label_font_name, label_bytes) =
                            label_for_run(&para.runs[0], &seen_fonts, &para.list_label);
                        current_content
                            .begin_text()
                            .set_font(Name(label_font_name.as_bytes()), font_size)
                            .next_line(label_x, baseline_y)
                            .show(Str(&label_bytes))
                            .end_text();
                    }

                    render_paragraph_lines(
                        &mut current_content,
                        &lines,
                        &para.alignment,
                        para_text_x,
                        para_text_width,
                        baseline_y,
                        line_h,
                        lines.len(),
                        0,
                    );
                }

                // Draw bottom border if present
                if let Some(bdr) = &para.border_bottom {
                    let line_y = slot_top - content_h - bdr.space_pt;
                    let [r, g, b] = bdr.color;
                    current_content
                        .set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0)
                        .rect(
                            doc.margin_left,
                            line_y - bdr.width_pt,
                            text_width,
                            bdr.width_pt,
                        )
                        .fill_nonzero()
                        .set_fill_rgb(0.0, 0.0, 0.0);
                }

                slot_top -= content_h;
                prev_space_after = effective_space_after;
            }

            Block::Table(table) => {
                render_table(
                    table,
                    doc,
                    &seen_fonts,
                    &mut current_content,
                    &mut all_contents,
                    &mut slot_top,
                    prev_space_after,
                );
                prev_space_after = 0.0;
            }
        }
    }
    all_contents.push(current_content);

    // Phase 3: allocate page and content IDs now that page count is known
    let n = all_contents.len();
    let page_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();
    let content_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();

    for (i, c) in all_contents.into_iter().enumerate() {
        pdf.stream(content_ids[i], &c.finish());
    }

    pdf.catalog(catalog_id).pages(pages_id);
    pdf.pages(pages_id)
        .kids(page_ids.iter().copied())
        .count(n as i32);

    let font_pairs: Vec<(String, Ref)> = font_order
        .iter()
        .map(|name| (seen_fonts[name].pdf_name.clone(), seen_fonts[name].font_ref))
        .collect();

    for i in 0..n {
        let mut page = pdf.page(page_ids[i]);
        page.media_box(Rect::new(0.0, 0.0, doc.page_width, doc.page_height))
            .parent(pages_id)
            .contents(content_ids[i]);
        {
            let mut resources = page.resources();
            {
                let mut fonts = resources.fonts();
                for (name, font_ref) in &font_pairs {
                    fonts.pair(Name(name.as_bytes()), *font_ref);
                }
            }
            if !image_xobjects.is_empty() {
                let mut xobjects = resources.x_objects();
                for (name, xobj_ref) in &image_xobjects {
                    xobjects.pair(Name(name.as_bytes()), *xobj_ref);
                }
            }
        }
    }

    Ok(pdf.finish())
}

fn label_for_run<'a>(
    run: &Run,
    seen_fonts: &'a HashMap<String, FontEntry>,
    label: &str,
) -> (&'a str, Vec<u8>) {
    let key = font_key(run);
    let entry = seen_fonts.get(&key).expect("font registered");
    (entry.pdf_name.as_str(), to_winansi_bytes(label))
}
