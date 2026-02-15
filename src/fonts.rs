use std::collections::HashMap;
use std::path::PathBuf;
use std::sync::OnceLock;

use pdf_writer::{Name, Pdf, Rect, Ref};
use ttf_parser::Face;

use crate::model::Run;

pub(crate) struct FontEntry {
    pub(crate) pdf_name: String,
    pub(crate) font_ref: Ref,
    pub(crate) widths_1000: Vec<f32>,
    pub(crate) line_h_ratio: Option<f32>,
    pub(crate) ascender_ratio: Option<f32>,
}

/// (lowercase family name, bold, italic) -> (file path, face index within TTC)
type FontLookup = HashMap<(String, bool, bool), (PathBuf, u32)>;

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

fn read_font_style(data: &[u8], face_index: u32) -> Option<(String, bool, bool)> {
    let face = Face::parse(data, face_index).ok()?;
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
            let is_collection = match path.extension().and_then(|e| e.to_str()) {
                Some("ttf" | "otf" | "TTF" | "OTF") => false,
                Some("ttc" | "TTC") => true,
                _ => continue,
            };
            let Ok(data) = std::fs::read(&path) else {
                continue;
            };
            let face_count = if is_collection {
                ttf_parser::fonts_in_collection(&data).unwrap_or(1)
            } else {
                1
            };
            for face_idx in 0..face_count {
                if let Some((family, bold, italic)) = read_font_style(&data, face_idx) {
                    index
                        .entry((family.to_lowercase(), bold, italic))
                        .or_insert((path.clone(), face_idx));
                }
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
fn find_font_file(font_name: &str, bold: bool, italic: bool) -> Option<(PathBuf, u32)> {
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
pub(crate) fn to_winansi_bytes(s: &str) -> Vec<u8> {
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
    face_index: u32,
) -> Option<(Vec<f32>, f32, f32)> {
    let face = Face::parse(font_data, face_index).ok()?;

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

pub(crate) fn primary_font_name(name: &str) -> &str {
    name.split(';').next().unwrap_or(name).trim()
}

pub(crate) fn font_key(run: &Run) -> String {
    let base = primary_font_name(&run.font_name);
    match (run.bold, run.italic) {
        (true, true) => format!("{}/BI", base),
        (true, false) => format!("{}/B", base),
        (false, true) => format!("{}/I", base),
        (false, false) => base.to_string(),
    }
}

pub(crate) type EmbeddedFonts = HashMap<(String, bool, bool), Vec<u8>>;

pub(crate) fn register_font(
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
            embed_truetype(pdf, font_ref, descriptor_ref, data_ref, font_name, data, 0)
        })
        .or_else(|| {
            find_font_file(font_name, bold, italic).and_then(|(path, face_index)| {
                let data = std::fs::read(&path).ok()?;
                embed_truetype(pdf, font_ref, descriptor_ref, data_ref, font_name, &data, face_index)
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
