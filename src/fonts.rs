use std::collections::{HashMap, HashSet};
use std::path::PathBuf;
use std::sync::OnceLock;

use memmap2::Mmap;
use pdf_writer::{Name, Pdf, Rect, Ref};
use ttf_parser::Face;

use crate::model::Run;

pub(crate) struct FontEntry {
    pub(crate) pdf_name: String,
    pub(crate) font_ref: Ref,
    pub(crate) widths_1000: Vec<f32>,
    pub(crate) line_h_ratio: Option<f32>,
    pub(crate) ascender_ratio: Option<f32>,
    pub(crate) char_to_gid: Option<HashMap<char, u16>>,
    pub(crate) char_widths_1000: Option<HashMap<char, f32>>,
}

impl FontEntry {
    /// Width of a single character in 1000-units. Uses the per-char cache (covers
    /// all Unicode chars seen in the document), falls back to the WinAnsi table.
    pub(crate) fn char_width_1000(&self, ch: char) -> f32 {
        if let Some(ref map) = self.char_widths_1000 {
            if let Some(&w) = map.get(&ch) {
                return w;
            }
        }
        // Fallback: WinAnsi lookup (Helvetica or chars not in used_chars)
        let byte = char_to_winansi(ch);
        if byte >= 32 {
            self.widths_1000[(byte - 32) as usize]
        } else {
            0.0
        }
    }

    pub(crate) fn word_width(&self, word: &str, font_size: f32) -> f32 {
        word.chars()
            .map(|ch| self.char_width_1000(ch) * font_size / 1000.0)
            .sum()
    }

    pub(crate) fn space_width(&self, font_size: f32) -> f32 {
        self.char_width_1000(' ') * font_size / 1000.0
    }
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
            dirs.push(PathBuf::from(&home).join("Library/Fonts"));
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

struct CachedFace {
    family: String,
    bold: bool,
    italic: bool,
    face_index: u32,
}

struct CachedFile {
    faces: Vec<CachedFace>,
}

struct FontCache {
    dir_mtimes: HashMap<PathBuf, i64>,
    files: HashMap<PathBuf, CachedFile>,
}

fn cache_path() -> Option<PathBuf> {
    let dir = if cfg!(target_os = "macos") {
        std::env::var("HOME")
            .ok()
            .map(|h| PathBuf::from(h).join("Library/Caches/docxide-pdf"))
    } else if cfg!(target_os = "windows") {
        std::env::var("LOCALAPPDATA")
            .ok()
            .map(|d| PathBuf::from(d).join("docxide-pdf/cache"))
    } else {
        std::env::var("XDG_CACHE_HOME")
            .ok()
            .map(PathBuf::from)
            .or_else(|| {
                std::env::var("HOME")
                    .ok()
                    .map(|h| PathBuf::from(h).join(".cache"))
            })
            .map(|d| d.join("docxide-pdf"))
    };
    dir.map(|d| d.join("font-index.tsv"))
}

const CACHE_VERSION: &str = "v1";

fn load_cache() -> FontCache {
    let mut fc = FontCache {
        dir_mtimes: HashMap::new(),
        files: HashMap::new(),
    };
    let Some(path) = cache_path() else {
        return fc;
    };
    let Ok(content) = std::fs::read_to_string(&path) else {
        return fc;
    };
    let mut lines = content.lines();
    if lines.next() != Some(CACHE_VERSION) {
        return fc;
    }
    for line in lines {
        let parts: Vec<&str> = line.split('\t').collect();
        match parts.first().copied() {
            Some("D") if parts.len() == 3 => {
                let Ok(mtime) = parts[2].parse::<i64>() else {
                    continue;
                };
                fc.dir_mtimes.insert(PathBuf::from(parts[1]), mtime);
            }
            Some("F") if parts.len() == 6 => {
                let file_path = PathBuf::from(parts[1]);
                let family = parts[2].to_string();
                let bold = parts[3] == "1";
                let italic = parts[4] == "1";
                let Ok(face_index) = parts[5].parse::<u32>() else {
                    continue;
                };
                let entry = fc
                    .files
                    .entry(file_path)
                    .or_insert(CachedFile { faces: Vec::new() });
                entry.faces.push(CachedFace {
                    family,
                    bold,
                    italic,
                    face_index,
                });
            }
            Some("F") if parts.len() == 3 && parts[2] == "-" => {
                fc.files
                    .entry(PathBuf::from(parts[1]))
                    .or_insert(CachedFile { faces: Vec::new() });
            }
            _ => {}
        }
    }
    fc
}

fn save_cache(cache: &FontCache) {
    let Some(path) = cache_path() else {
        return;
    };
    if let Some(dir) = path.parent() {
        let _ = std::fs::create_dir_all(dir);
    }
    let mut out = String::from(CACHE_VERSION);
    out.push('\n');
    for (dir_path, mtime) in &cache.dir_mtimes {
        out.push_str(&format!("D\t{}\t{}\n", dir_path.to_string_lossy(), mtime));
    }
    for (file_path, cached) in &cache.files {
        let path_str = file_path.to_string_lossy();
        if cached.faces.is_empty() {
            out.push_str(&format!("F\t{}\t-\n", path_str));
        } else {
            for face in &cached.faces {
                out.push_str(&format!(
                    "F\t{}\t{}\t{}\t{}\t{}\n",
                    path_str,
                    face.family,
                    if face.bold { "1" } else { "0" },
                    if face.italic { "1" } else { "0" },
                    face.face_index,
                ));
            }
        }
    }
    let _ = std::fs::write(&path, out);
}

fn dir_mtime(path: &std::path::Path) -> i64 {
    std::fs::metadata(path)
        .ok()
        .and_then(|m| m.modified().ok())
        .and_then(|t| t.duration_since(std::time::UNIX_EPOCH).ok())
        .map(|d| d.as_secs() as i64)
        .unwrap_or(0)
}

fn is_font_file(path: &std::path::Path) -> bool {
    matches!(
        path.extension()
            .and_then(|e| e.to_str())
            .map(|e| e.to_ascii_lowercase())
            .as_deref(),
        Some("ttf" | "otf" | "ttc")
    )
}

fn is_font_collection(path: &std::path::Path) -> bool {
    path.extension()
        .and_then(|e| e.to_str())
        .is_some_and(|e| e.eq_ignore_ascii_case("ttc"))
}

fn scan_font_dirs() -> FontLookup {
    let t0 = std::time::Instant::now();
    let mut index = FontLookup::new();
    let dirs = font_directories();

    let no_cache = std::env::var("DOCXSIDE_NO_FONT_CACHE").is_ok();

    let cache = if no_cache {
        FontCache {
            dir_mtimes: HashMap::new(),
            files: HashMap::new(),
        }
    } else {
        load_cache()
    };
    let mut new_cache = FontCache {
        dir_mtimes: HashMap::new(),
        files: HashMap::new(),
    };
    let mut files_scanned = 0u32;
    let mut dirs_cached = 0u32;
    let mut dirs_scanned = 0u32;
    let mut visited_dirs: std::collections::HashSet<PathBuf> = std::collections::HashSet::new();

    let mut stack: Vec<PathBuf> = dirs;
    while let Some(dir) = stack.pop() {
        if !visited_dirs.insert(dir.clone()) {
            continue;
        }

        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };

        // Collect directory listing: subdirs to recurse, font files to process
        let mut subdirs = Vec::new();
        let mut font_files = Vec::new();
        for entry in entries.flatten() {
            let path = entry.path();
            if path.is_dir() {
                subdirs.push(path);
            } else if is_font_file(&path) {
                font_files.push(path);
            }
        }
        stack.extend(subdirs);

        if font_files.is_empty() {
            continue;
        }

        let current_mtime = dir_mtime(&dir);

        if let Some(&cached_mtime) = cache.dir_mtimes.get(&dir)
            && cached_mtime == current_mtime
        {
            dirs_cached += 1;
            new_cache.dir_mtimes.insert(dir.clone(), current_mtime);
            for file_path in &font_files {
                if let Some(cached_file) = cache.files.get(file_path) {
                    for face in &cached_file.faces {
                        index
                            .entry((face.family.to_lowercase(), face.bold, face.italic))
                            .or_insert((file_path.clone(), face.face_index));
                    }
                    new_cache.files.insert(
                        file_path.clone(),
                        CachedFile {
                            faces: cached_file
                                .faces
                                .iter()
                                .map(|f| CachedFace {
                                    family: f.family.clone(),
                                    bold: f.bold,
                                    italic: f.italic,
                                    face_index: f.face_index,
                                })
                                .collect(),
                        },
                    );
                }
            }
            continue;
        }

        // Directory changed — scan all font files in it
        dirs_scanned += 1;
        new_cache.dir_mtimes.insert(dir, current_mtime);
        for file_path in font_files {
            files_scanned += 1;
            let Ok(file) = std::fs::File::open(&file_path) else {
                continue;
            };
            let Ok(data) = (unsafe { Mmap::map(&file) }) else {
                continue;
            };
            let face_count = if is_font_collection(&file_path) {
                ttf_parser::fonts_in_collection(&data).unwrap_or(1)
            } else {
                1
            };
            let mut faces = Vec::new();
            for face_idx in 0..face_count {
                if let Some((family, bold, italic)) = read_font_style(&data, face_idx) {
                    index
                        .entry((family.to_lowercase(), bold, italic))
                        .or_insert((file_path.clone(), face_idx));
                    faces.push(CachedFace {
                        family,
                        bold,
                        italic,
                        face_index: face_idx,
                    });
                }
            }
            new_cache.files.insert(file_path, CachedFile { faces });
        }
    }

    if !no_cache {
        save_cache(&new_cache);
    }

    log::info!(
        "Font scan: {:.1}ms, {} dirs cached / {} scanned, {} files parsed → {} entries",
        t0.elapsed().as_secs_f64() * 1000.0,
        dirs_cached,
        dirs_scanned,
        files_scanned,
        index.len(),
    );

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

/// Map a single Unicode char to its WinAnsi byte, or 0 if unmappable.
fn char_to_winansi(c: char) -> u8 {
    match c as u32 {
        0x0020..=0x007F => c as u8,
        0x00A0..=0x00FF => c as u8,
        0x20AC => 0x80,
        0x201A => 0x82,
        0x0192 => 0x83,
        0x201E => 0x84,
        0x2026 => 0x85,
        0x2020 => 0x86,
        0x2021 => 0x87,
        0x02C6 => 0x88,
        0x2030 => 0x89,
        0x0160 => 0x8A,
        0x2039 => 0x8B,
        0x0152 => 0x8C,
        0x017D => 0x8E,
        0x2018 => 0x91,
        0x2019 => 0x92,
        0x201C => 0x93,
        0x201D => 0x94,
        0x2022 => 0x95,
        0x2013 => 0x96,
        0x2014 => 0x97,
        0x02DC => 0x98,
        0x2122 => 0x99,
        0x0161 => 0x9A,
        0x203A => 0x9B,
        0x0153 => 0x9C,
        0x017E => 0x9E,
        0x0178 => 0x9F,
        _ => 0,
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

/// Encode UTF-8 text as big-endian 2-byte glyph IDs for CIDFont content streams.
pub(crate) fn encode_as_gids(text: &str, char_to_gid: &HashMap<char, u16>) -> Vec<u8> {
    let mut out = Vec::with_capacity(text.len() * 2);
    for ch in text.chars() {
        let gid = char_to_gid.get(&ch).copied().unwrap_or(0);
        out.push((gid >> 8) as u8);
        out.push((gid & 0xFF) as u8);
    }
    out
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

/// Embed a TrueType/OpenType font as a CIDFont (Type0 composite) with Identity-H encoding.
/// The font data is subsetted to only include glyphs used in the document.
fn embed_truetype(
    pdf: &mut Pdf,
    font_ref: Ref,
    descriptor_ref: Ref,
    data_ref: Ref,
    font_name: &str,
    font_data: &[u8],
    face_index: u32,
    used_chars: &HashSet<char>,
    alloc: &mut impl FnMut() -> Ref,
) -> Option<(Vec<f32>, f32, f32, HashMap<char, u16>, HashMap<char, f32>)> {
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

    // WinAnsi widths for layout (unchanged — layout uses these)
    let widths_1000: Vec<f32> = (32u8..=255u8)
        .map(|byte| {
            face.glyph_index(winansi_to_char(byte))
                .and_then(|gid| face.glyph_hor_advance(gid))
                .map(|adv| adv as f32 / units * 1000.0)
                .unwrap_or(0.0)
        })
        .collect();

    // Build GlyphRemapper, char_to_gid, and char_widths_1000 maps from used_chars
    let mut remapper = subsetter::GlyphRemapper::new();
    let mut char_to_gid = HashMap::new();
    let mut char_widths_1000 = HashMap::new();
    for &ch in used_chars {
        if let Some(gid) = face.glyph_index(ch) {
            let new_gid = remapper.remap(gid.0);
            char_to_gid.insert(ch, new_gid);
            let w = face
                .glyph_hor_advance(gid)
                .map(|adv| adv as f32 / units * 1000.0)
                .unwrap_or(0.0);
            char_widths_1000.insert(ch, w);
        }
    }

    // Subset the font
    let subset_data = subsetter::subset(font_data, face_index, &remapper)
        .unwrap_or_else(|e| {
            log::warn!("Font subsetting failed for {font_name}: {e} — embedding full font");
            font_data.to_vec()
        });

    let data_len = i32::try_from(subset_data.len()).ok()?;
    pdf.stream(data_ref, &subset_data)
        .pair(Name(b"Length1"), data_len);

    let ps_name = font_name.replace(' ', "");

    // FontDescriptor
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

    // CIDFont dict
    let cid_font_ref = alloc();
    let system_info = pdf_writer::types::SystemInfo {
        registry: pdf_writer::Str(b"Adobe"),
        ordering: pdf_writer::Str(b"Identity"),
        supplement: 0,
    };
    {
        let mut cid = pdf.cid_font(cid_font_ref);
        cid.subtype(pdf_writer::types::CidFontType::Type2);
        cid.base_font(Name(ps_name.as_bytes()));
        cid.system_info(system_info);
        cid.font_descriptor(descriptor_ref);
        cid.default_width(0.0);
        cid.cid_to_gid_map_predefined(Name(b"Identity"));
        // Write per-glyph widths
        let mut gid_widths: Vec<(u16, f32)> = char_to_gid
            .iter()
            .filter_map(|(&ch, &new_gid)| {
                face.glyph_index(ch)
                    .and_then(|gid| face.glyph_hor_advance(gid))
                    .map(|adv| (new_gid, adv as f32 / units * 1000.0))
            })
            .collect();
        gid_widths.sort_by_key(|&(gid, _)| gid);
        if !gid_widths.is_empty() {
            let mut w = cid.widths();
            for &(gid, width) in &gid_widths {
                w.consecutive(gid, [width]);
            }
        }
    }

    // ToUnicode CMap
    let tounicode_ref = alloc();
    let cmap_name = format!("{}-UTF16", ps_name);
    let mut cmap = pdf_writer::types::UnicodeCmap::new(
        Name(cmap_name.as_bytes()),
        pdf_writer::types::SystemInfo {
            registry: pdf_writer::Str(b"Adobe"),
            ordering: pdf_writer::Str(b"Identity"),
            supplement: 0,
        },
    );
    for (&ch, &new_gid) in &char_to_gid {
        cmap.pair(new_gid, ch);
    }
    let cmap_data = cmap.finish();
    pdf.stream(tounicode_ref, cmap_data.as_slice());

    // Type0 (composite) font dict
    pdf.type0_font(font_ref)
        .base_font(Name(ps_name.as_bytes()))
        .encoding_predefined(Name(b"Identity-H"))
        .descendant_font(cid_font_ref)
        .to_unicode(tounicode_ref);

    let line_gap = face.line_gap() as f32;
    let line_h_ratio = (face.ascender() as f32 - face.descender() as f32 + line_gap) / units;
    let ascender_ratio = face.ascender() as f32 / units;

    Some((widths_1000, line_h_ratio, ascender_ratio, char_to_gid, char_widths_1000))
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
    used_chars: &HashSet<char>,
) -> FontEntry {
    let t0 = std::time::Instant::now();
    let font_ref = alloc();
    let descriptor_ref = alloc();
    let data_ref = alloc();

    let font_candidates: Vec<&str> = font_name.split(';').map(|s| s.trim()).collect();

    let mut result = None;
    for candidate in &font_candidates {
        let embedded_key = (candidate.to_lowercase(), bold, italic);
        let embedded_data = embedded_fonts.get(&embedded_key);

        let found = embedded_data
            .and_then(|data| {
                embed_truetype(
                    pdf, font_ref, descriptor_ref, data_ref, candidate, data, 0,
                    used_chars, alloc,
                )
            })
            .or_else(|| {
                find_font_file(candidate, bold, italic).and_then(|(path, face_index)| {
                    let data = std::fs::read(&path).ok()?;
                    embed_truetype(
                        pdf,
                        font_ref,
                        descriptor_ref,
                        data_ref,
                        candidate,
                        &data,
                        face_index,
                        used_chars,
                        alloc,
                    )
                })
            });
        if let Some(metrics) = found {
            result = Some(metrics);
            break;
        }
    }

    let (widths, line_h_ratio, ascender_ratio, char_to_gid, char_widths_1000) = result
        .map(|(w, r, ar, m, cw)| (w, Some(r), Some(ar), Some(m), Some(cw)))
        .unwrap_or_else(|| {
            log::warn!("Font not found: {font_name} bold={bold} italic={italic} — using Helvetica");
            pdf.type1_font(font_ref)
                .base_font(Name(b"Helvetica"))
                .encoding_predefined(Name(b"WinAnsiEncoding"));
            (helvetica_widths(), None, None, None, None)
        });

    log::debug!(
        "register_font: {font_name} bold={bold} italic={italic} → {:.1}ms",
        t0.elapsed().as_secs_f64() * 1000.0,
    );

    FontEntry {
        pdf_name,
        font_ref,
        widths_1000: widths,
        line_h_ratio,
        ascender_ratio,
        char_to_gid,
        char_widths_1000,
    }
}
