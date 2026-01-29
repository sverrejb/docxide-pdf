use std::collections::HashMap;
use std::path::{Path, PathBuf};

use pdf_writer::{Content, Name, Pdf, Rect, Ref, Str};
use ttf_parser::Face;

use crate::error::Error;
use crate::model::Document;

struct FontEntry {
    pdf_name: String,
    font_ref: Ref,
}

/// Search the host system for a TTF/OTF file for the given font name.
/// Checks Microsoft Office bundled fonts and system font directories.
fn find_font_file(font_name: &str) -> Option<PathBuf> {
    // "Aptos Display" -> "AptosDisplay" for file-name lookups
    let normalized = font_name.replace(' ', "");

    // 1. Office CloudFonts (e.g. Aptos Display is downloaded here on Mac)
    if let Ok(home) = std::env::var("HOME") {
        let cloud_dir = PathBuf::from(&home)
            .join("Library/Group Containers/UBF8T346G9.Office/FontCache/4/CloudFonts")
            .join(font_name);
        if cloud_dir.is_dir() {
            if let Ok(entries) = std::fs::read_dir(&cloud_dir) {
                for entry in entries.flatten() {
                    let p = entry.path();
                    if matches!(
                        p.extension().and_then(|e| e.to_str()),
                        Some("ttf") | Some("otf")
                    ) {
                        return Some(p);
                    }
                }
            }
        }
    }

    // 2. Word app DFonts (Mac)
    let word_fonts =
        Path::new("/Applications/Microsoft Word.app/Contents/Resources/DFonts");
    for ext in &["ttf", "otf"] {
        let p = word_fonts.join(format!("{}.{}", normalized, ext));
        if p.exists() {
            return Some(p);
        }
    }

    // 3. System font directories
    let system_dirs = [
        "/Library/Fonts",
        "/System/Library/Fonts",
        "/System/Library/Fonts/Supplemental",
    ];
    for dir in &system_dirs {
        for ext in &["ttf", "otf"] {
            let p = Path::new(dir).join(format!("{}.{}", normalized, ext));
            if p.exists() {
                return Some(p);
            }
        }
    }

    None
}

/// Attempt to embed a TrueType/OpenType font into the PDF.
/// Returns Some(()) on success, None if the font couldn't be read or parsed.
fn embed_truetype(
    pdf: &mut Pdf,
    font_ref: Ref,
    descriptor_ref: Ref,
    data_ref: Ref,
    font_name: &str,
    font_path: &Path,
) -> Option<()> {
    let font_data = std::fs::read(font_path).ok()?;
    let face = Face::parse(&font_data, 0).ok()?;

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
            face.glyph_index(byte as char)
                .and_then(|gid| face.glyph_hor_advance(gid))
                .map(|adv| adv as f32 / units * 1000.0)
                .unwrap_or(0.0)
        })
        .collect();

    // Font data stream — Length1 is required for TrueType streams
    let data_len = i32::try_from(font_data.len()).ok()?;
    pdf.stream(data_ref, &font_data).pair(Name(b"Length1"), data_len);

    // PostScript names must not contain spaces
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

    // Write the font dict manually to use /Subtype /TrueType
    {
        let mut d = pdf.indirect(font_ref).dict();
        d.pair(Name(b"Type"), Name(b"Font"));
        d.pair(Name(b"Subtype"), Name(b"TrueType"));
        d.pair(Name(b"BaseFont"), Name(ps_name.as_bytes()));
        d.pair(Name(b"Encoding"), Name(b"WinAnsiEncoding"));
        d.pair(Name(b"FirstChar"), 32i32);
        d.pair(Name(b"LastChar"), 255i32);
        d.pair(Name(b"FontDescriptor"), descriptor_ref);
        d.insert(Name(b"Widths")).array().items(widths.iter().copied());
    }

    Some(())
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
    let page_id = alloc();
    let content_id = alloc();

    // Collect unique font names in document order
    let mut seen_fonts: HashMap<String, FontEntry> = HashMap::new();
    let mut font_order: Vec<String> = Vec::new();
    for para in &doc.paragraphs {
        for run in &para.runs {
            if !seen_fonts.contains_key(&run.font_name) {
                let idx = font_order.len() + 1;
                let pdf_name = format!("F{}", idx);
                let font_ref = alloc();
                // Reserve two extra IDs (descriptor + data stream) for embedded fonts
                let descriptor_ref = alloc();
                let data_ref = alloc();

                let embedded = find_font_file(&run.font_name)
                    .and_then(|path| {
                        embed_truetype(
                            &mut pdf,
                            font_ref,
                            descriptor_ref,
                            data_ref,
                            &run.font_name,
                            &path,
                        )
                    })
                    .is_some();

                if !embedded {
                    // Fallback: standard Helvetica (descriptor_ref and data_ref are unused)
                    pdf.type1_font(font_ref)
                        .base_font(Name(b"Helvetica"))
                        .encoding_predefined(Name(b"WinAnsiEncoding"));
                }

                seen_fonts.insert(
                    run.font_name.clone(),
                    FontEntry { pdf_name: pdf_name.clone(), font_ref },
                );
                font_order.push(run.font_name.clone());
            }
        }
    }

    // If document has no text, register a default font
    if seen_fonts.is_empty() {
        let font_ref = alloc();
        alloc(); // descriptor placeholder
        alloc(); // data placeholder
        pdf.type1_font(font_ref)
            .base_font(Name(b"Helvetica"))
            .encoding_predefined(Name(b"WinAnsiEncoding"));
        seen_fonts.insert(
            "Helvetica".into(),
            FontEntry { pdf_name: "F1".into(), font_ref },
        );
        font_order.push("Helvetica".into());
    }

    // Build content stream
    let mut content = Content::new();

    let mut slot_top = doc.page_height - doc.margin_top;
    let cursor_x = doc.margin_left;

    for para in &doc.paragraphs {
        slot_top -= para.space_before;

        let font_size = para.runs.first().map_or(12.0, |r| r.font_size);
        let baseline_y = slot_top - font_size;

        for run in &para.runs {
            let entry = seen_fonts.get(&run.font_name).expect("font registered");
            content
                .begin_text()
                .set_font(Name(entry.pdf_name.as_bytes()), run.font_size)
                .next_line(cursor_x, baseline_y)
                .show(Str(run.text.as_bytes()))
                .end_text();
        }

        let line_pitch = (font_size * 1.2_f32).max(doc.line_pitch);
        slot_top -= line_pitch;
        slot_top -= para.space_after;
    }

    pdf.stream(content_id, &content.finish());

    pdf.catalog(catalog_id).pages(pages_id);
    pdf.pages(pages_id).kids([page_id]).count(1);

    // Collect font pairs before entering the page builder scope
    let font_pairs: Vec<(String, Ref)> = font_order
        .iter()
        .map(|name| (seen_fonts[name].pdf_name.clone(), seen_fonts[name].font_ref))
        .collect();

    {
        let mut page = pdf.page(page_id);
        page.media_box(Rect::new(0.0, 0.0, doc.page_width, doc.page_height))
            .parent(pages_id)
            .contents(content_id);
        {
            let mut resources = page.resources();
            let mut fonts = resources.fonts();
            for (name, font_ref) in &font_pairs {
                fonts.pair(Name(name.as_bytes()), *font_ref);
            }
        }
    }

    Ok(pdf.finish())
}
