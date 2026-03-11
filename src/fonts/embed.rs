use std::collections::{HashMap, HashSet};

use pdf_writer::{Name, Pdf, Rect, Ref};
use ttf_parser::Face;

use super::FontMetrics;
use super::encoding::winansi_to_char;

/// Embed a TrueType/OpenType font as a CIDFont (Type0 composite) with Identity-H encoding.
/// The font data is subsetted to only include glyphs used in the document.
pub(super) fn embed_truetype(
    pdf: &mut Pdf,
    font_ref: Ref,
    descriptor_ref: Ref,
    data_ref: Ref,
    font_name: &str,
    font_data: &[u8],
    face_index: u32,
    used_chars: &HashSet<char>,
    alloc: &mut impl FnMut() -> Ref,
) -> Option<FontMetrics> {
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
    // For symbol fonts, try direct cmap subtable lookup as a fallback
    let symbol_cmap_lookup = |ch: char| -> Option<ttf_parser::GlyphId> {
        if let Some(cmap) = face.tables().cmap {
            for subtable in cmap.subtables {
                if let Some(gid) = subtable.glyph_index(ch as u32) {
                    return Some(gid);
                }
            }
        }
        None
    };
    for &ch in used_chars {
        let gid = face
            .glyph_index(ch)
            .or_else(|| {
                let cp = ch as u32;
                if (0xF000..=0xF0FF).contains(&cp) {
                    let low_ch = char::from_u32(cp - 0xF000)?;
                    face.glyph_index(low_ch)
                } else {
                    None
                }
            })
            .or_else(|| symbol_cmap_lookup(ch));
        if let Some(gid) = gid {
            let new_gid = remapper.remap(gid.0);
            char_to_gid.insert(ch, new_gid);
            let w = face
                .glyph_hor_advance(gid)
                .map(|adv| adv as f32 / units * 1000.0)
                .unwrap_or(0.0);
            char_widths_1000.insert(ch, w);
        }
    }

    // Extract kern pairs from kern table + GPOS PairAdjustment (keyed by remapped gids)
    let mut kern_pairs = HashMap::new();
    let char_gids: Vec<(char, ttf_parser::GlyphId, u16)> = char_to_gid
        .iter()
        .filter_map(|(&ch, &new_gid)| face.glyph_index(ch).map(|orig| (ch, orig, new_gid)))
        .collect();

    // Legacy kern table
    if let Some(kern) = face.tables().kern {
        for &(_, l_orig, l_new) in &char_gids {
            for &(_, r_orig, r_new) in &char_gids {
                let mut total: i16 = 0;
                for subtable in kern.subtables {
                    if subtable.horizontal && !subtable.variable {
                        if let Some(val) = subtable.glyphs_kerning(l_orig, r_orig) {
                            total += val;
                        }
                    }
                }
                if total != 0 {
                    kern_pairs.insert((l_new, r_new), total as f32 / units * 1000.0);
                }
            }
        }
    }

    // GPOS PairAdjustment (covers fonts with kerning only in GPOS, e.g. Cyrillic in Arial)
    if let Some(gpos) = face.tables().gpos {
        use ttf_parser::gpos::PairAdjustment;
        use ttf_parser::gpos::PositioningSubtable;
        for lookup_idx in 0..gpos.lookups.len() {
            let Some(lookup) = gpos.lookups.get(lookup_idx) else {
                continue;
            };
            for st_idx in 0..lookup.subtables.len() {
                let Some(PositioningSubtable::Pair(pair)) =
                    lookup.subtables.get::<PositioningSubtable>(st_idx)
                else {
                    continue;
                };
                match pair {
                    PairAdjustment::Format1 { coverage, sets } => {
                        for &(_, l_orig, l_new) in &char_gids {
                            let Some(cov_idx) = coverage.get(l_orig) else {
                                continue;
                            };
                            let Some(pair_set) = sets.get(cov_idx) else {
                                continue;
                            };
                            for &(_, r_orig, r_new) in &char_gids {
                                if let Some((val1, _)) = pair_set.get(r_orig) {
                                    if val1.x_advance != 0 {
                                        kern_pairs
                                            .entry((l_new, r_new))
                                            .or_insert(val1.x_advance as f32 / units * 1000.0);
                                    }
                                }
                            }
                        }
                    }
                    PairAdjustment::Format2 {
                        coverage,
                        classes,
                        matrix,
                    } => {
                        for &(_, l_orig, l_new) in &char_gids {
                            if coverage.get(l_orig).is_none() {
                                continue;
                            }
                            let c1 = classes.0.get(l_orig);
                            for &(_, r_orig, r_new) in &char_gids {
                                let c2 = classes.1.get(r_orig);
                                if let Some((val1, _)) = matrix.get((c1, c2)) {
                                    if val1.x_advance != 0 {
                                        kern_pairs
                                            .entry((l_new, r_new))
                                            .or_insert(val1.x_advance as f32 / units * 1000.0);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    if !kern_pairs.is_empty() {
        log::info!(
            "Kerning for {font_name}: {} pairs from {} chars",
            kern_pairs.len(),
            char_gids.len(),
        );
    }
    // Subset the font
    let subset_data = subsetter::subset(font_data, face_index, &remapper).unwrap_or_else(|e| {
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

    // Use OS/2 table metrics (what Word uses) instead of hhea metrics
    let (line_h_ratio, ascender_ratio) = if let Some(os2) = face.tables().os2 {
        if os2.use_typographic_metrics() {
            let asc = os2.typographic_ascender() as f32;
            let desc = os2.typographic_descender() as f32;
            let gap = os2.typographic_line_gap() as f32;
            ((asc - desc + gap) / units, asc / units)
        } else {
            let win_asc = os2.windows_ascender() as f32;
            let win_desc = os2.windows_descender() as f32;
            // usWinAscent/Descent define glyph clipping bounds; hhea lineGap
            // provides external leading that Word includes in line spacing
            let gap = face.line_gap() as f32;
            ((win_asc - win_desc + gap) / units, win_asc / units)
        }
    } else {
        let line_gap = face.line_gap() as f32;
        (
            (face.ascender() as f32 - face.descender() as f32 + line_gap) / units,
            face.ascender() as f32 / units,
        )
    };

    Some(FontMetrics {
        widths_1000,
        line_h_ratio,
        ascender_ratio,
        char_to_gid,
        char_widths_1000,
        kern_pairs,
        synthetic_bold: false,
    })
}
