use std::collections::{HashMap, HashSet};

use pdf_writer::types::{CidFontType, FontFlags, SystemInfo, UnicodeCmap};
use pdf_writer::{Name, Pdf, Rect, Ref, Str};
use ttf_parser::Face;
use ttf_parser::gpos::{PairAdjustment, PositioningSubtable};

use super::FontMetrics;
use super::encoding::winansi_to_char;

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
    let to_1000 = |v: f32| v / units * 1000.0;

    let ascent = to_1000(face.ascender() as f32);
    let descent = to_1000(face.descender() as f32);
    let cap_height = face
        .capital_height()
        .map(|h| to_1000(h as f32))
        .unwrap_or(700.0);

    let bb = face.global_bounding_box();
    let bbox = Rect::new(
        to_1000(bb.x_min as f32),
        to_1000(bb.y_min as f32),
        to_1000(bb.x_max as f32),
        to_1000(bb.y_max as f32),
    );

    let advance_1000 = |gid: ttf_parser::GlyphId| -> f32 {
        face.glyph_hor_advance(gid)
            .map(|adv| to_1000(adv as f32))
            .unwrap_or(0.0)
    };

    let widths_1000: Vec<f32> = (32u8..=255u8)
        .map(|byte| {
            face.glyph_index(winansi_to_char(byte))
                .map(&advance_1000)
                .unwrap_or(0.0)
        })
        .collect();

    let mut remapper = subsetter::GlyphRemapper::new();
    let mut char_to_gid = HashMap::new();
    let mut char_widths_1000 = HashMap::new();

    for &ch in used_chars {
        let gid = resolve_glyph(&face, ch);
        if let Some(gid) = gid {
            let new_gid = remapper.remap(gid.0);
            char_to_gid.insert(ch, new_gid);
            char_widths_1000.insert(ch, advance_1000(gid));
        }
    }

    let mut kern_pairs = HashMap::new();
    let char_gids: Vec<(ttf_parser::GlyphId, u16)> = char_to_gid
        .iter()
        .filter_map(|(&ch, &new_gid)| face.glyph_index(ch).map(|orig| (orig, new_gid)))
        .collect();

    extract_kern_pairs(&face, &char_gids, units, &mut kern_pairs);
    extract_gpos_pairs(&face, &char_gids, units, &mut kern_pairs);

    if !kern_pairs.is_empty() {
        log::info!(
            "Kerning for {font_name}: {} pairs from {} chars",
            kern_pairs.len(),
            char_gids.len(),
        );
    }

    let subset_data = subsetter::subset(font_data, face_index, &remapper).unwrap_or_else(|e| {
        log::warn!("Font subsetting failed for {font_name}: {e} — embedding full font");
        font_data.to_vec()
    });

    let data_len = i32::try_from(subset_data.len()).ok()?;
    pdf.stream(data_ref, &subset_data)
        .pair(Name(b"Length1"), data_len);

    let ps_name = font_name.replace(' ', "");
    let ps_name_ref = Name(ps_name.as_bytes());
    let system_info = SystemInfo {
        registry: Str(b"Adobe"),
        ordering: Str(b"Identity"),
        supplement: 0,
    };

    pdf.font_descriptor(descriptor_ref)
        .name(ps_name_ref)
        .flags(FontFlags::NON_SYMBOLIC)
        .bbox(bbox)
        .italic_angle(0.0)
        .ascent(ascent)
        .descent(descent)
        .cap_height(cap_height)
        .stem_v(80.0)
        .font_file2(data_ref);

    let cid_font_ref = alloc();
    {
        let mut cid = pdf.cid_font(cid_font_ref);
        cid.subtype(CidFontType::Type2);
        cid.base_font(ps_name_ref);
        cid.system_info(system_info);
        cid.font_descriptor(descriptor_ref);
        cid.default_width(0.0);
        cid.cid_to_gid_map_predefined(Name(b"Identity"));

        let mut gid_widths: Vec<(u16, f32)> = char_to_gid
            .iter()
            .filter_map(|(&ch, &new_gid)| {
                face.glyph_index(ch).map(|gid| (new_gid, advance_1000(gid)))
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

    let tounicode_ref = alloc();
    let cmap_name = format!("{}-UTF16", ps_name);
    let mut cmap = UnicodeCmap::new(Name(cmap_name.as_bytes()), system_info);
    for (&ch, &new_gid) in &char_to_gid {
        cmap.pair(new_gid, ch);
    }
    pdf.stream(tounicode_ref, cmap.finish().as_slice());

    pdf.type0_font(font_ref)
        .base_font(ps_name_ref)
        .encoding_predefined(Name(b"Identity-H"))
        .descendant_font(cid_font_ref)
        .to_unicode(tounicode_ref);

    let (line_h_ratio, ascender_ratio) = compute_line_metrics(&face, units);

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

fn resolve_glyph(face: &Face, ch: char) -> Option<ttf_parser::GlyphId> {
    face.glyph_index(ch)
        .or_else(|| {
            // Symbol fonts use Private Use Area (0xF000-0xF0FF); try the low byte
            let cp = ch as u32;
            if (0xF000..=0xF0FF).contains(&cp) {
                face.glyph_index(char::from_u32(cp - 0xF000)?)
            } else {
                None
            }
        })
        .or_else(|| {
            // Fallback: direct cmap subtable lookup for symbol fonts
            face.tables()
                .cmap?
                .subtables
                .into_iter()
                .find_map(|st| st.glyph_index(ch as u32))
        })
}

fn extract_kern_pairs(
    face: &Face,
    char_gids: &[(ttf_parser::GlyphId, u16)],
    units: f32,
    kern_pairs: &mut HashMap<(u16, u16), f32>,
) {
    let Some(kern) = face.tables().kern else {
        return;
    };
    for &(l_orig, l_new) in char_gids {
        for &(r_orig, r_new) in char_gids {
            let total: i16 = kern
                .subtables
                .into_iter()
                .filter(|st| st.horizontal && !st.variable)
                .filter_map(|st| st.glyphs_kerning(l_orig, r_orig))
                .sum();
            if total != 0 {
                kern_pairs.insert((l_new, r_new), total as f32 / units * 1000.0);
            }
        }
    }
}

fn extract_gpos_pairs(
    face: &Face,
    char_gids: &[(ttf_parser::GlyphId, u16)],
    units: f32,
    kern_pairs: &mut HashMap<(u16, u16), f32>,
) {
    let Some(gpos) = face.tables().gpos else {
        return;
    };
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
                    for &(l_orig, l_new) in char_gids {
                        let Some(cov_idx) = coverage.get(l_orig) else {
                            continue;
                        };
                        let Some(pair_set) = sets.get(cov_idx) else {
                            continue;
                        };
                        for &(r_orig, r_new) in char_gids {
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
                    for &(l_orig, l_new) in char_gids {
                        if coverage.get(l_orig).is_none() {
                            continue;
                        }
                        let c1 = classes.0.get(l_orig);
                        for &(r_orig, r_new) in char_gids {
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

fn compute_line_metrics(face: &Face, units: f32) -> (f32, f32) {
    if let Some(os2) = face.tables().os2 {
        if os2.use_typographic_metrics() {
            let asc = os2.typographic_ascender() as f32;
            let desc = os2.typographic_descender() as f32;
            let gap = os2.typographic_line_gap() as f32;
            return ((asc - desc + gap) / units, asc / units);
        }
        let win_asc = os2.windows_ascender() as f32;
        let win_desc = os2.windows_descender() as f32;
        // usWinAscent/Descent define glyph clipping bounds; hhea lineGap
        // provides external leading that Word includes in line spacing
        let gap = face.line_gap() as f32;
        return ((win_asc - win_desc + gap) / units, win_asc / units);
    }

    let line_gap = face.line_gap() as f32;
    (
        (face.ascender() as f32 - face.descender() as f32 + line_gap) / units,
        face.ascender() as f32 / units,
    )
}
