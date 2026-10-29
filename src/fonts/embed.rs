use std::collections::{BTreeMap, HashMap};

use pdf_writer::types::{CidFontType, FontFlags, SystemInfo, UnicodeCmap};
use pdf_writer::{Filter, Name, Pdf, Rect, Ref, Str};
use ttf_parser::Face;
use ttf_parser::gpos::{PairAdjustment, PositioningSubtable};

use super::encoding::winansi_to_char;
use super::{FontContext, FontMetrics, FontRefs};

pub(super) fn embed_truetype(
    pdf: &mut Pdf,
    refs: FontRefs,
    font_name: &str,
    font_data: &[u8],
    face_index: u32,
    ctx: &FontContext,
    alloc: &mut impl FnMut() -> Ref,
) -> Option<FontMetrics> {
    let used_chars = ctx.used_chars;
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

    let widths_1000 = (32u8..=255u8)
        .map(|byte| {
            face.glyph_index(winansi_to_char(byte))
                .map(&advance_1000)
                .unwrap_or(0.0)
        })
        .collect();

    let mut remapper = subsetter::GlyphRemapper::new();
    let mut char_to_gid = HashMap::with_capacity(used_chars.len());
    let mut char_widths_1000 = HashMap::with_capacity(used_chars.len());
    let mut char_ink_1000 = HashMap::with_capacity(used_chars.len());

    // HashSet order differs per process, and the remapper numbers glyphs in visit order:
    // sorting keeps the subset, and so the PDF bytes, the same from run to run.
    let mut chars: Vec<char> = used_chars.iter().copied().collect();
    chars.sort_unstable();
    for ch in chars {
        let gid = resolve_glyph(&face, ch);
        if let Some(gid) = gid {
            let new_gid = remapper.remap(gid.0);
            char_to_gid.insert(ch, new_gid);
            char_widths_1000.insert(ch, advance_1000(gid));
            if let Some(r) = face.glyph_bounding_box(gid) {
                char_ink_1000.insert(ch, (to_1000(r.y_max as f32), to_1000(-r.y_min as f32)));
            }
        }
    }

    let mut gid_remap = HashMap::new();
    let mut extra_widths = Vec::new();
    for (&gid, &ch) in ctx.extra_glyphs.into_iter().flatten() {
        let new_gid = remapper.remap(gid);
        gid_remap.insert(gid, new_gid);
        extra_widths.push((new_gid, advance_1000(ttf_parser::GlyphId(gid)), ch));
    }

    let hyphen_fallback = char_to_gid
        .get(&'\u{2011}')
        .copied()
        .filter(|_| face.glyph_index('\u{2011}').is_none());
    if hyphen_fallback.is_some() {
        log::info!("Font glyph substitution: {font_name}: U+2011 uses an available hyphen glyph");
    }
    // Keep U+2011 in the text when its glyph also draws '-' or U+2010, the
    // stand-in above or a font mapping both to one glyph (Cambria Math): Word's
    // PDF keeps both, and Identity CIDs would merge them into one CID.
    let hyphen_alias = char_to_gid
        .get(&'\u{2011}')
        .copied()
        .filter(|&gid| char_to_gid.values().filter(|&&g| g == gid).count() > 1)
        .map(|gid| {
            // The first CID past every remapped glyph, extra glyphs included.
            let cid = remapper.num_gids();
            char_to_gid.insert('\u{2011}', cid);
            (cid, gid)
        });

    let mut kern_pairs = HashMap::new();
    let char_gids: Vec<(ttf_parser::GlyphId, u16)> = char_to_gid
        .iter()
        .filter_map(|(&ch, &cid)| {
            // resolve_glyph tries U+2011 itself first, then its hyphen fallback.
            let glyph = if ch == '\u{2011}' {
                resolve_glyph(&face, ch)
            } else {
                face.glyph_index(ch)
            };
            glyph.map(|original| (original, cid))
        })
        .collect();

    extract_kern_pairs(&face, &char_gids, &to_1000, &mut kern_pairs);
    extract_gpos_pairs(&face, &char_gids, &to_1000, &mut kern_pairs);

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
    let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&subset_data, 6);
    {
        let mut stream = pdf.stream(refs.data, &compressed);
        stream.filter(Filter::FlateDecode);
        stream.pair(Name(b"Length1"), data_len);
    }

    let ps_name = font_name.replace(' ', "");
    let ps_name_ref = Name(ps_name.as_bytes());
    let system_info = SystemInfo {
        registry: Str(b"Adobe"),
        ordering: Str(b"Identity"),
        supplement: 0,
    };

    pdf.font_descriptor(refs.descriptor)
        .name(ps_name_ref)
        .flags(FontFlags::NON_SYMBOLIC)
        .bbox(bbox)
        .italic_angle(0.0)
        .ascent(ascent)
        .descent(descent)
        .cap_height(cap_height)
        .stem_v(80.0)
        .font_file2(refs.data);

    let cid_to_gid_ref = hyphen_alias.map(|(cid, gid)| {
        let mapping: Vec<u8> = (0..cid).chain([gid]).flat_map(u16::to_be_bytes).collect();
        let reference = alloc();
        pdf.stream(reference, &mapping);
        reference
    });
    let cid_font_ref = alloc();
    {
        let mut cid = pdf.cid_font(cid_font_ref);
        cid.subtype(CidFontType::Type2);
        cid.base_font(ps_name_ref);
        cid.system_info(system_info);
        cid.font_descriptor(refs.descriptor);
        cid.default_width(0.0);
        if let Some(reference) = cid_to_gid_ref {
            cid.cid_to_gid_map_stream(reference);
        } else {
            cid.cid_to_gid_map_predefined(Name(b"Identity"));
        }

        // The width of the glyph each character maps to (`resolve_glyph`): a
        // symbol font's character missing here got /DW 0 while drawn at its
        // real width (7.21.5).
        let mut gid_widths: Vec<(u16, f32)> = char_to_gid
            .iter()
            .map(|(&ch, &new_gid)| (new_gid, char_widths_1000[&ch]))
            .chain(extra_widths.iter().map(|&(gid, w, _)| (gid, w)))
            .collect();
        gid_widths.sort_by_key(|&(gid, _)| gid);
        gid_widths.dedup_by_key(|&mut (gid, _)| gid);
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
    let pua_unicode: fn(u32) -> Option<char> = if font_name.eq_ignore_ascii_case("symbol") {
        symbol_font_unicode
    } else if font_name.eq_ignore_ascii_case("wingdings") {
        wingdings_unicode
    } else if font_name.eq_ignore_ascii_case("wingdings 3") {
        wingdings3_unicode
    } else {
        |_| None
    };
    // Several characters can share a glyph (hyphen variants, no-break space)
    // but a CID maps to one character: keep the lowest code point, the plain
    // form, instead of whichever the HashMap happened to yield last.
    let mut cid_unicode: BTreeMap<u16, char> = BTreeMap::new();
    for (&ch, &new_gid) in &char_to_gid {
        let uni = pua_unicode(symbol_code(ch)).unwrap_or(ch);
        cid_unicode
            .entry(new_gid)
            .and_modify(|c| *c = (*c).min(uni))
            .or_insert(uni);
    }
    for &(cid, _, ch) in &extra_widths {
        cid_unicode.entry(cid).or_insert(ch);
    }
    for (cid, uni) in cid_unicode {
        match crate::docx::arabic::source_letters(uni) {
            Some((letter, None)) => cmap.pair(cid, letter),
            Some((lam, Some(alef))) => cmap.pair_with_multiple(cid, [lam, alef]),
            None => cmap.pair(cid, uni),
        }
    }
    pdf.stream(tounicode_ref, cmap.finish().as_slice());

    pdf.type0_font(refs.font)
        .base_font(ps_name_ref)
        .encoding_predefined(Name(b"Identity-H"))
        .descendant_font(cid_font_ref)
        .to_unicode(tounicode_ref);

    let lm = compute_line_metrics(&face, units);

    Some(FontMetrics {
        widths_1000,
        line_h_ratio: lm.line_h_ratio,
        ascender_ratio: lm.ascender_ratio,
        grid_line_ratio: lm.grid_line_ratio,
        plain_line_h_ratio: lm.plain_line_h_ratio,
        grid_baseline_shift: lm.grid_baseline_shift,
        superscript_ratio: face
            .superscript_metrics()
            .map(|m| m.y_size as f32 / units)
            .filter(|r| *r > 0.0),
        subscript_ratio: face
            .subscript_metrics()
            .map(|m| m.y_size as f32 / units)
            .filter(|r| *r > 0.0),
        underline: face
            .underline_metrics()
            .map(|m| (-(m.position as f32) / units, m.thickness as f32 / units))
            .filter(|&(_, t)| t > 0.0),
        strikeout: face
            .strikeout_metrics()
            .map(|m| (m.position as f32 / units, m.thickness as f32 / units))
            .filter(|&(_, t)| t > 0.0),
        east_asian: lm.east_asian,
        plain_ascender_ratio: lm.plain_ascender_ratio,
        leading_ratio: lm.leading_ratio,
        char_to_gid,
        char_widths_1000,
        char_ink_1000,
        kern_pairs,
        gid_remap,
        stand_ins: hyphen_fallback.map(|_| '\u{2011}').into_iter().collect(),
    })
}

/// A symbol font's private-use code (U+F0xx) for `ch`. Runs can hold the low
/// byte itself (welsh_palliative's U+00B7 bullet); the font maps both to one glyph.
fn symbol_code(ch: char) -> u32 {
    match ch as u32 {
        c @ 0x20..=0xFF => c | 0xF000,
        c => c,
    }
}

/// Adobe Symbol encoding for the Symbol font's codes that documents use as
/// bullets and signs. Word's export maps them the same way, so extracted text
/// reads "•" rather than U+F0B7 or "·".
// ponytail: common bullets/signs only; the full Symbol encoding when Greek/math text needs it
fn symbol_font_unicode(code: u32) -> Option<char> {
    Some(match code {
        0xF0B7 => '•',
        0xF02D => '−',
        0xF0B0 => '°',
        0xF0B1 => '±',
        0xF0B4 => '×',
        0xF0B8 => '÷',
        0xF0A3 => '≤',
        0xF0B3 => '≥',
        0xF0B9 => '≠',
        0xF0AB => '↔',
        0xF0AC => '←',
        0xF0AD => '↑',
        0xF0AE => '→',
        0xF0AF => '↓',
        0xF0A7 => '♣',
        0xF0A8 => '♦',
        0xF0A9 => '♥',
        0xF0AA => '♠',
        0xF0D7 => '⋅',
        0xF0D8 => '¬',
        0xF0E0 => '◊',
        _ => return None,
    })
}

/// Unicode for the Wingdings codes documents use as bullets and signs, from
/// the font's own glyph names (0xA7 `square4` → ▪). Word's export keeps the
/// private-use code in some documents (irish_school's U+F0A8 checkboxes),
/// which a screen reader skips, and maps to Unicode in others (samtale's ☺).
// ponytail: the corpus's codes only (no Wingdings 2/3, Webdings); add the
// full table when other symbols show up
fn wingdings_unicode(code: u32) -> Option<char> {
    Some(match code {
        0xF021 => '✏',
        0xF026 => '📖',
        0xF04A => '☺',
        0xF06C => '●',
        0xF06E => '■',
        0xF06F => '□',
        0xF071 => '❑',
        0xF076 => '❖',
        0xF09E => '·',
        0xF09F => '•',
        0xF0A7 => '▪',
        0xF0A8 => '◻',
        0xF0B2 => '⟡',
        0xF0D8 => '➢',
        0xF0E0 | 0xF0E8 => '→',
        0xF0E4 => '↗',
        0xF0FC => '✔',
        0xF0FE => '☑',
        _ => return None,
    })
}

/// Wingdings 3's triangles, by the font's glyph names (0x84 `trianglecentrt`
/// → ►), for the list bullets that use them (bosch_software); Word's export
/// keeps the private-use code.
// ponytail: the triangles only; its arrows when a document uses them
fn wingdings3_unicode(code: u32) -> Option<char> {
    Some(match code {
        0xF070 => '▲',
        0xF071 => '▼',
        0xF072 => '△',
        0xF073 => '▽',
        0xF074 => '◀',
        0xF075 => '▶',
        0xF076 => '◁',
        0xF077 => '▷',
        0xF081 => '▲',
        0xF082 => '▼',
        0xF083 => '◄',
        0xF084 => '►',
        _ => return None,
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
        // U+2011 keeps its no-break semantics in the source text. Older fonts
        // can lack it while still providing the same visible hyphen shape.
        .or_else(|| {
            (ch == '\u{2011}')
                .then(|| {
                    face.glyph_index('\u{2010}')
                        .or_else(|| face.glyph_index('-'))
                })
                .flatten()
        })
        // A Unicode space the font lacks (U+202F in macOS Arial 5.01 and Aptos
        // Italic) takes the font's own space instead of drawing .notdef.
        // ponytail: spaces only, at U+0020's width; Word rescues any missing
        // glyph from another font at its real width (Arial Italic in
        // learning_cultures): widen the CJK rescue in `register_font` to
        // non-CJK characters when other glyphs go missing
        .or_else(|| {
            matches!(ch, '\u{2000}'..='\u{200A}' | '\u{202F}' | '\u{205F}')
                .then(|| face.glyph_index(' '))
                .flatten()
        })
        // An Arabic form the font lacks draws its plain letter (see
        // `docx::arabic`); a lam-alef ligature draws only the lam.
        .or_else(|| face.glyph_index(crate::docx::arabic::source_letters(ch)?.0))
}

fn extract_kern_pairs(
    face: &Face,
    char_gids: &[(ttf_parser::GlyphId, u16)],
    to_1000: &impl Fn(f32) -> f32,
    kern_pairs: &mut HashMap<(u16, u16), f32>,
) {
    let Some(kern) = face.tables().kern else {
        return;
    };
    let subtables: Vec<_> = kern
        .subtables
        .into_iter()
        .filter(|st| st.horizontal && !st.variable)
        .collect();
    for &(l_orig, l_new) in char_gids {
        for &(r_orig, r_new) in char_gids {
            let total: i16 = subtables
                .iter()
                .filter_map(|st| st.glyphs_kerning(l_orig, r_orig))
                .sum();
            if total != 0 {
                kern_pairs.insert((l_new, r_new), to_1000(total as f32));
            }
        }
    }
}

fn extract_gpos_pairs(
    face: &Face,
    char_gids: &[(ttf_parser::GlyphId, u16)],
    to_1000: &impl Fn(f32) -> f32,
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
                            if let Some((val1, _)) = pair_set.get(r_orig)
                                && val1.x_advance != 0
                            {
                                kern_pairs
                                    .entry((l_new, r_new))
                                    .or_insert(to_1000(val1.x_advance as f32));
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
                            if let Some((val1, _)) = matrix.get((c1, c2))
                                && val1.x_advance != 0
                            {
                                kern_pairs
                                    .entry((l_new, r_new))
                                    .or_insert(to_1000(val1.x_advance as f32));
                            }
                        }
                    }
                }
            }
        }
    }
}

pub(super) struct LineMetrics {
    pub(super) line_h_ratio: f32,
    pub(super) ascender_ratio: f32,
    /// What Word counts docGrid cells with: sTypo metrics for Latin fonts (win +
    /// hhea lineGap would put Yu Mincho's 18pt-grid lines into two cells), the
    /// 1.3× line height for East Asian fonts.
    pub(super) grid_line_ratio: Option<f32>,
    /// The Latin-rule values, for whitespace-only runs and empty paragraph marks
    /// in an East Asian font (`pdf::layout::run_line_metrics`).
    pub(super) plain_line_h_ratio: f32,
    pub(super) plain_ascender_ratio: f32,
    pub(super) leading_ratio: f32,
    /// How far below a docGrid cell's centre Word puts the baseline, per em.
    pub(super) grid_baseline_shift: f32,
    /// Laid out by Word's East Asian rules (1.3× leading, see above).
    pub(super) east_asian: bool,
}

/// Has glyphs for CJK ideographs, Hangul or kana — what Word treats as an East Asian font.
fn is_east_asian_font(face: &Face) -> bool {
    ['一', '가', 'あ']
        .iter()
        .any(|&c| face.glyph_index(c).is_some())
}

/// Word lays out an East Asian font 1.3× taller than its glyph box, the hhea
/// ascent + descent (no lineGap), the extra leading split evenly above and
/// below the glyphs. Yu Gothic, whose hhea box is smaller than its win box,
/// steps 15.0/17.25 at 10.5/12pt in Word, not 17.6/20.1 (online export probe
/// 2026-10-05; Word's Yu Mincho and MS Mincho have hhea = win).
/// Whitespace-only runs keep the plain values. Measurements: roadmap, "CJK
/// Rendering Polish" item 3.
fn compute_line_metrics(face: &Face, units: f32) -> LineMetrics {
    let (plain_line_h_ratio, plain_ascender_ratio, typo_ratio) = plain_line_metrics(face, units);
    let hhea = face.tables().hhea;
    let east_asian = (face.tables().os2.is_some() && is_east_asian_font(face)).then(|| {
        let desc = -(hhea.descender as f32) / units;
        let box_h = (hhea.ascender - hhea.descender) as f32 / units;
        let line_h = box_h * 1.3;
        // The extra leading is split evenly above and below the glyphs:
        // online exports put SimSun/YaHei baselines (10.5-36pt) 0.15/0.19
        // em above where all-leading-above would (local Mac Word's way).
        (line_h, line_h - desc - (line_h - box_h) / 2.0)
    });
    // Word centres a grid-snapped line's glyph box (ascent + descent) in its
    // cell. A Latin font keeps its line gap above the ascent, as in its normal
    // lines; an East Asian font has none (its 1.3× leading is dropped too).
    let grid_baseline_shift = match face.tables().os2 {
        Some(os2) => {
            let typo = os2.use_typographic_metrics() && east_asian.is_none();
            let (asc, desc) = if typo {
                (os2.typographic_ascender(), os2.typographic_descender())
            } else if east_asian.is_some() {
                (hhea.ascender, hhea.descender)
            } else {
                (os2.windows_ascender(), os2.windows_descender())
            };
            let gap = if typo || east_asian.is_some() {
                0
            } else {
                external_leading(face, os2)
            };
            ((asc + desc) as f32 / 2.0 + gap as f32) / units
        }
        None => (face.ascender() + face.descender()) as f32 / 2.0 / units,
    };
    // ponytail: East Asian faces report none; their 1.3× leading is split
    // around the glyphs, so probe them before top alignment relies on it.
    let leading_ratio = match face.tables().os2 {
        Some(os2) if !os2.use_typographic_metrics() && east_asian.is_none() => {
            external_leading(face, os2) as f32 / units
        }
        _ => 0.0,
    };
    LineMetrics {
        leading_ratio,
        line_h_ratio: east_asian.map_or(plain_line_h_ratio, |(h, _)| h),
        ascender_ratio: east_asian.map_or(plain_ascender_ratio, |(_, a)| a),
        grid_line_ratio: east_asian.map(|(h, _)| h).or(typo_ratio),
        plain_line_h_ratio,
        plain_ascender_ratio,
        grid_baseline_shift,
        east_asian: east_asian.is_some(),
    }
}

/// The leading Word adds above the win box: what the hhea box plus lineGap
/// reaches past the win box, never negative (the OpenType OS/2 definition of
/// GDI's tmExternalLeading). Arial and Times New Roman get their whole lineGap
/// (67, 87); Calibri, Constantia, Candara and Palatino Linotype, whose hhea is
/// the typo box plus a gap that exactly fills the win box, get none. Measured
/// 2026-10-21 on 13 faces exported through Word's online service, with the
/// builds it embeds. An earlier form ignored the descents to fit a local Book
/// Antiqua Bold whose hhea descent exceeds its win descent; the build Word uses
/// has no such excess, and that form was wrong for Constantia.
fn external_leading(face: &Face, os2: ttf_parser::os2::Table) -> i16 {
    let hhea_box = face.ascender() as i32 - face.descender() as i32 + face.line_gap() as i32;
    // ttf-parser hands back usWinDescent negated, like hhea's descender.
    let win_box = os2.windows_ascender() as i32 - os2.windows_descender() as i32;
    (hhea_box - win_box).max(0) as i16
}

/// Returns (line_h_ratio, ascender_ratio, typo_line_ratio) by the Latin rules.
fn plain_line_metrics(face: &Face, units: f32) -> (f32, f32, Option<f32>) {
    if let Some(os2) = face.tables().os2 {
        let t_asc = os2.typographic_ascender() as f32;
        let t_desc = os2.typographic_descender() as f32;
        let t_gap = os2.typographic_line_gap() as f32;
        let typo_ratio = Some((t_asc - t_desc + t_gap) / units);
        if os2.use_typographic_metrics() {
            return ((t_asc - t_desc + t_gap) / units, t_asc / units, typo_ratio);
        }
        let win_asc = os2.windows_ascender() as f32;
        let win_desc = os2.windows_descender() as f32;
        // usWinAscent/Descent define glyph clipping bounds; hhea lineGap
        // provides external leading that Word includes in both line spacing
        // and baseline positioning (ascender offset from slot top)
        let gap = external_leading(face, os2) as f32;
        return (
            (win_asc - win_desc + gap) / units,
            (win_asc + gap) / units,
            typo_ratio,
        );
    }

    let line_gap = face.line_gap() as f32;
    (
        (face.ascender() as f32 - face.descender() as f32 + line_gap) / units,
        face.ascender() as f32 / units,
        None,
    )
}

#[cfg(test)]
mod tests {
    use super::*;

    const HYPHEN_ASCII: &[u8] = include_bytes!("testdata/nb-hyphen-ascii.ttf");
    const HYPHEN_UNICODE: &[u8] = include_bytes!("testdata/nb-hyphen-unicode.ttf");
    const HYPHEN_NATIVE: &[u8] = include_bytes!("testdata/nb-hyphen-native.ttf");
    const HYPHEN_NONE: &[u8] = include_bytes!("testdata/nb-hyphen-none.ttf");

    fn embed_hyphens(data: &[u8], chars: &str) -> (FontMetrics, String) {
        let mut pdf = Pdf::new();
        let refs = FontRefs {
            font: Ref::new(1),
            descriptor: Ref::new(2),
            data: Ref::new(3),
        };
        let mut next = 4;
        let metrics = embed_truetype(
            &mut pdf,
            refs,
            "Synthetic hyphen",
            data,
            0,
            &FontContext {
                embedded_fonts: &HashMap::new(),
                font_table: &HashMap::new(),
                used_chars: &chars.chars().collect(),
                word_text: true,
                extra_glyphs: None,
            },
            &mut || {
                let r = Ref::new(next);
                next += 1;
                r
            },
        )
        .unwrap();
        (metrics, String::from_utf8_lossy(&pdf.finish()).into_owned())
    }

    #[test]
    fn missing_nb_hyphen_prefers_unicode_hyphen_then_ascii() {
        for (data, expected) in [
            (HYPHEN_ASCII, '-'),
            (HYPHEN_UNICODE, '\u{2010}'),
            (HYPHEN_NATIVE, '\u{2011}'),
        ] {
            let face = Face::parse(data, 0).unwrap();
            assert_eq!(resolve_glyph(&face, '\u{2011}'), face.glyph_index(expected));
        }
        let face = Face::parse(HYPHEN_NONE, 0).unwrap();
        assert!(resolve_glyph(&face, '\u{2011}').is_none());
    }

    #[test]
    fn fallback_nb_hyphen_has_own_cid_and_advance() {
        for (data, chars, other, width, kern) in [
            (HYPHEN_ASCII, "1-2\u{2011}1", '-', 400.0, -50.0),
            (
                HYPHEN_UNICODE,
                "1-2\u{2010}1\u{2011}2",
                '\u{2010}',
                500.0,
                -75.0,
            ),
        ] {
            let (metrics, pdf) = embed_hyphens(data, chars);
            assert_ne!(
                metrics.char_to_gid[&'\u{2011}'],
                metrics.char_to_gid[&other]
            );
            assert_eq!(metrics.char_widths_1000[&'\u{2011}'], width);
            assert_eq!(
                metrics.kern_pairs[&(metrics.char_to_gid[&'\u{2011}'], metrics.char_to_gid[&'1'])],
                kern
            );
            assert!(pdf.contains("2011"));
            assert!(!pdf.contains("/CIDToGIDMap /Identity"));
        }
    }

    #[test]
    fn solitary_fallback_and_native_hyphen_keep_identity_map() {
        let (_, pdf) = embed_hyphens(HYPHEN_ASCII, "1\u{2011}2");
        assert!(pdf.contains("/CIDToGIDMap /Identity"));
        let (native, pdf) = embed_hyphens(HYPHEN_NATIVE, "1-2\u{2010}1\u{2011}2");
        assert_eq!(native.char_widths_1000[&'\u{2011}'], 700.0);
        assert!(pdf.contains("/CIDToGIDMap /Identity"));
    }

    #[test]
    fn wingdings_bullets_extract_as_unicode() {
        assert_eq!(wingdings_unicode(0xF0A7), Some('▪'));
        assert_eq!(wingdings_unicode(0xF0FC), Some('✔'));
        assert_eq!(wingdings_unicode(symbol_code('§')), Some('▪'));
        assert_eq!(wingdings_unicode(symbol_code(' ')), None);
        assert_eq!(wingdings_unicode(0xF0FA), None);
    }

    #[test]
    fn wingdings3_triangles_extract_as_unicode() {
        assert_eq!(wingdings3_unicode(0xF084), Some('►'));
        assert_eq!(wingdings3_unicode(0xF0FA), None);
    }

    #[test]
    fn symbol_low_bytes_extract_as_unicode() {
        assert_eq!(symbol_font_unicode(0xF0B7), Some('•'));
        assert_eq!(symbol_font_unicode(symbol_code('·')), Some('•'));
        assert_eq!(symbol_font_unicode(symbol_code('a')), None);
    }
}
