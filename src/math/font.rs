//! The math font: its MATH table, `ssty` alternates and glyph metrics.

use ttf_parser::gsub::{AlternateSubstitution, SingleSubstitution, SubstitutionSubtable};
use ttf_parser::math::{Constants, MathValue};
use ttf_parser::{Face, GlyphId, Tag};

pub(crate) struct MathFont<'a> {
    face: Face<'a>,
    constants: Constants<'a>,
    /// The `ssty` feature's alternate subtables, found once.
    ssty: Vec<AlternateSubstitution<'a>>,
    /// The `flac` feature's single substitutions (accents over tall bases).
    flac: Vec<SingleSubstitution<'a>>,
    upem: f32,
}

pub(super) enum Stretched {
    Glyph(u16),
    /// Parts and their offsets (points) along the stretch direction.
    Assembly(Vec<(u16, f32)>),
}

impl<'a> MathFont<'a> {
    /// The face, if it has a MATH table with constants.
    pub(crate) fn new(data: &'a [u8], index: u32) -> Option<Self> {
        let face = Face::parse(data, index).ok()?;
        let constants = face.tables().math?.constants?;
        let subtables = |tag: &[u8; 4]| -> Vec<SubstitutionSubtable<'a>> {
            face.tables()
                .gsub
                .and_then(|gsub| {
                    let feature = gsub.features.find(Tag::from_bytes(tag))?;
                    let lookups = feature.lookup_indices.into_iter();
                    Some(
                        lookups
                            .filter_map(|li| gsub.lookups.get(li))
                            .flat_map(|l| l.subtables.into_iter::<SubstitutionSubtable>())
                            .collect(),
                    )
                })
                .unwrap_or_default()
        };
        let ssty = subtables(b"ssty")
            .into_iter()
            .filter_map(|st| match st {
                SubstitutionSubtable::Alternate(alt) => Some(alt),
                _ => None,
            })
            .collect();
        let flac = subtables(b"flac")
            .into_iter()
            .filter_map(|st| match st {
                SubstitutionSubtable::Single(s) => Some(s),
                _ => None,
            })
            .collect();
        Some(Self {
            upem: face.units_per_em() as f32,
            face,
            constants,
            ssty,
            flac,
        })
    }

    fn pt(&self, units: f32, size: f32) -> f32 {
        units * size / self.upem
    }

    pub(super) fn constant(&self, f: fn(&Constants<'a>) -> MathValue<'a>, size: f32) -> f32 {
        self.pt(f(&self.constants).value as f32, size)
    }

    pub(super) fn script_percent(&self) -> f32 {
        self.constants.script_percent_scale_down() as f32
    }

    pub(super) fn script_script_percent(&self) -> f32 {
        self.constants.script_script_percent_scale_down() as f32
    }

    pub(super) fn display_operator_min_height(&self, size: f32) -> f32 {
        self.pt(self.constants.display_operator_min_height() as f32, size)
    }

    pub(super) fn radical_degree_bottom_raise_percent(&self) -> f32 {
        self.constants.radical_degree_bottom_raise_percent() as f32
    }

    /// The strikethrough's bottom above the baseline and its thickness at
    /// `size`: the OS/2 strikeout (not snapped to the quarter-point grid as a
    /// text run's is in `line_render::os2_strike`).
    pub(super) fn strikeout(&self, size: f32) -> (f32, f32) {
        // ponytail: a face without an OS/2 strikeout gets 0.3em / 0.05em.
        let (top, th) = self
            .face
            .strikeout_metrics()
            .map_or((0.3 * self.upem, 0.05 * self.upem), |m| {
                (m.position as f32, m.thickness as f32)
            });
        (self.pt(top - th, size), self.pt(th, size))
    }

    /// The font's line box per its hhea table at `size`: the ascent and
    /// descent Word gives a display-math line (Cambria Math's OS/2 win
    /// metrics are several ems tall and never used).
    pub(crate) fn line_metrics(&self, size: f32) -> (f32, f32) {
        let asc = self.face.ascender() as f32 + self.face.line_gap() as f32;
        (
            self.pt(asc, size),
            self.pt(-(self.face.descender() as f32), size),
        )
    }

    /// The glyph for `ch` at a script level (0 none, 1 script, 2 scriptscript):
    /// script levels take the `ssty` alternates.
    pub(super) fn glyph(&self, ch: char, level: usize) -> u16 {
        let gid = self.face.glyph_index(ch).map_or(0, |g| g.0);
        if level == 0 {
            return gid;
        }
        self.ssty(gid, level).unwrap_or(gid)
    }

    fn ssty(&self, gid: u16, level: usize) -> Option<u16> {
        self.ssty.iter().find_map(|alt| {
            let set = alt.alternate_sets.get(alt.coverage.get(GlyphId(gid))?)?;
            let n = set.alternates.len() as usize;
            set.alternates.get((level.min(n) - 1) as u16).map(|g| g.0)
        })
    }

    /// The flattened form of an accent, for bases taller than
    /// FlattenedAccentBaseHeight.
    pub(super) fn flac(&self, gid: u16) -> Option<u16> {
        self.flac.iter().find_map(|st| {
            let i = st.coverage().get(GlyphId(gid))?;
            match st {
                SingleSubstitution::Format1 { delta, .. } => {
                    Some((i32::from(gid) + i32::from(*delta)) as u16)
                }
                SingleSubstitution::Format2 { substitutes, .. } => substitutes.get(i).map(|g| g.0),
            }
        })
    }

    /// Advance, ink ascent and ink descent at `size`.
    pub(super) fn metrics(&self, gid: u16, size: f32) -> (f32, f32, f32) {
        let g = GlyphId(gid);
        let adv = self.face.glyph_hor_advance(g).unwrap_or(0) as f32;
        let (asc, desc) = self
            .face
            .glyph_bounding_box(g)
            .map_or((0.0, 0.0), |b| (b.y_max as f32, -(b.y_min as f32)));
        (self.pt(adv, size), self.pt(asc, size), self.pt(desc, size))
    }

    pub(super) fn italic_correction(&self, gid: u16, size: f32) -> f32 {
        self.face
            .tables()
            .math
            .and_then(|m| m.glyph_info)
            .and_then(|i| i.italic_corrections)
            .and_then(|c| c.get(GlyphId(gid)))
            .map_or(0.0, |v| self.pt(v.value as f32, size))
    }

    pub(super) fn top_accent_attachment(&self, gid: u16, size: f32) -> Option<f32> {
        let v = self
            .face
            .tables()
            .math?
            .glyph_info?
            .top_accent_attachments?
            .get(GlyphId(gid))?;
        Some(self.pt(v.value as f32, size))
    }

    /// The MATH table's variants and `gid`'s construction in one direction.
    fn construction(
        &self,
        gid: u16,
        vertical: bool,
    ) -> Option<(
        ttf_parser::math::Variants<'a>,
        ttf_parser::math::GlyphConstruction<'a>,
    )> {
        let variants = self.face.tables().math?.variants?;
        let constructions = if vertical {
            variants.vertical_constructions
        } else {
            variants.horizontal_constructions
        };
        Some((variants, constructions.get(GlyphId(gid))?))
    }

    /// The smallest variant of `gid` at least `target` points tall (or wide),
    /// else its assembly built to that length, else the largest variant.
    pub(super) fn stretch(&self, gid: u16, target: f32, size: f32, vertical: bool) -> Stretched {
        let Some((variants, c)) = self.construction(gid, vertical) else {
            return Stretched::Glyph(gid);
        };
        let mut largest = gid;
        for v in c.variants {
            largest = v.variant_glyph.0;
            if self.pt(v.advance_measurement as f32, size) >= target {
                return Stretched::Glyph(largest);
            }
        }
        match c.assembly {
            Some(a) => Stretched::Assembly(self.assembly(&variants, a, target, size, vertical)),
            None => Stretched::Glyph(largest),
        }
    }

    /// An accent at most `room` points wide: its assembly built to `room`
    /// when that fits, else its widest variant that fits (its narrowest when
    /// none does), in its `flac` form when `flat`.
    pub(super) fn accent(&self, gid: u16, room: f32, size: f32, flat: bool) -> Stretched {
        let flatten = |g: u16| if flat { self.flac(g).unwrap_or(g) } else { g };
        let Some((variants, c)) = self.construction(gid, false) else {
            return Stretched::Glyph(flatten(gid));
        };
        if let Some(a) = c.assembly {
            let parts = self.assembly(&variants, a, room, size, false);
            if parts
                .last()
                .is_some_and(|&(g, x)| x + self.metrics(g, size).0 <= room + 0.01)
            {
                return Stretched::Assembly(parts);
            }
        }
        let fits = |v: &ttf_parser::math::GlyphVariant| {
            self.pt(v.advance_measurement as f32, size) <= room
        };
        let best = c
            .variants
            .into_iter()
            .take_while(fits)
            .last()
            .or(c.variants.get(0))
            .map_or(gid, |v| v.variant_glyph.0);
        Stretched::Glyph(flatten(best))
    }

    /// An assembly's parts and their offsets, built to `target` points.
    fn assembly(
        &self,
        variants: &ttf_parser::math::Variants<'a>,
        assembly: ttf_parser::math::GlyphAssembly<'a>,
        target: f32,
        size: f32,
        vertical: bool,
    ) -> Vec<(u16, f32)> {
        let parts: Vec<_> = assembly.parts.into_iter().collect();
        let overlap = variants.min_connector_overlap as f32;
        let target_units = target * self.upem / size;
        // The fewest extender repeats whose parts, at the minimum overlap,
        // reach the target.
        let sequence = |repeats: usize| -> Vec<_> {
            parts
                .iter()
                .flat_map(|p| {
                    std::iter::repeat_n(*p, if p.part_flags.extender() { repeats } else { 1 })
                })
                .collect()
        };
        let length = |seq: &[ttf_parser::math::GlyphPart]| -> f32 {
            seq.iter().map(|p| p.full_advance as f32).sum::<f32>()
                - overlap * seq.len().saturating_sub(1) as f32
        };
        let mut repeats = 1;
        while length(&sequence(repeats)) < target_units && repeats < 100 {
            repeats += 1;
        }
        let sequence = sequence(repeats);
        // Spread any surplus over the joins, never below the minimum overlap.
        let joins = sequence.len().saturating_sub(1) as f32;
        let surplus = (length(&sequence) - target_units).max(0.0);
        let mut out = Vec::with_capacity(sequence.len());
        let mut pos = 0.0f32;
        for (i, p) in sequence.iter().enumerate() {
            if i > 0 {
                let prev = sequence[i - 1];
                let max_overlap =
                    (prev.end_connector_length.min(p.start_connector_length) as f32).max(overlap);
                let extra = if joins > 0.0 { surplus / joins } else { 0.0 };
                pos -= (overlap + extra).min(max_overlap);
            }
            out.push((p.glyph_id.0, self.pt(pos, size)));
            pos += p.full_advance as f32;
        }
        // Vertical parts stack bottom-up from 0: each sits on its position.
        if vertical {
            for o in &mut out {
                o.1 += self.metrics(o.0, size).2;
            }
        }
        out
    }
}
