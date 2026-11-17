//! What the canvas recorded (`canvas::Rec`), turned into the document's
//! display list (`crate::display`): text codes become glyphs of the fonts
//! they were written with, XObject names become pictures and forms, links
//! get their targets.

use std::collections::HashMap;

use crate::display::{
    DisplayDocument, FillRule, Font, FontSource, Glyph, GradientStop, Info, Item, LinearGradient,
    Link, LinkTarget, Matrix, Page, Paint, Path, Segment, TextRun,
};
use crate::fonts::FontEntry;
use crate::model::Document;

use super::GradientSpec;
use super::canvas::{Captured, IDENTITY, Rec, RecPaint, RecText, TjItem, concat, translate};
use super::layout::LinkAnnotation;

/// Forms drawn inside forms deeper than this are left out.
const MAX_FORM_DEPTH: usize = 8;

/// What one page drew, stream by stream, as `assemble_pdf_pages` puts the
/// streams together.
pub(super) struct RecordedPage {
    pub(super) width: f32,
    pub(super) height: f32,
    /// Headers and footers, zoomed with the body when the page shows comments.
    pub(super) header_footer: Option<Vec<Rec>>,
    pub(super) borders: Option<Vec<Rec>>,
    /// The comment zoom (`comments::page_zoom`), around headers, footers
    /// and the body.
    pub(super) zoom: Option<Matrix>,
    /// The vertical-alignment shift, around the zoomed body.
    pub(super) valign: Option<Matrix>,
    /// Shapes behind the text, the text layer, shapes in front: one list
    /// per stream.
    pub(super) body: Vec<Vec<Rec>>,
    pub(super) comment_pane: Option<Vec<Rec>>,
}

/// Builds the document from its recorded pages.
pub(super) struct Builder<'a> {
    fonts: HashMap<&'a str, &'a FontEntry>,
    page_xobjects: HashMap<&'a str, i32>,
    xobjects: &'a HashMap<i32, Captured>,
    font_index: HashMap<String, usize>,
    image_index: HashMap<i32, usize>,
    /// Where each text object of the stream being resolved left the pen.
    text_ends: HashMap<u32, Matrix>,
    /// The gradients of the page being built, by pattern name.
    gradients: HashMap<String, LinearGradient>,
    document: DisplayDocument,
}

impl<'a> Builder<'a> {
    pub(super) fn new(
        fonts: &'a HashMap<String, FontEntry>,
        page_xobjects: &'a [(String, pdf_writer::Ref)],
        xobjects: &'a HashMap<i32, Captured>,
    ) -> Self {
        Builder {
            fonts: fonts
                .values()
                .map(|entry| (entry.pdf_name.as_str(), entry))
                .collect(),
            page_xobjects: page_xobjects
                .iter()
                .map(|(name, r)| (name.as_str(), r.get()))
                .collect(),
            xobjects,
            font_index: HashMap::new(),
            image_index: HashMap::new(),
            text_ends: HashMap::new(),
            gradients: HashMap::new(),
            document: DisplayDocument::default(),
        }
    }

    pub(super) fn add_page(
        &mut self,
        page: RecordedPage,
        links: Vec<Link>,
        gradients: &[GradientSpec],
    ) {
        self.gradients = gradients
            .iter()
            .map(|spec| (spec.pattern_name.clone(), gradient(spec)))
            .collect();
        let mut items = Vec::new();
        if let Some(header_footer) = page.header_footer {
            let mut group: Vec<Item> = page.zoom.map(Item::Transform).into_iter().collect();
            group.extend(self.stream(header_footer));
            items.push(Item::Group(group));
        }
        if let Some(borders) = page.borders {
            items.push(Item::Group(self.stream(borders)));
        }
        let mut zoomed: Vec<Item> = page.zoom.map(Item::Transform).into_iter().collect();
        for stream in page.body {
            zoomed.push(Item::Group(self.stream(stream)));
        }
        let mut body: Vec<Item> = page.valign.map(Item::Transform).into_iter().collect();
        body.push(Item::Group(zoomed));
        items.push(Item::Group(body));
        if let Some(pane) = page.comment_pane {
            items.push(Item::Group(self.stream(pane)));
        }
        self.document.pages.push(Page {
            width: page.width,
            height: page.height,
            items,
            links,
        });
    }

    pub(super) fn finish(mut self, doc: &Document, language: &str) -> DisplayDocument {
        self.document.info = Info {
            title: doc.title.clone(),
            author: doc.author.clone(),
            subject: doc.subject.clone(),
            keywords: doc.keywords.clone(),
            language: (!language.is_empty()).then(|| language.to_string()),
        };
        self.document
    }

    fn unsupported(&mut self, what: &'static str) {
        *self.document.unsupported.entry(what).or_default() += 1;
    }

    /// One content stream: its text objects are its own.
    fn stream(&mut self, items: Vec<Rec>) -> Vec<Item> {
        let saved = std::mem::take(&mut self.text_ends);
        let resolved = self.items(items, None, 0);
        self.text_ends = saved;
        resolved
    }

    fn items(
        &mut self,
        items: Vec<Rec>,
        xobjects: Option<&HashMap<String, i32>>,
        depth: usize,
    ) -> Vec<Item> {
        let mut out = Vec::with_capacity(items.len());
        for item in items {
            match item {
                Rec::Group(items) => out.push(Item::Group(self.items(items, xobjects, depth))),
                Rec::Transform(m) => out.push(Item::Transform(m)),
                Rec::Clip { path, fill_rule } => out.push(Item::Clip { path, fill_rule }),
                Rec::Fill {
                    path,
                    fill_rule,
                    paint: RecPaint::Pattern(name),
                } => match self.gradients.get(&name) {
                    Some(gradient) => out.push(Item::FillGradient {
                        path,
                        fill_rule,
                        gradient: gradient.clone(),
                    }),
                    None => self.unsupported("pattern not described"),
                },
                Rec::Fill {
                    path,
                    fill_rule,
                    paint,
                } => {
                    if let Some(paint) = self.paint(paint) {
                        out.push(Item::Fill {
                            path,
                            fill_rule,
                            paint,
                        });
                    }
                }
                Rec::Stroke { path, paint, style } => {
                    if let Some(paint) = self.paint(paint) {
                        out.push(Item::Stroke { path, paint, style });
                    }
                }
                Rec::Text(text) => {
                    if let Some(run) = self.text(text) {
                        out.push(Item::Text(run));
                    }
                }
                Rec::XObject(name) => {
                    let reference = match xobjects {
                        Some(local) => local.get(&name).copied(),
                        None => self.page_xobjects.get(name.as_str()).copied(),
                    };
                    match reference {
                        Some(r) => out.extend(self.xobject(r, depth)),
                        None => self.unsupported("xobject without a name"),
                    }
                }
            }
        }
        out
    }

    /// A pattern stroked, or text painted with one, is not described yet: counted, not drawn.
    fn paint(&mut self, paint: RecPaint) -> Option<Paint> {
        match paint {
            RecPaint::Solid(paint) => Some(paint),
            RecPaint::Pattern(_) => {
                self.unsupported("gradient text or stroke");
                None
            }
        }
    }

    fn xobject(&mut self, reference: i32, depth: usize) -> Option<Item> {
        match self.xobjects.get(&reference) {
            Some(Captured::Image(..)) => self.image(reference).map(Item::Image),
            Some(Captured::Form {
                bbox,
                matrix,
                items,
                xobjects,
            }) => {
                if depth >= MAX_FORM_DEPTH {
                    self.unsupported("form nested too deep");
                    return None;
                }
                let local: HashMap<String, i32> = xobjects.iter().cloned().collect();
                let [x0, y0, x1, y1] = *bbox;
                let mut group = Vec::new();
                if *matrix != IDENTITY {
                    group.push(Item::Transform(*matrix));
                }
                group.push(Item::Clip {
                    path: Path(vec![Segment::Rect(x0, y0, x1 - x0, y1 - y0)]),
                    fill_rule: FillRule::NonZero,
                });
                let saved = std::mem::take(&mut self.text_ends);
                group.extend(self.items(items.clone(), Some(&local), depth + 1));
                self.text_ends = saved;
                Some(Item::Group(group))
            }
            None => {
                self.unsupported("xobject not recorded");
                None
            }
        }
    }

    /// The index of the picture written as object `reference`, its soft
    /// mask attached.
    fn image(&mut self, reference: i32) -> Option<usize> {
        if let Some(&index) = self.image_index.get(&reference) {
            return Some(index);
        }
        let Some(Captured::Image(image, mask)) = self.xobjects.get(&reference) else {
            return None;
        };
        let mut image = image.clone();
        if let Some(mask) = mask {
            match self.xobjects.get(mask) {
                Some(Captured::Image(mask, _)) => image.soft_mask = Some(Box::new(mask.clone())),
                _ => self.unsupported("soft mask not recorded"),
            }
        }
        let index = self.document.images.len();
        self.document.images.push(image);
        self.image_index.insert(reference, index);
        Some(index)
    }

    fn font(&mut self, entry: &FontEntry) -> Option<usize> {
        if let Some(&index) = self.font_index.get(&entry.pdf_name) {
            return Some(index);
        }
        let source = if let Some(name) = entry.standard_font {
            FontSource::Standard(name)
        } else if let Some(data) = &entry.embedded_data {
            FontSource::Data {
                data: data.clone(),
                index: entry.face_index,
            }
        } else if let Some(path) = &entry.font_path {
            FontSource::File {
                path: path.clone(),
                index: entry.face_index,
            }
        } else {
            self.unsupported("font without a source");
            return None;
        };
        let name = entry
            .ps_name
            .clone()
            .or_else(|| entry.standard_font.map(str::to_string))
            .unwrap_or_default();
        let index = self.document.fonts.len();
        self.document.fonts.push(Font { source, name });
        self.font_index.insert(entry.pdf_name.clone(), index);
        Some(index)
    }

    /// A show operator's glyphs, each at the pen a viewer puts it at: the
    /// font's `/W` advance, `Tc` and `TJ` moves, all scaled by `Tz`.
    fn text(&mut self, text: RecText) -> Option<TextRun> {
        let fonts = &self.fonts;
        let Some(&entry) = fonts.get(text.font.as_str()) else {
            self.unsupported("text in an unknown font");
            return None;
        };
        let font = self.font(entry)?;
        let matrix = if text.continues {
            self.text_ends
                .get(&text.text_object)
                .copied()
                .unwrap_or(IDENTITY)
        } else {
            text.matrix
        };
        let (size, scale) = (text.size, text.horizontal_scaling);
        let mut pen = 0.0;
        let mut glyphs = Vec::new();
        for item in text.items {
            match item {
                TjItem::Adjust(adjust) => pen -= adjust / 1000.0 * size * scale,
                TjItem::Text(bytes) => {
                    for (id, glyph_text, advance) in decode(entry, &bytes) {
                        glyphs.push(Glyph {
                            id,
                            x: pen,
                            advance,
                            text: glyph_text,
                        });
                        pen += (advance / 1000.0 * size + text.char_spacing) * scale;
                    }
                }
            }
        }
        self.text_ends
            .insert(text.text_object, concat(translate(pen, 0.0), matrix));
        let fill = self.text_paint(text.fill);
        let (stroke_paint, line_width) = text.stroke;
        let stroke = self.text_paint(stroke_paint);
        Some(TextRun {
            font,
            size,
            matrix,
            horizontal_scaling: scale,
            mode: text.mode,
            fill,
            stroke: Some((stroke, line_width)),
            glyphs,
        })
    }

    /// Text painted with a pattern (gradient text) is drawn black.
    fn text_paint(&mut self, paint: RecPaint) -> Paint {
        self.paint(paint).unwrap_or(Paint {
            color: crate::display::Color::Gray(0.0),
            alpha: 1.0,
        })
    }
}

/// A shape's gradient as the PDF export writes its pattern: an axial shading along
/// `GradientSpec::axis`, extended both ways, interpolated in linear light (`CalRGB`).
fn gradient(spec: &GradientSpec) -> LinearGradient {
    let [x0, y0, x1, y1] = spec.axis();
    LinearGradient {
        start: [x0, y0],
        end: [x1, y1],
        stops: spec
            .stops
            .iter()
            .map(|&(color, offset)| GradientStop {
                offset,
                color: color.map(|c| f32::from(c) / 255.0),
            })
            .collect(),
        extend: [true, true],
        linear_light: true,
    }
}

/// The glyphs `bytes` show in `entry`: two-byte codes of an embedded font
/// (`FontEntry::code_glyphs`), WinAnsi bytes of a standard one. A code the
/// font has no glyph for draws `.notdef`, which advances by nothing.
fn decode(entry: &FontEntry, bytes: &[u8]) -> Vec<(u16, String, f32)> {
    if entry.standard_font.is_some() || entry.char_to_gid.is_none() {
        return bytes
            .iter()
            .map(|&byte| {
                let width = byte
                    .checked_sub(32)
                    .and_then(|i| entry.widths_1000.get(usize::from(i)))
                    .copied()
                    .unwrap_or(0.0);
                let text = crate::fonts::winansi_to_char(byte).to_string();
                (u16::from(byte), text, width)
            })
            .collect();
    }
    bytes
        .chunks(2)
        .map(|pair| {
            let code = u16::from_be_bytes([pair[0], pair.get(1).copied().unwrap_or(0)]);
            match entry.code_glyphs.get(&code) {
                Some(glyph) => (glyph.gid, glyph.text.clone(), glyph.width_1000),
                None => (0, String::new(), 0.0),
            }
        })
        .collect()
}

/// The page's link areas, internal ones resolved through the bookmarks (an
/// internal link to a bookmark that does not exist gets none, as in the PDF).
pub(super) fn links(
    page_links: &[LinkAnnotation],
    bookmark_positions: &HashMap<String, (usize, f32)>,
    page_count: usize,
) -> Vec<Link> {
    page_links
        .iter()
        .filter_map(|link| {
            let target = match link.url.strip_prefix('#') {
                Some(anchor) => {
                    let &(page, y) = bookmark_positions.get(anchor)?;
                    LinkTarget::Page {
                        page: page.min(page_count.saturating_sub(1)),
                        y,
                    }
                }
                None => LinkTarget::Uri(link.url.clone()),
            };
            let r = link.rect;
            Some(Link {
                rect: [r.x1, r.y1, r.x2, r.y2],
                target,
                text: link.text.trim().to_string(),
            })
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn decodes_winansi_with_standard_widths() {
        let entry = FontEntry::stub();
        let glyphs = decode(&entry, b"A ");
        assert_eq!(glyphs[0].0, u16::from(b'A'));
        assert_eq!(glyphs[0].1, "A");
        assert_eq!(glyphs[1].2, 500.0);
    }

    #[test]
    fn continued_shows_start_where_the_pen_stopped() {
        let entry = FontEntry {
            standard_font: Some("Helvetica"),
            ..FontEntry::stub()
        };
        let fonts: HashMap<String, FontEntry> = [("F1".to_string(), entry)].into();
        let xobjects = HashMap::new();
        let mut builder = Builder::new(&fonts, &[], &xobjects);
        let show = |continues, bytes: &[u8]| {
            Rec::Text(RecText {
                font: "F1".into(),
                size: 10.0,
                matrix: translate(100.0, 200.0),
                continues,
                text_object: 0,
                char_spacing: 0.0,
                horizontal_scaling: 1.0,
                mode: crate::display::TextMode::Fill,
                fill: RecPaint::Solid(Paint {
                    color: crate::display::Color::Gray(0.0),
                    alpha: 1.0,
                }),
                stroke: (
                    RecPaint::Solid(Paint {
                        color: crate::display::Color::Gray(0.0),
                        alpha: 1.0,
                    }),
                    1.0,
                ),
                items: vec![TjItem::Text(bytes.to_vec()), TjItem::Adjust(-100.0)],
            })
        };
        let items = builder.stream(vec![show(false, b"ab"), show(true, b"c")]);
        let [Item::Text(first), Item::Text(second)] = items.as_slice() else {
            panic!("unexpected items: {items:?}");
        };
        // Every stub glyph is 500 units: 5pt at 10pt, plus the 1pt TJ move.
        assert_eq!(first.glyphs[1].x, 5.0);
        assert_eq!(second.matrix, translate(111.0, 200.0));
    }
}
