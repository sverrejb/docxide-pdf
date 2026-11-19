//! The content streams the layout draws into: `pdf_writer::Content`, which
//! writes the PDF operators unchanged, and, while a recording session is open
//! (`Recording`), what they draw as display items (`crate::display`).
//!
//! The layout keeps drawing PDF operators; this type interprets them as a
//! viewer would (graphics state stack, current path and clip, text state) and
//! records paths, text and XObjects with the state they are drawn in. Text is
//! recorded as the codes it shows: which glyphs and widths they stand for is
//! known from the fonts, which `display_list` resolves once the page is done.
//! Without a session the type is a plain `pdf_writer::Content`.

use std::cell::RefCell;
use std::collections::HashMap;

use pdf_writer::types::{ColorSpaceOperand, LineCapStyle, LineJoinStyle, TextRenderingMode};
use pdf_writer::{Buf, Name, Primitive, Ref, Str};

use crate::display::{
    Color, Document as DisplayDocument, FillRule, Image, LineCap, LineJoin, Matrix, Paint, Path,
    Segment, StrokeStyle, StructType, TextMode,
};

pub(crate) const IDENTITY: Matrix = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];

/// `a` then `b`: the matrix of applying `a` first (PDF's row-vector order).
pub(crate) fn concat(a: Matrix, b: Matrix) -> Matrix {
    [
        a[0] * b[0] + a[1] * b[2],
        a[0] * b[1] + a[1] * b[3],
        a[2] * b[0] + a[3] * b[2],
        a[2] * b[1] + a[3] * b[3],
        a[4] * b[0] + a[5] * b[2] + b[4],
        a[4] * b[1] + a[5] * b[3] + b[5],
    ]
}

pub(crate) fn translate(x: f32, y: f32) -> Matrix {
    [1.0, 0.0, 0.0, 1.0, x, y]
}

// ---------------------------------------------------------------------------
// Recording session

/// What a recording session collects besides the content streams: the
/// XObjects the pass writes, by object number, and the document built from
/// the last pass.
#[derive(Default)]
pub(crate) struct Session {
    pub(crate) xobjects: HashMap<i32, Captured>,
    pub(crate) document: Option<DisplayDocument>,
}

/// An XObject as the display list needs it.
pub(crate) enum Captured {
    /// A picture and the object number of its soft mask.
    Image(Image, Option<i32>),
    /// A form: its bounding box, what it draws and the XObjects its own
    /// resources name.
    Form {
        bbox: [f32; 4],
        items: Vec<Rec>,
        xobjects: Vec<(String, i32)>,
    },
}

thread_local! {
    static SESSION: RefCell<Option<Session>> = const { RefCell::new(None) };
}

/// Whether the content streams created now record what they draw.
pub(crate) fn recording() -> bool {
    SESSION.with(|s| s.borrow().is_some())
}

/// Runs `f` on the open session, if any.
pub(crate) fn with_session(f: impl FnOnce(&mut Session)) {
    SESSION.with(|s| {
        if let Some(session) = s.borrow_mut().as_mut() {
            f(session);
        }
    });
}

/// Records a written image XObject; `image` is only built while a session is
/// open, so plain conversions do not copy picture bytes.
pub(crate) fn capture_image(reference: Ref, image: impl FnOnce() -> Image, soft_mask: Option<Ref>) {
    with_session(|s| {
        s.xobjects.insert(
            reference.get(),
            Captured::Image(image(), soft_mask.map(Ref::get)),
        );
    });
}

/// Records a written form XObject: its bounding box, what it draws (see
/// `Content::take_recording`) and the XObjects its resources name.
pub(crate) fn capture_form(
    reference: Ref,
    bbox: [f32; 4],
    items: Vec<Rec>,
    xobjects: &[(&str, Ref)],
) {
    with_session(|s| {
        s.xobjects.insert(
            reference.get(),
            Captured::Form {
                bbox,
                items,
                xobjects: xobjects
                    .iter()
                    .map(|(name, r)| (name.to_string(), r.get()))
                    .collect(),
            },
        );
    });
}

/// An open recording session; dropping it closes the session.
pub(crate) struct Recording;

impl Recording {
    pub(crate) fn start() -> Self {
        SESSION.with(|s| *s.borrow_mut() = Some(Session::default()));
        Recording
    }

    /// The document the last pass recorded.
    pub(crate) fn finish(self) -> Option<DisplayDocument> {
        SESSION.with(|s| s.borrow_mut().as_mut()?.document.take())
    }
}

impl Drop for Recording {
    fn drop(&mut self) {
        SESSION.with(|s| *s.borrow_mut() = None);
    }
}

// ---------------------------------------------------------------------------
// Recorded items

/// A display item before its text and XObjects are resolved.
#[derive(Debug, Clone)]
pub(crate) enum Rec {
    Group(Vec<Rec>),
    Transform(Matrix),
    Clip {
        path: Path,
        fill_rule: FillRule,
    },
    Fill {
        path: Path,
        fill_rule: FillRule,
        paint: RecPaint,
    },
    Stroke {
        path: Path,
        paint: RecPaint,
        style: StrokeStyle,
    },
    Text(RecText),
    XObject(String),
    /// `BMC` with this tag.
    BeginMarked(String),
    /// `BDC` with an `/MCID`: content of a structure element.
    BeginTagged {
        kind: StructType,
        mcid: i32,
    },
    /// `EMC`.
    EndMarked,
}

#[derive(Debug, Clone, PartialEq)]
pub(crate) enum RecPaint {
    Solid(Paint),
    /// A pattern of the page's resources, by name.
    Pattern(String),
}

/// One show operator (`Tj` or `TJ`) and the text state it draws in.
#[derive(Debug, Clone)]
pub(crate) struct RecText {
    /// The font's resource name.
    pub(crate) font: String,
    pub(crate) size: f32,
    /// The text matrix at the first glyph, unless `continues`.
    pub(crate) matrix: Matrix,
    /// The show starts where the previous show of the same text object
    /// left the pen: its matrix is known once that one's glyphs are.
    pub(crate) continues: bool,
    pub(crate) text_object: u32,
    pub(crate) char_spacing: f32,
    /// `Tz` / 100.
    pub(crate) horizontal_scaling: f32,
    pub(crate) mode: TextMode,
    pub(crate) fill: RecPaint,
    pub(crate) stroke: (RecPaint, f32),
    pub(crate) items: Vec<TjItem>,
}

/// A piece of a `TJ` array.
#[derive(Debug, Clone, PartialEq)]
pub(crate) enum TjItem {
    Text(Vec<u8>),
    /// Thousandths of the font size, subtracted from the pen.
    Adjust(f32),
}

/// A TJ array without pdf-writer's spaces between strings and numbers,
/// `[(ab)9(cd)]` as Word writes it: about 3% of a text page.
struct TjArray<'a>(&'a [TjItem]);

impl Primitive for TjArray<'_> {
    fn write(self, buf: &mut Buf) {
        buf.extend([b'[']);
        for item in self.0 {
            match item {
                TjItem::Text(text) => Str(text).write(buf),
                TjItem::Adjust(adjust) => adjust.write(buf),
            }
        }
        buf.extend([b']']);
    }
}

// ---------------------------------------------------------------------------
// The content stream

/// A content stream: `pdf_writer::Content` plus, while a recording session
/// is open, the display items it draws.
pub(crate) struct Content {
    inner: pdf_writer::Content,
    rec: Option<Box<Recorder>>,
}

impl Content {
    pub(crate) fn new() -> Self {
        Content {
            inner: pdf_writer::Content::new(),
            rec: recording().then(|| Box::new(Recorder::new())),
        }
    }

    pub(crate) fn finish(self) -> Buf {
        self.inner.finish()
    }

    /// What was drawn so far, groups left open closed; recording starts
    /// again from a fresh state.
    pub(crate) fn take_recording(&mut self) -> Vec<Rec> {
        match self.rec.as_mut() {
            Some(rec) => std::mem::replace(rec.as_mut(), Recorder::new()).into_items(),
            None => Vec::new(),
        }
    }

    fn record(&mut self, f: impl FnOnce(&mut Recorder)) {
        if let Some(rec) = self.rec.as_mut() {
            f(rec);
        }
    }

    // Graphics state

    pub(crate) fn save_state(&mut self) -> &mut Self {
        self.inner.save_state();
        self.record(Recorder::save);
        self
    }

    pub(crate) fn restore_state(&mut self) -> &mut Self {
        self.inner.restore_state();
        self.record(Recorder::restore);
        self
    }

    pub(crate) fn transform(&mut self, matrix: [f32; 6]) -> &mut Self {
        self.inner.transform(matrix);
        self.record(|r| r.push(Rec::Transform(matrix)));
        self
    }

    pub(crate) fn set_line_width(&mut self, width: f32) -> &mut Self {
        self.inner.set_line_width(width);
        self.record(|r| r.state.line.width = width);
        self
    }

    pub(crate) fn set_line_cap(&mut self, cap: LineCapStyle) -> &mut Self {
        self.inner.set_line_cap(cap);
        self.record(|r| {
            r.state.line.cap = match cap {
                LineCapStyle::ButtCap => LineCap::Butt,
                LineCapStyle::RoundCap => LineCap::Round,
                LineCapStyle::ProjectingSquareCap => LineCap::Square,
            }
        });
        self
    }

    pub(crate) fn set_line_join(&mut self, join: LineJoinStyle) -> &mut Self {
        self.inner.set_line_join(join);
        self.record(|r| {
            r.state.line.join = match join {
                LineJoinStyle::MiterJoin => LineJoin::Miter,
                LineJoinStyle::RoundJoin => LineJoin::Round,
                LineJoinStyle::BevelJoin => LineJoin::Bevel,
            }
        });
        self
    }

    pub(crate) fn set_miter_limit(&mut self, limit: f32) -> &mut Self {
        self.inner.set_miter_limit(limit);
        self.record(|r| r.state.line.miter_limit = limit);
        self
    }

    pub(crate) fn set_dash_pattern(
        &mut self,
        array: impl IntoIterator<Item = f32>,
        phase: f32,
    ) -> &mut Self {
        let array: Vec<f32> = array.into_iter().collect();
        self.inner.set_dash_pattern(array.iter().copied(), phase);
        self.record(|r| {
            r.state.line.dash = array;
            r.state.line.dash_phase = phase;
        });
        self
    }

    /// `gs` with the page's `GSa{pct}` ExtGState (see assembly): fill at
    /// `pct` percent opacity.
    pub(crate) fn set_fill_alpha(&mut self, pct: u8) -> &mut Self {
        self.inner
            .set_parameters(Name(format!("GSa{pct}").as_bytes()));
        self.record(|r| r.state.fill_alpha = f32::from(pct) / 100.0);
        self
    }

    // Colour

    pub(crate) fn set_fill_rgb(&mut self, r: f32, g: f32, b: f32) -> &mut Self {
        self.inner.set_fill_rgb(r, g, b);
        self.record(|rec| rec.state.fill = FillColor::Color(Color::Rgb(r, g, b)));
        self
    }

    pub(crate) fn set_fill_gray(&mut self, gray: f32) -> &mut Self {
        self.inner.set_fill_gray(gray);
        self.record(|r| r.state.fill = FillColor::Color(Color::Gray(gray)));
        self
    }

    pub(crate) fn set_stroke_rgb(&mut self, r: f32, g: f32, b: f32) -> &mut Self {
        self.inner.set_stroke_rgb(r, g, b);
        self.record(|rec| rec.state.stroke = Color::Rgb(r, g, b));
        self
    }

    pub(crate) fn set_stroke_gray(&mut self, gray: f32) -> &mut Self {
        self.inner.set_stroke_gray(gray);
        self.record(|r| r.state.stroke = Color::Gray(gray));
        self
    }

    pub(crate) fn set_fill_color_space<'a>(
        &mut self,
        space: impl Into<ColorSpaceOperand<'a>>,
    ) -> &mut Self {
        self.inner.set_fill_color_space(space);
        self
    }

    pub(crate) fn set_fill_pattern(
        &mut self,
        tint: impl IntoIterator<Item = f32>,
        name: Name,
    ) -> &mut Self {
        self.inner.set_fill_pattern(tint, name);
        self.record(|r| r.state.fill = FillColor::Pattern(lossy(name)));
        self
    }

    // Paths

    pub(crate) fn move_to(&mut self, x: f32, y: f32) -> &mut Self {
        self.inner.move_to(x, y);
        self.record(|r| r.path.push(Segment::MoveTo(x, y)));
        self
    }

    pub(crate) fn line_to(&mut self, x: f32, y: f32) -> &mut Self {
        self.inner.line_to(x, y);
        self.record(|r| r.path.push(Segment::LineTo(x, y)));
        self
    }

    pub(crate) fn cubic_to(
        &mut self,
        x1: f32,
        y1: f32,
        x2: f32,
        y2: f32,
        x3: f32,
        y3: f32,
    ) -> &mut Self {
        self.inner.cubic_to(x1, y1, x2, y2, x3, y3);
        self.record(|r| r.path.push(Segment::CubicTo(x1, y1, x2, y2, x3, y3)));
        self
    }

    pub(crate) fn close_path(&mut self) -> &mut Self {
        self.inner.close_path();
        self.record(|r| r.path.push(Segment::Close));
        self
    }

    pub(crate) fn rect(&mut self, x: f32, y: f32, width: f32, height: f32) -> &mut Self {
        self.inner.rect(x, y, width, height);
        self.record(|r| r.path.push(Segment::Rect(x, y, width, height)));
        self
    }

    pub(crate) fn stroke(&mut self) -> &mut Self {
        self.inner.stroke();
        self.record(|r| r.paint(None, true));
        self
    }

    pub(crate) fn fill_nonzero(&mut self) -> &mut Self {
        self.inner.fill_nonzero();
        self.record(|r| r.paint(Some(FillRule::NonZero), false));
        self
    }

    pub(crate) fn fill_even_odd(&mut self) -> &mut Self {
        self.inner.fill_even_odd();
        self.record(|r| r.paint(Some(FillRule::EvenOdd), false));
        self
    }

    pub(crate) fn fill_nonzero_and_stroke(&mut self) -> &mut Self {
        self.inner.fill_nonzero_and_stroke();
        self.record(|r| r.paint(Some(FillRule::NonZero), true));
        self
    }

    pub(crate) fn fill_even_odd_and_stroke(&mut self) -> &mut Self {
        self.inner.fill_even_odd_and_stroke();
        self.record(|r| r.paint(Some(FillRule::EvenOdd), true));
        self
    }

    pub(crate) fn end_path(&mut self) -> &mut Self {
        self.inner.end_path();
        self.record(|r| r.paint(None, false));
        self
    }

    pub(crate) fn clip_nonzero(&mut self) -> &mut Self {
        self.inner.clip_nonzero();
        self.record(|r| r.clip = Some(FillRule::NonZero));
        self
    }

    pub(crate) fn clip_even_odd(&mut self) -> &mut Self {
        self.inner.clip_even_odd();
        self.record(|r| r.clip = Some(FillRule::EvenOdd));
        self
    }

    // Text

    pub(crate) fn begin_text(&mut self) -> &mut Self {
        self.inner.begin_text();
        self.record(Recorder::begin_text);
        self
    }

    pub(crate) fn end_text(&mut self) -> &mut Self {
        self.inner.end_text();
        self.record(|r| r.text = None);
        self
    }

    pub(crate) fn set_font(&mut self, font: Name, size: f32) -> &mut Self {
        self.inner.set_font(font, size);
        self.record(|r| r.state.font = Some((lossy(font), size)));
        self
    }

    pub(crate) fn set_char_spacing(&mut self, spacing: f32) -> &mut Self {
        self.inner.set_char_spacing(spacing);
        self.record(|r| r.state.char_spacing = spacing);
        self
    }

    pub(crate) fn set_horizontal_scaling(&mut self, scaling: f32) -> &mut Self {
        self.inner.set_horizontal_scaling(scaling);
        self.record(|r| r.state.horizontal_scaling = scaling / 100.0);
        self
    }

    pub(crate) fn set_text_rendering_mode(&mut self, mode: TextRenderingMode) -> &mut Self {
        self.inner.set_text_rendering_mode(mode);
        self.record(|r| {
            r.state.mode = match mode {
                TextRenderingMode::Fill | TextRenderingMode::FillClip => TextMode::Fill,
                TextRenderingMode::Stroke | TextRenderingMode::StrokeClip => TextMode::Stroke,
                TextRenderingMode::FillStroke | TextRenderingMode::FillStrokeClip => {
                    TextMode::FillStroke
                }
                TextRenderingMode::Invisible | TextRenderingMode::Clip => TextMode::Invisible,
            }
        });
        self
    }

    pub(crate) fn next_line(&mut self, x: f32, y: f32) -> &mut Self {
        self.inner.next_line(x, y);
        self.record(|r| {
            let text = r.text.get_or_insert_with(TextObject::default);
            text.line = concat(translate(x, y), text.line);
            text.start = Some(text.line);
        });
        self
    }

    pub(crate) fn set_text_matrix(&mut self, matrix: [f32; 6]) -> &mut Self {
        self.inner.set_text_matrix(matrix);
        self.record(|r| {
            let text = r.text.get_or_insert_with(TextObject::default);
            text.line = matrix;
            text.start = Some(matrix);
        });
        self
    }

    pub(crate) fn show(&mut self, text: Str) -> &mut Self {
        self.inner.show(text);
        self.record(|r| r.show(vec![TjItem::Text(text.0.to_vec())]));
        self
    }

    /// `TJ`, written compactly (`TjArray`).
    pub(crate) fn show_array(&mut self, items: &[TjItem]) -> &mut Self {
        self.inner.op("TJ").operand(TjArray(items));
        self.record(|r| r.show(items.to_vec()));
        self
    }

    /// `TJ` as pdf-writer writes it: items separated by spaces.
    pub(crate) fn show_positioned(&mut self, items: &[TjItem]) -> &mut Self {
        {
            let mut show = self.inner.show_positioned();
            let mut writer = show.items();
            for item in items {
                match item {
                    TjItem::Text(text) => writer.show(Str(text)),
                    TjItem::Adjust(adjust) => writer.adjust(*adjust),
                };
            }
        }
        self.record(|r| r.show(items.to_vec()));
        self
    }

    // XObjects and marked content

    pub(crate) fn x_object(&mut self, name: Name) -> &mut Self {
        self.inner.x_object(name);
        self.record(|r| r.push(Rec::XObject(lossy(name))));
        self
    }

    pub(crate) fn begin_marked_content(&mut self, tag: Name) -> &mut Self {
        self.inner.begin_marked_content(tag);
        self.record(|r| r.push(Rec::BeginMarked(lossy(tag))));
        self
    }

    /// `/kind <</MCID mcid>> BDC`: content of a structure element.
    pub(crate) fn begin_tagged(&mut self, kind: StructType, mcid: i32) -> &mut Self {
        self.inner
            .begin_marked_content_with_properties(Name(kind.name().as_bytes()))
            .properties()
            .identify(mcid);
        self.record(|r| r.push(Rec::BeginTagged { kind, mcid }));
        self
    }

    pub(crate) fn end_marked_content(&mut self) -> &mut Self {
        self.inner.end_marked_content();
        self.record(|r| r.push(Rec::EndMarked));
        self
    }
}

fn lossy(name: Name) -> String {
    String::from_utf8_lossy(name.0).into_owned()
}

// ---------------------------------------------------------------------------
// The interpreter

#[derive(Debug, Clone, PartialEq)]
enum FillColor {
    Color(Color),
    Pattern(String),
}

/// The graphics state the recorder tracks (`q` saves it, `Q` restores it).
#[derive(Debug, Clone)]
struct State {
    fill: FillColor,
    fill_alpha: f32,
    stroke: Color,
    line: StrokeStyle,
    font: Option<(String, f32)>,
    char_spacing: f32,
    horizontal_scaling: f32,
    mode: TextMode,
}

impl Default for State {
    fn default() -> Self {
        State {
            fill: FillColor::Color(Color::Gray(0.0)),
            fill_alpha: 1.0,
            stroke: Color::Gray(0.0),
            line: StrokeStyle {
                width: 1.0,
                cap: LineCap::Butt,
                join: LineJoin::Miter,
                miter_limit: 10.0,
                dash: Vec::new(),
                dash_phase: 0.0,
            },
            font: None,
            char_spacing: 0.0,
            horizontal_scaling: 1.0,
            mode: TextMode::Fill,
        }
    }
}

impl State {
    fn fill_paint(&self) -> RecPaint {
        match &self.fill {
            FillColor::Color(color) => RecPaint::Solid(Paint {
                color: *color,
                alpha: self.fill_alpha,
            }),
            FillColor::Pattern(name) => RecPaint::Pattern(name.clone()),
        }
    }

    fn stroke_paint(&self) -> RecPaint {
        RecPaint::Solid(Paint {
            color: self.stroke,
            alpha: 1.0,
        })
    }
}

/// The text object between `BT` and `ET`.
#[derive(Debug, Clone)]
struct TextObject {
    id: u32,
    /// The text line matrix (`Td`, `Tm`).
    line: Matrix,
    /// Where the next show starts, when a `Td` or `Tm` placed it; `None`
    /// once a show moved the pen.
    start: Option<Matrix>,
}

impl Default for TextObject {
    fn default() -> Self {
        TextObject {
            id: 0,
            line: IDENTITY,
            start: Some(IDENTITY),
        }
    }
}

struct Frame {
    items: Vec<Rec>,
    saved: State,
}

struct Recorder {
    /// Open groups, the stream's own first.
    frames: Vec<Frame>,
    state: State,
    path: Vec<Segment>,
    /// A `W` waiting for the painting operator that ends the path.
    clip: Option<FillRule>,
    text: Option<TextObject>,
    next_text_object: u32,
}

impl Recorder {
    fn new() -> Self {
        Recorder {
            frames: vec![Frame {
                items: Vec::new(),
                saved: State::default(),
            }],
            state: State::default(),
            path: Vec::new(),
            clip: None,
            text: None,
            next_text_object: 0,
        }
    }

    fn push(&mut self, item: Rec) {
        if let Some(frame) = self.frames.last_mut() {
            frame.items.push(item);
        }
    }

    fn save(&mut self) {
        self.frames.push(Frame {
            items: Vec::new(),
            saved: self.state.clone(),
        });
    }

    /// A `Q` without its `q` is ignored, as viewers do.
    fn restore(&mut self) {
        if self.frames.len() < 2 {
            return;
        }
        if let Some(frame) = self.frames.pop() {
            self.state = frame.saved;
            self.push(Rec::Group(frame.items));
        }
    }

    /// Ends the current path: fills it by `fill`, strokes it if `stroke`,
    /// then clips by it if a `W` asked for it.
    fn paint(&mut self, fill: Option<FillRule>, stroke: bool) {
        let path = Path(std::mem::take(&mut self.path));
        if let Some(fill_rule) = fill {
            let paint = self.state.fill_paint();
            self.push(Rec::Fill {
                path: path.clone(),
                fill_rule,
                paint,
            });
        }
        if stroke {
            let paint = self.state.stroke_paint();
            let style = self.state.line.clone();
            self.push(Rec::Stroke {
                path: path.clone(),
                paint,
                style,
            });
        }
        if let Some(fill_rule) = self.clip.take() {
            self.push(Rec::Clip { path, fill_rule });
        }
    }

    fn begin_text(&mut self) {
        self.text = Some(TextObject {
            id: self.next_text_object,
            ..TextObject::default()
        });
        self.next_text_object += 1;
    }

    fn show(&mut self, items: Vec<TjItem>) {
        let Some((font, size)) = self.state.font.clone() else {
            return;
        };
        let text = self.text.get_or_insert_with(TextObject::default);
        let (matrix, continues) = match text.start.take() {
            Some(start) => (start, false),
            None => (IDENTITY, true),
        };
        let record = RecText {
            font,
            size,
            matrix,
            continues,
            text_object: text.id,
            char_spacing: self.state.char_spacing,
            horizontal_scaling: self.state.horizontal_scaling,
            mode: self.state.mode,
            fill: self.state.fill_paint(),
            stroke: (self.state.stroke_paint(), self.state.line.width),
            items,
        };
        self.push(Rec::Text(record));
    }

    /// The stream's items, the groups still open closed.
    fn into_items(mut self) -> Vec<Rec> {
        while self.frames.len() > 1 {
            self.restore();
        }
        self.frames.pop().map(|f| f.items).unwrap_or_default()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn recorded(draw: impl FnOnce(&mut Content)) -> (Vec<u8>, Vec<Rec>) {
        let _session = Recording::start();
        let mut content = Content::new();
        draw(&mut content);
        let items = content.take_recording();
        (content.finish().to_vec(), items)
    }

    #[test]
    fn writes_what_pdf_writer_writes() {
        let draw = |c: &mut Content| {
            c.save_state();
            c.set_fill_rgb(1.0, 0.0, 0.0);
            c.rect(1.0, 2.0, 3.0, 4.0);
            c.fill_nonzero();
            c.restore_state();
            c.begin_text();
            c.set_font(Name(b"F1"), 12.0);
            c.next_line(10.0, 20.0);
            c.show(Str(b"ab"));
            c.end_text();
        };
        let mut plain = pdf_writer::Content::new();
        plain.save_state();
        plain.set_fill_rgb(1.0, 0.0, 0.0);
        plain.rect(1.0, 2.0, 3.0, 4.0);
        plain.fill_nonzero();
        plain.restore_state();
        plain.begin_text();
        plain.set_font(Name(b"F1"), 12.0);
        plain.next_line(10.0, 20.0);
        plain.show(Str(b"ab"));
        plain.end_text();
        let (bytes, _) = recorded(draw);
        assert_eq!(bytes, plain.finish().to_vec());
    }

    #[test]
    fn records_groups_paints_and_text() {
        let (_, items) = recorded(|c| {
            c.save_state();
            c.set_fill_rgb(1.0, 0.0, 0.0);
            c.rect(1.0, 2.0, 3.0, 4.0);
            c.fill_nonzero();
            c.restore_state();
            c.begin_text();
            c.set_font(Name(b"F1"), 12.0);
            c.next_line(10.0, 20.0);
            c.show(Str(b"ab"));
            c.show(Str(b"c"));
            c.end_text();
        });
        let [Rec::Group(group), Rec::Text(first), Rec::Text(second)] = items.as_slice() else {
            panic!("unexpected items: {items:?}");
        };
        let [Rec::Fill { paint, .. }] = group.as_slice() else {
            panic!("unexpected group: {group:?}");
        };
        assert_eq!(
            *paint,
            RecPaint::Solid(Paint {
                color: Color::Rgb(1.0, 0.0, 0.0),
                alpha: 1.0
            })
        );
        assert_eq!(first.matrix, translate(10.0, 20.0));
        assert!(!first.continues);
        assert!(second.continues);
        assert_eq!(first.text_object, second.text_object);
    }

    #[test]
    fn clip_applies_after_the_painting_operator() {
        let (_, items) = recorded(|c| {
            c.rect(0.0, 0.0, 1.0, 1.0);
            c.clip_nonzero();
            c.end_path();
        });
        assert!(matches!(items.as_slice(), [Rec::Clip { .. }]));
    }

    #[test]
    fn records_marked_content_as_written() {
        let (bytes, items) = recorded(|c| {
            c.begin_marked_content(Name(b"Artifact"));
            c.end_marked_content();
            c.begin_tagged(StructType::P, 3);
            c.end_marked_content();
        });
        let mut plain = pdf_writer::Content::new();
        plain.begin_marked_content(Name(b"Artifact"));
        plain.end_marked_content();
        plain
            .begin_marked_content_with_properties(Name(b"P"))
            .properties()
            .identify(3);
        plain.end_marked_content();
        assert_eq!(bytes, plain.finish().to_vec());
        let [
            Rec::BeginMarked(artifact),
            Rec::EndMarked,
            Rec::BeginTagged {
                kind: StructType::P,
                mcid: 3,
            },
            Rec::EndMarked,
        ] = items.as_slice()
        else {
            panic!("unexpected items: {items:?}");
        };
        assert_eq!(artifact, "Artifact");
    }

    #[test]
    fn records_nothing_without_a_session() {
        let mut content = Content::new();
        content.rect(0.0, 0.0, 1.0, 1.0);
        content.fill_nonzero();
        assert!(content.take_recording().is_empty());
    }
}
