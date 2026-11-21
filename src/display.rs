//! Laid-out pages, independent of PDF: what each page of a converted document
//! draws, in painting order, with the fonts and pictures it draws them with.
//!
//! [`crate::convert`] returns a [`Document`] in [`crate::Conversion::display`]
//! (with [`crate::ConvertOptions::display_list`]): the same layout it writes
//! as PDF, for a caller that writes or renders pages itself. Coordinates are PDF user space:
//! points, origin at the page's bottom-left corner, y up.

use std::collections::BTreeMap;
use std::path::PathBuf;
use std::sync::Arc;

/// A 2-D affine matrix `[a, b, c, d, e, f]`, as PDF writes it:
/// `x' = a·x + c·y + e`, `y' = b·x + d·y + f`.
pub type Matrix = [f32; 6];

/// The whole document, laid out.
#[non_exhaustive]
#[derive(Debug, Clone, Default)]
pub struct Document {
    /// The pages, in order.
    pub pages: Vec<Page>,
    /// The fonts [`TextRun::font`] indexes.
    pub fonts: Vec<Font>,
    /// The pictures [`Item::Image`] indexes.
    pub images: Vec<Image>,
    /// Title, author and other metadata.
    pub info: Info,
    /// The logical structure (tagged PDF) tying the pages' marked content
    /// ([`Item::BeginMarked`]) and links to what they are, when the document
    /// is tagged.
    pub structure: Option<Structure>,
    /// The bookmarks, in reading order: one per heading.
    pub outline: Vec<OutlineEntry>,
    /// What the layout drew that this description does not carry yet, by
    /// kind, with how many times: a writer of these pages draws less than
    /// the PDF export where this is not empty.
    pub unsupported: BTreeMap<Unsupported, usize>,
}

/// A kind of drawing [`Document::unsupported`] counts.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash)]
pub enum Unsupported {
    /// A fill with a pattern other than a shape gradient.
    Pattern,
    /// Text or a stroke painted with a gradient (drawn black, or left out).
    GradientTextOrStroke,
    /// An XObject drawn by a name the page's resources do not have.
    UnnamedXObject,
    /// An XObject that was written but not recorded.
    UnrecordedXObject,
    /// A picture's soft mask that was written but not recorded.
    UnrecordedSoftMask,
    /// Forms nested deeper than the display list follows.
    FormTooDeep,
    /// A font with neither a file, a program nor a standard-14 name.
    FontWithoutSource,
    /// Text in a font the page's resources do not have.
    UnknownFont,
    /// A marked sequence that is neither a structure element's content nor
    /// an artifact (left out, what it draws kept).
    MarkedContent,
}

/// What the document says about itself.
#[non_exhaustive]
#[derive(Debug, Clone, Default)]
pub struct Info {
    /// Title, from the document's core properties.
    pub title: Option<String>,
    /// Author, from the document's core properties.
    pub author: Option<String>,
    /// Subject, from the document's core properties.
    pub subject: Option<String>,
    /// Keywords, from the document's core properties.
    pub keywords: Option<String>,
    /// BCP 47 language tag of the document's text.
    pub language: Option<String>,
    /// Whether the export claims PDF/UA-1: the document is tagged and passes
    /// the checks that depend on it (a title, alternate text on every
    /// figure, headings in order, every font embedded, no `.notdef` drawn).
    pub pdf_ua: bool,
}

/// One page.
#[non_exhaustive]
#[derive(Debug, Clone, Default)]
pub struct Page {
    /// Width and height in points.
    pub width: f32,
    /// Height in points.
    pub height: f32,
    /// What the page draws, in painting order.
    pub items: Vec<Item>,
    /// The page's link areas.
    pub links: Vec<Link>,
}

/// What a page draws. Items are painted in order; a [`Item::Transform`] or
/// [`Item::Clip`] applies to the items after it, up to the end of the
/// enclosing [`Item::Group`] (or page).
#[non_exhaustive]
#[derive(Debug, Clone)]
pub enum Item {
    /// Items drawn in a state of their own: transforms and clips inside end
    /// with the group.
    Group(Vec<Item>),
    /// Concatenated to the current transform.
    Transform(Matrix),
    /// Intersected with the current clip.
    Clip {
        /// The clipping outline.
        path: Path,
        /// Which points are inside the path.
        fill_rule: FillRule,
    },
    /// The path filled with a solid colour.
    Fill {
        /// The outline to fill.
        path: Path,
        /// Which points are inside the path.
        fill_rule: FillRule,
        /// The fill colour and opacity.
        paint: Paint,
    },
    /// The path's outline stroked.
    Stroke {
        /// The path to stroke.
        path: Path,
        /// The stroke colour and opacity.
        paint: Paint,
        /// Line width, caps, joins and dashes.
        style: StrokeStyle,
    },
    /// The path filled with a gradient.
    FillGradient {
        /// The outline to fill.
        path: Path,
        /// Which points are inside the path.
        fill_rule: FillRule,
        /// The gradient the path is filled with.
        gradient: LinearGradient,
    },
    /// Glyphs of one font on one baseline.
    Text(TextRun),
    /// A picture of [`Document::images`], drawn into the unit square
    /// of the current transform (as a PDF image XObject is).
    Image(usize),
    /// What is drawn from here to the matching [`Item::EndMarked`] is marked:
    /// content of a structure element, or an artifact. Marked sequences nest,
    /// and may begin and end in different groups, as PDF's do.
    BeginMarked(Mark),
    /// Ends the innermost open marked sequence.
    EndMarked,
}

/// What a marked sequence of a page is.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Mark {
    /// Content of a structure element of type `kind`, which lists it as
    /// [`Kid::Content`] by `mcid`, unique on its page.
    Content {
        /// The type of the structure element the content belongs to.
        kind: StructType,
        /// Marked-content id, unique on its page.
        mcid: u32,
    },
    /// Not part of the document's content, of a type or of none (the
    /// export's borders, shading and comment pane).
    Artifact(Option<ArtifactKind>),
}

/// The type of an artifact.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ArtifactKind {
    /// Headers and footers.
    Pagination,
}

/// The document's logical structure: its elements, the first the root
/// ([`StructType::Document`]), each listing what it holds in reading order.
#[non_exhaustive]
#[derive(Debug, Clone, Default)]
pub struct Structure {
    /// Every element; [`Kid::Element`] indexes them, the first is the root.
    pub elements: Vec<Element>,
}

/// An element of the structure.
#[non_exhaustive]
#[derive(Debug, Clone)]
pub struct Element {
    /// The element's structure type.
    pub kind: StructType,
    /// What the element holds, in reading order.
    pub kids: Vec<Kid>,
    /// Alternate text (a figure's, a formula's).
    pub alt: Option<String>,
    /// The text the element's content stands for (letters drawn as small
    /// caps, for one).
    pub actual_text: Option<String>,
    /// BCP 47 tag of the element's language, when not the document's.
    pub lang: Option<String>,
    /// The element's identifier (each note has one).
    pub id: Option<String>,
    /// A table cell's attributes, when it has any to say (a header scope, or a span of more than one column).
    pub table_cell: Option<TableCell>,
    /// For a list ([`StructType::L`]), how its items are numbered.
    pub list_numbering: Option<ListNumbering>,
}

/// What an element holds.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kid {
    /// An element of [`Structure::elements`], by index.
    Element(usize),
    /// The page's sequence marked [`Mark::Content`] with this `mcid`.
    Content {
        /// Index into [`Document::pages`].
        page: usize,
        /// Marked-content id on that page.
        mcid: u32,
    },
    /// A link area: the index in its page's [`Page::links`].
    Link {
        /// Index into [`Document::pages`].
        page: usize,
        /// Index into that page's [`Page::links`].
        link: usize,
    },
}

/// A table cell's attributes.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TableCell {
    /// For a header cell, the cells it heads.
    pub header_scope: Option<HeaderScope>,
    /// Columns the cell spans, 1 or more.
    pub col_span: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
/// The cells a table header cell heads.
pub enum HeaderScope {
    /// The cells of its row.
    Row,
    /// The cells of its column.
    Column,
    /// The cells of its row and of its column.
    Both,
}

/// How a list's items are numbered, as PDF names it.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ListNumbering {
    /// Bullets of an unspecified shape.
    Unordered,
    /// Numbers of an unspecified system.
    Ordered,
    /// Labels that are terms described by the item's body, not bullets or numbers.
    Description,
    /// Solid round bullets.
    Disc,
    /// Open round bullets.
    Circle,
    /// Solid square bullets.
    Square,
    /// Arabic numerals: 1, 2, 3.
    Decimal,
    /// Lowercase Roman numerals: i, ii, iii.
    LowerRoman,
    /// Uppercase Roman numerals: I, II, III.
    UpperRoman,
    /// Lowercase letters: a, b, c.
    LowerAlpha,
    /// Uppercase letters: A, B, C.
    UpperAlpha,
}

/// A bookmark: a heading and where it is.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq)]
pub struct OutlineEntry {
    /// The heading's level: an entry nests under the nearest entry before it
    /// of a smaller level.
    pub level: u8,
    /// The heading's text.
    pub title: String,
    /// Index into [`Document::pages`].
    pub page: usize,
    /// The heading's top, in the page's user space.
    pub top: f32,
}

/// A path in the current user space.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct Path(pub Vec<Segment>);

#[derive(Debug, Clone, Copy, PartialEq)]
/// A piece of a [`Path`], in the path's user space.
pub enum Segment {
    /// Starts a new subpath at x, y.
    MoveTo(f32, f32),
    /// A straight line to x, y.
    LineTo(f32, f32),
    /// A cubic Bézier curve: two control points, then the end point.
    CubicTo(f32, f32, f32, f32, f32, f32),
    /// A closed rectangle: corner, width, height.
    Rect(f32, f32, f32, f32),
    /// Closes the current subpath with a straight line to its start.
    Close,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
/// Which points a path encloses, for filling and clipping.
pub enum FillRule {
    /// Inside where the nonzero winding number rule says so.
    NonZero,
    /// Inside where the even-odd rule says so.
    EvenOdd,
}

/// A colour and its opacity.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Paint {
    /// The colour, painted at `alpha`.
    pub color: Color,
    /// 0 (transparent) to 1 (opaque).
    pub alpha: f32,
}

/// Device colours, components from 0 to 1.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq)]
pub enum Color {
    /// DeviceGray: 0 black, 1 white.
    Gray(f32),
    /// DeviceRGB: red, green, blue.
    Rgb(f32, f32, f32),
}

/// A linear (axial) gradient. Its axis is in the page's default user space, as a PDF pattern's
/// is: the transforms in force where it fills do not move it.
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq)]
pub struct LinearGradient {
    /// Where the axis starts (offset 0), in the page's default user space.
    pub start: [f32; 2],
    /// Where the axis ends (offset 1), in the page's default user space.
    pub end: [f32; 2],
    /// Offsets from 0 (`start`) to 1 (`end`), ascending.
    pub stops: Vec<GradientStop>,
    /// Whether the first and last colours go on before `start` and past `end`.
    pub extend: [bool; 2],
    /// Interpolated in linear light, as Word draws gradients, rather than between sRGB values.
    pub linear_light: bool,
}

/// A colour of a gradient, sRGB components from 0 to 1.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct GradientStop {
    /// Position along the axis, from 0 (`start`) to 1 (`end`).
    pub offset: f32,
    /// Red, green, blue.
    pub color: [f32; 3],
}

#[non_exhaustive]
#[derive(Debug, Clone, PartialEq)]
/// How a path is stroked.
pub struct StrokeStyle {
    /// Line width, in the current user space.
    pub width: f32,
    /// How the ends of open subpaths are drawn.
    pub cap: LineCap,
    /// How corners are drawn.
    pub join: LineJoin,
    /// Ratio of miter length to line width beyond which a miter join is drawn bevelled.
    pub miter_limit: f32,
    /// Dash lengths; empty draws a solid line.
    pub dash: Vec<f32>,
    /// How far into the dash pattern the line starts.
    pub dash_phase: f32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
/// The shape of the ends of a stroked open subpath.
pub enum LineCap {
    /// Squared off at the endpoint.
    Butt,
    /// A semicircle around the endpoint.
    Round,
    /// Squared off half the line width past the endpoint.
    Square,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
/// The shape of a stroked path's corners.
pub enum LineJoin {
    /// The outer edges extended until they meet.
    Miter,
    /// An arc around the corner.
    Round,
    /// Cut off straight across the corner.
    Bevel,
}

/// Glyphs of one font on one baseline.
#[non_exhaustive]
#[derive(Debug, Clone)]
pub struct TextRun {
    /// Index into [`Document::fonts`].
    pub font: usize,
    /// Font size, in text space units.
    pub size: f32,
    /// Text space to the current user space at the run's origin. Synthetic
    /// italics are a shear in it.
    pub matrix: Matrix,
    /// Horizontal scaling of the glyphs, 1 = 100%.
    pub horizontal_scaling: f32,
    /// Whether the glyphs are filled, outlined, both or invisible.
    pub mode: TextMode,
    /// The glyphs' fill colour and opacity.
    pub fill: Paint,
    /// The outline, drawn in [`TextMode::Stroke`] and
    /// [`TextMode::FillStroke`] (synthetic bold).
    pub stroke: TextStroke,
    /// The glyphs, in the order the run shows them.
    pub glyphs: Vec<Glyph>,
}

/// How a [`TextRun`]'s glyph outlines are stroked.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct TextStroke {
    /// The outline's colour and opacity.
    pub paint: Paint,
    /// Line width, in the run's user space.
    pub width: f32,
}

#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
/// How a [`TextRun`]'s glyphs are painted.
pub enum TextMode {
    /// Filled.
    Fill,
    /// Outlined.
    Stroke,
    /// Filled, then outlined.
    FillStroke,
    /// Drawn for its text only: searchable, not visible.
    Invisible,
}

/// One glyph of a [`TextRun`].
#[non_exhaustive]
#[derive(Debug, Clone, PartialEq)]
pub struct Glyph {
    /// Which glyph of the font it is.
    pub id: GlyphId,
    /// Where the glyph's origin is along the baseline, in text space units
    /// from the run's origin (horizontal scaling included).
    pub x: f32,
    /// How far the glyph moves the pen, in thousandths of the font size
    /// (before horizontal scaling): the font's advance as the PDF export
    /// rounds it.
    pub advance: f32,
    /// The text the glyph stands for (several characters for a ligature,
    /// none for a glyph that stands for nothing).
    pub text: String,
}

/// Which glyph of a [`Font`] a [`Glyph`] draws.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GlyphId {
    /// The glyph id in the font file.
    Gid(u16),
    /// A WinAnsi character code, for a font the export writes as a
    /// single-byte simple font (every [`FontSource::Standard`] font).
    WinAnsi(u8),
}

/// A font, as the layout measured it.
#[non_exhaustive]
#[derive(Debug, Clone)]
pub struct Font {
    /// Where the font's program comes from.
    pub source: FontSource,
    /// The PostScript-like name the PDF export gives it.
    pub name: String,
}

#[non_exhaustive]
#[derive(Debug, Clone)]
/// Where a [`Font`]'s program comes from.
pub enum FontSource {
    /// A font file on disk; `index` is the face in a collection.
    File {
        /// The font file's path.
        path: PathBuf,
        /// The face in a collection, 0 for a single font.
        index: u32,
    },
    /// A font program the document embeds.
    Data {
        /// The font program's bytes.
        data: Arc<Vec<u8>>,
        /// The face in a collection, 0 for a single font.
        index: u32,
    },
    /// One of the PDF standard 14 fonts, by its PostScript name, drawn with
    /// WinAnsi codes.
    Standard(&'static str),
}

/// A picture, encoded as the PDF export encodes it.
#[non_exhaustive]
#[derive(Debug, Clone)]
pub struct Image {
    /// Width in pixels.
    pub width: u32,
    /// Height in pixels.
    pub height: u32,
    /// The colour space of the samples.
    pub color_space: ImageColorSpace,
    /// Bits per colour component of a sample.
    pub bits_per_component: u8,
    /// How `data` is encoded.
    pub filter: ImageFilter,
    /// The samples, encoded by `filter`.
    pub data: Vec<u8>,
    /// Opacity per pixel, a DeviceGray picture.
    pub soft_mask: Option<Box<Image>>,
    /// Whether a viewer smooths the picture when scaling it.
    pub interpolate: bool,
}

#[non_exhaustive]
#[derive(Debug, Clone, PartialEq)]
/// The colour space of an [`Image`]'s samples.
pub enum ImageColorSpace {
    /// DeviceGray: one component per pixel.
    Gray,
    /// DeviceRGB: three components per pixel.
    Rgb,
    /// DeviceCMYK: four components per pixel.
    Cmyk,
    /// A palette of RGB triples, `palette.len() / 3` entries.
    IndexedRgb {
        /// RGB triples, one byte per component.
        palette: Vec<u8>,
    },
}

#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
/// How an [`Image`]'s samples are encoded.
pub enum ImageFilter {
    /// JPEG (DCTDecode).
    Jpeg,
    /// zlib (FlateDecode).
    Flate,
    /// Raw samples.
    None,
}

/// A standard structure type of tagged PDF (ISO 32000-1, 14.8.4): what an
/// element of the document's logical structure is.
#[non_exhaustive]
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum StructType {
    /// The root: the whole document.
    Document,
    /// A section: a group of elements that belong together (a table cell's content, a text box).
    Sect,
    /// A table of contents, of [`StructType::Toci`] items.
    Toc,
    /// An item of a table of contents.
    Toci,
    /// Paragraph.
    P,
    /// Heading, level 1.
    H1,
    /// Heading, level 2.
    H2,
    /// Heading, level 3.
    H3,
    /// Heading, level 4.
    H4,
    /// Heading, level 5.
    H5,
    /// Heading, level 6.
    H6,
    /// A list, of [`StructType::Li`] items, each a [`StructType::Lbl`] and a
    /// [`StructType::LBody`].
    L,
    /// A list item: a [`StructType::Lbl`] and a [`StructType::LBody`].
    Li,
    /// A list item's label: bullet or number.
    Lbl,
    /// A list item's content.
    LBody,
    /// A table: [`StructType::Tr`] rows, in a [`StructType::THead`] and a [`StructType::TBody`].
    Table,
    /// A table's header rows.
    THead,
    /// A table's body rows.
    TBody,
    /// Table row.
    Tr,
    /// Table header cell.
    Th,
    /// Table data cell.
    Td,
    /// A picture or drawing, described by its alternate text.
    Figure,
    /// A mathematical formula, read by its alternate text.
    Formula,
    /// A footnote or endnote.
    Note,
    /// Inline text with its own language or actual text.
    Span,
    /// A link: its text and its link area.
    Link,
}

impl StructType {
    /// The type's name in PDF: `"P"`, `"H1"`, `"TOCI"`...
    pub fn name(self) -> &'static str {
        match self {
            Self::Document => "Document",
            Self::Sect => "Sect",
            Self::Toc => "TOC",
            Self::Toci => "TOCI",
            Self::P => "P",
            Self::H1 => "H1",
            Self::H2 => "H2",
            Self::H3 => "H3",
            Self::H4 => "H4",
            Self::H5 => "H5",
            Self::H6 => "H6",
            Self::L => "L",
            Self::Li => "LI",
            Self::Lbl => "Lbl",
            Self::LBody => "LBody",
            Self::Table => "Table",
            Self::THead => "THead",
            Self::TBody => "TBody",
            Self::Tr => "TR",
            Self::Th => "TH",
            Self::Td => "TD",
            Self::Figure => "Figure",
            Self::Formula => "Formula",
            Self::Note => "Note",
            Self::Span => "Span",
            Self::Link => "Link",
        }
    }
}

/// A link area on a page.
#[non_exhaustive]
#[derive(Debug, Clone)]
pub struct Link {
    /// Left, bottom, right, top.
    pub rect: [f32; 4],
    /// Where the link goes.
    pub target: LinkTarget,
    /// The linked text, a description of the link.
    pub text: String,
}

#[non_exhaustive]
#[derive(Debug, Clone)]
/// Where a [`Link`] goes.
pub enum LinkTarget {
    /// An external address (URL).
    Uri(String),
    /// A page index and the y to show at the top.
    Page {
        /// Index into [`Document::pages`].
        page: usize,
        /// The y to show at the top of the window, in the page's user space.
        y: f32,
    },
}
