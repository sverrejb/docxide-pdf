#[derive(Clone, Copy, Debug, PartialEq)]
pub enum Alignment {
    Left,
    Center,
    Right,
    Justify,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum TabAlignment {
    Left,
    Center,
    Right,
    Decimal,
}

#[derive(Clone, Debug)]
pub struct TabStop {
    pub position: f32,
    pub alignment: TabAlignment,
    pub leader: Option<char>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum VertAlign {
    Baseline,
    Superscript,
    Subscript,
}

pub struct HeaderFooter {
    pub paragraphs: Vec<Paragraph>,
}

pub struct Footnote {
    pub paragraphs: Vec<Paragraph>,
}

#[derive(Clone, Copy, Debug)]
pub enum LineSpacing {
    Auto(f32),     // multiplier (e.g. 1.0 = single, 1.15 = default)
    Exact(f32),    // fixed height in points
    AtLeast(f32),  // minimum height in points
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum SectionBreakType {
    NextPage,
    Continuous,
    OddPage,
    EvenPage,
}

pub struct ColumnDef {
    pub width: f32, // points
    pub space: f32, // gap after this column, in points
}

pub struct ColumnsConfig {
    pub columns: Vec<ColumnDef>,
    pub sep: bool,
}

pub struct SectionProperties {
    pub page_width: f32,
    pub page_height: f32,
    pub margin_top: f32,
    pub margin_bottom: f32,
    pub margin_left: f32,
    pub margin_right: f32,
    pub header_margin: f32,
    pub footer_margin: f32,
    pub header_default: Option<HeaderFooter>,
    pub header_first: Option<HeaderFooter>,
    pub footer_default: Option<HeaderFooter>,
    pub footer_first: Option<HeaderFooter>,
    pub different_first_page: bool,
    pub line_pitch: f32,
    pub break_type: SectionBreakType,
    pub columns: Option<ColumnsConfig>,
}

pub struct Section {
    pub properties: SectionProperties,
    pub blocks: Vec<Block>,
}

pub struct Document {
    pub sections: Vec<Section>,
    pub line_spacing: LineSpacing,
    /// Fonts embedded in the DOCX (deobfuscated TTF/OTF bytes).
    /// Key: (lowercase_font_name, bold, italic)
    pub embedded_fonts: std::collections::HashMap<(String, bool, bool), Vec<u8>>,
    pub footnotes: std::collections::HashMap<u32, Footnote>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum ImageFormat {
    Jpeg,
    Png,
}

#[derive(Clone)]
pub struct EmbeddedImage {
    pub data: Vec<u8>,
    pub format: ImageFormat,
    pub pixel_width: u32,
    pub pixel_height: u32,
    pub display_width: f32,  // points
    pub display_height: f32, // points
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum HorizontalPosition {
    Offset(f32),
    AlignCenter,
    AlignLeft,
    AlignRight,
}

#[derive(Clone)]
pub struct FloatingImage {
    pub image: EmbeddedImage,
    pub h_position: HorizontalPosition,
    pub h_relative_from: &'static str,
    pub v_offset_pt: f32,
    pub v_relative_from: &'static str,
}

#[derive(Clone)]
pub struct ParagraphBorder {
    pub width_pt: f32,  // line thickness in points
    pub space_pt: f32,  // gap between text and border in points
    pub color: [u8; 3], // RGB
}

#[derive(Clone, Default)]
pub struct ParagraphBorders {
    pub top: Option<ParagraphBorder>,
    pub bottom: Option<ParagraphBorder>,
    pub left: Option<ParagraphBorder>,
    pub right: Option<ParagraphBorder>,
    pub between: Option<ParagraphBorder>,
}

pub struct Paragraph {
    pub runs: Vec<Run>,
    pub space_before: f32,
    pub space_after: f32,
    pub content_height: f32,
    pub alignment: Alignment,
    pub indent_left: f32,
    pub indent_right: f32,
    pub indent_hanging: f32,
    pub indent_first_line: f32,
    pub list_label: String,
    pub contextual_spacing: bool,
    pub keep_next: bool,
    pub keep_lines: bool,
    pub line_spacing: Option<LineSpacing>,
    pub image: Option<EmbeddedImage>,
    pub borders: ParagraphBorders,
    pub shading: Option<[u8; 3]>,
    pub page_break_before: bool,
    pub column_break_before: bool,
    pub tab_stops: Vec<TabStop>,
    pub extra_line_breaks: u32,
    pub floating_images: Vec<FloatingImage>,
}

#[derive(Clone)]
pub struct Run {
    pub text: String,
    pub font_size: f32,
    pub font_name: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub dstrike: bool,
    pub char_spacing: f32,
    pub text_scale: f32, // percentage, 100.0 = normal
    pub caps: bool,
    pub small_caps: bool,
    pub vanish: bool,
    pub color: Option<[u8; 3]>, // None = automatic (black)
    pub highlight: Option<[u8; 3]>,
    pub is_tab: bool,
    pub vertical_align: VertAlign,
    pub field_code: Option<FieldCode>,
    pub hyperlink_url: Option<String>,
    pub inline_image: Option<EmbeddedImage>,
    pub footnote_id: Option<u32>,
    pub is_footnote_ref_mark: bool,
}

#[derive(Clone, Debug, PartialEq)]
pub enum FieldCode {
    Page,
    NumPages,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum VMerge {
    None,
    Restart,
    Continue,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum CellVAlign {
    Top,
    Center,
    Bottom,
}

#[derive(Clone, Copy, Debug)]
pub struct CellBorder {
    pub present: bool,
    pub color: Option<[u8; 3]>,
    pub width: f32,
}

impl Default for CellBorder {
    fn default() -> Self {
        Self {
            present: false,
            color: None,
            width: 0.5,
        }
    }
}

impl CellBorder {
    pub fn visible(color: Option<[u8; 3]>, width: f32) -> Self {
        Self {
            present: true,
            color,
            width,
        }
    }
}

#[derive(Clone, Copy, Debug, Default)]
pub struct CellBorders {
    pub top: CellBorder,
    pub bottom: CellBorder,
    pub left: CellBorder,
    pub right: CellBorder,
}

#[derive(Clone, Copy, Debug)]
pub struct CellMargins {
    pub top: f32,
    pub left: f32,
    pub bottom: f32,
    pub right: f32,
}

impl Default for CellMargins {
    fn default() -> Self {
        Self {
            top: 0.0,
            left: 5.4,
            bottom: 0.0,
            right: 5.4,
        }
    }
}

pub struct Table {
    pub col_widths: Vec<f32>, // points
    pub rows: Vec<TableRow>,
    pub table_indent: f32,
    pub cell_margins: CellMargins,
}

pub struct TableRow {
    pub cells: Vec<TableCell>,
    pub height: Option<f32>,
    pub height_exact: bool,
}

pub struct TableCell {
    pub width: f32, // points
    pub paragraphs: Vec<Paragraph>,
    pub borders: CellBorders,
    pub shading: Option<[u8; 3]>,
    pub grid_span: u16,
    pub v_merge: VMerge,
    pub v_align: CellVAlign,
}

pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
}
