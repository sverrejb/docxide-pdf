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

pub struct Document {
    pub page_width: f32,
    pub page_height: f32,
    pub margin_top: f32,
    pub margin_bottom: f32,
    pub margin_left: f32,
    pub margin_right: f32,
    pub line_pitch: f32,
    pub line_spacing: f32, // auto line spacing factor (e.g. 278/240)
    pub blocks: Vec<Block>,
    /// Fonts embedded in the DOCX (deobfuscated TTF/OTF bytes).
    /// Key: (lowercase_font_name, bold, italic)
    pub embedded_fonts: std::collections::HashMap<(String, bool, bool), Vec<u8>>,
    pub header_default: Option<HeaderFooter>,
    pub header_first: Option<HeaderFooter>,
    pub footer_default: Option<HeaderFooter>,
    pub footer_first: Option<HeaderFooter>,
    pub header_margin: f32,
    pub footer_margin: f32,
    pub different_first_page: bool,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum ImageFormat {
    Jpeg,
    Png,
}

pub struct EmbeddedImage {
    pub data: Vec<u8>,
    pub format: ImageFormat,
    pub pixel_width: u32,
    pub pixel_height: u32,
    pub display_width: f32,  // points
    pub display_height: f32, // points
}

#[derive(Clone)]
pub struct BorderBottom {
    pub width_pt: f32,     // line thickness in points
    pub space_pt: f32,     // gap between text and border in points
    pub color: [u8; 3],    // RGB
}

pub struct Paragraph {
    pub runs: Vec<Run>,
    pub space_before: f32,
    pub space_after: f32,
    pub content_height: f32,
    pub alignment: Alignment,
    pub indent_left: f32,
    pub indent_hanging: f32,
    pub list_label: String,
    pub contextual_spacing: bool,
    pub keep_next: bool,
    pub line_spacing: Option<f32>, // per-paragraph override (e.g. 240/240 = 1.0)
    pub image: Option<EmbeddedImage>,
    pub border_bottom: Option<BorderBottom>,
    pub page_break_before: bool,
    pub tab_stops: Vec<TabStop>,
    pub extra_line_breaks: u32,
}

pub struct Run {
    pub text: String,
    pub font_size: f32,
    pub font_name: String,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub color: Option<[u8; 3]>, // None = automatic (black)
    pub is_tab: bool,
    pub vertical_align: VertAlign,
    pub field_code: Option<FieldCode>,
    pub hyperlink_url: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub enum FieldCode {
    Page,
    NumPages,
}

#[derive(Clone, Copy, Debug, Default)]
pub struct CellBorders {
    pub top: bool,
    pub bottom: bool,
    pub left: bool,
    pub right: bool,
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
        Self { top: 0.0, left: 5.4, bottom: 0.0, right: 5.4 }
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
}

pub struct TableCell {
    pub width: f32, // points
    pub paragraphs: Vec<Paragraph>,
    pub borders: CellBorders,
}

pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
}
