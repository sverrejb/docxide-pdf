mod chart;
mod drawing;
mod table;

use std::collections::HashMap;

// Re-export all public types so consumers keep using `crate::model::*`
pub use chart::*;
pub use drawing::*;
pub use table::*;

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum Alignment {
    #[default]
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

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum VertAlign {
    #[default]
    Baseline,
    Superscript,
    Subscript,
}

pub struct HeaderFooter {
    pub blocks: Vec<Block>,
}

pub struct Footnote {
    pub paragraphs: Vec<Paragraph>,
}

#[derive(Clone, Copy, Debug)]
pub enum LineSpacing {
    Auto(f32),    // multiplier (e.g. 1.0 = single, 1.15 = default)
    Exact(f32),   // fixed height in points
    AtLeast(f32), // minimum height in points
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum DocGridType {
    #[default]
    Default,
    Lines,
    LinesAndChars,
    SnapToChars,
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
    pub header_even: Option<HeaderFooter>,
    pub footer_default: Option<HeaderFooter>,
    pub footer_first: Option<HeaderFooter>,
    pub footer_even: Option<HeaderFooter>,
    pub different_first_page: bool,
    pub line_pitch: f32,
    pub grid_type: DocGridType,
    pub break_type: SectionBreakType,
    pub columns: Option<ColumnsConfig>,
    pub page_num_start: Option<u32>,
    pub page_num_format: Option<String>,
}

pub struct Section {
    pub properties: SectionProperties,
    pub blocks: Vec<Block>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum FontFamily {
    Auto,
    Roman,
    Swiss,
    Modern,
    Script,
    Decorative,
}

#[derive(Clone, Debug)]
pub struct FontTableEntry {
    pub alt_name: Option<String>,
    pub family: FontFamily,
}

pub type FontTable = HashMap<String, FontTableEntry>;

pub struct Document {
    pub sections: Vec<Section>,
    pub line_spacing: LineSpacing,
    /// Fonts embedded in the DOCX (deobfuscated TTF/OTF bytes).
    /// Key: (lowercase_font_name, bold, italic)
    pub embedded_fonts: HashMap<(String, bool, bool), Vec<u8>>,
    pub footnotes: HashMap<u32, Footnote>,
    pub font_table: FontTable,
    pub even_and_odd_headers: bool,
    pub default_tab_stop: f32,
    /// Maps style IDs to display names (for STYLEREF resolution)
    pub style_id_to_name: HashMap<String, String>,
    /// Theme minor (body) font — used as the default chart label font
    pub chart_font_name: String,
    pub title: Option<String>,
    pub author: Option<String>,
    pub subject: Option<String>,
    pub keywords: Option<String>,
    pub auto_hyphenation: bool,
    pub default_lang: Option<String>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum HorizontalPosition {
    Offset(f32),
    AlignCenter,
    AlignLeft,
    AlignRight,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum VerticalPosition {
    Offset(f32),
    AlignTop,
    AlignCenter,
    AlignBottom,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum HRelativeFrom {
    Page,
    Margin,
    Column,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum VRelativeFrom {
    Page,
    Margin,
    TopMargin,
    Paragraph,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum WrapType {
    None,
    Square,
    Tight,
    Through,
    TopAndBottom,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum WrapText {
    BothSides,
    Left,
    Right,
    Largest,
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

#[derive(Default)]
pub struct Paragraph {
    pub runs: Vec<Run>,
    pub style_id: Option<String>,
    pub space_before: f32,
    pub space_after: f32,
    pub content_height: f32,
    pub alignment: Alignment,
    pub indent_left: f32,
    pub indent_right: f32,
    pub indent_hanging: f32,
    pub indent_first_line: f32,
    pub list_label: String,
    pub list_label_font: Option<String>,
    pub list_label_font_size: Option<f32>,
    pub list_label_bold: bool,
    pub list_label_color: Option<[u8; 3]>,
    /// Tab stop from the numbering level definition (pts from paragraph left edge).
    /// Used to position text after the list label on the first line.
    pub num_level_tab_stop: Option<f32>,
    pub contextual_spacing: bool,
    pub keep_next: bool,
    pub keep_lines: bool,
    pub widow_control: bool,
    pub line_spacing: Option<LineSpacing>,
    pub image: Option<EmbeddedImage>,
    pub borders: ParagraphBorders,
    pub shading: Option<[u8; 3]>,
    pub page_break_before: bool,
    pub page_break_after: bool,
    pub column_break_before: bool,
    pub tab_stops: Vec<TabStop>,
    pub floating_images: Vec<FloatingImage>,
    pub textboxes: Vec<Textbox>,
    pub connectors: Vec<ConnectorShape>,
    pub inline_chart: Option<InlineChart>,
    pub smartart: Option<SmartArtDiagram>,
    pub horizontal_rule: Option<HorizontalRule>,
    pub is_section_break: bool,
    pub bookmarks: Vec<String>,
    pub outline_level: Option<u8>,
    pub paragraph_mark_vanish: bool,
    pub snap_to_grid: bool,
    pub suppress_auto_hyphens: bool,
}

pub struct HorizontalRule {
    pub height_pt: f32,
    pub fill_color: [u8; 3],
    pub width_pct: f32,
    /// Standard HR (o:hrstd="t"): render as thin line, height_pt is just spacing
    pub is_standard: bool,
}

#[derive(Clone)]
pub struct Run {
    pub text: String,
    pub font_size: f32,
    pub font_name: String,
    pub east_asia_font_name: Option<String>,
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
    pub is_line_break: bool,
    pub vertical_align: VertAlign,
    pub field_code: Option<FieldCode>,
    pub hyperlink_url: Option<String>,
    pub inline_image: Option<EmbeddedImage>,
    pub footnote_id: Option<u32>,
    pub is_footnote_ref_mark: bool,
    pub kern_threshold: Option<f32>,
    pub char_style_id: Option<String>,
    pub text_outline: Option<TextOutline>,
    pub text_fill: Option<TextFill>,
    pub text_shadow: Option<TextShadow>,
    pub text_glow: Option<TextGlow>,
    pub lang: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct TextOutline {
    pub width_pt: f32,
    pub color: [u8; 3],
}

#[derive(Clone, Debug, PartialEq)]
pub enum TextFill {
    Solid([u8; 3]),
    Gradient {
        stops: Vec<([u8; 3], f32)>,
        angle_deg: f32,
    },
    NoFill,
}

#[derive(Clone, Debug, PartialEq)]
pub struct TextShadow {
    pub color: [u8; 3],
    pub offset_x: f32,
    pub offset_y: f32,
    pub alpha: f32,
}

#[derive(Clone, Debug, PartialEq)]
pub struct TextGlow {
    pub color: [u8; 3],
    pub radius_pt: f32,
}

impl Default for Run {
    fn default() -> Self {
        Self {
            text: String::new(),
            font_size: 0.0,
            font_name: String::new(),
            east_asia_font_name: None,
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            dstrike: false,
            char_spacing: 0.0,
            text_scale: 100.0,
            caps: false,
            small_caps: false,
            vanish: false,
            color: None,
            highlight: None,
            is_tab: false,
            is_line_break: false,
            vertical_align: VertAlign::Baseline,
            field_code: None,
            hyperlink_url: None,
            inline_image: None,
            footnote_id: None,
            is_footnote_ref_mark: false,
            kern_threshold: None,
            char_style_id: None,
            text_outline: None,
            text_fill: None,
            text_shadow: None,
            text_glow: None,
            lang: None,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum FieldCode {
    Page,
    NumPages,
    StyleRef(String),
}

pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
}
