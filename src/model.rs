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
    pub break_type: SectionBreakType,
    pub columns: Option<ColumnsConfig>,
    pub page_num_start: Option<u32>,
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

pub type FontTable = std::collections::HashMap<String, FontTableEntry>;

pub struct Document {
    pub sections: Vec<Section>,
    pub line_spacing: LineSpacing,
    /// Fonts embedded in the DOCX (deobfuscated TTF/OTF bytes).
    /// Key: (lowercase_font_name, bold, italic)
    pub embedded_fonts: std::collections::HashMap<(String, bool, bool), Vec<u8>>,
    pub footnotes: std::collections::HashMap<u32, Footnote>,
    pub font_table: FontTable,
    pub even_and_odd_headers: bool,
    /// Maps style IDs to display names (for STYLEREF resolution)
    pub style_id_to_name: std::collections::HashMap<String, String>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum ImageFormat {
    Jpeg,
    Png,
}

#[derive(Clone)]
pub struct EmbeddedImage {
    pub data: std::sync::Arc<Vec<u8>>,
    pub format: ImageFormat,
    pub pixel_width: u32,
    pub pixel_height: u32,
    pub display_width: f32,  // points
    pub display_height: f32, // points
    pub jpeg_components: u8,
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

#[derive(Clone)]
pub struct FloatingImage {
    pub image: EmbeddedImage,
    pub h_position: HorizontalPosition,
    pub h_relative_from: HRelativeFrom,
    pub v_position: VerticalPosition,
    pub v_relative_from: VRelativeFrom,
    pub wrap_type: WrapType,
    pub behind_doc: bool,
}

/// Geometry definition for a shape — either a preset name or custom paths.
/// Supports all 187 OOXML preset shapes and arbitrary custom geometry (a:custGeom).
#[derive(Clone, Debug)]
pub struct ShapeGeometry {
    pub preset: Option<String>,
    pub adjustments: Vec<(String, i64)>,
    pub custom: Option<CustomGeometry>,
}

impl Default for ShapeGeometry {
    fn default() -> Self {
        Self {
            preset: Some("rect".to_string()),
            adjustments: Vec::new(),
            custom: None,
        }
    }
}

#[derive(Clone, Debug)]
pub struct CustomGeometry {
    pub adjust_defaults: Vec<(String, i64)>,
    pub guides: Vec<CustomGuideDef>,
    pub paths: Vec<CustomPathDef>,
}

#[derive(Clone, Debug)]
pub struct CustomGuideDef {
    pub name: String,
    pub op: crate::geometry::FormulaOp,
    pub x: String,
    pub y: String,
    pub z: String,
}

#[derive(Clone, Debug)]
pub struct CustomPathDef {
    pub commands: Vec<CustomPathCommand>,
    pub w: Option<i64>,
    pub h: Option<i64>,
    pub fill: crate::geometry::PathFill,
    pub stroke: bool,
}

#[derive(Clone, Debug)]
pub enum CustomPathCommand {
    MoveTo { x: String, y: String },
    LineTo { x: String, y: String },
    ArcTo { wr: String, hr: String, st_ang: String, sw_ang: String },
    CubicBezTo { x1: String, y1: String, x2: String, y2: String, x3: String, y3: String },
    QuadBezTo { x1: String, y1: String, x2: String, y2: String },
    Close,
}

pub struct SmartArtShape {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub shape_type: ShapeGeometry,
    pub fill: Option<[u8; 3]>,
    pub stroke_color: Option<[u8; 3]>,
    pub stroke_width: f32,
    pub text: String,
    pub font_size: f32,
    pub text_color: Option<[u8; 3]>,
}

pub struct SmartArtDiagram {
    #[allow(dead_code)]
    pub display_width: f32,
    #[allow(dead_code)]
    pub display_height: f32,
    pub shapes: Vec<SmartArtShape>,
}

pub enum ConnectorType {
    Line { flip_h: bool, flip_v: bool },
    Arc { start_angle: f32, end_angle: f32, rotation: f32 },
}

pub struct ConnectorShape {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub stroke_color: [u8; 3],
    pub stroke_width: f32,
    pub connector_type: ConnectorType,
}

pub enum ShapeFill {
    Solid([u8; 3]),
    LinearGradient {
        stops: Vec<([u8; 3], f32)>,
        angle_deg: f32,
    },
}

pub struct Textbox {
    pub paragraphs: Vec<Paragraph>,
    pub width_pt: f32,
    pub height_pt: f32,
    pub h_position: HorizontalPosition,
    pub h_relative_from: HRelativeFrom,
    pub v_offset_pt: f32,
    pub v_relative_from: VRelativeFrom,
    pub fill: Option<ShapeFill>,
    pub shape_type: ShapeGeometry,
    pub margin_left: f32,
    pub margin_right: f32,
    pub margin_top: f32,
    #[allow(dead_code)]
    pub margin_bottom: f32,
    pub wrap_type: WrapType,
    #[allow(dead_code)]
    pub dist_top: f32,
    pub dist_bottom: f32,
    pub behind_doc: bool,
    pub no_text_wrap: bool,
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
    pub contextual_spacing: bool,
    pub keep_next: bool,
    pub keep_lines: bool,
    pub line_spacing: Option<LineSpacing>,
    pub image: Option<EmbeddedImage>,
    pub borders: ParagraphBorders,
    pub shading: Option<[u8; 3]>,
    pub page_break_before: bool,
    pub page_break_after: bool,
    pub column_break_before: bool,
    pub tab_stops: Vec<TabStop>,
    pub extra_line_breaks: u32,
    pub floating_images: Vec<FloatingImage>,
    pub textboxes: Vec<Textbox>,
    pub connectors: Vec<ConnectorShape>,
    pub inline_chart: Option<InlineChart>,
    pub smartart: Option<SmartArtDiagram>,
    pub is_section_break: bool,
}

impl Default for Paragraph {
    fn default() -> Self {
        Self {
            runs: Vec::new(),
            style_id: None,
            space_before: 0.0,
            space_after: 0.0,
            content_height: 0.0,
            alignment: Alignment::Left,
            indent_left: 0.0,
            indent_right: 0.0,
            indent_hanging: 0.0,
            indent_first_line: 0.0,
            list_label: String::new(),
            list_label_font: None,
            contextual_spacing: false,
            keep_next: false,
            keep_lines: false,
            line_spacing: None,
            image: None,
            borders: ParagraphBorders::default(),
            shading: None,
            page_break_before: false,
            page_break_after: false,
            column_break_before: false,
            tab_stops: Vec::new(),
            extra_line_breaks: 0,
            floating_images: Vec::new(),
            textboxes: Vec::new(),
            connectors: Vec::new(),
            inline_chart: None,
            smartart: None,
            is_section_break: false,
        }
    }
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
    pub vertical_align: VertAlign,
    pub field_code: Option<FieldCode>,
    pub hyperlink_url: Option<String>,
    pub inline_image: Option<EmbeddedImage>,
    pub footnote_id: Option<u32>,
    pub is_footnote_ref_mark: bool,
    pub kern_threshold: Option<f32>,
    pub char_style_id: Option<String>,
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
            vertical_align: VertAlign::Baseline,
            field_code: None,
            hyperlink_url: None,
            inline_image: None,
            footnote_id: None,
            is_footnote_ref_mark: false,
            kern_threshold: None,
            char_style_id: None,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum FieldCode {
    Page,
    NumPages,
    StyleRef(String),
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

pub struct TablePosition {
    pub h_position: HorizontalPosition,
    pub h_anchor: &'static str, // "page", "margin", or "column"
    pub v_offset_pt: f32,
    pub v_anchor: &'static str, // "page", "margin", or "text"
}

pub struct Table {
    pub col_widths: Vec<f32>, // points
    pub rows: Vec<TableRow>,
    pub table_indent: f32,
    pub cell_margins: CellMargins,
    pub position: Option<TablePosition>,
}

pub struct TableRow {
    pub cells: Vec<TableCell>,
    pub height: Option<f32>,
    pub height_exact: bool,
    pub is_header: bool,
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

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum MarkerSymbol {
    Circle,
    Square,
    Diamond,
    Triangle,
    Plus,
    Star,
    X,
    Dash,
    Dot,
    None,
}

pub struct ChartSeries {
    pub label: String,
    pub color: Option<[u8; 3]>,
    pub fill_alpha: Option<f32>,
    pub values: Vec<f32>,
    pub x_values: Option<Vec<f32>>,
    pub bubble_sizes: Option<Vec<f32>>,
    pub marker: Option<MarkerSymbol>,
}

pub enum ChartType {
    Bar { horizontal: bool, #[allow(dead_code)] stacked: bool },
    Line,
    Pie,
    Area,
    Scatter,
    Bubble,
    Doughnut { hole_size_pct: f32 },
    Radar,
}

#[derive(Clone)]
pub struct ChartAxis {
    pub labels: Vec<String>,
    #[allow(dead_code)]
    pub delete: bool,
    pub gridline_color: Option<[u8; 3]>,
    pub line_color: Option<[u8; 3]>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum LegendPosition {
    Right,
    Bottom,
    Top,
    Left,
}

pub struct ChartLegend {
    pub position: LegendPosition,
}

pub struct Chart {
    pub chart_type: ChartType,
    pub series: Vec<ChartSeries>,
    pub cat_axis: Option<ChartAxis>,
    pub val_axis: Option<ChartAxis>,
    pub legend: Option<ChartLegend>,
    pub gap_width_pct: f32,
    pub plot_border_color: Option<[u8; 3]>,
    pub accent_colors: Vec<[u8; 3]>,
}

pub struct InlineChart {
    pub chart: Chart,
    pub display_width: f32,
    pub display_height: f32,
}
