mod chart;
mod drawing;
mod table;

use std::collections::HashMap;

pub use chart::*;
pub use drawing::*;
pub use table::*;

/// `w:suff`: what follows a list label before the text.
#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum LabelSuffix {
    #[default]
    Tab,
    Space,
    Nothing,
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum Alignment {
    #[default]
    Left,
    Center,
    Right,
    Justify,
    /// `w:jc val="distribute"` — "Distribute All Characters Equally"
    /// (§17.18.44). Like Justify, but the last line is stretched too and the
    /// slack is spread between characters rather than only between words.
    Distribute,
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

#[derive(Clone, Debug)]
pub struct Comment {
    pub initials: String,
    pub text: String,
    /// 1-based ordinal in document encounter order. Word renumbers densely
    /// for display ("[R1]", "[R2]", …) regardless of the raw `w:id`.
    pub display_index: u32,
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
    pub width: f32,
    pub space: f32,
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
    /// The margin was written negative: the header (footer) never pushes the
    /// body away from it.
    pub margin_top_fixed: bool,
    pub margin_bottom_fixed: bool,
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
    pub page_borders: Option<PageBorders>,
    /// §17.6.23 `w:vAlign` — vertical alignment of body text between the top and
    /// bottom margins on each page of the section.
    pub vertical_align: PageVerticalAlign,
    /// §17.6.8 `w:lnNumType` — margin line numbering for this section.
    pub line_numbering: Option<LineNumbering>,
    /// §17.11.11/.5 `w:sectPr/w:footnotePr|w:endnotePr` `w:numFmt @val` — section-wide
    /// footnote/endnote mark numbering format. None → built-in default (footnote
    /// decimal, endnote lowerRoman). Word reads the format here, NOT from the
    /// doc-wide settings.xml bag.
    pub footnote_num_fmt: Option<String>,
    pub endnote_num_fmt: Option<String>,
}

impl SectionProperties {
    /// Width of the text area between the side margins.
    pub fn text_width(&self) -> f32 {
        self.page_width - self.margin_left - self.margin_right
    }

    /// The docGrid line pitch when the grid snaps lines, else None.
    pub fn line_grid_pitch(&self) -> Option<f32> {
        (matches!(
            self.grid_type,
            DocGridType::Lines | DocGridType::LinesAndChars | DocGridType::SnapToChars
        ) && self.line_pitch > 0.0)
            .then_some(self.line_pitch)
    }
}

/// §17.6.8 `w:lnNumType` — line numbers shown in the margin (legal/contract docs).
#[derive(Clone, Copy)]
pub struct LineNumbering {
    /// `@countBy` — show a number on every Nth line (default 1 = every line).
    pub count_by: u32,
    /// `@start` — starting line-number value (default 1).
    pub start: i32,
    /// `@distance` (points) from the text margin to the numbers; `None` = Word's
    /// auto default of 0.25".
    pub distance: Option<f32>,
    /// `@restart` — when the counter resets (ST_LineNumberRestart §17.18.47).
    pub restart: LineNumberRestart,
}

#[derive(Clone, Copy, Default, PartialEq)]
pub enum LineNumberRestart {
    #[default]
    NewPage,
    NewSection,
    Continuous,
}

/// §17.6.23 ST_VerticalJc — how the section's body content is positioned
/// vertically within the text region of each page.
#[derive(Clone, Copy, Default, PartialEq)]
pub enum PageVerticalAlign {
    #[default]
    Top,
    Center,
    /// `both` (vertical justify) — falls back to Top; true inter-paragraph
    /// distribution is not implemented (unexercised in the corpus).
    Both,
    Bottom,
}

#[derive(Clone, Copy, Default, PartialEq)]
pub enum PageBorderDisplay {
    #[default]
    AllPages,
    FirstPage,
    NotFirstPage,
}

#[derive(Clone, Default)]
pub struct PageBorders {
    pub top: Option<ParagraphBorder>,
    pub bottom: Option<ParagraphBorder>,
    pub left: Option<ParagraphBorder>,
    pub right: Option<ParagraphBorder>,
    /// `@offsetFrom="page"` → each edge's `space` is the distance from the page
    /// edge to the border; `"text"` (the default) → distance outward from the
    /// text margins (§17.6.10 / ST_PageBorderOffset §17.18.63).
    pub offset_from_page: bool,
    pub display: PageBorderDisplay,
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
    /// Windows charset byte (`w:charset`, hex in the XML): 0x80 Shift-JIS, 0x81 Hangul,
    /// 0x86 GB2312, 0x88 Big5. Word keys missing-font substitution on it.
    pub charset: Option<u8>,
}

pub type FontTable = HashMap<String, FontTableEntry>;

pub struct Document {
    pub sections: Vec<Section>,
    pub line_spacing: LineSpacing,
    /// Fonts embedded in the DOCX (deobfuscated TTF/OTF bytes).
    /// Key: (lowercase_font_name, bold, italic)
    pub embedded_fonts: HashMap<(String, bool, bool), Vec<u8>>,
    pub footnotes: HashMap<u32, Footnote>,
    /// footnotes.xml's separator paragraph: Word lays it out above a page's
    /// notes like any paragraph and draws the rule as its strikethrough.
    pub footnote_separator: Option<Paragraph>,
    pub endnotes: HashMap<u32, Footnote>,
    pub comments: HashMap<u32, Comment>,
    pub font_table: FontTable,
    pub even_and_odd_headers: bool,
    /// `w:mirrorMargins`: like evenAndOddHeaders, makes odd/even section breaks
    /// insert filler pages even when the section restarts its numbering.
    pub mirror_margins: bool,
    pub default_tab_stop: f32,
    /// Maps style IDs to display names (for STYLEREF resolution)
    pub style_id_to_name: HashMap<String, String>,
    /// Theme minor (body) font: the default chart label font, and what Word
    /// substitutes for a missing font with no declared family.
    pub theme_minor_font: String,
    pub title: Option<String>,
    pub author: Option<String>,
    pub subject: Option<String>,
    pub keywords: Option<String>,
    pub default_lang: Option<String>,
    /// Word's `compressPunctuation` character-spacing control (see
    /// `docx::settings`); drives full-width punctuation squeezing in line breaking.
    pub compress_punctuation: bool,
    /// Word's `compatibilityMode` (see `docx::settings`).
    pub compat_mode: u32,
    /// Word's `doNotExpandShiftReturn` (see `docx::settings`).
    pub do_not_expand_shift_return: bool,
    /// Word's `adjustLineHeightInTable` (see `docx::settings`).
    pub adjust_line_height_in_table: bool,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum HorizontalPosition {
    Offset(f32),
    AlignCenter,
    AlignLeft,
    AlignRight,
}

impl Default for HorizontalPosition {
    fn default() -> Self {
        Self::Offset(0.0)
    }
}

impl HorizontalPosition {
    /// The explicit offset; 0 for the alignment variants.
    pub fn offset_or_zero(self) -> f32 {
        match self {
            Self::Offset(o) => o,
            _ => 0.0,
        }
    }

    /// Left edge of an object `obj_w` wide placed in the area `area_w` wide
    /// that starts at `origin`.
    pub fn place(self, origin: f32, area_w: f32, obj_w: f32) -> f32 {
        match self {
            HorizontalPosition::AlignCenter => origin + (area_w - obj_w) / 2.0,
            HorizontalPosition::AlignRight => origin + area_w - obj_w,
            HorizontalPosition::AlignLeft => origin,
            HorizontalPosition::Offset(o) => origin + o,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum VerticalPosition {
    Offset(f32),
    AlignTop,
    AlignCenter,
    AlignBottom,
}

impl Default for VerticalPosition {
    fn default() -> Self {
        Self::Offset(0.0)
    }
}

impl VerticalPosition {
    /// The explicit offset; 0 for the alignment variants.
    pub fn offset_or_zero(self) -> f32 {
        match self {
            Self::Offset(o) => o,
            _ => 0.0,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Default)]
pub enum HRelativeFrom {
    Page,
    Margin,
    #[default]
    Column,
}

#[derive(Clone, Copy, Debug, PartialEq, Default)]
pub enum VRelativeFrom {
    Page,
    Margin,
    TopMargin,
    #[default]
    Paragraph,
}

#[derive(Clone, Copy, Debug, PartialEq, Default)]
pub enum WrapType {
    #[default]
    None,
    Square,
    Tight,
    Through,
    TopAndBottom,
}

impl WrapType {
    /// Text flows beside the object (square, tight or through wrapping).
    pub fn wraps_beside(self) -> bool {
        matches!(self, WrapType::Square | WrapType::Tight | WrapType::Through)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Default)]
pub enum WrapText {
    #[default]
    BothSides,
    Left,
    Right,
    Largest,
}

/// Consecutive paragraphs with equal properties share one frame; Word keeps
/// two frames apart even when only their hSpace differs.
#[derive(Clone, Debug, PartialEq)]
pub struct FrameProperties {
    pub h_relative_from: HRelativeFrom,
    pub h_position: HorizontalPosition,
    pub v_relative_from: VRelativeFrom,
    /// `w:y`, or `w:yAlign` within the vAnchor area.
    pub v_position: VerticalPosition,
    /// Frame width `w:w` in points (0 = auto/unspecified).
    pub width: f32,
    /// Frame height `w:h` in points (0 = auto/unspecified).
    pub height: f32,
    /// `w:wrap` none/notBeside: body text may not flow beside the frame, so
    /// in-flow text starts below its bottom edge.
    pub text_below: bool,
    /// `w:hSpace` / `w:vSpace` in points.
    pub h_space: f32,
    pub v_space: f32,
}

#[derive(Clone, Debug, PartialEq)]
pub struct ParagraphBorder {
    pub width_pt: f32,
    pub space_pt: f32,
    pub color: [u8; 3],
}

#[derive(Clone, Default)]
pub struct ParagraphBorders {
    pub top: Option<ParagraphBorder>,
    pub bottom: Option<ParagraphBorder>,
    pub left: Option<ParagraphBorder>,
    pub right: Option<ParagraphBorder>,
    pub between: Option<ParagraphBorder>,
}

/// A numbered or bulleted paragraph's place in its list.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct ListItem {
    /// `w:ilvl`.
    pub level: u8,
    /// Abstract list id: numIds sharing a definition are one list.
    pub list_id: u32,
    /// How the level's labels look (the tagged L's `/ListNumbering`).
    pub numbering: pdf_writer::types::ListNumbering,
}

#[derive(Default)]
pub struct Paragraph {
    pub runs: Vec<Run>,
    pub style_id: Option<String>,
    pub space_before: f32,
    pub space_after: f32,
    /// The before/after came from HTML auto spacing, which a table cell drops
    /// at its top and bottom edge.
    pub space_before_auto: bool,
    pub space_after_auto: bool,
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
    /// `w:lvlJc`: the label is left-, centre- or right-aligned on its position.
    pub list_label_jc: Alignment,
    pub list_label_suff: LabelSuffix,
    /// For L/LI tagging.
    pub list_item: Option<ListItem>,
    /// A `TOC` field begins here. Word tags "toc N" paragraphs as TOC/TOCI
    /// only inside such a field; hand-styled ones stay paragraphs.
    pub starts_toc_field: bool,
    /// Numbering level tab stop (pts from paragraph left edge)
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
    /// True when `page_break_before` originated from an explicit
    /// `<w:br w:type="page"/>` at the start of the paragraph (as opposed to
    /// the `<w:pageBreakBefore/>` style property). An explicit break must
    /// advance the page even if we're already at the top of one — Word
    /// emits a blank page in that case; the style property is idempotent.
    pub page_break_before_explicit: bool,
    pub page_break_after: bool,
    /// Run index where a page break inside the paragraph splits it; the
    /// body parser turns the rest into a continuation paragraph.
    pub page_break_at: Option<usize>,
    pub column_break_before: bool,
    /// A column break with nothing after it: the next paragraph starts in the
    /// next column (or page, in a one-column section).
    pub column_break_after: bool,
    /// §17.3.3.1 `w:br w:type="textWrapping" w:clear="all"` — content after
    /// this paragraph restarts below any floating objects.
    pub clears_floats: bool,
    pub tab_stops: Vec<TabStop>,
    pub floating_images: Vec<FloatingImage>,
    pub textboxes: Vec<Textbox>,
    pub connectors: Vec<ConnectorShape>,
    pub inline_chart: Option<InlineChart>,
    pub smartart: Vec<SmartArtDiagram>,
    /// Diagrams drawn at their anchor over the text (`SmartArtDiagram::anchor`).
    pub floating_smartart: Vec<SmartArtDiagram>,
    pub horizontal_rule: Option<HorizontalRule>,
    pub is_section_break: bool,
    pub bookmarks: Vec<String>,
    pub outline_level: Option<u8>,
    pub paragraph_mark_vanish: bool,
    pub paragraph_mark_font_size: Option<f32>,
    pub paragraph_mark_font_name: Option<String>,
    /// The mark's `w:position` in points (negative = lowered).
    pub paragraph_mark_position: f32,
    pub snap_to_grid: bool,
    pub auto_space_de: bool,
    pub auto_space_dn: bool,
    pub frame_props: Option<FrameProperties>,
}

pub struct HorizontalRule {
    pub height_pt: f32,
    pub fill_color: [u8; 3],
    pub width_pct: f32,
    /// When true (o:hrstd), render as thin line; height_pt is spacing only
    pub is_standard: bool,
}

#[derive(Clone)]
pub struct Run {
    pub text: String,
    pub font_size: f32,
    pub font_name: String,
    pub east_asia_font_name: Option<String>,
    /// `w:rFonts@cs`: the font for complex-script letters (Arabic, Hebrew, Thai, …).
    pub cs_font_name: Option<String>,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub double_underline: bool,
    pub strikethrough: bool,
    pub dstrike: bool,
    pub char_spacing: f32,
    pub text_scale: f32,
    pub caps: bool,
    pub small_caps: bool,
    pub vanish: bool,
    pub color: Option<[u8; 3]>, // None = DOCX "automatic" (typically black)
    pub highlight: Option<[u8; 3]>,
    /// Background shading from `w:rPr/w:shd`. Distinct from `highlight`
    /// (which is `w:highlight` — predefined named colors). Both can coexist.
    pub shading: Option<[u8; 3]>,
    pub border: Option<ParagraphBorder>,
    pub is_tab: bool,
    /// `Some(alignment)` marks this tab as a positional tab (`w:ptab`): it is
    /// resolved against the margin box with its own alignment rather than the
    /// paragraph's tab stops. `is_tab` is also true so it splits line segments.
    pub ptab_alignment: Option<TabAlignment>,
    pub is_line_break: bool,
    pub vertical_align: VertAlign,
    pub field_code: Option<FieldCode>,
    pub hyperlink_url: Option<String>,
    pub inline_image: Option<EmbeddedImage>,
    pub footnote_id: Option<u32>,
    pub is_footnote_ref_mark: bool,
    pub endnote_id: Option<u32>,
    pub is_endnote_ref_mark: bool,
    /// `w:kern` in points; see `Run::kerns_at`.
    pub kern_threshold: Option<f32>,
    /// Points the run is raised above the baseline (`w:position`; negative: lowered).
    pub position: f32,
    pub char_style_id: Option<String>,
    pub text_outline: Option<TextOutline>,
    pub text_fill: Option<TextFill>,
    pub text_shadow: Option<TextShadow>,
    pub text_glow: Option<TextGlow>,
    pub lang: Option<String>,
    /// The language of the run's Latin, East Asian and complex-script text, inherited (run,
    /// character style, paragraph style, docDefaults); `lang` is the run's own
    /// value only, which run merging compares.
    pub text_lang: Option<std::sync::Arc<str>>,
    pub text_lang_east_asia: Option<std::sync::Arc<str>>,
    pub text_lang_bidi: Option<std::sync::Arc<str>>,
    /// True when font_size was inherited from defaults, not set by inline rPr or char style.
    pub font_size_from_default: bool,
    /// True when font_name was inherited from defaults, not set by inline rPr or char style.
    pub font_name_from_default: bool,
    /// True when the run's own rPr sets w:b / w:i. Direct formatting is absolute
    /// for toggle properties (§17.7.3), so a table style's bold or italic must not
    /// override it.
    pub bold_is_direct: bool,
    pub italic_is_direct: bool,
    /// Active comment IDs covering this run (empty for the common no-comments case).
    /// Multiple IDs when comment ranges overlap.
    pub comment_ids: Vec<u32>,
    /// True for runs synthesized from Office Math (OMML). The math font (e.g.
    /// Cambria Math) has very tall metrics for big operators; such runs must not
    /// inflate the surrounding text line height.
    pub is_math: bool,
    /// The Office Math zone's spoken form, shared by the zone's runs: the
    /// /Alt of the Formula they are tagged as. None when it says nothing.
    pub formula: Option<std::sync::Arc<str>>,
    /// A legacy FORMCHECKBOX field, drawn as a square in place of `text`.
    pub checkbox: Option<FormCheckbox>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct FormCheckbox {
    /// `w:size` in points, or the run's size for `w:sizeAuto`.
    pub size: f32,
    pub checked: bool,
}

impl FormCheckbox {
    /// The check box's run text: a ballot box that only carries it through
    /// line layout, which sizes it; the renderer draws a square instead.
    pub const TEXT: &'static str = "\u{2610}";
}

impl Run {
    /// `w:kern`: pair kerning applies at this size and above, and a value of 0
    /// switches it off — Word writes 0 for an unticked "Kerning for fonts", so
    /// a Normal style's 0 overrides docDefaults' 1pt (slovak_pedagogical).
    pub fn kerns_at(&self, font_size: f32) -> bool {
        self.kern_threshold
            .is_some_and(|t| t > 0.0 && font_size >= t)
    }
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
            cs_font_name: None,
            bold: false,
            italic: false,
            underline: false,
            double_underline: false,
            strikethrough: false,
            dstrike: false,
            char_spacing: 0.0,
            text_scale: 100.0,
            caps: false,
            small_caps: false,
            vanish: false,
            color: None,
            highlight: None,
            shading: None,
            border: None,
            is_tab: false,
            ptab_alignment: None,
            is_line_break: false,
            vertical_align: VertAlign::Baseline,
            field_code: None,
            hyperlink_url: None,
            inline_image: None,
            footnote_id: None,
            is_footnote_ref_mark: false,
            endnote_id: None,
            is_endnote_ref_mark: false,
            kern_threshold: None,
            position: 0.0,
            char_style_id: None,
            text_outline: None,
            text_fill: None,
            text_shadow: None,
            text_glow: None,
            lang: None,
            text_lang: None,
            text_lang_east_asia: None,
            text_lang_bidi: None,
            font_size_from_default: false,
            font_name_from_default: false,
            bold_is_direct: false,
            italic_is_direct: false,
            comment_ids: Vec::new(),
            is_math: false,
            formula: None,
            checkbox: None,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum FieldCode {
    Page,
    NumPages,
    /// `number` is the `\n` switch: the referenced paragraph's list number.
    StyleRef {
        name: String,
        number: bool,
    },
    PageRef(String),
    /// An `IF` over nested fields, evaluated per page: legislation running
    /// heads read `IF {STYLEREF X \n} = 0 "{STYLEREF X}" "Part {STYLEREF X \n}"`,
    /// so the cached result is one page's value.
    If(Vec<IfPart>),
}

#[derive(Clone, Debug, PartialEq)]
pub enum IfPart {
    Text(String),
    Field(FieldCode),
}

pub enum Block {
    Paragraph(Paragraph),
    Table(Table),
}
