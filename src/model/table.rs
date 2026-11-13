use super::{Block, HorizontalPosition, Paragraph};

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

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum TextDirection {
    #[default]
    LrTb,
    TbRl,
    BtLr,
}

#[derive(Clone, Copy, Debug, Default, PartialEq)]
pub enum BorderStyle {
    #[default]
    Single,
    Dotted,
    Dashed,
    DashSmallGap,
    DashDot,
    DashDotDot,
    Double,
    Wave,
    /// Thin line inside, thick line outside (`thinThick*Gap`).
    ThinThick,
    /// Thick line inside, thin line outside (`thickThin*Gap`).
    ThickThin,
    ThinThickThin,
}

/// How paragraph, run and page borders draw (`pdf::color::stroke_border`);
/// cell borders keep their own conventions in `pdf::table::draw_border`.
/// From Word's online export of `paragraph-border-styles` (sz 12, butt caps).
impl BorderStyle {
    /// The lines of the band as (offset, width) in line widths, offsets
    /// positive toward what the border encloses: a double border is two
    /// lines a line apart, thinThick*Gap a full line inside, a half gap and
    /// a half line outside (the online converter puts the thick line inside;
    /// desktop Word the thin one, as the spec reads). Dashed styles are one
    /// line; the wave is a path.
    pub fn layers(self) -> &'static [(f32, f32)] {
        match self {
            BorderStyle::Double => &[(1.0, 1.0), (-1.0, 1.0)],
            BorderStyle::ThinThick => &[(0.5, 1.0), (-0.75, 0.5)],
            BorderStyle::ThickThin => &[(0.75, 0.5), (-0.5, 1.0)],
            BorderStyle::ThinThickThin => &[(1.25, 0.5), (0.0, 1.0), (-1.25, 0.5)],
            _ => &[(0.0, 1.0)],
        }
    }

    /// The dash array in line widths; empty for a solid line. ponytail:
    /// dashSmallGap and dotDotDash unmeasured.
    pub fn dashes(self) -> &'static [f32] {
        match self {
            BorderStyle::Dotted => &[1.0, 1.0],
            BorderStyle::Dashed => &[3.0, 1.0],
            BorderStyle::DashSmallGap => &[2.0, 1.0],
            BorderStyle::DashDot => &[3.0, 1.0, 1.0, 1.0],
            BorderStyle::DashDotDot => &[3.0, 1.0, 1.0, 1.0, 1.0, 1.0],
            _ => &[],
        }
    }
}

#[derive(Clone, Copy, Debug)]
pub struct CellBorder {
    pub present: bool,
    pub color: Option<[u8; 3]>,
    pub width: f32,
    pub style: BorderStyle,
    pub is_override: bool,
}

impl Default for CellBorder {
    fn default() -> Self {
        Self {
            present: false,
            color: None,
            width: 0.5,
            style: BorderStyle::Single,
            is_override: false,
        }
    }
}

impl CellBorder {
    pub fn visible(color: Option<[u8; 3]>, width: f32, style: BorderStyle) -> Self {
        Self {
            present: true,
            color,
            width,
            style,
            is_override: false,
        }
    }

    /// Thickness of the band the border paints: two lines and the gap between
    /// them for a double border (see `draw_border`), one line otherwise.
    pub fn band(&self) -> f32 {
        match (self.present, self.style) {
            (false, _) => 0.0,
            (true, BorderStyle::Double) => 3.0 * self.width.max(0.25),
            (true, _) => self.width,
        }
    }
}

#[derive(Clone, Copy, Debug, Default)]
pub struct CellBorders {
    pub top: CellBorder,
    pub bottom: CellBorder,
    pub left: CellBorder,
    pub right: CellBorder,
    /// The cell's own top before the edge it shares with the row above was
    /// resolved into `top` (None when unchanged): a row continued on a new
    /// page draws this one at the page top.
    pub own_top: Option<CellBorder>,
    /// A vertical merge that runs to the table's last row keeps its first
    /// cell's own bottom here while `bottom` takes the table's to draw (None
    /// otherwise): the row it starts in is sized by its own.
    pub own_bottom: Option<CellBorder>,
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
    pub allow_overlap: bool,
    pub h_position: HorizontalPosition,
    pub h_anchor: &'static str, // "page", "margin", or "column"
    pub v_offset_pt: f32,
    pub v_anchor: &'static str, // "page", "margin", or "text"
    pub top_from_text: f32,     // min distance above table (pts)
    pub bottom_from_text: f32,  // min distance below table (pts)
    pub left_from_text: f32,    // min distance left of table (pts)
    pub right_from_text: f32,   // min distance right of table (pts)
}

#[derive(Clone, Copy, Debug, PartialEq, Default)]
pub enum TableAlignment {
    #[default]
    Left,
    Center,
    Right,
}

pub struct Table {
    pub col_widths: Vec<f32>,
    pub rows: Vec<TableRow>,
    pub table_indent: f32,
    pub table_indent_explicit: bool,
    /// The table has a table style (its own `w:tblStyle` or the document's
    /// default table style).
    pub has_table_style: bool,
    pub cell_margins: CellMargins,
    pub position: Option<TablePosition>,
    pub alignment: TableAlignment,
    pub fixed_layout: bool,
    pub auto_width: bool,
    /// `w:tblW w:type="pct"`: table width as a fraction of the available
    /// content width (5000ths in the XML; 1.0 = full width).
    pub width_pct: Option<f32>,
    /// True when the required `w:tblGrid` was missing and col_widths were
    /// synthesized from row cell widths.
    pub grid_inferred: bool,
    /// No cell names a preferred width (`tcW` absent, auto or 0).
    pub cells_unsized: bool,
    /// Every cell's `tcW` is in twips or auto (none absent or in percent).
    pub cells_twip_sized: bool,
    /// `w:bidiVisual`: the columns run right to left (already mirrored in
    /// `rows` and `col_widths`), and the table starts at the right margin.
    pub bidi_visual: bool,
    /// `tblLook` firstRow / firstColumn for tagging: Word's tagged export
    /// treats them as on when `tblLook` is absent (header row → THead/TH,
    /// first column → TH).
    pub header_first_row: bool,
    pub header_first_col: bool,
}

impl Table {
    pub fn first_cell(&self) -> Option<&TableCell> {
        self.rows.first()?.cells.first()
    }

    pub fn last_cell_of_first_row(&self) -> Option<&TableCell> {
        self.rows.first()?.cells.last()
    }

    /// Every paragraph in the table, nested tables included.
    pub fn all_paragraphs(&self) -> impl Iterator<Item = &Paragraph> {
        self.rows
            .iter()
            .flat_map(|r| &r.cells)
            .flat_map(|c| c.all_paragraphs())
    }

    /// How far an AutoFit table may reach past the text area before Word
    /// squeezes it: before compat 15 by its outdent on the left and its last
    /// cell's right margin on the right (parish_housing's 9854 tw tables on a
    /// 9639 tw page stay whole; Word probes on case100 squeeze its 0-margin
    /// table to the text width, a 108 tw margin lets it 5.4pt past). Word
    /// 2013+ keeps the outer border bands inside the text area instead
    /// (cote_dazur: 0.5pt borders, 9062 tw of table on a 9072 tw page).
    pub fn autofit_overhang(&self, compat_mode: u32) -> f32 {
        if compat_mode >= 15 {
            let left = self.first_cell().map_or(0.0, |c| c.borders.left.band());
            let right = self
                .last_cell_of_first_row()
                .map_or(0.0, |c| c.borders.right.band());
            return -(left + right) / 2.0;
        }
        let outdent = if self.has_table_style || self.table_indent_explicit {
            self.first_cell_left_margin()
        } else {
            0.0
        };
        outdent + self.last_cell_right_margin()
    }

    /// The left margin of the first row's first cell, its own or the table's:
    /// before compat 15 it puts that cell's text at the indent (Word's import
    /// of croatian_regulations' HTML table: 4.5pt cell margins over a 0.75pt
    /// table default, text at the margin).
    pub fn first_cell_left_margin(&self) -> f32 {
        self.first_cell()
            .and_then(|c| c.cell_margins)
            .map_or(self.cell_margins.left, |m| m.left)
    }

    /// The right margin of the first row's last cell, its own or the table's.
    pub fn last_cell_right_margin(&self) -> f32 {
        self.last_cell_of_first_row()
            .and_then(|c| c.cell_margins)
            .map_or(self.cell_margins.right, |m| m.right)
    }
}

pub struct TableRow {
    pub cells: Vec<TableCell>,
    /// Grid columns left empty before the first cell (`w:gridBefore`).
    pub grid_before: usize,
    pub height: Option<f32>,
    pub height_exact: bool,
    pub is_header: bool,
    pub cant_split: bool,
}

impl TableRow {
    /// Each cell with the grid column it starts at and the columns it spans.
    pub fn grid_cells(&self) -> impl Iterator<Item = (usize, usize, &TableCell)> {
        self.cells.iter().scan(self.grid_before, |col, cell| {
            let span = cell.grid_span.max(1) as usize;
            let start = *col;
            *col += span;
            Some((start, span, cell))
        })
    }
}

/// A `w:shd` line/cross pattern fill (e.g. `thinDiagStripe`). Rendered as real
/// hatching rather than collapsed into a solid tint, so diagonal-stripe cells
/// (e.g. ground/earth symbols) look "dithered" like Word's output.
#[derive(Clone, Copy, Debug)]
pub enum HatchKind {
    Horz,
    Vert,
    DiagFwd,
    DiagBack,
    CrossHorzVert,
    CrossDiag,
}

#[derive(Clone, Copy, Debug)]
pub struct HatchPattern {
    pub kind: HatchKind,
    pub fg: [u8; 3],
    pub bg: [u8; 3],
}

pub struct TableCell {
    pub width: f32,
    /// `w:tcW w:type="pct"`: the preferred width as a fraction of the table
    /// width (`width` then holds the raw 5000ths, read as twips).
    pub width_pct: Option<f32>,
    pub content: Vec<Block>,
    pub borders: CellBorders,
    pub shading: Option<[u8; 3]>,
    /// Line/cross pattern fill. When present, the renderer draws hatching
    /// instead of `shading`'s solid approximation.
    pub hatch: Option<HatchPattern>,
    pub grid_span: u16,
    pub v_merge: VMerge,
    pub v_align: CellVAlign,
    pub text_direction: TextDirection,
    pub cell_margins: Option<CellMargins>,
    pub hide_mark: bool,
}

impl TableCell {
    /// Includes paragraphs inside nested tables.
    pub fn all_paragraphs(&self) -> Vec<&Paragraph> {
        let mut result = Vec::new();
        for block in &self.content {
            match block {
                Block::Paragraph(p) => result.push(p),
                Block::Table(t) => result.extend(t.all_paragraphs()),
            }
        }
        result
    }
}

impl Table {
    /// Every paragraph in the table's cells, nested tables included.
    pub fn all_paragraphs(&self) -> impl Iterator<Item = &Paragraph> {
        self.rows
            .iter()
            .flat_map(|r| r.cells.iter())
            .flat_map(|c| c.all_paragraphs())
    }
}
