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
    pub col_widths: Vec<f32>, // points
    pub rows: Vec<TableRow>,
    pub table_indent: f32,
    pub cell_margins: CellMargins,
    pub position: Option<TablePosition>,
    pub alignment: TableAlignment,
    pub fixed_layout: bool,
}

pub struct TableRow {
    pub cells: Vec<TableCell>,
    pub height: Option<f32>,
    pub height_exact: bool,
    pub is_header: bool,
}

pub struct TableCell {
    pub width: f32, // points
    pub content: Vec<Block>,
    pub borders: CellBorders,
    pub shading: Option<[u8; 3]>,
    pub grid_span: u16,
    pub v_merge: VMerge,
    pub v_align: CellVAlign,
    pub text_direction: TextDirection,
    pub cell_margins: Option<CellMargins>,
}

impl TableCell {
    /// Recursively collect all paragraphs from this cell's content,
    /// including paragraphs inside nested tables.
    pub fn all_paragraphs(&self) -> Vec<&Paragraph> {
        fn collect<'a>(blocks: &'a [Block], out: &mut Vec<&'a Paragraph>) {
            for block in blocks {
                match block {
                    Block::Paragraph(p) => out.push(p),
                    Block::Table(t) => {
                        for row in &t.rows {
                            for cell in &row.cells {
                                collect(&cell.content, out);
                            }
                        }
                    }
                }
            }
        }
        let mut result = Vec::new();
        collect(&self.content, &mut result);
        result
    }
}
