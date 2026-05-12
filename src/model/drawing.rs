use std::sync::Arc;

use crate::geometry::{FormulaOp, PathFill};

use super::{
    HRelativeFrom, HorizontalPosition, Paragraph, VRelativeFrom, VerticalPosition, WrapText,
    WrapType,
};

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum ImageFormat {
    Jpeg,
    Png,
    Bmp,
    /// Enhanced Metafile — a vector graphics format. Embedded as a PDF Form
    /// XObject (PDF drawing operators) rather than a raster Image XObject.
    Emf,
}

#[derive(Clone)]
pub struct EmbeddedImage {
    pub data: Arc<Vec<u8>>,
    pub format: ImageFormat,
    pub pixel_width: u32,
    pub pixel_height: u32,
    pub display_width: f32,
    pub display_height: f32,
    pub jpeg_components: u8,
    /// effectExtent + distT/distB combined (points)
    pub layout_extra_height: f32,
    /// effectExtent.t + distT (points)
    pub layout_extra_top: f32,
    pub stroke_color: Option<[u8; 3]>,
    pub stroke_width: f32,
    pub shadow: Option<ImageShadow>,
    pub soft_edge: Option<SoftEdge>,
    pub glow: Option<ImageGlow>,
    pub inner_shadow: Option<InnerShadow>,
    pub reflection: Option<ImageReflection>,
    /// Non-rectangular clip shape (e.g. "ellipse"). None means standard rect (no clipping).
    pub clip_geometry: Option<ShapeGeometry>,
}

#[derive(Clone, Debug)]
pub struct ImageShadow {
    pub offset_x: f32,    // points, positive = right
    pub offset_y: f32,    // points, positive = down (screen coords, not PDF)
    pub blur_radius: f32, // points
    pub color: [u8; 3],
    pub alpha: f32,       // 0.0–1.0
}

#[derive(Clone, Debug)]
pub struct SoftEdge {
    pub radius: f32, // points
}

#[derive(Clone, Debug)]
pub struct ImageGlow {
    pub radius: f32,    // points
    pub color: [u8; 3],
    pub alpha: f32,     // 0.0–1.0
}

#[derive(Clone, Debug)]
pub struct InnerShadow {
    pub offset_x: f32,
    pub offset_y: f32,
    pub blur_radius: f32,
    pub color: [u8; 3],
    pub alpha: f32,
}

#[derive(Clone, Debug)]
pub struct ImageReflection {
    pub start_alpha: f32, // 0.0–1.0
    pub end_alpha: f32,   // 0.0–1.0
    pub distance: f32,    // points (gap between image and reflection)
    #[allow(dead_code)]
    pub blur_radius: f32, // points
    /// Fraction of image height visible in the reflection (0.0–1.0). 1.0 = full image.
    pub end_pos: f32,
}

#[derive(Clone)]
pub struct FloatingImage {
    pub image: EmbeddedImage,
    pub h_position: HorizontalPosition,
    pub h_relative_from: HRelativeFrom,
    pub v_position: VerticalPosition,
    pub v_relative_from: VRelativeFrom,
    pub wrap_type: WrapType,
    pub wrap_text: WrapText,
    /// Vertices in 1/21600 of extent (OOXML wrapPolygon coordinate space)
    pub wrap_polygon: Option<Vec<(i32, i32)>>,
    pub behind_doc: bool,
    pub dist_top: f32,
    pub dist_bottom: f32,
    pub dist_left: f32,
    pub dist_right: f32,
}

/// Covers all 187 OOXML preset shapes and arbitrary custom geometry (a:custGeom).
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
    pub op: FormulaOp,
    pub x: String,
    pub y: String,
    pub z: String,
}

#[derive(Clone, Debug)]
pub struct CustomPathDef {
    pub commands: Vec<CustomPathCommand>,
    pub w: Option<i64>,
    pub h: Option<i64>,
    pub fill: PathFill,
    pub stroke: bool,
}

#[derive(Clone, Debug)]
pub enum CustomPathCommand {
    MoveTo {
        x: String,
        y: String,
    },
    LineTo {
        x: String,
        y: String,
    },
    ArcTo {
        wr: String,
        hr: String,
        st_ang: String,
        sw_ang: String,
    },
    CubicBezTo {
        x1: String,
        y1: String,
        x2: String,
        y2: String,
        x3: String,
        y3: String,
    },
    QuadBezTo {
        x1: String,
        y1: String,
        x2: String,
        y2: String,
    },
    Close,
}

#[derive(Clone, Copy, Default)]
pub enum SmartArtTextAnchor {
    #[default]
    Top,
    Center,
    Bottom,
}

#[derive(Clone, Copy, Default)]
pub enum SmartArtTextAlign {
    #[default]
    Left,
    Center,
    Right,
}

pub struct SmartArtRun {
    pub text: String,
    pub font_name: Option<String>,
    pub font_size: f32,
    pub bold: bool,
    pub italic: bool,
    pub color: Option<[u8; 3]>,
    pub underline: bool,
    pub strikethrough: bool,
    /// Baseline shift in 1000ths of %. >0 = superscript, <0 = subscript.
    pub baseline: i32,
    pub highlight: Option<[u8; 3]>,
}

pub struct SmartArtPara {
    pub runs: Vec<SmartArtRun>,
    pub bullet: Option<String>,
    pub align: SmartArtTextAlign,
    /// Line spacing as a multiplier (e.g. 0.9 for 90%). 0.0 = use default 1.2.
    pub line_spacing_pct: f32,
}

pub struct SmartArtShape {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
    /// Rotation in degrees (clockwise in OOXML, converted for PDF)
    pub rotation_deg: f32,
    pub shape_type: ShapeGeometry,
    pub fill: Option<[u8; 3]>,
    pub image_fill: Option<EmbeddedImage>,
    pub stroke_color: Option<[u8; 3]>,
    pub stroke_width: f32,
    pub paragraphs: Vec<SmartArtPara>,
    /// Separate text rectangle from dsp:txXfrm (x, y, w, h in pts)
    pub text_rect: Option<(f32, f32, f32, f32)>,
    /// Body insets from a:bodyPr (top, right, bottom, left) in pts
    pub text_insets: (f32, f32, f32, f32),
    pub text_anchor: SmartArtTextAnchor,
}

pub struct SmartArtDiagram {
    #[allow(dead_code)]
    pub display_width: f32,
    #[allow(dead_code)]
    pub display_height: f32,
    pub shapes: Vec<SmartArtShape>,
}

#[derive(Clone)]
pub enum ConnectorType {
    Line {
        flip_h: bool,
        flip_v: bool,
    },
    Arc {
        start_angle: f32,
        end_angle: f32,
        rotation: f32,
    },
}

#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum ArrowEnd {
    #[default]
    None,
    Arrow,
    Triangle,
    Stealth,
    Diamond,
    Oval,
}

impl ArrowEnd {
    pub fn from_attr(s: &str) -> Self {
        match s {
            "triangle" => Self::Triangle,
            "stealth" => Self::Stealth,
            "diamond" => Self::Diamond,
            "oval" => Self::Oval,
            "arrow" => Self::Arrow,
            _ => Self::None,
        }
    }
}

#[derive(Clone)]
pub struct ConnectorShape {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub stroke_color: [u8; 3],
    pub stroke_width: f32,
    pub connector_type: ConnectorType,
    pub head_end: ArrowEnd,
    pub tail_end: ArrowEnd,
}

#[derive(Clone, Copy, Default)]
pub enum TextAnchor {
    #[default]
    Top,
    Middle,
    Bottom,
}

pub enum ShapeFill {
    Solid([u8; 3]),
    LinearGradient {
        stops: Vec<([u8; 3], f32)>,
        angle_deg: f32,
    },
}

#[derive(Clone, Debug)]
pub struct TextWarp {
    pub preset: String,
    pub adjustments: Vec<(String, i64)>,
}

#[derive(Clone, Debug, Default)]
pub enum AutoFit {
    #[default]
    None,
    #[allow(dead_code)]
    Normal {
        font_scale: Option<f32>,
        line_space_reduction: Option<f32>,
    },
    Shape,
}

pub struct Textbox {
    pub paragraphs: Vec<Paragraph>,
    pub width_pt: f32,
    pub height_pt: f32,
    pub h_position: HorizontalPosition,
    pub h_relative_from: HRelativeFrom,
    /// Resolved offset for paths that need a single f32 (content-height
    /// reservation, paragraph-relative tb_y_top). For `AlignTop/Center/Bottom`
    /// this is 0 — the alignment is preserved in `v_position`, which the
    /// y_top resolver consults when it has the section dimensions.
    pub v_offset_pt: f32,
    /// Full vertical anchoring intent. `Offset(o)` matches `v_offset_pt = o`;
    /// the `AlignTop`/`AlignCenter`/`AlignBottom` variants need section
    /// geometry to compute their final y, which the renderer does.
    pub v_position: VerticalPosition,
    pub v_relative_from: VRelativeFrom,
    pub fill: Option<ShapeFill>,
    pub shape_type: ShapeGeometry,
    pub stroke_color: Option<[u8; 3]>,
    pub stroke_width: f32,
    pub text_anchor: TextAnchor,
    pub margin_left: f32,
    pub margin_right: f32,
    pub margin_top: f32,
    pub margin_bottom: f32,
    pub wrap_type: WrapType,
    #[allow(dead_code)]
    pub dist_top: f32,
    pub dist_bottom: f32,
    pub behind_doc: bool,
    pub no_text_wrap: bool,
    #[allow(dead_code)]
    pub is_wordart: bool,
    pub text_warp: Option<TextWarp>,
    pub auto_fit: AutoFit,
}
