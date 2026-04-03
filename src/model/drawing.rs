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
}

#[derive(Clone, Debug)]
pub struct ImageShadow {
    pub offset_x: f32,    // points, positive = right
    pub offset_y: f32,    // points, positive = down (screen coords, not PDF)
    pub blur_radius: f32, // points
    pub color: [u8; 3],
    pub alpha: f32,       // 0.0–1.0
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

pub struct ConnectorShape {
    pub x: f32,
    pub y: f32,
    pub width: f32,
    pub height: f32,
    pub stroke_color: [u8; 3],
    pub stroke_width: f32,
    pub connector_type: ConnectorType,
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
    pub v_offset_pt: f32,
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
    pub is_wordart: bool,
    pub text_warp: Option<TextWarp>,
    pub auto_fit: AutoFit,
}
