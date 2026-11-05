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
    /// `c:spPr/a:ln@w` in points.
    pub line_width: Option<f32>,
    /// `c:marker/c:size`, its fill and outline width.
    pub marker_size: Option<f32>,
    pub marker_color: Option<[u8; 3]>,
    pub marker_line_width: Option<f32>,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum BarGrouping {
    Clustered,
    /// A 3-D chart's series one behind another, each in its own depth row.
    Standard,
    Stacked,
    PercentStacked,
}

#[derive(Clone)]
pub enum ChartType {
    Bar {
        horizontal: bool,
        grouping: BarGrouping,
    },
    Line,
    Pie,
    Area,
    Scatter,
    Bubble,
    Doughnut {
        hole_size_pct: f32,
    },
    Radar,
}

#[derive(Clone)]
pub struct ChartAxis {
    pub labels: Vec<String>,
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
    pub val_format_code: Option<String>,
    /// `c:title`'s text, when it has one of its own.
    pub title: Option<String>,
    /// `c:dLbls/c:showVal`: each bar carries its value.
    pub show_values: bool,
    /// `c:dTable`: a table of the values under the plot, in place of the
    /// category axis labels.
    pub data_table: bool,
    /// The chart area's outline: Word draws a grey 0.75pt one when
    /// `c:chartSpace` sets no `c:spPr` of its own.
    pub frame: Option<[u8; 3]>,
    /// `c:view3D` of a `c:bar3DChart`.
    pub view_3d: Option<View3D>,
}

/// A 3-D bar chart's view: rotations in degrees, depth and the gap between
/// depth rows as fractions of a bar's width (`depthPercent`, `gapDepth`).
#[derive(Clone, Copy)]
pub struct View3D {
    pub rot_x: f32,
    pub rot_y: f32,
    pub depth: f32,
    pub gap_depth: f32,
    /// `hPercent`: a 3-D pie's thickness against its default.
    pub height: f32,
    pub shape: BarShape,
    /// `c:serAx` shown: series names along the floor's right edge.
    pub series_axis: bool,
}

/// `c:shape`: a 3-D bar's solid. The `ToMax` kinds are cut from a cone or
/// pyramid whose tip sits at the value axis maximum.
#[derive(Clone, Copy, PartialEq)]
pub enum BarShape {
    Box,
    Cylinder,
    Cone,
    ConeToMax,
    Pyramid,
    PyramidToMax,
}

pub struct InlineChart {
    pub chart: Chart,
    pub display_width: f32,
    pub display_height: f32,
}
