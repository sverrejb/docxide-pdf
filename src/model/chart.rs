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

#[derive(Default)]
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
    /// Office 2016 charts. A waterfall's `totals` are the points drawn
    /// from zero as running totals.
    Waterfall {
        totals: Vec<usize>,
    },
    Funnel,
    Treemap,
    Sunburst,
    /// Quartiles by Excel's exclusive method (rank p·(n+1)), else inclusive
    /// (rank p·(n−1)+1).
    BoxWhisker {
        exclusive: bool,
    },
    /// Bars sorted largest first under the line of their cumulative share.
    Pareto,
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
    /// A treemap's or sunburst's outer category of each point.
    pub parent_labels: Vec<String>,
    /// An Office 2016 chart's text, which Word takes from the chartStyle part.
    pub ex_text: Option<ChartExText>,
}

/// Size and colour of one kind of chart text.
#[derive(Clone, Copy)]
pub struct ChartText {
    pub size: f32,
    pub color: [u8; 3],
}

#[derive(Clone, Copy)]
pub struct ChartExText {
    pub title: ChartText,
    pub axis: ChartText,
    pub data_label: ChartText,
    pub legend: ChartText,
}

/// An Office 2016 chart's default grey (tx1 at 15%) for its frame, axis
/// lines and gridlines.
pub const CHARTEX_GRID: [u8; 3] = [0xD9, 0xD9, 0xD9];

/// A waterfall's legend: its three kinds of bar.
pub const WATERFALL_LEGEND: [&str; 3] = ["Increase", "Decrease", "Total"];

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
    /// `perspective` when `rAngAx` is 0: the scene seen through a camera
    /// whose vertical field of view is half this many degrees. `None` with
    /// right-angle axes, where depth draws obliquely.
    pub perspective: Option<f32>,
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
