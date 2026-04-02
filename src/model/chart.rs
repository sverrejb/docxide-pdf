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

#[derive(Clone, Copy, Debug, PartialEq)]
pub enum BarGrouping {
    Clustered,
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
    pub val_format_code: Option<String>,
}

pub struct InlineChart {
    pub chart: Chart,
    pub display_width: f32,
    pub display_height: f32,
}
