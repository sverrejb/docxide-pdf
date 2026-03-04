use std::collections::HashMap;
use std::io::Read;

use crate::model::{
    Chart, ChartAxis, ChartLegend, ChartSeries, ChartType, InlineChart, LegendPosition,
    MarkerSymbol,
};

use super::{DML_NS, read_zip_text};

const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";

fn chart_child<'a>(parent: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    parent
        .children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(CHART_NS))
}

fn chart_attr<'a>(parent: roxmltree::Node<'a, 'a>, child: &str) -> Option<&'a str> {
    chart_child(parent, child).and_then(|n| n.attribute("val"))
}

fn parse_hex_color_dml(val: &str) -> Option<[u8; 3]> {
    if val.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&val[0..2], 16).ok()?;
    let g = u8::from_str_radix(&val[2..4], 16).ok()?;
    let b = u8::from_str_radix(&val[4..6], 16).ok()?;
    Some([r, g, b])
}

fn dml_child<'a>(parent: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    parent
        .children()
        .find(|n| n.tag_name().name() == name && n.tag_name().namespace() == Some(DML_NS))
}

fn extract_srgb_fill(sp_pr: roxmltree::Node) -> Option<[u8; 3]> {
    let srgb = dml_child(dml_child(sp_pr, "solidFill")?, "srgbClr")?;
    srgb.attribute("val").and_then(parse_hex_color_dml)
}

fn extract_fill_alpha(sp_pr: roxmltree::Node) -> Option<f32> {
    let srgb = dml_child(dml_child(sp_pr, "solidFill")?, "srgbClr")?;
    let alpha_node = dml_child(srgb, "alpha")?;
    let val: f32 = alpha_node.attribute("val")?.parse().ok()?;
    Some(val / 100_000.0)
}

fn extract_line_color(sp_pr: roxmltree::Node) -> Option<[u8; 3]> {
    let ln = dml_child(sp_pr, "ln")?;
    // noFill means no line
    if dml_child(ln, "noFill").is_some() {
        return None;
    }
    extract_srgb_fill(ln)
}

fn parse_num_cache(parent: roxmltree::Node, child_name: &str) -> Vec<f32> {
    let num_cache = chart_child(parent, child_name)
        .and_then(|c| chart_child(c, "numRef"))
        .and_then(|nr| chart_child(nr, "numCache"));
    let Some(num_cache) = num_cache else {
        return Vec::new();
    };
    let mut pts: Vec<(usize, f32)> = num_cache
        .children()
        .filter(|n| n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(CHART_NS))
        .filter_map(|pt| {
            let idx = pt.attribute("idx")?.parse::<usize>().ok()?;
            let v = chart_child(pt, "v")?.text()?.parse::<f32>().ok()?;
            Some((idx, v))
        })
        .collect();
    pts.sort_by_key(|(idx, _)| *idx);
    let count = pts.last().map(|(idx, _)| idx + 1).unwrap_or(0);
    let mut values = vec![0.0; count];
    for (idx, v) in pts {
        values[idx] = v;
    }
    values
}

fn parse_series(ser_node: roxmltree::Node) -> (ChartSeries, Vec<String>) {
    let label = chart_child(ser_node, "tx")
        .and_then(|tx| chart_child(tx, "strRef"))
        .and_then(|sr| chart_child(sr, "strCache"))
        .and_then(|cache| {
            cache
                .descendants()
                .find(|n| n.tag_name().name() == "v" && n.tag_name().namespace() == Some(CHART_NS))
                .and_then(|v| v.text())
        })
        .unwrap_or("")
        .to_string();

    let marker_node = chart_child(ser_node, "marker");

    let sp_pr = chart_child(ser_node, "spPr");
    let color = sp_pr
        .and_then(extract_srgb_fill)
        .or_else(|| sp_pr.and_then(extract_line_color))
        .or_else(|| {
            marker_node
                .and_then(|m| chart_child(m, "spPr"))
                .and_then(extract_srgb_fill)
        });
    let fill_alpha = sp_pr.and_then(extract_fill_alpha);

    let marker = marker_node
        .and_then(|m| chart_attr(m, "symbol"))
        .and_then(|val| match val {
            "circle" => Some(MarkerSymbol::Circle),
            "square" => Some(MarkerSymbol::Square),
            "diamond" => Some(MarkerSymbol::Diamond),
            "triangle" => Some(MarkerSymbol::Triangle),
            "plus" => Some(MarkerSymbol::Plus),
            "star" => Some(MarkerSymbol::Star),
            "x" => Some(MarkerSymbol::X),
            "dash" => Some(MarkerSymbol::Dash),
            "dot" => Some(MarkerSymbol::Dot),
            "none" => Some(MarkerSymbol::None),
            _ => Option::None,
        });

    // Scatter/bubble use yVal instead of val
    let y_vals = parse_num_cache(ser_node, "yVal");
    let values = if !y_vals.is_empty() {
        y_vals
    } else {
        parse_num_cache(ser_node, "val")
    };

    let x_values = non_empty_vec(parse_num_cache(ser_node, "xVal"));
    let bubble_sizes = non_empty_vec(parse_num_cache(ser_node, "bubbleSize"));

    let cat_labels = parse_str_cache(chart_child(ser_node, "cat"));

    (
        ChartSeries {
            label,
            color,
            fill_alpha,
            values,
            x_values,
            bubble_sizes,
            marker,
        },
        cat_labels,
    )
}

fn non_empty_vec(v: Vec<f32>) -> Option<Vec<f32>> {
    if v.is_empty() { None } else { Some(v) }
}

fn parse_str_cache(container: Option<roxmltree::Node>) -> Vec<String> {
    let str_cache = container
        .and_then(|c| chart_child(c, "strRef"))
        .and_then(|sr| chart_child(sr, "strCache"));
    let Some(str_cache) = str_cache else {
        return Vec::new();
    };
    let mut pts: Vec<(usize, String)> = str_cache
        .children()
        .filter(|n| n.tag_name().name() == "pt" && n.tag_name().namespace() == Some(CHART_NS))
        .filter_map(|pt| {
            let idx = pt.attribute("idx")?.parse::<usize>().ok()?;
            let v = chart_child(pt, "v")?.text()?.to_string();
            Some((idx, v))
        })
        .collect();
    pts.sort_by_key(|(idx, _)| *idx);
    let count = pts.last().map(|(idx, _)| idx + 1).unwrap_or(0);
    let mut labels = vec![String::new(); count];
    for (idx, v) in pts {
        labels[idx] = v;
    }
    labels
}

fn parse_axis(ax_node: roxmltree::Node) -> ChartAxis {
    let delete = chart_attr(ax_node, "delete") == Some("1");

    let gridline_color = chart_child(ax_node, "majorGridlines")
        .and_then(|gl| chart_child(gl, "spPr"))
        .and_then(extract_line_color);

    let line_color = chart_child(ax_node, "spPr").and_then(extract_line_color);

    ChartAxis {
        labels: Vec::new(),
        delete,
        gridline_color,
        line_color,
    }
}

fn parse_legend(legend_node: roxmltree::Node) -> ChartLegend {
    let position = match chart_attr(legend_node, "legendPos") {
        Some("b") => LegendPosition::Bottom,
        Some("t") => LegendPosition::Top,
        Some("l") => LegendPosition::Left,
        _ => LegendPosition::Right,
    };
    ChartLegend { position }
}

fn collect_series(type_node: roxmltree::Node) -> (Vec<ChartSeries>, Vec<String>) {
    let mut series_list = Vec::new();
    let mut cat_labels = Vec::new();
    for ser_node in type_node
        .children()
        .filter(|n| n.tag_name().name() == "ser" && n.tag_name().namespace() == Some(CHART_NS))
    {
        let (series, labels) = parse_series(ser_node);
        if cat_labels.is_empty() && !labels.is_empty() {
            cat_labels = labels;
        }
        series_list.push(series);
    }
    (series_list, cat_labels)
}

fn assign_cat_labels(cat_axis: &mut Option<ChartAxis>, cat_labels: Vec<String>) {
    if let Some(ref mut ax) = *cat_axis {
        if ax.labels.is_empty() {
            ax.labels = cat_labels;
        }
    } else if !cat_labels.is_empty() {
        *cat_axis = Some(ChartAxis {
            labels: cat_labels,
            delete: true,
            gridline_color: None,
            line_color: None,
        });
    }
}

fn extract_plot_border(plot_area: roxmltree::Node) -> Option<[u8; 3]> {
    plot_area
        .children()
        .find(|n| n.tag_name().name() == "spPr" && n.tag_name().namespace() == Some(CHART_NS))
        .and_then(extract_line_color)
}

fn parse_chart_space(xml_content: &str, accent_colors: Vec<[u8; 3]>) -> Option<Chart> {
    let doc = roxmltree::Document::parse(xml_content).ok()?;
    let chart_space = doc.root_element();
    let chart_node = chart_child(chart_space, "chart")?;
    let plot_area = chart_child(chart_node, "plotArea")?;

    let (type_node, chart_type, gap_width_pct) = detect_chart_type(plot_area)?;

    let (series_list, cat_labels) = collect_series(type_node);
    let is_scatter_like = matches!(chart_type, ChartType::Scatter | ChartType::Bubble);

    // Scatter/bubble have two valAx; use first as cat_axis, second as val_axis
    let (mut cat_axis, val_axis) = if is_scatter_like {
        let val_axes: Vec<_> = plot_area
            .children()
            .filter(|n| {
                n.tag_name().name() == "valAx" && n.tag_name().namespace() == Some(CHART_NS)
            })
            .map(parse_axis)
            .collect();
        (val_axes.first().cloned(), val_axes.get(1).cloned())
    } else {
        (
            chart_child(plot_area, "catAx").map(parse_axis),
            chart_child(plot_area, "valAx").map(parse_axis),
        )
    };

    assign_cat_labels(&mut cat_axis, cat_labels);

    Some(Chart {
        chart_type,
        series: series_list,
        cat_axis,
        val_axis,
        legend: chart_child(chart_node, "legend").map(parse_legend),
        gap_width_pct,
        plot_border_color: extract_plot_border(plot_area),
        accent_colors,
    })
}

fn detect_chart_type<'a>(
    plot_area: roxmltree::Node<'a, 'a>,
) -> Option<(roxmltree::Node<'a, 'a>, ChartType, f32)> {
    if let Some(n) = chart_child(plot_area, "barChart") {
        let horizontal = chart_attr(n, "barDir") == Some("bar");
        let stacked = matches!(
            chart_attr(n, "grouping"),
            Some("stacked") | Some("percentStacked")
        );
        let gap_width_pct = chart_attr(n, "gapWidth")
            .and_then(|v| v.parse::<f32>().ok())
            .unwrap_or(150.0);
        return Some((
            n,
            ChartType::Bar {
                horizontal,
                stacked,
            },
            gap_width_pct,
        ));
    }

    let default_gap = 150.0;
    if let Some(n) = chart_child(plot_area, "lineChart") {
        return Some((n, ChartType::Line, default_gap));
    }
    if let Some(n) =
        chart_child(plot_area, "pieChart").or_else(|| chart_child(plot_area, "pie3DChart"))
    {
        return Some((n, ChartType::Pie, default_gap));
    }
    if let Some(n) = chart_child(plot_area, "areaChart") {
        return Some((n, ChartType::Area, default_gap));
    }
    if let Some(n) = chart_child(plot_area, "doughnutChart") {
        let hole_size_pct = chart_attr(n, "holeSize")
            .and_then(|v| v.parse::<f32>().ok())
            .unwrap_or(50.0);
        return Some((n, ChartType::Doughnut { hole_size_pct }, default_gap));
    }
    if let Some(n) = chart_child(plot_area, "radarChart") {
        return Some((n, ChartType::Radar, default_gap));
    }
    if let Some(n) = chart_child(plot_area, "scatterChart") {
        return Some((n, ChartType::Scatter, default_gap));
    }
    if let Some(n) = chart_child(plot_area, "bubbleChart") {
        return Some((n, ChartType::Bubble, default_gap));
    }
    None
}

pub(super) fn parse_chart_from_zip<R: Read + std::io::Seek>(
    r_id: &str,
    rels: &HashMap<String, String>,
    zip: &mut zip::ZipArchive<R>,
    display_w: f32,
    display_h: f32,
    accent_colors: Vec<[u8; 3]>,
) -> Option<InlineChart> {
    let target = rels.get(r_id)?;
    let zip_path = target
        .strip_prefix('/')
        .map(String::from)
        .unwrap_or_else(|| format!("word/{}", target));
    let xml_content = read_zip_text(zip, &zip_path)?;
    let chart = parse_chart_space(&xml_content, accent_colors)?;
    Some(InlineChart {
        chart,
        display_width: display_w,
        display_height: display_h,
    })
}
