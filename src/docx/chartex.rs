//! Office 2016 charts (`cx:chartSpace`, the `chartEx` part): waterfall,
//! funnel, histogram, treemap, sunburst and box & whisker, read into the
//! chart model. Histograms become column charts of their bins.

use crate::model::{
    BarGrouping, CHARTEX_GRID, Chart, ChartAxis, ChartExText, ChartLegend, ChartSeries, ChartText,
    ChartType, LegendPosition,
};

use super::color::resolve_dml_color;
use super::styles::ThemeFonts;
use super::{find_child, find_children};

pub(super) const CHARTEX_NS: &str = "http://schemas.microsoft.com/office/drawing/2014/chartex";

fn cx<'a>(node: roxmltree::Node<'a, 'a>, name: &str) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, CHARTEX_NS)
}

fn cx_all<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> impl Iterator<Item = roxmltree::Node<'a, 'a>> {
    find_children(node, name, CHARTEX_NS)
}

/// A `cx:lvl`'s points in index order.
fn level_points(lvl: roxmltree::Node) -> Vec<String> {
    let n: usize = lvl
        .attribute("ptCount")
        .and_then(|v| v.parse().ok())
        .unwrap_or(0);
    let mut out = vec![String::new(); n];
    for pt in cx_all(lvl, "pt") {
        let idx: usize = pt
            .attribute("idx")
            .and_then(|v| v.parse().ok())
            .unwrap_or(0);
        if idx >= out.len() {
            out.resize(idx + 1, String::new());
        }
        out[idx] = pt.text().unwrap_or("").to_string();
    }
    out
}

struct Data {
    /// Category levels, the leaves first.
    levels: Vec<Vec<String>>,
    values: Vec<f32>,
}

fn parse_data(data: roxmltree::Node) -> Data {
    let levels = cx(data, "strDim")
        .map(|d| cx_all(d, "lvl").map(level_points).collect())
        .unwrap_or_default();
    let values = cx(data, "numDim")
        .and_then(|d| cx(d, "lvl"))
        .map(|l| {
            level_points(l)
                .iter()
                .map(|v| v.trim().parse().unwrap_or(0.0))
                .collect()
        })
        .unwrap_or_default();
    Data { levels, values }
}

fn text_of(node: roxmltree::Node) -> String {
    node.descendants()
        .filter(|n| n.tag_name().name() == "v" || n.tag_name().name() == "t")
        .filter_map(|n| n.text())
        .collect()
}

/// Excel's histogram bins over `values`: `binCount` equal bins, `binSize`
/// wide ones, or by default Scott's rule (3.5 σ n^-1/3). Bins close on the
/// right (`intervalClosed="r"`) except the first, which takes the minimum.
fn histogram(values: &[f32], binning: roxmltree::Node) -> (Vec<String>, Vec<f32>) {
    let (Some(min), Some(max)) = (
        values.iter().copied().reduce(f32::min),
        values.iter().copied().reduce(f32::max),
    ) else {
        return (Vec::new(), Vec::new());
    };
    let n = values.len() as f32;
    let count_attr = cx(binning, "binCount")
        .and_then(|c| c.attribute("val"))
        .and_then(|v| v.parse::<usize>().ok());
    let size_attr = cx(binning, "binSize")
        .and_then(|c| c.attribute("val"))
        .and_then(|v| v.parse::<f32>().ok());
    let width = match (count_attr, size_attr) {
        (Some(k), _) if k > 0 => (max - min) / k as f32,
        (_, Some(w)) if w > 0.0 => w,
        _ => {
            let mean = values.iter().sum::<f32>() / n;
            let sd = (values.iter().map(|v| (v - mean).powi(2)).sum::<f32>() / (n - 1.0).max(1.0))
                .sqrt();
            3.5 * sd / n.cbrt()
        }
    };
    let bins = if width > 0.0 {
        (((max - min) / width).ceil() as usize).max(1)
    } else {
        1
    };
    let fmt = |v: f32| {
        let r = (v * 10.0).round() / 10.0;
        if r.fract() == 0.0 {
            format!("{}", r as i64)
        } else {
            format!("{r}")
        }
    };
    let mut labels = Vec::with_capacity(bins);
    let mut counts = vec![0.0f32; bins];
    for b in 0..bins {
        let (lo, hi) = (min + b as f32 * width, min + (b + 1) as f32 * width);
        let open = if b == 0 { '[' } else { '(' };
        labels.push(format!("{open}{}, {}]", fmt(lo), fmt(hi)));
    }
    for &v in values {
        let b = if width > 0.0 {
            (((v - min) / width).ceil() as usize)
                .saturating_sub(1)
                .min(bins - 1)
        } else {
            0
        };
        counts[b] += 1.0;
    }
    (labels, counts)
}

const CHART_STYLE_NS: &str = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";

/// The chart's text as its chartStyle part sets it: each element's
/// `fontRef` colour and `defRPr` size, rounded to Word's half points (the
/// probes' 11.97pt labels print at 12, 18.62pt titles at 18.5). Without the
/// part, Word's own chartEx style: 9pt labels and a 14pt title in 65% grey,
/// data labels in 75%.
fn style_text(style: Option<&str>, theme: &ThemeFonts) -> ChartExText {
    let doc = style.and_then(|s| roxmltree::Document::parse(s).ok());
    let text = |name: &str, size: f32, grey: u8| {
        let el = doc.as_ref().and_then(|d| {
            d.root_element()
                .children()
                .find(|n| n.has_tag_name((CHART_STYLE_NS, name)))
        });
        let child = |n: &str| el.and_then(|e| find_child(e, n, CHART_STYLE_NS));
        ChartText {
            size: child("defRPr")
                .and_then(|r| r.attribute("sz")?.parse::<f32>().ok())
                .map_or(size, |sz| (sz / 50.0).round() / 2.0),
            color: child("fontRef")
                .and_then(|f| resolve_dml_color(f, theme))
                .unwrap_or([grey; 3]),
        }
    };
    ChartExText {
        title: text("title", 14.0, 0x59),
        axis: text("valueAxis", 9.0, 0x59),
        data_label: text("dataLabel", 9.0, 0x40),
        legend: text("legend", 9.0, 0x59),
    }
}

pub(super) fn parse_chartex_space(
    xml: &str,
    accent_colors: Vec<[u8; 3]>,
    theme: &ThemeFonts,
    style: Option<&str>,
) -> Option<Chart> {
    let doc = roxmltree::Document::parse(xml).ok()?;
    let space = doc.root_element();
    let data_by_id: Vec<(String, Data)> = cx(space, "chartData")
        .map(|cd| {
            cx_all(cd, "data")
                .map(|d| (d.attribute("id").unwrap_or("0").to_string(), parse_data(d)))
                .collect()
        })
        .unwrap_or_default();
    let chart = cx(space, "chart")?;
    let plot = cx(chart, "plotArea")?;
    let region = cx(plot, "plotAreaRegion")?;
    let series = cx_all(region, "series").find(|s| s.attribute("hidden") != Some("1"))?;
    let pareto = cx_all(region, "series").any(|s| s.attribute("layoutId") == Some("paretoLine"));
    let layout = series.attribute("layoutId")?;
    let data_id = cx(series, "dataId")
        .and_then(|d| d.attribute("val"))
        .unwrap_or("0");
    let data = &data_by_id.iter().find(|(id, _)| id == data_id)?.1;
    let layout_pr = cx(series, "layoutPr");
    let label = cx(series, "tx").map(text_of).unwrap_or_default();
    let show_values = cx(series, "dataLabels")
        .and_then(|d| cx(d, "visibility"))
        .is_some_and(|v| v.attribute("value") == Some("1"));
    let gap_width_pct = cx_all(plot, "axis")
        .find_map(|a| cx(a, "catScaling"))
        .and_then(|s| s.attribute("gapWidth"))
        .and_then(|g| g.parse::<f32>().ok())
        .map_or(150.0, |g| g * 100.0);

    let leaves = data.levels.first().cloned().unwrap_or_default();
    let mut box_index: Option<Vec<f32>> = None;
    let (chart_type, labels, values) = match layout {
        "waterfall" => {
            let totals = layout_pr
                .and_then(|p| cx(p, "subtotals"))
                .map(|s| {
                    cx_all(s, "idx")
                        .filter_map(|i| i.attribute("val")?.parse().ok())
                        .collect()
                })
                .unwrap_or_default();
            (ChartType::Waterfall { totals }, leaves, data.values.clone())
        }
        "funnel" => (ChartType::Funnel, leaves, data.values.clone()),
        "treemap" => (ChartType::Treemap, leaves, data.values.clone()),
        "sunburst" => (ChartType::Sunburst, leaves, data.values.clone()),
        "boxWhisker" => {
            // One box per distinct category; each point keeps its box's index
            // in `x_values`.
            let mut boxes: Vec<String> = Vec::new();
            let mut index = Vec::with_capacity(leaves.len());
            for l in &leaves {
                let i = boxes.iter().position(|b| b == l).unwrap_or_else(|| {
                    boxes.push(l.clone());
                    boxes.len() - 1
                });
                index.push(i as f32);
            }
            box_index = Some(index);
            let exclusive = layout_pr
                .and_then(|p| cx(p, "statistics"))
                .and_then(|s| s.attribute("quartileMethod"))
                != Some("inclusive");
            (
                ChartType::BoxWhisker { exclusive },
                boxes,
                data.values.clone(),
            )
        }
        "clusteredColumn" => {
            let column = ChartType::Bar {
                horizontal: false,
                grouping: BarGrouping::Clustered,
            };
            match layout_pr.and_then(|p| cx(p, "binning")) {
                Some(binning) => {
                    let (labels, counts) = histogram(&data.values, binning);
                    (column, labels, counts)
                }
                // A pareto's bars stand largest first.
                None if pareto => {
                    let mut pts: Vec<(String, f32)> =
                        leaves.into_iter().zip(data.values.clone()).collect();
                    pts.sort_by(|a, b| b.1.total_cmp(&a.1));
                    let (labels, values) = pts.into_iter().unzip();
                    (ChartType::Pareto, labels, values)
                }
                None => (column, leaves, data.values.clone()),
            }
        }
        // ponytail: region maps (`regionMap`) are not drawn.
        _ => return None,
    };
    let legend = cx(chart, "legend").map(|l| ChartLegend {
        position: match l.attribute("pos") {
            Some("r") => LegendPosition::Right,
            Some("l") => LegendPosition::Left,
            Some("t") => LegendPosition::Top,
            _ => LegendPosition::Bottom,
        },
    });
    Some(Chart {
        chart_type,
        series: vec![ChartSeries {
            label,
            color: None,
            fill_alpha: None,
            values,
            x_values: box_index,
            bubble_sizes: None,
            marker: None,
            line_width: None,
            marker_size: None,
            marker_color: None,
            marker_line_width: None,
        }],
        cat_axis: Some(ChartAxis {
            labels,
            gridline_color: None,
            line_color: None,
        }),
        val_axis: None,
        legend,
        gap_width_pct,
        plot_border_color: None,
        accent_colors,
        val_format_code: None,
        title: cx(chart, "title")
            .and_then(|t| cx(t, "tx"))
            .map(text_of)
            .filter(|t| !t.is_empty()),
        show_values,
        data_table: false,
        frame: Some(CHARTEX_GRID),
        view_3d: None,
        parent_labels: data.levels.get(1).cloned().unwrap_or_default(),
        ex_text: Some(style_text(style, theme)),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn histogram_bins_close_on_the_right() {
        let xml = r#"<b xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:binCount val="5"/></b>"#;
        let doc = roxmltree::Document::parse(xml).unwrap();
        let values = [
            3., 7., 8., 12., 14., 15., 18., 21., 22., 25., 27., 31., 33., 38., 41.,
        ];
        let (labels, counts) = histogram(&values, doc.root_element());
        assert_eq!(labels[0], "[3, 10.6]");
        assert_eq!(labels[1], "(10.6, 18.2]");
        // Word's histogram probe: 3, 4, 3, 3, 2.
        assert_eq!(counts, vec![3., 4., 3., 3., 2.]);
    }
}
