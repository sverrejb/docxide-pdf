//! XML plumbing shared by the part parsers: namespaces, child finders, attribute and unit readers.

pub(super) const WML_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(super) const DML_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
pub(super) const WPD_NS: &str =
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
pub(super) const WPS_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
pub(super) const REL_NS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(super) const MC_NS_TOP: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
pub(super) const CHART_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/chart";
pub(super) const DSP_NS: &str = "http://schemas.microsoft.com/office/drawing/2008/diagram";
pub(super) const W14_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
pub(super) const VML_NS: &str = "urn:schemas-microsoft-com:vml";
pub(super) const W10_NS: &str = "urn:schemas-microsoft-com:office:word";
pub(super) const OFFICE_NS: &str = "urn:schemas-microsoft-com:office:office";
pub(super) const PIC_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/picture";
pub(super) const DIAGRAM_NS: &str = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
pub(super) const MATH_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
pub(super) const WPC_NS: &str =
    "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
pub(super) const WPG_NS: &str = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";

pub(super) fn twips_to_pts(twips: f32) -> f32 {
    twips / 20.0
}

pub(super) fn emu_to_pts(emu: f32) -> f32 {
    emu / 12700.0
}

/// A numeric attribute; `None` when absent or unparsable.
pub(super) fn f32_attr<'n, 'm>(
    node: roxmltree::Node,
    attr: impl Into<roxmltree::ExpandedName<'n, 'm>>,
) -> Option<f32> {
    node.attribute(attr).and_then(|v| v.parse::<f32>().ok())
}

/// An EMU attribute in points; `None` when absent.
pub(super) fn emu_attr_opt<'n, 'm>(
    node: roxmltree::Node,
    attr: impl Into<roxmltree::ExpandedName<'n, 'm>>,
) -> Option<f32> {
    f32_attr(node, attr).map(emu_to_pts)
}

/// An EMU attribute in points; 0 when absent.
pub(super) fn emu_attr(node: roxmltree::Node, attr: &str) -> f32 {
    emu_attr_opt(node, attr).unwrap_or(0.0)
}

/// A DrawingML angle attribute (60000ths of a degree) in degrees.
pub(super) fn angle_attr<'n, 'm>(
    node: roxmltree::Node,
    attr: impl Into<roxmltree::ExpandedName<'n, 'm>>,
) -> Option<f32> {
    f32_attr(node, attr).map(|v| v / 60_000.0)
}

/// A DrawingML percentage attribute (`ST_Percentage` as 1/1000 of a percent)
/// as a fraction.
pub(super) fn frac_attr<'n, 'm>(
    node: roxmltree::Node,
    attr: impl Into<roxmltree::ExpandedName<'n, 'm>>,
) -> Option<f32> {
    f32_attr(node, attr).map(|v| v / 100_000.0)
}

/// ST_OnOff truthiness (§17.17.4): "1", "true", and "on" all mean on.
pub(super) fn parse_on_off(v: &str) -> bool {
    matches!(v, "1" | "true" | "on")
}

/// Parse a VML/CSS length like "12.5pt" — the "pt" suffix and surrounding
/// whitespace are optional.
/// A VML length in points: `12pt`, `2in`, `5cm`, `10mm`, `3pc`, `96px` or a
/// bare number (points). An OLE preview sized `height:2in` fell back to its
/// `dyaOrig` (240pt for a 144pt chart) and moved to the next page.
pub(super) fn parse_pt(s: &str) -> Option<f32> {
    let s = s.trim();
    let (num, per_unit) = [
        ("pt", 1.0),
        ("in", 72.0),
        ("cm", 72.0 / 2.54),
        ("mm", 72.0 / 25.4),
        ("pc", 12.0),
        ("px", 0.75),
    ]
    .into_iter()
    .find_map(|(unit, k)| s.strip_suffix(unit).map(|n| (n, k)))
    .unwrap_or((s, 1.0));
    num.trim().parse::<f32>().ok().map(|v| v * per_unit)
}

/// Find a child element by name and namespace.
pub(super) fn find_child<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
    namespace: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    node.children().find(|n| n.has_tag_name((namespace, name)))
}

pub(super) fn dml<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, DML_NS)
}

/// All child elements with the given name and namespace.
pub(super) fn find_children<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
    namespace: &str,
) -> impl Iterator<Item = roxmltree::Node<'a, 'a>> {
    node.children()
        .filter(move |n| n.has_tag_name((namespace, name)))
}

pub(super) fn dml_children<'a>(
    parent: roxmltree::Node<'a, 'a>,
    name: &str,
) -> impl Iterator<Item = roxmltree::Node<'a, 'a>> {
    find_children(parent, name, DML_NS)
}

pub(super) fn wpd<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, WPD_NS)
}

pub(super) fn wps<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, WPS_NS)
}

pub(super) fn dsp<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, DSP_NS)
}

pub(super) fn chart_ns<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, CHART_NS)
}

pub(super) fn chart_ns_attr<'a>(node: roxmltree::Node<'a, 'a>, child: &str) -> Option<&'a str> {
    chart_ns(node, child).and_then(|n| n.attribute("val"))
}

pub(super) fn chart_ns_children<'a>(
    parent: roxmltree::Node<'a, 'a>,
    name: &str,
) -> impl Iterator<Item = roxmltree::Node<'a, 'a>> {
    find_children(parent, name, CHART_NS)
}

pub(super) fn wml_bool(parent: roxmltree::Node, name: &str) -> Option<bool> {
    wml(parent, name).map(|n| {
        n.attribute((WML_NS, "val"))
            .is_none_or(|v| v != "0" && v != "false")
    })
}

pub(super) fn wml<'a>(
    node: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(node, name, WML_NS)
}

pub(super) fn wml_attr<'a>(node: roxmltree::Node<'a, 'a>, child: &str) -> Option<&'a str> {
    wml(node, child).and_then(|n| n.attribute((WML_NS, "val")))
}

/// First math-namespace child element with the given local name.
pub(super) fn math_child<'a>(
    parent: roxmltree::Node<'a, 'a>,
    name: &str,
) -> Option<roxmltree::Node<'a, 'a>> {
    find_child(parent, name, MATH_NS)
}

/// A math property's `m:val`: `<pr><name m:val="…"/></pr>` under `node`.
pub(super) fn math_val<'a>(node: roxmltree::Node<'a, 'a>, pr: &str, name: &str) -> Option<&'a str> {
    math_child(math_child(node, pr)?, name)?.attribute((MATH_NS, "val"))
}

/// A math property flag (`<pr><name m:val="…"/></pr>`): None when absent; an
/// element without `m:val` is on (ST_OnOff).
pub(super) fn math_bool(node: roxmltree::Node, pr: &str, name: &str) -> Option<bool> {
    let n = math_child(math_child(node, pr)?, name)?;
    Some(n.attribute((MATH_NS, "val")).is_none_or(parse_on_off))
}

/// A math run's text (its `m:t` children).
pub(super) fn math_run_text(r: roxmltree::Node) -> String {
    find_children(r, "t", MATH_NS)
        .filter_map(|t| t.text())
        .collect()
}

/// A WordprocessingML twips attribute in points. ST_TwipsMeasure also
/// takes a universal measure such as "10mm" (§22.9.2.15): a page written
/// with `w:left="10mm"` lost its margins to the 1in default.
pub(super) fn twips_attr(node: roxmltree::Node, attr: &str) -> Option<f32> {
    let v = node.attribute((WML_NS, attr))?;
    v.parse::<f32>()
        .ok()
        .map(twips_to_pts)
        .or_else(|| universal_measure_pt(v))
}

fn universal_measure_pt(v: &str) -> Option<f32> {
    let per_unit = [
        ("mm", 72.0 / 25.4),
        ("cm", 72.0 / 2.54),
        ("in", 72.0),
        ("pt", 1.0),
        ("pc", 12.0),
        ("pi", 12.0),
    ];
    per_unit
        .iter()
        .find_map(|(unit, k)| Some(v.strip_suffix(unit)?.parse::<f32>().ok()? * k))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn vml_lengths_convert_to_points() {
        let pt = super::parse_pt;
        assert_eq!(pt("12pt"), Some(12.0));
        assert_eq!(pt(" 2in "), Some(144.0));
        assert_eq!(pt("186.8pt"), Some(186.8));
        assert!((pt("2.54cm").unwrap() - 72.0).abs() < 1e-4);
        assert!((pt("25.4mm").unwrap() - 72.0).abs() < 1e-4);
        assert_eq!(pt("3pc"), Some(36.0));
        assert_eq!(pt("96px"), Some(72.0));
        assert_eq!(pt("7"), Some(7.0));
        assert_eq!(pt("auto"), None);
    }

    #[test]
    fn twips_take_universal_measures() {
        let doc = roxmltree::Document::parse(
            r#"<w:m xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:a="1440" w:b="10mm" w:c="1in" w:d="-2.5pt"/>"#,
        )
        .unwrap();
        let n = doc.root_element();
        assert_eq!(super::twips_attr(n, "a"), Some(72.0));
        assert!((super::twips_attr(n, "b").unwrap() - 28.346).abs() < 0.01);
        assert_eq!(super::twips_attr(n, "c"), Some(72.0));
        assert_eq!(super::twips_attr(n, "d"), Some(-2.5));
    }

    #[test]
    fn test_twips_to_pts() {
        assert_eq!(twips_to_pts(0.0), 0.0);
        assert_eq!(twips_to_pts(20.0), 1.0);
        assert_eq!(twips_to_pts(360.0), 18.0);
        assert_eq!(twips_to_pts(1440.0), 72.0); // 1 inch
    }

    #[test]
    fn test_emu_to_pts() {
        assert_eq!(emu_to_pts(0.0), 0.0);
        assert_eq!(emu_to_pts(12700.0), 1.0);
        assert_eq!(emu_to_pts(914400.0), 72.0); // 1 inch
    }

    // --- XML-based utility tests ---

    #[test]
    fn test_wml_bool() {
        let xml = format!(
            r#"<w:pPr xmlns:w="{}">
                <w:b/>
                <w:i w:val="0"/>
                <w:caps w:val="true"/>
                <w:vanish w:val="false"/>
            </w:pPr>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        // Present with no val attribute → true
        assert_eq!(wml_bool(node, "b"), Some(true));
        // val="0" → false
        assert_eq!(wml_bool(node, "i"), Some(false));
        // val="true" → true
        assert_eq!(wml_bool(node, "caps"), Some(true));
        // val="false" → false
        assert_eq!(wml_bool(node, "vanish"), Some(false));
        // Missing element → None
        assert_eq!(wml_bool(node, "strike"), None);
    }

    #[test]
    fn test_wml_attr() {
        let xml = format!(
            r#"<w:pPr xmlns:w="{}">
                <w:jc w:val="center"/>
                <w:sz w:val="24"/>
            </w:pPr>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        assert_eq!(wml_attr(node, "jc"), Some("center"));
        assert_eq!(wml_attr(node, "sz"), Some("24"));
        assert_eq!(wml_attr(node, "missing"), None);
    }

    #[test]
    fn test_twips_attr() {
        let xml = format!(
            r#"<w:ind xmlns:w="{}" w:left="720" w:right="360"/>"#,
            WML_NS
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        let node = doc.root_element();
        assert_eq!(twips_attr(node, "left"), Some(36.0)); // 720/20=36pt
        assert_eq!(twips_attr(node, "right"), Some(18.0)); // 360/20=18pt
        assert_eq!(twips_attr(node, "hanging"), None);
    }
}
