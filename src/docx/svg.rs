//! Minimal SVG reader for `asvg:svgBlip` pictures. Office 2016+ stores the SVG
//! beside the PNG that older Words draw, and Word's PDF export draws the vector.
//! Shapes, paths, solid paints, strokes, class stylesheets and transforms cover
//! icons and plots; anything else (text, images, gradients, opacity, `use`)
//! makes `parse` return None so the picture keeps its PNG fallback.

use std::collections::HashMap;

use roxmltree::{Document, Node};

const SVG_NS: &str = "http://www.w3.org/2000/svg";
pub(crate) const IDENTITY: [f32; 6] = [1.0, 0.0, 0.0, 1.0, 0.0, 0.0];
/// Control-point distance, as a fraction of the radius, of a quarter-circle Bézier.
const KAPPA: f32 = 0.552_284_8;

#[derive(Clone, Debug, PartialEq)]
pub(crate) enum Seg {
    Move(f32, f32),
    Line(f32, f32),
    Cubic(f32, f32, f32, f32, f32, f32),
    Close,
}

#[derive(Clone, Debug)]
pub(crate) struct SvgItem {
    pub segs: Vec<Seg>,
    /// Maps the item's user space into viewBox space (`a b c d e f`).
    pub transform: [f32; 6],
    pub fill: Option<[u8; 3]>,
    pub even_odd: bool,
    pub stroke: Option<[u8; 3]>,
    pub stroke_width: f32,
    /// 0 butt, 1 round, 2 square (PDF numbering).
    pub cap: u8,
    /// 0 miter, 1 round, 2 bevel (PDF numbering).
    pub join: u8,
    pub miter_limit: f32,
    pub dash: Vec<f32>,
}

/// The viewBox is stretched to the picture extent whatever
/// `preserveAspectRatio` says: Word resizes an SVG like any picture (probe
/// s07, a 160×80 viewBox in a 166×166pt frame fills it).
pub(crate) struct SvgDrawing {
    pub view_box: [f32; 4],
    pub items: Vec<SvgItem>,
}

pub(crate) fn looks_like_svg(data: &[u8]) -> bool {
    let start = data.strip_prefix(b"\xef\xbb\xbf").unwrap_or(data);
    start
        .iter()
        .find(|b| !b.is_ascii_whitespace())
        .is_some_and(|&b| b == b'<')
        && !data.starts_with(b"<?xml-stylesheet")
}

/// None means "not an SVG we can draw": the caller falls back to the raster.
pub(crate) fn parse(data: &[u8]) -> Option<SvgDrawing> {
    if !looks_like_svg(data) {
        return None;
    }
    let text = std::str::from_utf8(data)
        .ok()?
        .trim_start_matches('\u{feff}');
    let doc = Document::parse(text).ok()?;
    let root = doc.root_element();
    if !root.has_tag_name((SVG_NS, "svg")) {
        return None;
    }
    let view_box = match root.attribute("viewBox").map(numbers) {
        Some(Some(v)) if v.len() == 4 && v[2] > 0.0 && v[3] > 0.0 => [v[0], v[1], v[2], v[3]],
        Some(_) => return None,
        None => {
            let (w, h) = (root_length(root, "width")?, root_length(root, "height")?);
            if w <= 0.0 || h <= 0.0 {
                return None;
            }
            [0.0, 0.0, w, h]
        }
    };
    let mut parser = Parser {
        classes: parse_stylesheet(root)?,
        items: Vec::new(),
    };
    parser.element(root, &Style::default(), IDENTITY)?;
    Some(SvgDrawing {
        view_box,
        items: parser.items,
    })
}

#[derive(Clone, Copy)]
enum Paint {
    None,
    Rgb([u8; 3]),
    Current,
}

#[derive(Clone)]
struct Style {
    fill: Paint,
    stroke: Paint,
    color: [u8; 3],
    stroke_width: f32,
    cap: u8,
    join: u8,
    miter_limit: f32,
    dash: Vec<f32>,
    even_odd: bool,
    /// `visibility: hidden` — inherited, a child may turn it back on.
    hidden: bool,
    /// `display: none` — the whole subtree is skipped.
    skip: bool,
}

impl Default for Style {
    fn default() -> Self {
        Style {
            fill: Paint::Rgb([0, 0, 0]),
            stroke: Paint::None,
            color: [0, 0, 0],
            stroke_width: 1.0,
            cap: 0,
            join: 0,
            miter_limit: 4.0,
            dash: Vec::new(),
            even_odd: false,
            hidden: false,
            skip: false,
        }
    }
}

impl Style {
    fn resolve(&self, paint: Paint) -> Option<[u8; 3]> {
        match paint {
            Paint::None => None,
            Paint::Rgb(c) => Some(c),
            Paint::Current => Some(self.color),
        }
    }
}

struct Parser {
    classes: HashMap<String, Vec<(String, String)>>,
    items: Vec<SvgItem>,
}

impl Parser {
    fn element(&mut self, n: Node, parent: &Style, parent_t: [f32; 6]) -> Option<()> {
        if !n.is_element() || n.tag_name().namespace() != Some(SVG_NS) {
            return Some(());
        }
        let name = n.tag_name().name();
        match name {
            "title" | "desc" | "metadata" | "defs" | "style" | "clipPath" | "mask" | "marker"
            | "pattern" | "linearGradient" | "radialGradient" | "symbol" | "filter" | "script" => {
                return Some(());
            }
            // ponytail: text needs fonts, `use` needs defs lookup; raster fallback until a fixture needs them.
            "text" | "tspan" | "image" | "use" | "foreignObject" | "switch" => return None,
            "svg" if !n.parent().is_some_and(|p| p.is_root()) => return None,
            _ => {}
        }
        let mut style = parent.clone();
        for a in n.attributes() {
            if a.namespace().is_none() && a.name() != "style" {
                apply_prop(&mut style, a.name(), a.value())?;
            }
        }
        for class in n.attribute("class").unwrap_or("").split_whitespace() {
            for (k, v) in self.classes.get(class).into_iter().flatten() {
                apply_prop(&mut style, k, v)?;
            }
        }
        for (k, v) in declarations(n.attribute("style").unwrap_or("")) {
            apply_prop(&mut style, &k, &v)?;
        }
        if style.skip {
            return Some(());
        }
        let t = match n.attribute("transform") {
            Some(s) => multiply(parent_t, parse_transform(s)?),
            None => parent_t,
        };
        let segs = match name {
            "svg" | "g" | "a" => {
                for child in n.children() {
                    self.element(child, &style, t)?;
                }
                return Some(());
            }
            "rect" => rect(n)?,
            "circle" => {
                let r = length(n.attribute("r"))?;
                ellipse(
                    length(n.attribute("cx")).unwrap_or(0.0),
                    length(n.attribute("cy")).unwrap_or(0.0),
                    r,
                    r,
                )
            }
            "ellipse" => ellipse(
                length(n.attribute("cx")).unwrap_or(0.0),
                length(n.attribute("cy")).unwrap_or(0.0),
                length(n.attribute("rx"))?,
                length(n.attribute("ry"))?,
            ),
            "line" => vec![
                Seg::Move(
                    length(n.attribute("x1")).unwrap_or(0.0),
                    length(n.attribute("y1")).unwrap_or(0.0),
                ),
                Seg::Line(
                    length(n.attribute("x2")).unwrap_or(0.0),
                    length(n.attribute("y2")).unwrap_or(0.0),
                ),
            ],
            "polyline" | "polygon" => {
                let pts = numbers(n.attribute("points").unwrap_or(""))?;
                let mut segs: Vec<Seg> = pts
                    .chunks_exact(2)
                    .enumerate()
                    .map(|(i, p)| {
                        if i == 0 {
                            Seg::Move(p[0], p[1])
                        } else {
                            Seg::Line(p[0], p[1])
                        }
                    })
                    .collect();
                if name == "polygon" && !segs.is_empty() {
                    segs.push(Seg::Close);
                }
                segs
            }
            "path" => path_data(n.attribute("d").unwrap_or(""))?,
            _ => return Some(()),
        };
        if style.hidden || segs.is_empty() {
            return Some(());
        }
        let stroke = style
            .resolve(style.stroke)
            .filter(|_| style.stroke_width > 0.0);
        self.items.push(SvgItem {
            segs,
            transform: t,
            fill: style.resolve(style.fill),
            even_odd: style.even_odd,
            stroke,
            stroke_width: style.stroke_width,
            cap: style.cap,
            join: style.join,
            miter_limit: style.miter_limit,
            dash: style.dash.clone(),
        });
        Some(())
    }
}

/// One presentation property, from an attribute, a class rule or an inline
/// style. None = a value we cannot draw (gradient, opacity, clip, filter).
fn apply_prop(style: &mut Style, name: &str, value: &str) -> Option<()> {
    let value = value.trim();
    if value == "inherit" {
        return Some(());
    }
    match name {
        "fill" => style.fill = paint(value)?,
        "stroke" => style.stroke = paint(value)?,
        "color" => {
            if let Paint::Rgb(c) = paint(value)? {
                style.color = c;
            }
        }
        "stroke-width" => style.stroke_width = length(Some(value))?,
        "stroke-linecap" => {
            style.cap = match value {
                "butt" => 0,
                "round" => 1,
                "square" => 2,
                _ => return None,
            }
        }
        "stroke-linejoin" => {
            style.join = match value {
                "miter" | "miter-clip" | "arcs" => 0,
                "round" => 1,
                "bevel" => 2,
                _ => return None,
            }
        }
        "stroke-miterlimit" => style.miter_limit = value.parse().ok()?,
        "stroke-dasharray" => {
            style.dash = if value == "none" {
                Vec::new()
            } else {
                let d = numbers(value)?;
                if d.iter().any(|&x| x < 0.0) {
                    return None;
                }
                if d.iter().all(|&x| x == 0.0) {
                    Vec::new()
                } else {
                    d
                }
            }
        }
        "fill-rule" => {
            style.even_odd = match value {
                "evenodd" => true,
                "nonzero" => false,
                _ => return None,
            }
        }
        "display" => style.skip = value == "none",
        "visibility" => style.hidden = matches!(value, "hidden" | "collapse"),
        // ponytail: opacity would need an ExtGState per value; raster fallback until a fixture needs it.
        "opacity" | "fill-opacity" | "stroke-opacity" => {
            if value.parse::<f32>().ok()? < 1.0 {
                return None;
            }
        }
        "clip-path" | "mask" | "filter" | "marker" | "marker-start" | "marker-mid"
        | "marker-end"
            if value != "none" =>
        {
            return None;
        }
        _ => {}
    }
    Some(())
}

fn paint(value: &str) -> Option<Paint> {
    let lower = value.to_ascii_lowercase();
    match lower.as_str() {
        "none" | "transparent" => Some(Paint::None),
        "currentcolor" => Some(Paint::Current),
        _ => color(&lower).map(Paint::Rgb),
    }
}

fn color(lower: &str) -> Option<[u8; 3]> {
    if let Some(hex) = lower.strip_prefix('#') {
        if !hex.is_ascii() {
            return None;
        }
        let byte = |s: &str| u8::from_str_radix(s, 16).ok();
        return match hex.len() {
            3 => Some([
                byte(&hex[0..1])? * 17,
                byte(&hex[1..2])? * 17,
                byte(&hex[2..3])? * 17,
            ]),
            6 => Some([byte(&hex[0..2])?, byte(&hex[2..4])?, byte(&hex[4..6])?]),
            _ => None,
        };
    }
    if let Some(args) = lower.strip_prefix("rgb(").and_then(|s| s.strip_suffix(')')) {
        let parts: Vec<&str> = args.split([',', ' ']).filter(|s| !s.is_empty()).collect();
        if parts.len() != 3 {
            return None;
        }
        let channel = |s: &str| -> Option<u8> {
            let v = match s.strip_suffix('%') {
                Some(p) => p.parse::<f32>().ok()? * 2.55,
                None => s.parse::<f32>().ok()?,
            };
            Some(v.round().clamp(0.0, 255.0) as u8)
        };
        return Some([channel(parts[0])?, channel(parts[1])?, channel(parts[2])?]);
    }
    // ponytail: the CSS basic names; the other 130 fall back to the raster.
    Some(match lower {
        "black" => [0, 0, 0],
        "white" => [255, 255, 255],
        "red" => [255, 0, 0],
        "lime" => [0, 255, 0],
        "blue" => [0, 0, 255],
        "yellow" => [255, 255, 0],
        "cyan" | "aqua" => [0, 255, 255],
        "magenta" | "fuchsia" => [255, 0, 255],
        "gray" | "grey" => [128, 128, 128],
        "silver" => [192, 192, 192],
        "maroon" => [128, 0, 0],
        "green" => [0, 128, 0],
        "navy" => [0, 0, 128],
        "olive" => [128, 128, 0],
        "purple" => [128, 0, 128],
        "teal" => [0, 128, 128],
        "orange" => [255, 165, 0],
        "darkgray" | "darkgrey" => [169, 169, 169],
        "lightgray" | "lightgrey" => [211, 211, 211],
        _ => return None,
    })
}

/// A CSS length in user units (px). Relative units are unsupported.
fn length(value: Option<&str>) -> Option<f32> {
    let v = value?.trim();
    let split = v
        .find(|c: char| c.is_ascii_alphabetic() || c == '%')
        .unwrap_or(v.len());
    let number: f32 = v[..split].trim().parse().ok()?;
    let per_unit = match &v[split..] {
        "" | "px" => 1.0,
        "pt" => 4.0 / 3.0,
        "pc" => 16.0,
        "mm" => 96.0 / 25.4,
        "cm" => 96.0 / 2.54,
        "in" => 96.0,
        _ => return None,
    };
    Some(number * per_unit)
}

/// Root width/height; a percentage counts as absent.
fn root_length(root: Node, name: &str) -> Option<f32> {
    length(
        root.attribute(name)
            .filter(|v| !v.trim_end().ends_with('%')),
    )
}

fn numbers(s: &str) -> Option<Vec<f32>> {
    s.split(|c: char| c.is_whitespace() || c == ',')
        .filter(|t| !t.is_empty())
        .map(|t| t.parse::<f32>().ok())
        .collect()
}

fn declarations(css: &str) -> Vec<(String, String)> {
    css.split(';')
        .filter_map(|d| d.split_once(':'))
        .map(|(k, v)| (k.trim().to_string(), v.trim().to_string()))
        .collect()
}

/// `<style>` sheets of plain `.class { … }` rules; anything else is unsupported.
fn parse_stylesheet(root: Node) -> Option<HashMap<String, Vec<(String, String)>>> {
    let mut css = String::new();
    for style in root
        .descendants()
        .filter(|n| n.has_tag_name((SVG_NS, "style")))
    {
        css.extend(style.children().filter_map(|c| c.text()));
    }
    while let Some(start) = css.find("/*") {
        let end = css[start..].find("*/").map_or(css.len(), |e| start + e + 2);
        css.replace_range(start..end, " ");
    }
    let mut rules: HashMap<String, Vec<(String, String)>> = HashMap::new();
    for block in css.split('}') {
        let Some((selectors, decls)) = block.split_once('{') else {
            if block.trim().is_empty() {
                continue;
            }
            return None;
        };
        let decls = declarations(decls);
        for selector in selectors.split(',') {
            let class = selector.trim().strip_prefix('.')?;
            if class.is_empty()
                || !class
                    .chars()
                    .all(|c| c.is_alphanumeric() || c == '-' || c == '_')
            {
                return None;
            }
            rules
                .entry(class.to_string())
                .or_default()
                .extend(decls.iter().cloned());
        }
    }
    Some(rules)
}

/// `p ∘ l`: apply `l` first, then `p`.
pub(crate) fn multiply(p: [f32; 6], l: [f32; 6]) -> [f32; 6] {
    [
        p[0] * l[0] + p[2] * l[1],
        p[1] * l[0] + p[3] * l[1],
        p[0] * l[2] + p[2] * l[3],
        p[1] * l[2] + p[3] * l[3],
        p[0] * l[4] + p[2] * l[5] + p[4],
        p[1] * l[4] + p[3] * l[5] + p[5],
    ]
}

fn parse_transform(s: &str) -> Option<[f32; 6]> {
    let mut m = IDENTITY;
    let mut rest = s.trim();
    while !rest.is_empty() {
        let open = rest.find('(')?;
        let close = open + rest[open..].find(')')?;
        let a = numbers(&rest[open + 1..close])?;
        let rotation = |deg: f32| {
            let (sin, cos) = deg.to_radians().sin_cos();
            [cos, sin, -sin, cos, 0.0, 0.0]
        };
        let local = match (rest[..open].trim(), a.len()) {
            ("matrix", 6) => [a[0], a[1], a[2], a[3], a[4], a[5]],
            ("translate", 1) => [1.0, 0.0, 0.0, 1.0, a[0], 0.0],
            ("translate", 2) => [1.0, 0.0, 0.0, 1.0, a[0], a[1]],
            ("scale", 1) => [a[0], 0.0, 0.0, a[0], 0.0, 0.0],
            ("scale", 2) => [a[0], 0.0, 0.0, a[1], 0.0, 0.0],
            ("rotate", 1) => rotation(a[0]),
            ("rotate", 3) => multiply(
                multiply([1.0, 0.0, 0.0, 1.0, a[1], a[2]], rotation(a[0])),
                [1.0, 0.0, 0.0, 1.0, -a[1], -a[2]],
            ),
            ("skewX", 1) => [1.0, 0.0, a[0].to_radians().tan(), 1.0, 0.0, 0.0],
            ("skewY", 1) => [1.0, a[0].to_radians().tan(), 0.0, 1.0, 0.0, 0.0],
            _ => return None,
        };
        m = multiply(m, local);
        rest = rest[close + 1..].trim_start_matches(|c: char| c.is_whitespace() || c == ',');
    }
    Some(m)
}

fn rect(n: Node) -> Option<Vec<Seg>> {
    let x = length(n.attribute("x")).unwrap_or(0.0);
    let y = length(n.attribute("y")).unwrap_or(0.0);
    let w = length(n.attribute("width"))?;
    let h = length(n.attribute("height"))?;
    if w <= 0.0 || h <= 0.0 {
        return Some(Vec::new());
    }
    let (rx, ry) = (length(n.attribute("rx")), length(n.attribute("ry")));
    let rx = rx.or(ry).unwrap_or(0.0).min(w / 2.0);
    let ry = ry.or(rx.into()).unwrap_or(0.0).min(h / 2.0);
    if rx <= 0.0 || ry <= 0.0 {
        return Some(vec![
            Seg::Move(x, y),
            Seg::Line(x + w, y),
            Seg::Line(x + w, y + h),
            Seg::Line(x, y + h),
            Seg::Close,
        ]);
    }
    let (kx, ky) = (rx * KAPPA, ry * KAPPA);
    let (r, b) = (x + w, y + h);
    Some(vec![
        Seg::Move(x + rx, y),
        Seg::Line(r - rx, y),
        Seg::Cubic(r - rx + kx, y, r, y + ry - ky, r, y + ry),
        Seg::Line(r, b - ry),
        Seg::Cubic(r, b - ry + ky, r - rx + kx, b, r - rx, b),
        Seg::Line(x + rx, b),
        Seg::Cubic(x + rx - kx, b, x, b - ry + ky, x, b - ry),
        Seg::Line(x, y + ry),
        Seg::Cubic(x, y + ry - ky, x + rx - kx, y, x + rx, y),
        Seg::Close,
    ])
}

fn ellipse(cx: f32, cy: f32, rx: f32, ry: f32) -> Vec<Seg> {
    if rx <= 0.0 || ry <= 0.0 {
        return Vec::new();
    }
    let (kx, ky) = (rx * KAPPA, ry * KAPPA);
    vec![
        Seg::Move(cx + rx, cy),
        Seg::Cubic(cx + rx, cy + ky, cx + kx, cy + ry, cx, cy + ry),
        Seg::Cubic(cx - kx, cy + ry, cx - rx, cy + ky, cx - rx, cy),
        Seg::Cubic(cx - rx, cy - ky, cx - kx, cy - ry, cx, cy - ry),
        Seg::Cubic(cx + kx, cy - ry, cx + rx, cy - ky, cx + rx, cy),
        Seg::Close,
    ]
}

struct Tokens<'a> {
    b: &'a [u8],
    i: usize,
}

impl Tokens<'_> {
    fn skip(&mut self) {
        while self
            .b
            .get(self.i)
            .is_some_and(|c| c.is_ascii_whitespace() || *c == b',')
        {
            self.i += 1;
        }
    }

    fn done(&mut self) -> bool {
        self.skip();
        self.i >= self.b.len()
    }

    fn command(&mut self) -> Option<u8> {
        self.skip();
        let c = *self.b.get(self.i)?;
        c.is_ascii_alphabetic().then(|| {
            self.i += 1;
            c
        })
    }

    /// One number; `1.5.5` and `-5-3` split where a second number starts.
    fn num(&mut self) -> Option<f32> {
        self.skip();
        let b = self.b;
        let start = self.i;
        let mut i = start;
        if matches!(b.get(i), Some(b'-' | b'+')) {
            i += 1;
        }
        let mut digits = 0;
        while b.get(i).is_some_and(u8::is_ascii_digit) {
            i += 1;
            digits += 1;
        }
        if b.get(i) == Some(&b'.') {
            i += 1;
            while b.get(i).is_some_and(u8::is_ascii_digit) {
                i += 1;
                digits += 1;
            }
        }
        if digits == 0 {
            return None;
        }
        if matches!(b.get(i), Some(b'e' | b'E')) {
            let mut j = i + 1;
            if matches!(b.get(j), Some(b'-' | b'+')) {
                j += 1;
            }
            let exp_start = j;
            while b.get(j).is_some_and(u8::is_ascii_digit) {
                j += 1;
            }
            if j > exp_start {
                i = j;
            }
        }
        self.i = i;
        std::str::from_utf8(&b[start..i]).ok()?.parse().ok()
    }

    fn point(&mut self, rel: bool, cx: f32, cy: f32) -> Option<(f32, f32)> {
        let (x, y) = (self.num()?, self.num()?);
        Some(if rel { (cx + x, cy + y) } else { (x, y) })
    }

    fn flag(&mut self) -> Option<bool> {
        self.skip();
        let c = *self.b.get(self.i)?;
        self.i += 1;
        match c {
            b'0' => Some(false),
            b'1' => Some(true),
            _ => None,
        }
    }
}

fn quad(out: &mut Vec<Seg>, (x0, y0): (f32, f32), (qx, qy): (f32, f32), (x, y): (f32, f32)) {
    out.push(Seg::Cubic(
        x0 + 2.0 / 3.0 * (qx - x0),
        y0 + 2.0 / 3.0 * (qy - y0),
        x + 2.0 / 3.0 * (qx - x),
        y + 2.0 / 3.0 * (qy - y),
        x,
        y,
    ));
}

pub(crate) fn path_data(d: &str) -> Option<Vec<Seg>> {
    let mut t = Tokens {
        b: d.as_bytes(),
        i: 0,
    };
    let mut out = Vec::new();
    let (mut cx, mut cy, mut sx, mut sy) = (0.0f32, 0.0f32, 0.0f32, 0.0f32);
    let mut ctrl: Option<(f32, f32)> = None;
    let mut qctrl: Option<(f32, f32)> = None;
    let mut cmd = 0u8;
    while !t.done() {
        if let Some(c) = t.command() {
            cmd = c;
        } else if cmd == 0 || cmd.eq_ignore_ascii_case(&b'z') {
            return None;
        }
        let rel = cmd.is_ascii_lowercase();
        let (mut next_ctrl, mut next_qctrl) = (None, None);
        match cmd.to_ascii_uppercase() {
            b'M' => {
                let (x, y) = t.point(rel, cx, cy)?;
                out.push(Seg::Move(x, y));
                (cx, cy, sx, sy) = (x, y, x, y);
                cmd = if rel { b'l' } else { b'L' };
            }
            b'L' => {
                let (x, y) = t.point(rel, cx, cy)?;
                out.push(Seg::Line(x, y));
                (cx, cy) = (x, y);
            }
            b'H' => {
                let x = t.num()?;
                cx = if rel { cx + x } else { x };
                out.push(Seg::Line(cx, cy));
            }
            b'V' => {
                let y = t.num()?;
                cy = if rel { cy + y } else { y };
                out.push(Seg::Line(cx, cy));
            }
            b'C' | b'S' => {
                let (x1, y1) = if cmd.eq_ignore_ascii_case(&b'C') {
                    t.point(rel, cx, cy)?
                } else {
                    ctrl.map_or((cx, cy), |(px, py)| (2.0 * cx - px, 2.0 * cy - py))
                };
                let (x2, y2) = t.point(rel, cx, cy)?;
                let (x, y) = t.point(rel, cx, cy)?;
                out.push(Seg::Cubic(x1, y1, x2, y2, x, y));
                next_ctrl = Some((x2, y2));
                (cx, cy) = (x, y);
            }
            b'Q' | b'T' => {
                let q = if cmd.eq_ignore_ascii_case(&b'Q') {
                    t.point(rel, cx, cy)?
                } else {
                    qctrl.map_or((cx, cy), |(px, py)| (2.0 * cx - px, 2.0 * cy - py))
                };
                let (x, y) = t.point(rel, cx, cy)?;
                quad(&mut out, (cx, cy), q, (x, y));
                next_qctrl = Some(q);
                (cx, cy) = (x, y);
            }
            b'A' => {
                let (rx, ry, rot) = (t.num()?, t.num()?, t.num()?);
                let (large, sweep) = (t.flag()?, t.flag()?);
                let (x, y) = t.point(rel, cx, cy)?;
                arc(&mut out, (cx, cy), rx, ry, rot, large, sweep, (x, y));
                (cx, cy) = (x, y);
            }
            b'Z' => {
                out.push(Seg::Close);
                (cx, cy) = (sx, sy);
            }
            _ => return None,
        }
        ctrl = next_ctrl;
        qctrl = next_qctrl;
    }
    Some(out)
}

/// SVG endpoint arc → centre form → ≤90° cubic pieces (SVG 1.1 §F.6.5).
#[allow(clippy::too_many_arguments)]
fn arc(
    out: &mut Vec<Seg>,
    (x0, y0): (f32, f32),
    rx: f32,
    ry: f32,
    rot_deg: f32,
    large: bool,
    sweep: bool,
    (x, y): (f32, f32),
) {
    use std::f32::consts::{FRAC_PI_2, PI};
    if (x0, y0) == (x, y) {
        return;
    }
    let (mut rx, mut ry) = (rx.abs(), ry.abs());
    if rx == 0.0 || ry == 0.0 {
        out.push(Seg::Line(x, y));
        return;
    }
    let (sin_p, cos_p) = rot_deg.to_radians().sin_cos();
    let (dx, dy) = ((x0 - x) / 2.0, (y0 - y) / 2.0);
    let x1 = cos_p * dx + sin_p * dy;
    let y1 = -sin_p * dx + cos_p * dy;
    let lambda = x1 * x1 / (rx * rx) + y1 * y1 / (ry * ry);
    if lambda > 1.0 {
        rx *= lambda.sqrt();
        ry *= lambda.sqrt();
    }
    let num = rx * rx * ry * ry - rx * rx * y1 * y1 - ry * ry * x1 * x1;
    let den = rx * rx * y1 * y1 + ry * ry * x1 * x1;
    let coef = (num / den).max(0.0).sqrt() * if large != sweep { 1.0 } else { -1.0 };
    let (cx1, cy1) = (coef * rx * y1 / ry, -coef * ry * x1 / rx);
    let cx = cos_p * cx1 - sin_p * cy1 + (x0 + x) / 2.0;
    let cy = sin_p * cx1 + cos_p * cy1 + (y0 + y) / 2.0;
    let angle = |(ux, uy): (f32, f32), (vx, vy): (f32, f32)| {
        let dot = ux * vx + uy * vy;
        let len = ((ux * ux + uy * uy) * (vx * vx + vy * vy)).sqrt();
        let a = (dot / len).clamp(-1.0, 1.0).acos();
        if ux * vy - uy * vx < 0.0 { -a } else { a }
    };
    let v1 = ((x1 - cx1) / rx, (y1 - cy1) / ry);
    let v2 = ((-x1 - cx1) / rx, (-y1 - cy1) / ry);
    let theta1 = angle((1.0, 0.0), v1);
    let mut dtheta = angle(v1, v2);
    if !sweep && dtheta > 0.0 {
        dtheta -= 2.0 * PI;
    } else if sweep && dtheta < 0.0 {
        dtheta += 2.0 * PI;
    }
    let n = ((dtheta.abs() / FRAC_PI_2).ceil() as usize).max(1);
    let step = dtheta / n as f32;
    let alpha = 4.0 / 3.0 * (step / 4.0).tan();
    let point = |th: f32| {
        let (s, c) = th.sin_cos();
        (
            cx + rx * c * cos_p - ry * s * sin_p,
            cy + rx * c * sin_p + ry * s * cos_p,
        )
    };
    let tangent = |th: f32| {
        let (s, c) = th.sin_cos();
        (
            -rx * s * cos_p - ry * c * sin_p,
            -rx * s * sin_p + ry * c * cos_p,
        )
    };
    let mut th = theta1;
    for i in 0..n {
        let th2 = th + step;
        let (p0, d0) = (point(th), tangent(th));
        let (p3, d3) = (point(th2), tangent(th2));
        let end = if i + 1 == n { (x, y) } else { p3 };
        out.push(Seg::Cubic(
            p0.0 + alpha * d0.0,
            p0.1 + alpha * d0.1,
            p3.0 - alpha * d3.0,
            p3.1 - alpha * d3.1,
            end.0,
            end.1,
        ));
        th = th2;
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const PLOT: &str = r##"<svg xmlns="http://www.w3.org/2000/svg" width="160" height="80" viewBox="0 0 160 80"><rect x="1" y="1" width="158" height="78" fill="white" stroke="#333333"/><path d="M10 65 L45 45 L78 50 L116 22 L150 10" fill="none" stroke="#1674bf" stroke-width="2"/></svg>"##;

    #[test]
    fn issue_142_plot_parses_to_rect_and_polyline() {
        let d = parse(PLOT.as_bytes()).unwrap();
        assert_eq!(d.view_box, [0.0, 0.0, 160.0, 80.0]);
        assert_eq!(d.items.len(), 2);
        assert_eq!(
            (d.items[0].fill, d.items[0].stroke, d.items[0].stroke_width),
            (Some([255, 255, 255]), Some([0x33, 0x33, 0x33]), 1.0)
        );
        assert_eq!(d.items[0].segs[0], Seg::Move(1.0, 1.0));
        assert_eq!(
            (d.items[1].fill, d.items[1].stroke, d.items[1].stroke_width),
            (None, Some([0x16, 0x74, 0xbf]), 2.0)
        );
        assert_eq!(d.items[1].segs.len(), 5);
    }

    #[test]
    fn class_rules_and_transforms_apply_and_unsupported_content_falls_back() {
        let svg = r##"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><style>.a{fill:#ff0000;stroke:none;}</style><g transform="translate(1 2) scale(2)"><circle class="a" cx="1" cy="1" r="1" style="stroke:blue;stroke-width:0.5pt"/></g></svg>"##;
        let d = parse(svg.as_bytes()).unwrap();
        let it = &d.items[0];
        assert_eq!((it.fill, it.stroke), (Some([255, 0, 0]), Some([0, 0, 255])));
        assert!((it.stroke_width - 2.0 / 3.0).abs() < 1e-5);
        assert_eq!(it.transform, [2.0, 0.0, 0.0, 2.0, 1.0, 2.0]);
        assert_eq!(it.segs.len(), 6);
        for bad in [
            r#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"><text>x</text></svg>"#,
            r#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"><rect width="1" height="1" fill="url(#g)"/></svg>"#,
            r#"<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1"><rect width="1" height="1" opacity="0.5"/></svg>"#,
            r#"<svg xmlns="http://www.w3.org/2000/svg"><rect width="1" height="1"/></svg>"#,
        ] {
            assert!(parse(bad.as_bytes()).is_none(), "{bad}");
        }
    }

    #[test]
    fn path_data_handles_relative_shorthand_and_arcs() {
        let segs = path_data("m1,1 h2v2H1z M5 5c1 0 1 1 1 1s1-1 2 0 q1 1 2 0t1 0").unwrap();
        assert_eq!(
            &segs[..5],
            &[
                Seg::Move(1.0, 1.0),
                Seg::Line(3.0, 1.0),
                Seg::Line(3.0, 3.0),
                Seg::Line(1.0, 3.0),
                Seg::Close
            ]
        );
        assert_eq!(segs.len(), 10);
        assert_eq!(segs[7], Seg::Cubic(6.0, 6.0, 7.0, 5.0, 8.0, 6.0));
        let arc = path_data("M0 0 A1 1 0 0 1 2 0").unwrap();
        assert_eq!(arc.len(), 3);
        let Seg::Cubic(_, _, _, _, x, y) = arc[2] else {
            panic!()
        };
        assert_eq!((x, y), (2.0, 0.0));
        let Seg::Cubic(_, _, _, _, mx, my) = arc[1] else {
            panic!()
        };
        assert!(
            (mx - 1.0).abs() < 1e-5 && (my + 1.0).abs() < 1e-5,
            "apex {mx} {my}"
        );
        assert!(path_data("M0 0 1 1 Z 5").is_none());
        assert!(path_data("1 1").is_none());
    }
}
