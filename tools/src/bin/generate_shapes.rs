/// Code generator: parses presetShapeDefinitions.xml (ECMA-376 Annex D) and emits
/// `src/geometry/definitions.rs` with all 187 preset shape definitions as Rust const data.
///
/// Usage:
///   generate-shapes [input.xml] [output.rs]
///
/// Defaults:
///   input  = tools/data/presetShapeDefinitions.xml
///   output = (stdout)

use roxmltree::{Document, Node};
use std::fmt::Write;
use std::fs;

fn main() {
    let xml_path = std::env::args()
        .nth(1)
        .unwrap_or_else(|| "tools/data/presetShapeDefinitions.xml".into());
    let out_path = std::env::args().nth(2);

    let xml = fs::read_to_string(&xml_path).unwrap_or_else(|e| panic!("Failed to read {xml_path}: {e}"));
    let doc = Document::parse(&xml).unwrap_or_else(|e| panic!("Failed to parse XML: {e}"));
    let root = doc.root_element();

    let shapes: Vec<ShapeDef> = root
        .children()
        .filter(|n| n.is_element())
        .map(parse_shape)
        .collect();

    eprintln!("Parsed {} shapes", shapes.len());

    let code = generate(&shapes);

    match out_path {
        Some(p) => {
            fs::write(&p, &code).unwrap_or_else(|e| panic!("Failed to write {p}: {e}"));
            eprintln!("Wrote {p}");
        }
        None => print!("{code}"),
    }
}

// ---------------------------------------------------------------------------
// Data model for parsed shapes
// ---------------------------------------------------------------------------

struct ShapeDef {
    name: String,
    rust_name: String,
    adjusts: Vec<(String, i64)>,
    guides: Vec<GuideParsed>,
    paths: Vec<PathParsed>,
    text_rect: Option<[String; 4]>,
}

struct GuideParsed {
    name: String,
    op: String,
    x: String,
    y: String,
    z: String,
}

struct PathParsed {
    commands: Vec<CmdParsed>,
    w: Option<i64>,
    h: Option<i64>,
    fill: String,
    stroke: bool,
}

enum CmdParsed {
    MoveTo(String, String),
    LineTo(String, String),
    ArcTo(String, String, String, String),
    CubicBezTo(String, String, String, String, String, String),
    QuadBezTo(String, String, String, String),
    Close,
}

// ---------------------------------------------------------------------------
// XML parsing
// ---------------------------------------------------------------------------

fn parse_shape(node: Node) -> ShapeDef {
    let name = node.tag_name().name().to_string();
    let rust_name = to_screaming_snake(&name);
    let mut adjusts = Vec::new();
    let mut guides = Vec::new();
    let mut paths = Vec::new();
    let mut text_rect = None;

    for child in node.children().filter(|n| n.is_element()) {
        match child.tag_name().name() {
            "avLst" => {
                for gd in child.children().filter(|n| n.is_element() && n.tag_name().name() == "gd") {
                    let gd_name = gd.attribute("name").unwrap_or("");
                    let fmla = gd.attribute("fmla").unwrap_or("");
                    let parts: Vec<&str> = fmla.split_whitespace().collect();
                    if parts.len() >= 2 && parts[0] == "val" {
                        adjusts.push((gd_name.to_string(), parts[1].parse::<i64>().unwrap_or(0)));
                    }
                }
            }
            "gdLst" => {
                for gd in child.children().filter(|n| n.is_element() && n.tag_name().name() == "gd") {
                    let gd_name = gd.attribute("name").unwrap_or("").to_string();
                    let fmla = gd.attribute("fmla").unwrap_or("");
                    let parts: Vec<&str> = fmla.split_whitespace().collect();
                    if parts.is_empty() {
                        continue;
                    }
                    guides.push(GuideParsed {
                        name: gd_name,
                        op: fmla_to_op(parts[0]).to_string(),
                        x: parts.get(1).unwrap_or(&"").to_string(),
                        y: parts.get(2).unwrap_or(&"").to_string(),
                        z: parts.get(3).unwrap_or(&"").to_string(),
                    });
                }
            }
            "pathLst" => {
                for p in child.children().filter(|n| n.is_element() && n.tag_name().name() == "path") {
                    paths.push(parse_path(p));
                }
            }
            "rect" => {
                text_rect = Some([
                    child.attribute("l").unwrap_or("l").to_string(),
                    child.attribute("t").unwrap_or("t").to_string(),
                    child.attribute("r").unwrap_or("r").to_string(),
                    child.attribute("b").unwrap_or("b").to_string(),
                ]);
            }
            _ => {} // ignore ahLst, cxnLst
        }
    }

    ShapeDef { name, rust_name, adjusts, guides, paths, text_rect }
}

fn parse_path(node: Node) -> PathParsed {
    let w = node.attribute("w").and_then(|v| v.parse().ok());
    let h = node.attribute("h").and_then(|v| v.parse().ok());
    let fill = match node.attribute("fill") {
        Some("none") => "None",
        _ => "Norm",
    };
    let stroke = !matches!(node.attribute("stroke"), Some("false"));

    let mut commands = Vec::new();
    for cmd in node.children().filter(|n| n.is_element()) {
        match cmd.tag_name().name() {
            "moveTo" => {
                if let Some(pt) = cmd.children().find(|n| n.is_element()) {
                    commands.push(CmdParsed::MoveTo(
                        pt.attribute("x").unwrap_or("0").into(),
                        pt.attribute("y").unwrap_or("0").into(),
                    ));
                }
            }
            "lnTo" => {
                if let Some(pt) = cmd.children().find(|n| n.is_element()) {
                    commands.push(CmdParsed::LineTo(
                        pt.attribute("x").unwrap_or("0").into(),
                        pt.attribute("y").unwrap_or("0").into(),
                    ));
                }
            }
            "arcTo" => {
                commands.push(CmdParsed::ArcTo(
                    cmd.attribute("wR").unwrap_or("0").into(),
                    cmd.attribute("hR").unwrap_or("0").into(),
                    cmd.attribute("stAng").unwrap_or("0").into(),
                    cmd.attribute("swAng").unwrap_or("0").into(),
                ));
            }
            "cubicBezTo" => {
                let pts: Vec<_> = cmd.children().filter(|n| n.is_element()).collect();
                if pts.len() >= 3 {
                    commands.push(CmdParsed::CubicBezTo(
                        pts[0].attribute("x").unwrap_or("0").into(),
                        pts[0].attribute("y").unwrap_or("0").into(),
                        pts[1].attribute("x").unwrap_or("0").into(),
                        pts[1].attribute("y").unwrap_or("0").into(),
                        pts[2].attribute("x").unwrap_or("0").into(),
                        pts[2].attribute("y").unwrap_or("0").into(),
                    ));
                }
            }
            "quadBezTo" => {
                let pts: Vec<_> = cmd.children().filter(|n| n.is_element()).collect();
                if pts.len() >= 2 {
                    commands.push(CmdParsed::QuadBezTo(
                        pts[0].attribute("x").unwrap_or("0").into(),
                        pts[0].attribute("y").unwrap_or("0").into(),
                        pts[1].attribute("x").unwrap_or("0").into(),
                        pts[1].attribute("y").unwrap_or("0").into(),
                    ));
                }
            }
            "close" => commands.push(CmdParsed::Close),
            _ => {}
        }
    }

    PathParsed { commands, w, h, fill: fill.to_string(), stroke }
}

// ---------------------------------------------------------------------------
// Rust code generation
// ---------------------------------------------------------------------------

fn generate(shapes: &[ShapeDef]) -> String {
    let mut o = String::with_capacity(512 * 1024);

    emit_header(&mut o, shapes.len());
    emit_lookup(&mut o, shapes);

    for s in shapes {
        emit_shape(&mut o, s);
    }

    o
}

fn emit_header(o: &mut String, count: usize) {
    writeln!(o, "// Auto-generated by tools/src/bin/generate_shapes.rs — do not edit manually.").unwrap();
    writeln!(o, "// Source: presetShapeDefinitions.xml (ECMA-376, Annex D)").unwrap();
    writeln!(o, "// Shapes: {count}").unwrap();
    writeln!(o).unwrap();
    writeln!(o, "use super::formulas::{{FormulaOp, GuideDef}};").unwrap();
    writeln!(o, "use super::path::{{PathCommandDef, PathDef, PathFill}};").unwrap();
    writeln!(o).unwrap();

    // Struct definitions
    writeln!(o, "pub struct TextRectDef {{").unwrap();
    writeln!(o, "    pub l: &'static str,").unwrap();
    writeln!(o, "    pub t: &'static str,").unwrap();
    writeln!(o, "    pub r: &'static str,").unwrap();
    writeln!(o, "    pub b: &'static str,").unwrap();
    writeln!(o, "}}").unwrap();
    writeln!(o).unwrap();

    writeln!(o, "pub struct PresetDef {{").unwrap();
    writeln!(o, "    pub adjust_defaults: &'static [(&'static str, i64)],").unwrap();
    writeln!(o, "    pub guides: &'static [GuideDef],").unwrap();
    writeln!(o, "    pub paths: &'static [PathDef],").unwrap();
    writeln!(o, "    pub text_rect: Option<TextRectDef>,").unwrap();
    writeln!(o, "}}").unwrap();
    writeln!(o).unwrap();

    // gd! macro
    writeln!(o, "macro_rules! gd {{").unwrap();
    writeln!(
        o,
        "    ($name:expr, $op:ident, $x:expr, $y:expr, $z:expr) => {{"
    )
    .unwrap();
    writeln!(
        o,
        "        GuideDef {{ name: $name, op: FormulaOp::$op, x: $x, y: $y, z: $z }}"
    )
    .unwrap();
    writeln!(o, "    }};").unwrap();
    writeln!(o, "}}").unwrap();
    writeln!(o).unwrap();
}

fn emit_lookup(o: &mut String, shapes: &[ShapeDef]) {
    writeln!(o, "pub fn lookup(name: &str) -> Option<&'static PresetDef> {{").unwrap();
    writeln!(o, "    match name {{").unwrap();
    for s in shapes {
        writeln!(o, "        \"{}\" => Some(&{}),", s.name, s.rust_name).unwrap();
    }
    writeln!(o, "        _ => None,").unwrap();
    writeln!(o, "    }}").unwrap();
    writeln!(o, "}}").unwrap();
    writeln!(o).unwrap();
}

fn emit_shape(o: &mut String, s: &ShapeDef) {
    // Path command statics
    for (i, path) in s.paths.iter().enumerate() {
        let cmds_name = if s.paths.len() == 1 {
            format!("{}_CMDS", s.rust_name)
        } else {
            format!("{}_P{}_CMDS", s.rust_name, i)
        };

        writeln!(o, "static {cmds_name}: &[PathCommandDef] = &[").unwrap();
        for cmd in &path.commands {
            match cmd {
                CmdParsed::MoveTo(x, y) => {
                    writeln!(o, "    PathCommandDef::MoveTo {{ x: \"{x}\", y: \"{y}\" }},").unwrap();
                }
                CmdParsed::LineTo(x, y) => {
                    writeln!(o, "    PathCommandDef::LineTo {{ x: \"{x}\", y: \"{y}\" }},").unwrap();
                }
                CmdParsed::ArcTo(wr, hr, st, sw) => {
                    writeln!(
                        o,
                        "    PathCommandDef::ArcTo {{ wr: \"{wr}\", hr: \"{hr}\", st_ang: \"{st}\", sw_ang: \"{sw}\" }},"
                    )
                    .unwrap();
                }
                CmdParsed::CubicBezTo(x1, y1, x2, y2, x3, y3) => {
                    writeln!(
                        o,
                        "    PathCommandDef::CubicBezTo {{ x1: \"{x1}\", y1: \"{y1}\", x2: \"{x2}\", y2: \"{y2}\", x3: \"{x3}\", y3: \"{y3}\" }},"
                    )
                    .unwrap();
                }
                CmdParsed::QuadBezTo(x1, y1, x2, y2) => {
                    writeln!(
                        o,
                        "    PathCommandDef::QuadBezTo {{ x1: \"{x1}\", y1: \"{y1}\", x2: \"{x2}\", y2: \"{y2}\" }},"
                    )
                    .unwrap();
                }
                CmdParsed::Close => {
                    writeln!(o, "    PathCommandDef::Close,").unwrap();
                }
            }
        }
        writeln!(o, "];").unwrap();
    }

    // Paths static
    let paths_name = format!("{}_PATHS", s.rust_name);
    writeln!(o, "static {paths_name}: &[PathDef] = &[").unwrap();
    for (i, path) in s.paths.iter().enumerate() {
        let cmds_name = if s.paths.len() == 1 {
            format!("{}_CMDS", s.rust_name)
        } else {
            format!("{}_P{}_CMDS", s.rust_name, i)
        };
        let w_str = match path.w {
            Some(v) => format!("Some({v})"),
            None => "None".into(),
        };
        let h_str = match path.h {
            Some(v) => format!("Some({v})"),
            None => "None".into(),
        };
        writeln!(
            o,
            "    PathDef {{ commands: {cmds_name}, w: {w_str}, h: {h_str}, fill: PathFill::{}, stroke: {} }},",
            path.fill, path.stroke
        )
        .unwrap();
    }
    writeln!(o, "];").unwrap();

    // Guides static
    let has_guides = !s.guides.is_empty();
    let guides_name = format!("{}_GUIDES", s.rust_name);
    if has_guides {
        writeln!(o, "static {guides_name}: &[GuideDef] = &[").unwrap();
        for g in &s.guides {
            writeln!(
                o,
                "    gd!(\"{}\", {}, \"{}\", \"{}\", \"{}\"),",
                g.name, g.op, g.x, g.y, g.z
            )
            .unwrap();
        }
        writeln!(o, "];").unwrap();
    }

    // PresetDef static
    writeln!(o, "static {}: PresetDef = PresetDef {{", s.rust_name).unwrap();

    // adjust_defaults
    if s.adjusts.is_empty() {
        writeln!(o, "    adjust_defaults: &[],").unwrap();
    } else {
        write!(o, "    adjust_defaults: &[").unwrap();
        for (i, (name, val)) in s.adjusts.iter().enumerate() {
            if i > 0 {
                write!(o, ", ").unwrap();
            }
            write!(o, "(\"{name}\", {val})").unwrap();
        }
        writeln!(o, "],").unwrap();
    }

    // guides
    if has_guides {
        writeln!(o, "    guides: {guides_name},").unwrap();
    } else {
        writeln!(o, "    guides: &[],").unwrap();
    }

    // paths
    writeln!(o, "    paths: {paths_name},").unwrap();

    // text_rect
    match &s.text_rect {
        Some([l, t, r, b]) => {
            writeln!(
                o,
                "    text_rect: Some(TextRectDef {{ l: \"{l}\", t: \"{t}\", r: \"{r}\", b: \"{b}\" }}),"
            )
            .unwrap();
        }
        None => writeln!(o, "    text_rect: None,").unwrap(),
    }

    writeln!(o, "}};").unwrap();
    writeln!(o).unwrap();
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

fn fmla_to_op(s: &str) -> &str {
    match s {
        "val" => "Val",
        "*/" => "MulDiv",
        "+-" => "AddSub",
        "+/" => "AddDiv",
        "?:" => "IfElse",
        "abs" => "Abs",
        "sqrt" => "Sqrt",
        "min" => "Min",
        "max" => "Max",
        "pin" => "Pin",
        "sin" => "Sin",
        "cos" => "Cos",
        "tan" => "Tan",
        "at2" => "Atan2",
        "cat2" => "CosAtan2",
        "sat2" => "SinAtan2",
        "mod" => "Mod",
        other => panic!("Unknown formula operator: '{other}'"),
    }
}

fn to_screaming_snake(s: &str) -> String {
    let mut out = String::new();
    for (i, c) in s.chars().enumerate() {
        if c.is_uppercase() && i > 0 {
            out.push('_');
        }
        for uc in c.to_uppercase() {
            out.push(uc);
        }
    }
    out
}
