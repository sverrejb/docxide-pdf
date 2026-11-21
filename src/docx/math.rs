//! OMML (§22.1) → `MathZone` trees, and the document's `m:mathPr`.

use crate::model::{FracKind, MathNode, MathRow, MathRun, MathScript, MathSettings, MathStyle};

use super::{MATH_NS, math_bool, math_child, math_children, math_run_text, math_val, revision_of};

/// Builds math rows. `format` resolves a math run's `w:rPr` (its colour)
/// against the paragraph's formatting; `revisions` says whether deleted math
/// shows and in which author colour.
pub(super) struct MathParser<'f> {
    pub(super) format: &'f dyn Fn(Option<roxmltree::Node>) -> crate::model::Run,
    pub(super) revisions: &'f super::Revisions,
}

/// A math run's `w:rPr`, also inside the revision wrapper Word saves revised
/// math runs with (`<m:r><w:ins><w:rPr/><m:t/></w:ins></m:r>`).
pub(super) fn run_rpr<'a>(r: roxmltree::Node<'a, 'a>) -> Option<roxmltree::Node<'a, 'a>> {
    math_children(r, true).find(|c| c.has_tag_name((super::WML_NS, "rPr")))
}

fn flag(node: roxmltree::Node, pr: &str, name: &str) -> bool {
    math_bool(node, pr, name).unwrap_or(false)
}

/// A character property's first character, `default` when absent; an empty
/// value (`m:endChr m:val=""`) is no character.
fn chr(node: roxmltree::Node, pr: &str, name: &str, default: char) -> Option<char> {
    match math_child(node, pr).and_then(|p| math_child(p, name)) {
        Some(n) => n.attribute((MATH_NS, "val")).and_then(|v| v.chars().next()),
        None => Some(default),
    }
}

impl MathParser<'_> {
    pub(super) fn row(&self, node: roxmltree::Node) -> MathRow {
        math_children(node, self.revisions.markup)
            .filter(|c| c.tag_name().namespace() == Some(MATH_NS))
            .flat_map(|c| self.node(c))
            .collect()
    }

    fn arg(&self, node: roxmltree::Node, name: &str) -> MathRow {
        math_child(node, name).map_or_else(Vec::new, |c| self.row(c))
    }

    /// An argument that may be absent; present but empty, it still keeps its
    /// room (see `math::Ctx::row`).
    fn opt_arg(&self, node: roxmltree::Node, name: &str) -> Option<MathRow> {
        math_child(node, name).map(|c| self.row(c))
    }

    fn node(&self, n: roxmltree::Node) -> Vec<MathNode> {
        let one = |m| vec![m];
        match n.tag_name().name() {
            "r" => self.run(n).map_or_else(Vec::new, one),
            "f" => one(MathNode::Frac {
                kind: match math_val(n, "fPr", "type") {
                    Some("skw") => FracKind::Skewed,
                    Some("lin") => FracKind::Linear,
                    Some("noBar") => FracKind::NoBar,
                    _ => FracKind::Bar,
                },
                num: self.arg(n, "num"),
                den: self.arg(n, "den"),
            }),
            "sSub" | "sSup" | "sSubSup" | "sPre" => one(MathNode::Scripts {
                base: self.arg(n, "e"),
                sub: self.opt_arg(n, "sub"),
                sup: self.opt_arg(n, "sup"),
                pre: n.tag_name().name() == "sPre",
            }),
            "rad" => one(MathNode::Rad {
                // Without degHide even an absent degree keeps its room
                // (Word, probe `mp26_empty`).
                degree: (!flag(n, "radPr", "degHide")).then(|| self.arg(n, "deg")),
                body: self.arg(n, "e"),
            }),
            "nary" => {
                // A hidden sub's content shows before the sup's, a hidden
                // sup's at the start of the operand (probe `mp23_nary`).
                let mut sup = self.arg(n, "sup");
                let mut body = self.arg(n, "e");
                let sub = if flag(n, "naryPr", "subHide") {
                    sup.splice(0..0, self.arg(n, "sub"));
                    None
                } else {
                    Some(self.arg(n, "sub"))
                };
                let sup = if flag(n, "naryPr", "supHide") {
                    body.splice(0..0, sup);
                    None
                } else {
                    Some(sup)
                };
                one(MathNode::Nary {
                    chr: chr(n, "naryPr", "chr", '∫').unwrap_or('∫'),
                    under_over: math_val(n, "naryPr", "limLoc").map(|v| v == "undOvr"),
                    sub,
                    sup,
                    grow: flag(n, "naryPr", "grow"),
                    body,
                })
            }
            "d" => one(MathNode::Delim {
                open: chr(n, "dPr", "begChr", '('),
                close: chr(n, "dPr", "endChr", ')'),
                sep: chr(n, "dPr", "sepChr", '|'),
                // m:grow is on unless turned off (§22.1.2.43).
                grow: math_bool(n, "dPr", "grow").unwrap_or(true),
                items: n
                    .children()
                    .filter(|c| c.has_tag_name((MATH_NS, "e")))
                    .map(|c| self.row(c))
                    .collect(),
            }),
            "func" => one(MathNode::Func {
                name: self.arg(n, "fName"),
                body: self.arg(n, "e"),
            }),
            "acc" => one(MathNode::Accent {
                chr: chr(n, "accPr", "chr", '\u{302}').unwrap_or('\u{302}'),
                body: self.arg(n, "e"),
            }),
            "bar" => one(MathNode::Bar {
                top: math_val(n, "barPr", "pos") == Some("top"),
                body: self.arg(n, "e"),
            }),
            // An m:chr without a value draws no character; m:vertJc is top
            // when absent but bot when it has no value (Word).
            "groupChr" => match chr(n, "groupChrPr", "chr", '\u{23DF}') {
                Some(c) => one(MathNode::GroupChr {
                    chr: c,
                    top: math_val(n, "groupChrPr", "pos") == Some("top"),
                    vert_top: math_child(n, "groupChrPr")
                        .and_then(|p| math_child(p, "vertJc"))
                        .is_none_or(|v| v.attribute((MATH_NS, "val")) == Some("top")),
                    body: self.arg(n, "e"),
                }),
                None => self.arg(n, "e"),
            },
            "limLow" | "limUpp" => one(MathNode::Limit {
                upper: n.tag_name().name() == "limUpp",
                body: self.arg(n, "e"),
                limit: self.arg(n, "lim"),
            }),
            "m" => one(MathNode::Matrix {
                rows: n
                    .children()
                    .filter(|c| c.has_tag_name((MATH_NS, "mr")))
                    .map(|mr| {
                        mr.children()
                            .filter(|c| c.has_tag_name((MATH_NS, "e")))
                            .map(|e| self.row(e))
                            .collect()
                    })
                    .collect(),
            }),
            "eqArr" => one(MathNode::EqArr {
                rows: n
                    .children()
                    .filter(|c| c.has_tag_name((MATH_NS, "e")))
                    .map(|e| self.row(e))
                    .collect(),
            }),
            "box" => one(MathNode::Boxed(self.arg(n, "e"))),
            "borderBox" => one(MathNode::BorderBox {
                hide: ["hideTop", "hideBot", "hideLeft", "hideRight"]
                    .map(|h| flag(n, "borderBoxPr", h)),
                body: self.arg(n, "e"),
            }),
            "phant" => one(MathNode::Phantom {
                // m:show defaults on: an m:phant without it is visible.
                show: math_bool(n, "phantPr", "show").unwrap_or(true),
                zero_width: flag(n, "phantPr", "zeroWid"),
                zero_ascent: flag(n, "phantPr", "zeroAsc"),
                zero_descent: flag(n, "phantPr", "zeroDesc"),
                body: self.arg(n, "e"),
            }),
            // Properties and control characters draw nothing.
            name if name.ends_with("Pr") => Vec::new(),
            // Arguments reached directly (m:e and the like): their content.
            _ => self.row(n),
        }
    }

    fn run(&self, r: roxmltree::Node) -> Option<MathNode> {
        let markup = self.revisions.markup;
        let text = math_run_text(r, markup);
        if text.is_empty() {
            return None;
        }
        let fmt = (self.format)(run_rpr(r));
        // Markup view: revised math takes the author's colour, deleted math
        // is struck through and inserted math is not underlined (Word for
        // Mac, corpus redlines 970e9bbbcc and d0281643d2).
        let revision = markup
            .then(|| r.descendants().find(|t| t.has_tag_name((MATH_NS, "t"))))
            .flatten()
            .and_then(revision_of);
        let val = |name| math_val(r, "rPr", name);
        Some(MathNode::Run(MathRun {
            text,
            style: match val("sty") {
                Some("p") => Some(MathStyle::Plain),
                Some("b") => Some(MathStyle::Bold),
                Some("i") => Some(MathStyle::Italic),
                Some("bi") => Some(MathStyle::BoldItalic),
                _ => None,
            },
            script: match val("scr") {
                Some("script") => MathScript::Script,
                Some("fraktur") => MathScript::Fraktur,
                Some("double-struck") => MathScript::DoubleStruck,
                Some("sans-serif") => MathScript::SansSerif,
                Some("monospace") => MathScript::Monospace,
                _ => MathScript::Roman,
            },
            normal: flag(r, "rPr", "nor").then(|| crate::model::TextFont {
                name: fmt.font_name.clone(),
                bold: fmt.bold,
                italic: fmt.italic,
                size: fmt.font_size,
            }),
            color: revision.map_or(fmt.color, |(_, rev)| {
                Some(self.revisions.revision_color(rev))
            }),
            strike: revision.is_some_and(|(deleted, _)| deleted),
        }))
    }
}

/// `m:mathPr` (§22.1.2.62) from settings.xml.
pub(super) fn parse_math_settings(settings_root: roxmltree::Node) -> MathSettings {
    let mut s = MathSettings::default();
    let Some(pr) = settings_root
        .descendants()
        .find(|n| n.has_tag_name((MATH_NS, "mathPr")))
    else {
        return s;
    };
    let val = |name| math_child(pr, name)?.attribute((MATH_NS, "val"));
    if let Some(font) = val("mathFont") {
        s.font = font.to_string();
    }
    if let Some(n) = math_child(pr, "smallFrac") {
        s.small_frac = n
            .attribute((MATH_NS, "val"))
            .is_none_or(super::parse_on_off);
    }
    if let Some(v) = val("intLim") {
        s.int_under_over = v == "undOvr";
    }
    if let Some(v) = val("naryLim") {
        s.nary_under_over = v == "undOvr";
    }
    if let Some(twips) = val("wrapIndent").and_then(|v| v.parse().ok()) {
        s.wrap_indent = super::twips_to_pts(twips);
    }
    s.break_after = val("brkBin") == Some("after");
    s
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parse(xml: &str) -> MathRow {
        parse_in(xml, &crate::docx::Revisions::default())
    }

    fn parse_in(xml: &str, revisions: &crate::docx::Revisions) -> MathRow {
        let doc = format!(
            r#"<m:oMath xmlns:m="{MATH_NS}" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{xml}</m:oMath>"#
        );
        let tree = roxmltree::Document::parse(&doc).unwrap();
        MathParser {
            format: &|_| crate::model::Run::default(),
            revisions,
        }
        .row(tree.root_element())
    }

    /// How Word saves revised math: the revision inside the run, around its
    /// properties and text.
    const REVISED: &str = r#"<m:r><w:del w:id="1" w:author="A"><w:rPr/><m:t>a</m:t></w:del></m:r><m:r><w:ins w:id="2" w:author="A"><w:rPr/><m:t>b</m:t></w:ins></m:r>"#;

    fn runs(row: &MathRow) -> Vec<(String, bool, bool)> {
        row.iter()
            .filter_map(|n| match n {
                MathNode::Run(r) => Some((r.text.clone(), r.strike, r.color.is_some())),
                _ => None,
            })
            .collect()
    }

    #[test]
    fn revised_math_final_view() {
        // Deleted math is gone, inserted math stays, plain.
        assert_eq!(runs(&parse(REVISED)), [("b".into(), false, false)]);
    }

    #[test]
    fn revised_math_markup_view() {
        // Both in the author's colour; only the deletion is struck through.
        let markup = crate::docx::Revisions {
            markup: true,
            ..Default::default()
        };
        assert_eq!(
            runs(&parse_in(REVISED, &markup)),
            [("a".into(), true, true), ("b".into(), false, true)]
        );
    }

    #[test]
    fn fraction_under_radical() {
        let row = parse(
            "<m:r><m:t>T=2π</m:t></m:r><m:rad><m:radPr/><m:deg><m:r><m:t>2</m:t></m:r></m:deg>\
             <m:e><m:f><m:fPr/><m:num><m:r><m:t>l</m:t></m:r></m:num><m:den><m:r><m:t>g</m:t></m:r></m:den></m:f></m:e></m:rad>",
        );
        assert_eq!(row.len(), 2);
        let MathNode::Rad { degree, body } = &row[1] else {
            panic!("{row:?}")
        };
        assert!(degree.is_some());
        assert!(
            matches!(&body[0], MathNode::Frac { kind: FracKind::Bar, num, den } if num.len() == 1 && den.len() == 1)
        );
    }

    #[test]
    fn hidden_degree_and_delimiter_defaults() {
        let row = parse(
            r#"<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>
               <m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val=""/></m:dPr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:d>"#,
        );
        assert!(matches!(&row[0], MathNode::Rad { degree: None, .. }));
        let MathNode::Delim {
            open,
            close,
            sep,
            grow,
            items,
        } = &row[1]
        else {
            panic!("{row:?}")
        };
        assert_eq!(
            (*open, *close, *sep, *grow, items.len()),
            (Some('['), None, Some('|'), true, 2)
        );
    }

    #[test]
    fn run_styles() {
        let row = parse(
            r#"<m:r><m:rPr><m:sty m:val="bi"/></m:rPr><m:t>T</m:t></m:r><m:r><m:rPr><m:scr m:val="double-struck"/><m:sty m:val="p"/></m:rPr><m:t>R</m:t></m:r>"#,
        );
        let style = |n: &MathNode| match n {
            MathNode::Run(r) => (r.style, r.script),
            _ => panic!(),
        };
        assert_eq!(
            style(&row[0]),
            (Some(MathStyle::BoldItalic), MathScript::Roman)
        );
        assert_eq!(
            style(&row[1]),
            (Some(MathStyle::Plain), MathScript::DoubleStruck)
        );
    }
}
