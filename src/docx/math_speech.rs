//! Office Math (OMML) read aloud, for the `/Alt` of the Formula a math zone is
//! tagged as (PDF/UA 7.7-1): Word's export speaks `T=2π√(l/g)` as "cap T
//! equals 2 pi root 2 of l over g". The rules follow what Word writes.
// ponytail: the structures and symbols the corpus uses; matrices and accents
// are read as their contents in order

use super::{
    MATH_NS, find_children, math_bool, math_child, math_children, math_run_text, math_val,
};

/// The spoken form of an `m:oMath` zone; empty when it has nothing to say.
pub(super) fn speak(omath: roxmltree::Node) -> String {
    let mut words = Vec::new();
    speak_into(omath, &mut words);
    words.join(" ")
}

fn speak_into(node: roxmltree::Node, out: &mut Vec<String>) {
    for c in math_children(node, false).filter(|n| n.tag_name().namespace() == Some(MATH_NS)) {
        let part = |name: &str, out: &mut Vec<String>| {
            if let Some(n) = math_child(c, name) {
                speak_into(n, out);
            }
        };
        // `word` and the part after it, when the part has text (an empty
        // `m:deg` or `m:sub` says nothing); returns whether it spoke.
        let said = |word: &str, name: &str, out: &mut Vec<String>| {
            let filled = math_child(c, name).filter(|n| n.descendants().any(|t| t.is_text()));
            if let Some(n) = filled {
                out.push(word.into());
                speak_into(n, out);
            }
            filled.is_some()
        };
        match c.tag_name().name() {
            "r" => speak_run(c, out),
            "f" => {
                part("num", out);
                out.push("over".into());
                part("den", out);
            }
            "rad" => {
                let hidden = math_bool(c, "radPr", "degHide").unwrap_or(false);
                if hidden || !said("root", "deg", out) {
                    out.push("square root of".into());
                } else {
                    out.push("of".into());
                }
                part("e", out);
            }
            "sSub" | "sSup" | "sSubSup" => {
                part("e", out);
                said("sub", "sub", out);
                said("to the", "sup", out);
            }
            "d" => {
                let beg = math_val(c, "dPr", "begChr").unwrap_or("(");
                let end = math_val(c, "dPr", "endChr").unwrap_or(")");
                out.extend(delimiter(beg, "open"));
                for (i, e) in find_children(c, "e", MATH_NS).enumerate() {
                    if i > 0 {
                        out.push("comma".into());
                    }
                    speak_into(e, out);
                }
                out.extend(delimiter(end, "close"));
            }
            "nary" => {
                let op = match math_val(c, "naryPr", "chr").unwrap_or("∫") {
                    "∑" => "sum",
                    "∏" => "product",
                    "∮" => "contour integral",
                    "∫" => "integral",
                    other => other,
                };
                out.push(op.into());
                said("from", "sub", out);
                said("to", "sup", out);
                out.push("of".into());
                part("e", out);
            }
            // One equation after another, with a pause between them.
            "eqArr" => {
                for (i, e) in find_children(c, "e", MATH_NS).enumerate() {
                    if i > 0
                        && let Some(last) = out.last_mut()
                    {
                        last.push(',');
                    }
                    speak_into(e, out);
                }
            }
            name if name.ends_with("Pr") => {}
            _ => speak_into(c, out),
        }
    }
}

fn delimiter(chr: &str, side: &str) -> Option<String> {
    let name = match chr {
        "" => return None,
        "(" | ")" => "paren",
        "[" | "]" => "bracket",
        "{" | "}" => "brace",
        "|" => return Some("vertical bar".into()),
        other => return Some(other.into()),
    };
    Some(format!("{side} {name}"))
}

/// One `m:r`: its characters with the run's math style (`m:sty`), which Word
/// says before each letter ("bold italic cap T") and digit ("bold 1").
fn speak_run(r: roxmltree::Node, out: &mut Vec<String>) {
    // The formula as it stands: deleted math is not spoken.
    let text = math_run_text(r, false);
    let sty = math_val(r, "rPr", "sty");
    // Plain (upright) text is a word: a function name or a unit.
    if sty == Some("p")
        || math_child(r, "rPr")
            .and_then(|p| math_child(p, "nor"))
            .is_some()
    {
        out.extend(text.split_whitespace().map(str::to_string));
        return;
    }
    let bold = matches!(sty, Some("b" | "bi"));
    let mut number = String::new();
    for ch in text.chars() {
        if ch.is_ascii_digit() && !bold {
            number.push(ch);
            continue;
        }
        if !number.is_empty() {
            out.push(std::mem::take(&mut number));
        }
        // `&` only marks an alignment point.
        if ch.is_whitespace() || ch == '&' {
            continue;
        }
        let (word, prefix) = if let Some(w) = symbol_word(ch) {
            (w.to_string(), None)
        } else if let Some(w) = letter_word(ch) {
            let style = if sty == Some("bi") {
                "bold italic"
            } else {
                "bold"
            };
            (w, bold.then_some(style))
        } else {
            (
                ch.to_string(),
                (bold && ch.is_ascii_digit()).then_some("bold"),
            )
        };
        out.extend(prefix.map(str::to_string));
        out.push(word);
    }
    if !number.is_empty() {
        out.push(number);
    }
}

/// A Latin or Greek letter as itself or its name, capitals with "cap" before.
fn letter_word(ch: char) -> Option<String> {
    const GREEK: [&str; 25] = [
        "alpha",
        "beta",
        "gamma",
        "delta",
        "epsilon",
        "zeta",
        "eta",
        "theta",
        "iota",
        "kappa",
        "lambda",
        "mu",
        "nu",
        "xi",
        "omicron",
        "pi",
        "rho",
        "final sigma",
        "sigma",
        "tau",
        "upsilon",
        "phi",
        "chi",
        "psi",
        "omega",
    ];
    let (name, cap) = match ch {
        'a'..='z' => (ch.to_string(), false),
        'A'..='Z' => (ch.to_string(), true),
        'α'..='ω' => (GREEK[ch as usize - 'α' as usize].to_string(), false),
        // U+03A2 is unassigned, so the capitals line up with GREEK.
        'Α'..='Ω' if ch != '\u{3A2}' => (GREEK[ch as usize - 'Α' as usize].to_string(), true),
        'ϑ' => ("script theta".into(), false),
        'ϕ' => ("phi".into(), false),
        'ϵ' => ("epsilon".into(), false),
        _ => return None,
    };
    Some(if cap { format!("cap {name}") } else { name })
}

fn symbol_word(ch: char) -> Option<&'static str> {
    Some(match ch {
        '=' => "equals",
        '+' => "plus",
        '-' | '−' => "minus",
        '±' => "plus or minus",
        '∓' => "minus or plus",
        '×' => "times",
        '·' | '⋅' | '∙' => "dot",
        '÷' => "divided by",
        '/' => "slash",
        '<' => "less than",
        '>' => "greater than",
        '≤' => "less than or equal to",
        '≥' => "greater than or equal to",
        '≪' => "much less than",
        '≫' => "much greater than",
        '≠' => "not equal to",
        '≈' => "almost equal to",
        '≡' => "identical to",
        '∝' => "proportional to",
        '∞' => "infinity",
        '∆' => "increment",
        '∂' => "partial",
        '∇' => "nabla",
        '→' => "right arrow",
        '√' => "square root",
        '%' => "percent",
        '°' => "degrees",
        '′' => "prime",
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn spoken(omml: &str) -> String {
        let xml = format!(
            r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">{omml}</m:oMath>"#
        );
        let doc = roxmltree::Document::parse(&xml).unwrap();
        speak(doc.root_element())
    }

    #[test]
    fn equation_array_rows_pause_and_skip_alignment_marks() {
        assert_eq!(
            spoken(
                "<m:eqArr><m:e><m:r><m:t>x&amp;=1</m:t></m:r></m:e><m:e><m:r><m:t>y&amp;=2</m:t></m:r></m:e></m:eqArr>"
            ),
            "x equals 1, y equals 2"
        );
    }

    // Expected strings are Word's own /Alt in pendulum_mechanics_oscillation_lab.
    #[test]
    fn speaks_like_word() {
        assert_eq!(
            spoken(
                "<m:r><m:t>T=2π</m:t></m:r><m:rad><m:deg><m:r><m:t>2</m:t></m:r></m:deg>\
                 <m:e><m:f><m:num><m:r><m:t>l</m:t></m:r></m:num>\
                 <m:den><m:r><m:t>g</m:t></m:r></m:den></m:f></m:e></m:rad>"
            ),
            "cap T equals 2 pi root 2 of l over g"
        );
        assert_eq!(
            spoken(
                "<m:r><m:t>ϑ≪</m:t></m:r><m:f><m:num><m:r><m:t>π</m:t></m:r></m:num>\
                 <m:den><m:r><m:t>2</m:t></m:r></m:den></m:f>"
            ),
            "script theta much less than pi over 2"
        );
        let bi = r#"<m:rPr><m:sty m:val="bi"/></m:rPr>"#;
        let b = r#"<m:rPr><m:sty m:val="b"/></m:rPr>"#;
        assert_eq!(
            spoken(&format!("<m:r>{bi}<m:t>l±∆l</m:t></m:r>")),
            "bold italic l plus or minus increment bold italic l"
        );
        let t10 = format!(
            "<m:sSub><m:e><m:r>{bi}<m:t>T</m:t></m:r></m:e><m:sub><m:r>{bi}<m:t>10</m:t></m:r></m:sub></m:sSub>"
        );
        assert_eq!(
            spoken(&format!(
                "{t10}<m:r>{bi}<m:t>±</m:t></m:r><m:r>{b}<m:t>Δ</m:t></m:r>{t10}"
            )),
            "bold italic cap T sub bold 1 bold 0 plus or minus bold cap delta \
             bold italic cap T sub bold 1 bold 0"
        );
        assert_eq!(
            spoken(&format!(
                "<m:sSubSup><m:e><m:r>{bi}<m:t>T</m:t></m:r></m:e><m:sub><m:r>{bi}<m:t>1</m:t></m:r></m:sub>\
                 <m:sup><m:r>{bi}<m:t>2</m:t></m:r></m:sup></m:sSubSup>"
            )),
            "bold italic cap T sub bold 1 to the bold 2"
        );
        assert_eq!(spoken("<m:r><m:t>l </m:t></m:r>"), "l");
        assert_eq!(spoken("<m:r><m:t> </m:t></m:r>"), "");
    }
}
