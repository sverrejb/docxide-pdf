//! Field instructions: the codes the renderer re-evaluates, their switches and arguments.

use crate::model::{FieldCode, FormCheckbox, SeqStep, StyleRefForm};

use super::styles::half_points;
use super::{wml, wml_bool};

/// A field the renderer re-evaluates per page, from its instruction.
pub(super) fn parse_field_code(instr: &str) -> Option<FieldCode> {
    let keyword = instr.split_whitespace().next()?;
    if keyword.eq_ignore_ascii_case("PAGE") {
        Some(FieldCode::Page(number_format_switch(instr)))
    } else if keyword.eq_ignore_ascii_case("NUMPAGES") {
        Some(FieldCode::NumPages(number_format_switch(instr)))
    } else if keyword.eq_ignore_ascii_case("SECTIONPAGES") {
        Some(FieldCode::SectionPages(number_format_switch(instr)))
    } else if keyword.eq_ignore_ascii_case("STYLEREF") {
        parse_styleref_arg(instr).map(|name| FieldCode::StyleRef {
            name,
            // `\t` only qualifies a number switch.
            form: match (has_switch(instr, "\\n"), has_switch(instr, "\\t")) {
                (true, true) => StyleRefForm::Number,
                (true, false) => StyleRefForm::Label,
                _ => StyleRefForm::Text,
            },
        })
    } else if keyword.eq_ignore_ascii_case("PAGEREF") {
        instr
            .split_whitespace()
            .nth(1)
            .map(|s| FieldCode::PageRef(s.to_string()))
    } else if keyword.eq_ignore_ascii_case("SEQ") {
        let name = instr
            .split_whitespace()
            .nth(1)
            .filter(|w| !w.starts_with('\\'))?;
        let step = if let Some(n) = switch_args(instr, "\\r").next() {
            SeqStep::Restart(n.parse().ok()?)
        } else if has_switch(instr, "\\c") {
            SeqStep::Repeat
        } else if let Some(n) = switch_args(instr, "\\s").next() {
            SeqStep::Heading(n.parse().ok()?)
        } else {
            SeqStep::Next
        };
        Some(FieldCode::Seq {
            name: name.to_string(),
            format: number_format_switch(instr),
            step,
            hidden: has_switch(instr, "\\h"),
        })
    } else {
        None
    }
}

/// `EQ \o\ac(□,√)` — Word's overstrike equation, how 方案模板 draw a checked
/// checkbox: the arguments are painted on top of each other and the first sets
/// the advance. The field carries **no cached result** (Word computes it when
/// the document is laid out), so without this the whole checkbox disappeared.
///
/// Only the overstrike switch (`\o`) is understood — fractions (`\f`), radicals
/// (`\r`) and the rest keep no result, same as before. Checkbox shapes map to
/// their precomposed characters (`□`/`☐` + `√` → `☑`); any other overstrike
/// keeps the base character, which is the part that stays legible when the
/// accent is dropped.
pub(super) fn parse_eq_overstrike(instr: &str) -> Option<String> {
    let keyword = instr.split_whitespace().next()?;
    if !keyword.eq_ignore_ascii_case("EQ") {
        return None;
    }
    let lower = instr.to_ascii_lowercase();
    let over = lower.find("\\o")?;
    let open = instr[over..].find('(')? + over;
    let close = instr.rfind(')')?;
    if close <= open {
        return None;
    }
    let parts: Vec<&str> = instr[open + 1..close].split(',').map(str::trim).collect();
    let base = parts.first().copied().unwrap_or("");
    let marks: String = parts.iter().skip(1).copied().collect();
    if base.chars().count() == 1 && matches!(base, "□" | "☐") && marks.contains('√') {
        return Some("☑".to_string());
    }
    (!base.is_empty()).then(|| base.to_string())
}

/// The words of a field instruction, each with whether it was quoted: a
/// quoted argument is one word (without its quotes) and never a switch.
fn field_words(instr: &str) -> impl Iterator<Item = (&str, bool)> + '_ {
    let mut rest = instr;
    std::iter::from_fn(move || {
        rest = rest.trim_start();
        if let Some(quoted) = rest.strip_prefix('"') {
            let mut escaped = false;
            let close = quoted.find(|c| {
                let closes = c == '"' && !escaped;
                escaped = c == '\\' && !escaped;
                closes
            });
            let (word, after) = close.map_or((quoted, ""), |i| (&quoted[..i], &quoted[i + 1..]));
            rest = after;
            return Some((word, true));
        }
        let end = rest
            .find(|c: char| c.is_whitespace() || c == '"')
            .unwrap_or(rest.len());
        let (word, after) = rest.split_at(end);
        rest = after;
        (!word.is_empty()).then_some((word, false))
    })
}

fn is_switch((word, quoted): (&str, bool), switch: &str) -> bool {
    !quoted && word.eq_ignore_ascii_case(switch)
}

pub(super) fn has_switch(instr: &str, switch: &str) -> bool {
    field_words(instr).any(|w| is_switch(w, switch))
}

/// The word after each `switch` in a field instruction (`\* ROMAN \*
/// MERGEFORMAT` → "ROMAN", "MERGEFORMAT").
pub(super) fn switch_args<'a>(
    instr: &'a str,
    switch: &'a str,
) -> impl Iterator<Item = &'a str> + 'a {
    field_words(instr)
        .zip(field_words(instr).skip(1))
        .filter(move |&(word, _)| is_switch(word, switch))
        .map(|(_, (arg, _))| arg)
}

/// `\* ROMAN`, `\* alphabetic` …: the numFmt a numeric field is shown in.
/// Like Word, the case of the format word picks upper or lower case.
fn number_format_switch(instr: &str) -> Option<&'static str> {
    switch_args(instr, "\\*").find_map(|fmt| {
        let upper = fmt.chars().all(|c| c.is_ascii_uppercase());
        match fmt.to_ascii_lowercase().as_str() {
            "roman" if upper => Some("upperRoman"),
            "roman" => Some("lowerRoman"),
            "alphabetic" if upper => Some("upperLetter"),
            "alphabetic" => Some("lowerLetter"),
            "arabic" => Some("decimal"),
            _ => None,
        }
    })
}

/// The display text of a `MACROBUTTON` instruction once its macro name is
/// complete: everything after the name. The field has no separator — the text
/// or graphic after the macro name is its value (ISO/IEC 29500-1 §17.16.5.34).
pub(super) fn macrobutton_display(instr: &str) -> Option<&str> {
    let (keyword, rest) = instr.trim_start().split_once(char::is_whitespace)?;
    if !keyword.eq_ignore_ascii_case("MACROBUTTON") {
        return None;
    }
    rest.trim_start()
        .split_once(char::is_whitespace)
        .map(|(_, display)| display)
}

/// A legacy check box form field (`w:ffData/w:checkBox` on its begin
/// `fldChar`); `w:checked` overrides `w:default`.
pub(super) fn parse_checkbox(fld_char: roxmltree::Node, run_size: f32) -> Option<FormCheckbox> {
    let cb = wml(wml(fld_char, "ffData")?, "checkBox")?;
    let size = half_points(cb, "size").unwrap_or(run_size);
    let checked = wml_bool(cb, "checked")
        .or_else(|| wml_bool(cb, "default"))
        .unwrap_or(false);
    Some(FormCheckbox { size, checked })
}

fn parse_styleref_arg(instr: &str) -> Option<String> {
    let trimmed = instr.trim();
    let kw = trimmed.split_whitespace().next()?;
    if !kw.eq_ignore_ascii_case("styleref") {
        return None;
    }
    quoted_or_word(&trimmed[kw.len()..]).map(str::to_string)
}

/// A field argument: a quoted string, or one bare word that is no switch.
pub(super) fn quoted_or_word(rest: &str) -> Option<&str> {
    let rest = rest.trim_start();
    match rest.strip_prefix('"') {
        Some(quoted) => quoted.split('"').next(),
        None => rest
            .split_whitespace()
            .next()
            .filter(|w| !w.starts_with('\\')),
    }
}

/// The instruction of a TOC field.
pub(super) fn is_toc_instr(instr: &str) -> bool {
    instr
        .split_whitespace()
        .next()
        .is_some_and(|k| k.eq_ignore_ascii_case("TOC"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::docx::WML_NS;

    #[test]
    fn quoted_text_holds_no_switch() {
        let instr = r#" IF 1 = 1 "a \* MERGEFORMAT b" "\h"\* Upper \* "x" "#;
        assert_eq!(
            switch_args(instr, "\\*").collect::<Vec<_>>(),
            ["Upper", "x"]
        );
        assert!(!has_switch(instr, "\\h"));
        let escaped = r#" IF 1 = 1 "a \" \* MERGEFORMAT \" b" "c" "#;
        assert!(!has_switch(escaped, "\\*"));
    }

    #[test]
    fn seq_field_reads_its_name_and_switches() {
        let Some(FieldCode::Seq {
            name,
            format,
            step,
            hidden,
        }) = parse_field_code(" SEQ 自定义 \\* ROMAN \\r 4 \\h ")
        else {
            panic!("not a SEQ field");
        };
        assert_eq!(
            (name.as_str(), format, step, hidden),
            ("自定义", Some("upperRoman"), SeqStep::Restart(4), true)
        );
        assert!(matches!(
            parse_field_code(" SEQ Figure \\c "),
            Some(FieldCode::Seq {
                step: SeqStep::Repeat,
                ..
            })
        ));
        assert!(matches!(
            parse_field_code(" SEQ Figure \\* ARABIC \\s 1 "),
            Some(FieldCode::Seq {
                step: SeqStep::Heading(1),
                ..
            })
        ));
        assert!(parse_field_code(" SEQ \\* ARABIC ").is_none());
    }

    #[test]
    fn eq_overstrike_checkbox_and_other_fields() {
        assert_eq!(parse_eq_overstrike(r"EQ \o\ac(□,√)").as_deref(), Some("☑"));
        assert_eq!(
            parse_eq_overstrike(r" eq \O\ac( ☐ , √ )").as_deref(),
            Some("☑")
        );
        assert_eq!(parse_eq_overstrike(r"EQ \o(a,´)").as_deref(), Some("a"));
        assert_eq!(parse_eq_overstrike(r"EQ \f(1,2)"), None);
        assert_eq!(parse_eq_overstrike(r"PAGE \* MERGEFORMAT"), None);
        assert_eq!(parse_eq_overstrike(r"EQ \o()"), None);
    }

    #[test]
    fn macrobutton_value_follows_the_macro_name() {
        assert_eq!(macrobutton_display(" MACROBUTTON AutoNew "), Some(""));
        assert_eq!(
            macrobutton_display(" MACROBUTTON NoMacro [Click here] "),
            Some("[Click here] ")
        );
        assert_eq!(macrobutton_display(" MACROBUTTON AutoNew"), None);
        assert_eq!(macrobutton_display(r" PAGE \* MERGEFORMAT "), None);
    }

    #[test]
    fn checkbox_size_and_state() {
        let parse = |check_box: &str| {
            let xml = format!(
                r#"<w:fldChar xmlns:w="{WML_NS}" w:fldCharType="begin"><w:ffData><w:checkBox>{check_box}</w:checkBox></w:ffData></w:fldChar>"#
            );
            let doc = roxmltree::Document::parse(&xml).unwrap();
            parse_checkbox(doc.root_element(), 11.0)
        };
        let cb = |size, checked| Some(FormCheckbox { size, checked });
        assert_eq!(
            parse(r#"<w:size w:val="20"/><w:default w:val="0"/>"#),
            cb(10.0, false)
        );
        assert_eq!(
            parse(r#"<w:sizeAuto/><w:default w:val="1"/>"#),
            cb(11.0, true)
        );
        assert_eq!(
            parse(r#"<w:sizeAuto/><w:default w:val="1"/><w:checked w:val="0"/>"#),
            cb(11.0, false)
        );
        assert_eq!(parse(r#"<w:sizeAuto/><w:checked/>"#), cb(11.0, true));
        let text_field = format!(
            r#"<w:fldChar xmlns:w="{WML_NS}" w:fldCharType="begin"><w:ffData><w:textInput/></w:ffData></w:fldChar>"#
        );
        let doc = roxmltree::Document::parse(&text_field).unwrap();
        assert_eq!(parse_checkbox(doc.root_element(), 11.0), None);
    }
}
