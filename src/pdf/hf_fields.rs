//! Header and footer field results: PAGE, NUMPAGES, SECTIONPAGES, STYLEREF and IF.

use crate::model::{FieldCode, IfPart, Run, styleref_key};

use super::table_layout::HfSubstitution;

/// `runs` with their field codes replaced by this page's values.
pub(super) fn substitute_hf_runs(runs: &[Run], sub: &HfSubstitution) -> Vec<Run> {
    runs.iter()
        .map(|run| {
            let mut r = run.clone();
            if let Some(ref fc) = run.field_code {
                r.field_code = None;
                r.text = sub.eval(fc).unwrap_or_else(|| run.text.clone());
                if matches!(fc, FieldCode::Page(_) | FieldCode::NumPages(_)) {
                    r.caps = false;
                    r.small_caps = false;
                }
            }
            r
        })
        .collect()
}

/// A PAGE, NUMPAGES or SECTIONPAGES value in its numFmt (None = decimal).
pub(super) fn page_field_text(n: usize, fmt: Option<&str>) -> String {
    fmt.map_or_else(
        || n.to_string(),
        |fmt| crate::docx::numbering::format_number(n as u32, fmt),
    )
}

impl HfSubstitution<'_> {
    /// None keeps the cached result.
    fn eval(&self, fc: &FieldCode) -> Option<String> {
        Some(match fc {
            FieldCode::Page(fmt) => page_field_text(self.page_num, fmt.or(self.page_num_format)),
            FieldCode::NumPages(fmt) => page_field_text(self.total_pages, *fmt),
            FieldCode::SectionPages(fmt) => page_field_text(self.section_pages, *fmt),
            FieldCode::StyleRef { name, form } => self
                .styleref_values
                .get(&styleref_key(name, *form))
                .cloned()
                .unwrap_or_default(),
            FieldCode::PageRef(_) | FieldCode::Seq { .. } => return None,
            FieldCode::If(parts) => return self.eval_if(parts),
        })
    }

    /// `IF expr1 op expr2 "true" "false"` after substituting the nested
    /// fields.
    fn eval_if(&self, parts: &[IfPart]) -> Option<String> {
        let mut instr = String::new();
        for part in parts {
            match part {
                IfPart::Text(t) => instr.push_str(t),
                IfPart::Field(fc) => instr.push_str(&self.eval(fc)?),
            }
        }
        let args = field_args(instr.trim_start().get(2..)?);
        let [a, op, b, t, f] = args.as_slice() else {
            return None;
        };
        let ord = match (a.trim().parse::<f64>(), b.trim().parse::<f64>()) {
            (Ok(x), Ok(y)) => x.partial_cmp(&y)?,
            _ => a.cmp(b),
        };
        let holds = match op.as_str() {
            "=" => ord.is_eq(),
            "<>" => ord.is_ne(),
            "<" => ord.is_lt(),
            "<=" => ord.is_le(),
            ">" => ord.is_gt(),
            ">=" => ord.is_ge(),
            _ => return None,
        };
        Some(if holds { t } else { f }.clone())
    }
}

/// Field arguments: whitespace-separated, double quotes group (and are
/// dropped), comparison operators stand alone even without spaces.
fn field_args(s: &str) -> Vec<String> {
    let mut args = Vec::new();
    let mut chars = s.chars().peekable();
    while let Some(&c) = chars.peek() {
        if c.is_whitespace() {
            chars.next();
        } else if c == '"' {
            chars.next();
            args.push(chars.by_ref().take_while(|&c| c != '"').collect());
        } else if "=<>".contains(c) {
            let mut op = String::new();
            while let Some(&c) = chars.peek().filter(|c| "=<>".contains(**c)) {
                op.push(c);
                chars.next();
            }
            args.push(op);
        } else {
            let mut word = String::new();
            while let Some(&c) = chars
                .peek()
                .filter(|c| !c.is_whitespace() && !"\"=<>".contains(**c))
            {
                word.push(c);
                chars.next();
            }
            args.push(word);
        }
    }
    args
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::model::StyleRefForm;
    use std::collections::HashMap;

    #[test]
    fn if_over_stylerefs_takes_the_page_value() {
        let sr = |form| {
            IfPart::Field(FieldCode::StyleRef {
                name: "CharPartNo".into(),
                form,
            })
        };
        let text = |t: &str| IfPart::Text(t.into());
        // IF {STYLEREF CharPartNo \n} = 0 "{STYLEREF CharPartNo}" "Part {STYLEREF CharPartNo \n}"
        let parts = [
            text(" IF "),
            sr(StyleRefForm::Label),
            text(" = 0 \""),
            sr(StyleRefForm::Text),
            text("\" \"Part "),
            sr(StyleRefForm::Label),
            text("\""),
        ];
        let mut styleref_values = HashMap::new();
        let eval = |values: &HashMap<String, String>| {
            HfSubstitution {
                page_num: 1,
                total_pages: 1,
                section_pages: 1,
                styleref_values: values,
                page_num_format: None,
            }
            .eval(&FieldCode::If(parts.to_vec()))
        };
        assert_eq!(eval(&styleref_values), None);
        let key = |form| styleref_key("charpartno", form);
        styleref_values.insert(key(StyleRefForm::Text), "Part 3".to_string());
        styleref_values.insert(key(StyleRefForm::Label), "0".to_string());
        assert_eq!(eval(&styleref_values).as_deref(), Some("Part 3"));
        styleref_values.insert(key(StyleRefForm::Label), "4".to_string());
        assert_eq!(eval(&styleref_values).as_deref(), Some("Part 4"));
        assert_eq!(field_args("a<>\"b c\" 2"), ["a", "<>", "b c", "2"]);
    }

    #[test]
    fn page_count_results_drop_caps_in_headers() {
        let cased = |field_code| Run {
            text: "Z".into(),
            caps: true,
            small_caps: true,
            underline: true,
            field_code: Some(field_code),
            ..Default::default()
        };
        let runs = [
            cased(FieldCode::Page(Some("lowerLetter"))),
            cased(FieldCode::NumPages(Some("lowerLetter"))),
            cased(FieldCode::SectionPages(Some("lowerLetter"))),
        ];
        let sub = HfSubstitution {
            page_num: 1,
            total_pages: 2,
            section_pages: 2,
            styleref_values: &HashMap::new(),
            page_num_format: None,
        };
        let out = substitute_hf_runs(&runs, &sub);
        assert_eq!(
            out.iter().map(|r| r.text.as_str()).collect::<Vec<_>>(),
            ["a", "b", "b"]
        );
        assert!(!out[0].caps && !out[0].small_caps && out[0].underline);
        assert!(!out[1].caps && !out[1].small_caps);
        assert!(out[2].caps && out[2].small_caps);
    }
}
