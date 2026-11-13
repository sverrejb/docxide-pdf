//! Chart axis ranges, tick and number formatting, and label measuring and wrapping.

use crate::fonts::FontEntry;

fn ceil_nice(val: f32) -> f32 {
    if val <= 0.0 {
        return 1.0;
    }
    let mag = 10.0f32.powf(val.log10().floor());
    let norm = val / mag;
    let nice = if norm <= 1.5 {
        1.0
    } else if norm <= 3.0 {
        2.0
    } else if norm <= 7.0 {
        5.0
    } else {
        10.0
    };
    nice * mag
}

fn nice_tick_step(max_val: f32, target_ticks: usize) -> f32 {
    let raw_step = max_val / target_ticks as f32;
    ceil_nice(raw_step)
}

fn text_width_approx(text: &str, font_size: f32) -> f32 {
    text.len() as f32 * font_size * 0.5
}

pub(super) fn text_width(text: &str, font_size: f32, font: Option<&FontEntry>) -> f32 {
    match font {
        Some(f) => f.word_width(text, font_size, false),
        None => text_width_approx(text, font_size),
    }
}

pub(super) fn format_tick_label(val: f32, step: f32, format_code: Option<&str>) -> String {
    match format_code {
        Some(fmt) => apply_number_format(val, fmt),
        None => {
            if step.fract() == 0.0 {
                format!("{}", val as i64)
            } else {
                format!("{:.1}", val)
            }
        }
    }
}

pub(super) fn apply_number_format(val: f32, fmt: &str) -> String {
    let has_comma = fmt.contains(',');
    let has_percent = fmt.contains('%');
    let effective_val = if has_percent { val * 100.0 } else { val };

    let decimals = if let Some(dot_pos) = fmt.rfind('.') {
        fmt[dot_pos + 1..]
            .chars()
            .take_while(|&c| c == '0' || c == '#')
            .count()
    } else {
        0
    };

    let formatted = format!("{:.prec$}", effective_val as f64, prec = decimals);

    let result = if has_comma {
        insert_thousands_separator(&formatted)
    } else {
        formatted
    };

    if has_percent {
        format!("{}%", result)
    } else {
        result
    }
}

fn insert_thousands_separator(s: &str) -> String {
    let (int_part, dec_part) = match s.find('.') {
        Some(pos) => (&s[..pos], Some(&s[pos..])),
        None => (s, None),
    };
    let negative = int_part.starts_with('-');
    let digits = if negative { &int_part[1..] } else { int_part };
    let mut result = String::new();
    for (i, ch) in digits.chars().rev().enumerate() {
        if i > 0 && i % 3 == 0 {
            result.push(',');
        }
        result.push(ch);
    }
    let mut out: String = result.chars().rev().collect();
    if negative {
        out.insert(0, '-');
    }
    if let Some(d) = dec_part {
        out.push_str(d);
    }
    out
}

pub(super) fn wrap_label(
    text: &str,
    max_width: f32,
    font_size: f32,
    font: Option<&FontEntry>,
) -> Vec<String> {
    let full_w = text_width(text, font_size, font);
    if full_w <= max_width || !text.contains(' ') {
        return vec![text.to_string()];
    }
    let words: Vec<&str> = text.split_whitespace().collect();
    let mut lines: Vec<String> = Vec::new();
    let mut current_line = String::new();
    for word in &words {
        let test = if current_line.is_empty() {
            word.to_string()
        } else {
            format!("{} {}", current_line, word)
        };
        if text_width(&test, font_size, font) > max_width && !current_line.is_empty() {
            lines.push(current_line);
            current_line = word.to_string();
        } else {
            current_line = test;
        }
    }
    if !current_line.is_empty() {
        lines.push(current_line);
    }
    lines
}

pub(super) fn compute_axis_range(
    max_val: f32,
    target_ticks: usize,
    headroom_threshold: f32,
) -> (f32, f32) {
    let ts = nice_tick_step(max_val, target_ticks);
    let mut am = (max_val / ts).ceil() * ts;
    if max_val >= am * headroom_threshold {
        am += ts;
    }
    (ts, am)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_ceil_nice_zero_and_negative() {
        assert_eq!(ceil_nice(0.0), 1.0);
        assert_eq!(ceil_nice(-5.0), 1.0);
    }

    #[test]
    fn test_ceil_nice_values() {
        assert_eq!(ceil_nice(0.3), 0.2); // mag=0.1, norm=3.0, nice=2.0
        assert_eq!(ceil_nice(0.8), 1.0); // mag=0.1, norm=8.0, nice=10.0
        assert_eq!(ceil_nice(1.5), 1.0); // mag=1.0, norm=1.5, nice=1.0
        assert_eq!(ceil_nice(25.0), 20.0); // mag=10, norm=2.5, nice=2.0
        assert_eq!(ceil_nice(70.0), 50.0); // mag=10, norm=7.0, nice=5.0 (<=7.0)
        assert_eq!(ceil_nice(15.0), 10.0); // mag=10, norm=1.5, nice=1.0
        assert_eq!(ceil_nice(800.0), 1000.0);
        assert_eq!(ceil_nice(3500.0), 5000.0);
    }

    #[test]
    fn test_nice_tick_step() {
        assert_eq!(nice_tick_step(100.0, 5), 20.0);
        assert_eq!(nice_tick_step(1000.0, 5), 200.0);
        assert_eq!(nice_tick_step(80.0, 5), 20.0);
    }

    #[test]
    fn test_compute_axis_range_with_headroom() {
        let (step, max) = compute_axis_range(100.0, 5, 0.9);
        assert_eq!(step, 20.0);
        assert_eq!(max, 120.0);
    }

    #[test]
    fn test_compute_axis_range_tight() {
        let (step, max) = compute_axis_range(85.0, 10, 0.98);
        assert_eq!(step, 10.0);
        assert_eq!(max, 90.0);
    }

    #[test]
    fn test_format_tick_label_integer_step() {
        assert_eq!(format_tick_label(100.0, 20.0, None), "100");
        assert_eq!(format_tick_label(0.0, 10.0, None), "0");
    }

    #[test]
    fn test_format_tick_label_fractional_step() {
        assert_eq!(format_tick_label(1.5, 0.5, None), "1.5");
    }

    #[test]
    fn test_format_tick_label_with_format_code() {
        assert_eq!(format_tick_label(1234.5, 100.0, Some("#,##0")), "1,234");
        assert_eq!(format_tick_label(0.15, 0.1, Some("0.0%")), "15.0%");
    }

    #[test]
    fn test_apply_number_format_basic() {
        assert_eq!(apply_number_format(1234.56, "0.00"), "1234.56");
        assert_eq!(apply_number_format(1234.0, "0"), "1234");
    }

    #[test]
    fn test_apply_number_format_comma() {
        assert_eq!(apply_number_format(1234567.0, "#,##0"), "1,234,567");
        assert_eq!(apply_number_format(999.0, "#,##0"), "999");
    }

    #[test]
    fn test_apply_number_format_percent() {
        assert_eq!(apply_number_format(0.15, "0%"), "15%");
        assert_eq!(apply_number_format(1.0, "0%"), "100%");
    }

    #[test]
    fn test_insert_thousands_separator() {
        assert_eq!(insert_thousands_separator("1234"), "1,234");
        assert_eq!(insert_thousands_separator("1234567"), "1,234,567");
        assert_eq!(insert_thousands_separator("999"), "999");
        assert_eq!(insert_thousands_separator("1234.56"), "1,234.56");
        assert_eq!(insert_thousands_separator("-1234"), "-1,234");
    }

    #[test]
    fn test_text_width_approx() {
        assert_eq!(text_width_approx("Hello", 12.0), 30.0);
        assert_eq!(text_width_approx("", 12.0), 0.0);
        assert_eq!(text_width_approx("A", 10.0), 5.0);
    }
}
