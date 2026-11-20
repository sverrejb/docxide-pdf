use super::*;
use crate::model::{MathNode, MathRun, MathScript};

fn cambria_math() -> Vec<u8> {
    std::fs::read(concat!(env!("CARGO_MANIFEST_DIR"), "/fonts/Cambria.ttc")).unwrap()
}

fn run(text: &str) -> MathNode {
    MathNode::Run(MathRun {
        text: text.to_string(),
        style: None,
        script: MathScript::Roman,
        normal: None,
        color: None,
        strike: false,
    })
}

fn glyphs(b: &MathBox) -> Vec<(char, f32, f32, f32)> {
    b.items
        .iter()
        .filter_map(|it| match it {
            MathItem::Glyph { ch, x, y, size, .. } => Some((*ch, *x, *y, *size)),
            _ => None,
        })
        .collect()
}

fn layout_at(root: Vec<MathNode>, display: bool, size: f32) -> MathBox {
    let data = cambria_math();
    let font = MathFont::new(&data, 1).expect("Cambria Math face");
    layout(
        &root,
        display,
        size,
        &font,
        &MathSettings::default(),
        &|_, _, _| None,
    )
}

#[test]
fn script_sizes_floor_to_half_points() {
    // Word: 9→6.5/5, 12→8.5/7, 16→11.5/9.5, 20→14.5/12 (probe mp_sizes).
    for (base, script, scriptscript) in [
        (9.0, 6.5, 5.0),
        (12.0, 8.5, 7.0),
        (16.0, 11.5, 9.5),
        (20.0, 14.5, 12.0),
    ] {
        let root = vec![MathNode::Scripts {
            base: vec![run("x")],
            sub: None,
            sup: Some(vec![MathNode::Scripts {
                base: vec![run("y")],
                sub: None,
                sup: Some(vec![run("2")]),
                pre: false,
            }]),
            pre: false,
        }];
        let sizes: Vec<f32> = glyphs(&layout_at(root, false, base))
            .iter()
            .map(|g| g.3)
            .collect();
        assert_eq!(sizes, vec![base, script, scriptscript], "base {base}");
    }
}

#[test]
fn pendulum_radical_fraction() {
    // T=2π ²√(l/g) at 16pt, inline: Word draws l and g at 11.5pt, 9.5pt above
    // and 8.02pt below the baseline, the degree at 9.5pt.
    let root = vec![
        run("T=2π"),
        MathNode::Rad {
            degree: Some(vec![run("2")]),
            body: vec![MathNode::Frac {
                kind: crate::model::FracKind::Bar,
                num: vec![run("l")],
                den: vec![run("g")],
            }],
        },
    ];
    let b = layout_at(root, false, 16.0);
    let g = glyphs(&b);
    let find = |c: char| {
        *g.iter()
            .find(|g| g.0 == c)
            .unwrap_or_else(|| panic!("{c} in {g:?}"))
    };
    let l = find('\u{1D459}');
    let gg = find('\u{1D454}');
    assert_eq!(l.3, 11.5);
    assert!((l.2 - 9.5).abs() < 0.3, "numerator shift {}", l.2);
    assert!((gg.2 + 8.02).abs() < 0.3, "denominator shift {}", gg.2);
    assert_eq!(find('2').3, 16.0);
    let degree = g
        .iter()
        .filter(|g| g.0 == '2')
        .map(|g| g.3)
        .fold(99.0f32, f32::min);
    assert_eq!(degree, 9.5);
    // T's italic correction and a thick space before "=": Word 4.75pt.
    let t = find('\u{1D447}');
    let eq = find('=');
    let t_adv = 9.5234375;
    assert!(
        (eq.1 - t.1 - t_adv - 4.75).abs() < 0.2,
        "gap {}",
        eq.1 - t.1 - t_adv
    );
}

#[test]
fn radical_surd_drop_and_display_gap() {
    // Word at 12pt (probes mp3_rad/mp7_rad): the bar over x stands 10.0pt
    // up, above the half-surplus place; √y takes the 8.91pt-wide surd and
    // its bar 10.56pt up; display √(a/b) keeps the gap (bar at 17.55pt).
    let rad = |body: Vec<MathNode>, display: bool| {
        let b = layout_at(vec![MathNode::Rad { degree: None, body }], display, 12.0);
        let bar = b
            .items
            .iter()
            .filter_map(|it| match it {
                MathItem::Rule { y, h, .. } => Some(y + h),
                _ => None,
            })
            .fold(f32::MIN, f32::max);
        let body_x = glyphs(&b).iter().find(|g| g.0 != '√').unwrap().1;
        (bar, body_x)
    };
    let (bar, _) = rad(vec![run("x")], false);
    assert!((bar - 10.0).abs() < 0.15, "√x bar {bar}");
    let (bar, body_x) = rad(vec![run("y")], false);
    assert!((body_x - 8.91).abs() < 0.05, "√y surd width {body_x}");
    assert!((bar - 10.56).abs() < 0.15, "√y bar {bar}");
    let frac = MathNode::Frac {
        kind: crate::model::FracKind::Bar,
        num: vec![run("a")],
        den: vec![run("b")],
    };
    let (bar, _) = rad(vec![frac.clone()], true);
    assert!((bar - 17.55).abs() < 0.15, "display √(a/b) bar {bar}");
    // A body past the largest variant takes an assembled surd at once.
    let tall = (0..3).fold(frac, |n, _| MathNode::Frac {
        kind: crate::model::FracKind::Bar,
        num: vec![n.clone()],
        den: vec![n],
    });
    let (bar, _) = rad(vec![tall], true);
    assert!(bar > 40.0 && bar < 120.0, "tall radical bar {bar}");
}

#[test]
fn group_chr_limit_and_side() {
    // Word at 12pt (probes mp9_gc/mp10_gc): a pos="top" ⏟ draws as ⏞ on the
    // baseline with its script-size content 7.0pt below; a pos="bot" ⏞ with
    // m:vertJc="bot" draws as ⏟ with the content 9.5pt above.
    let gc = |chr: char, top: bool, vert_top: bool| {
        let root = vec![MathNode::GroupChr {
            chr,
            top,
            vert_top,
            body: vec![run("a+b+c")],
        }];
        let g = glyphs(&layout_at(root, false, 12.0));
        let brace = g.iter().find(|g| g.3 == 12.0).unwrap().0;
        let content = g.iter().find(|g| g.0 == '\u{1D44E}').unwrap();
        (brace, content.2, content.3)
    };
    let (brace, y, size) = gc('⏟', true, true);
    assert_eq!((brace, size), ('⏞', 8.5));
    assert!((y + 7.0).abs() < 0.15, "limit below {y}");
    let (brace, y, _) = gc('⏞', false, false);
    assert_eq!(brace, '⏟');
    assert!((y - 9.5).abs() < 0.15, "limit above {y}");
}

#[test]
fn eq_arr_alignment_points() {
    // Word (probe mp11_eqarr): x&=1 over yyy&=22 puts both = at one x, the
    // left parts flush right to it; the & draws nothing. In 2&x+&3&y+ over
    // 4&x+&&y− the y's line up although the second row's third segment is
    // empty (the space after + stays with the segment before).
    let eq_x = |rows: &[&str], c: char| -> Vec<f32> {
        let root = vec![MathNode::EqArr {
            rows: rows.iter().map(|r| vec![run(r)]).collect(),
        }];
        let g = glyphs(&layout_at(root, true, 12.0));
        assert!(g.iter().all(|g| g.0 != '&'));
        g.iter().filter(|g| g.0 == c).map(|g| g.1).collect()
    };
    let eq = eq_x(&["x&=1", "yyy&=22"], '=');
    assert!((eq[0] - eq[1]).abs() < 0.01, "= at {eq:?}");
    let y = eq_x(&["2&x+&3&y+", "4&x+&&y−"], '\u{1D466}');
    assert!((y[0] - y[1]).abs() < 0.01, "y at {y:?}");
}

#[test]
fn normal_text_in_its_own_font() {
    // Word (probe mp14_nor): e^max sets "max" in the text font at script size,
    // one atom, upright.
    let data = cambria_math();
    let font = MathFont::new(&data, 1).expect("Cambria Math face");
    let normal = MathNode::Run(MathRun {
        text: "max".into(),
        style: None,
        script: MathScript::Roman,
        normal: Some(crate::model::TextFont {
            name: "Calibri".into(),
            bold: false,
            italic: false,
            size: 12.0,
        }),
        color: None,
        strike: false,
    });
    let root = vec![MathNode::Scripts {
        base: vec![run("e")],
        sub: None,
        sup: Some(vec![normal]),
        pre: false,
    }];
    let measure = |tf: &crate::model::TextFont, text: &str, size: f32| {
        assert_eq!((tf.name.as_str(), text), ("Calibri", "max"));
        Some(TextBox {
            w: size,
            asc: 0.7 * size,
            desc: 0.0,
            font: "F9".into(),
        })
    };
    let b = layout(
        &root,
        false,
        12.0,
        &font,
        &MathSettings::default(),
        &measure,
    );
    let texts: Vec<_> = b
        .items
        .iter()
        .filter_map(|it| match it {
            MathItem::Text {
                text, font, size, ..
            } => Some((text.as_str(), font.as_str(), *size)),
            _ => None,
        })
        .collect();
    assert_eq!(texts, [("max", "F9", 8.5)]);
}

#[test]
fn display_breaks_prefer_relations_past_the_wrap_indent() {
    // Word (probes mp20_relpos/mp21_relmin): y=a+b+cd=… breaks before its
    // second = once that = stands 72pt into the line, even with room for
    // more; y=a+b= keeps going to the last + that fits.
    let terms: String = (1..12).map(|i| format!("+a{i}x{i}")).collect();
    let head = |pre: &str| {
        let b = layout_at(vec![run(&format!("{pre}=a0x0{terms}"))], true, 12.0);
        let (head, tail) = b.fit(200.0).unwrap();
        (head.width, glyphs(&tail)[0].0)
    };
    let (w, first) = head("y=a+b+cd");
    assert_eq!(first, '=');
    assert!(w > 72.0 && w < 100.0, "head {w}");
    let (w, first) = head("y=a+b");
    assert_eq!(first, '+');
    assert!(w > 150.0, "head {w}");
}

#[test]
fn accent_fits_within_its_base() {
    // Word (probe mp22_acc): the widest hat that fits on both sides of the
    // attachment point; 𝐴's apex leaves less room than 𝐻 gives, and a hat
    // over xyz stays narrower than the three letters.
    let data = cambria_math();
    let font = MathFont::new(&data, 1).expect("Cambria Math face");
    let hat_width = |body: &str| {
        let root = vec![MathNode::Accent {
            chr: '\u{302}',
            body: vec![run(body)],
        }];
        let b = layout(
            &root,
            false,
            12.0,
            &font,
            &MathSettings::default(),
            &|_, _, _| None,
        );
        let gid = b
            .items
            .iter()
            .find_map(|it| match it {
                MathItem::Glyph {
                    gid, ch: '\u{302}', ..
                } => Some(*gid),
                _ => None,
            })
            .unwrap();
        (font.metrics(gid, 12.0).0, b.width)
    };
    let (a, _) = hat_width("A");
    let (h, _) = hat_width("H");
    let (x3, base) = hat_width("xyz");
    assert!(a < 4.0 && (h - 6.66).abs() < 0.05, "A {a} H {h}");
    assert!((x3 - 11.34).abs() < 0.05 && x3 < base, "xyz hat {x3}");
}

#[test]
fn empty_arguments_and_fraction_spacing() {
    // Word (probes mp24_empty/mp25_empty): an empty superscript keeps a ⬚'s
    // room; a stacked fraction stands a thin space from its neighbours.
    let width = |root: Vec<MathNode>| layout_at(root, false, 12.0).width;
    let x = width(vec![run("x")]);
    let x_empty_sup = width(vec![MathNode::Scripts {
        base: vec![run("x")],
        sub: None,
        sup: Some(vec![]),
        pre: false,
    }]);
    assert!(x_empty_sup - x > 7.5, "empty sup adds {}", x_empty_sup - x);
    let frac = |kind| MathNode::Frac {
        kind,
        num: vec![run("1")],
        den: vec![run("2")],
    };
    let stacked = width(vec![run("a"), frac(crate::model::FracKind::Bar), run("b")]);
    let alone = width(vec![frac(crate::model::FracKind::Bar)]) + width(vec![run("ab")]);
    assert!(
        (stacked - alone - 4.0).abs() < 0.5,
        "two thin spaces: {}",
        stacked - alone
    );
}

#[test]
fn display_fraction_shifts() {
    // a/b on its own line at 12pt: numerator 9.05 up, denominator 7.95 down.
    let root = vec![MathNode::Frac {
        kind: crate::model::FracKind::Bar,
        num: vec![run("a")],
        den: vec![run("b")],
    }];
    let b = layout_at(root, true, 12.0);
    let g = glyphs(&b);
    assert_eq!(g[0].3, 12.0);
    assert!((g[0].2 - 9.05).abs() < 0.2, "{g:?}");
    assert!((g[1].2 + 7.95).abs() < 0.2, "{g:?}");
}

#[test]
fn binary_and_relation_spacing() {
    // a+b=c at 12pt: medium (2.67) around +, thick (3.33) around =, plus a's
    // and b's italic corrections before the upright operators.
    let b = layout_at(vec![run("a+b=c")], false, 12.0);
    let g = glyphs(&b);
    let gap = |i: usize, adv: f32| g[i + 1].1 - g[i].1 - adv;
    let data = cambria_math();
    let font = MathFont::new(&data, 1).unwrap();
    let adv = |c: char| font.metrics(font.glyph(c, 0), 12.0).0;
    let plus = adv('+');
    assert!(
        (gap(1, plus) - 2.67).abs() < 0.05,
        "after + {}",
        gap(1, plus)
    );
    let eq = adv('=');
    assert!((gap(3, eq) - 3.33).abs() < 0.05, "after = {}", gap(3, eq));
    // A leading minus is unary: no space after it.
    let b = layout_at(vec![run("-a")], false, 12.0);
    let g = glyphs(&b);
    assert_eq!(g[0].0, '\u{2212}');
    assert!((g[1].1 - adv('\u{2212}')).abs() < 0.01);
}

#[test]
fn inline_zones_break_after_operators() {
    // Word: "… + 𝑎5𝑥5 +" / "𝑎6𝑥6 + …": a break after each top-level binary
    // or relation operator, none inside a fraction; display math breaks
    // before them instead.
    let root = vec![
        run("a+b="),
        MathNode::Frac {
            kind: crate::model::FracKind::Bar,
            num: vec![run("x+y")],
            den: vec![run("2")],
        },
    ];
    let b = layout_at(root.clone(), false, 12.0);
    assert_eq!(b.breaks.len(), 2, "{:?}", b.breaks);
    let display = layout_at(root, true, 12.0);
    assert_eq!(display.breaks.len(), 2);
    assert!(display.breaks.iter().zip(&b.breaks).all(|(d, i)| d.1 < i.1));
    let room = b.breaks[0].0;
    let (head, tail) = b.fit(room).unwrap();
    assert_eq!(glyphs(&head).len(), 2);
    assert_eq!(glyphs(&tail).len(), 2 + 4);
    assert!((head.width + tail.width - b.width).abs() < b.breaks[0].1 - b.breaks[0].0 + 0.01);
    assert!(glyphs(&tail).iter().all(|g| g.1 >= -0.01));
    // Each part is as tall as its own atoms: the fraction's height stays on
    // its line (Word, probe mp16_brk).
    assert!(head.ascent < tail.ascent && head.descent < tail.descent);
    assert_eq!((tail.ascent, tail.descent), (b.ascent, b.descent));
}

#[test]
fn struck_run_strikes_through_operator_spacing() {
    // Word strikes deleted math in one line per run, across the space around
    // the relation (corpus redline 970e9bbbcc).
    let struck = MathNode::Run(MathRun {
        text: "x=y".into(),
        style: None,
        script: MathScript::Roman,
        normal: None,
        color: Some([1, 2, 3]),
        strike: true,
    });
    let b = layout_at(vec![struck], false, 12.0);
    let rules: Vec<_> = b
        .items
        .iter()
        .filter_map(|it| match it {
            MathItem::Rule { x, w, color, .. } => Some((*x, *w, *color)),
            _ => None,
        })
        .collect();
    assert_eq!(rules.len(), 1, "one strike for the run: {rules:?}");
    let (x, w, color) = rules[0];
    assert!(
        x.abs() < 0.01 && (x + w - b.width).abs() < 0.5,
        "{rules:?} in {}",
        b.width
    );
    assert_eq!(color, Some([1, 2, 3]));
}
