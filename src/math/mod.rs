//! Office Math layout: an OMML tree (`model::MathZone`) to a box of placed
//! glyphs and rules, driven by the math font's OpenType MATH table the way
//! Word does it (Word's math engine is where that table comes from).
//!
//! Rules measured from Word's exports (online and local; see
//! `current_focus/office-math.md`):
//! - script sizes are the base size × ScriptPercentScaleDown (73%) and
//!   × ScriptScriptPercentScaleDown (60%), floored to the half point
//!   (16pt → 11.5 / 9.5, 12pt → 8.5 / 7.0);
//! - script-level glyphs take the font's `ssty` alternates (1st and 2nd);
//! - letters map to the Mathematical Alphanumeric Symbols (italic by default);
//! - atoms are spaced with TeX's table (thin 3/18, medium 4/18, thick 5/18 em);
//! - an italic glyph followed by an upright one keeps its italic correction.

mod alphanumeric;
mod font;

use crate::model::{FracKind, MathNode, MathRow, MathRun, MathSettings, TextFont};
pub(crate) use font::MathFont;
use ttf_parser::math::{Constants, MathValue};

/// A laid-out zone. Coordinates are points from the box's left end on its
/// baseline, y up.
#[derive(Debug)]
pub struct MathBox {
    pub width: f32,
    pub ascent: f32,
    pub descent: f32,
    pub items: Vec<MathItem>,
    /// The font's MathLeading at the zone's size: ink higher than the text's
    /// ascent less this gap raises the line (OpenType MATH, `MathLeading`).
    pub leading: f32,
    /// Where the zone may break, as (end of the line's part, start of the
    /// next, at a relation): see `assemble`.
    pub breaks: Vec<(f32, f32, bool)>,
    /// Each top-level atom's start, ascent and descent: a broken zone's parts
    /// take only their own atoms' height.
    atoms: Vec<(f32, f32, f32)>,
    /// A display zone (`m:oMathPara`), and how far its continuation lines
    /// are indented.
    pub display: bool,
    pub wrap_indent: f32,
    /// The math font's own line box at the zone's size (`MathFont::line_metrics`):
    /// what a line of nothing but plain math is sized by.
    pub line_ascent: f32,
    pub line_descent: f32,
}

#[derive(Debug, Clone, PartialEq)]
pub enum MathItem {
    /// A glyph of the math font by its original id; `ch` is what it reads as.
    Glyph {
        gid: u16,
        ch: char,
        x: f32,
        y: f32,
        size: f32,
        color: Option<[u8; 3]>,
    },
    /// A filled rectangle (fraction bar, radical overbar) in the text colour,
    /// `y` its bottom.
    Rule { x: f32, y: f32, w: f32, h: f32 },
    /// An `m:nor` run's text in its own font, by the font's PDF name.
    Text {
        text: String,
        font: String,
        x: f32,
        y: f32,
        size: f32,
        color: Option<[u8; 3]>,
    },
}

/// How `m:nor` text measures in its font: width, ascent, descent and the
/// font's PDF name.
pub(crate) struct TextBox {
    pub(crate) w: f32,
    pub(crate) asc: f32,
    pub(crate) desc: f32,
    pub(crate) font: String,
}

/// Measures `m:nor` text at a size; None when its font isn't available.
pub(crate) type TextMeasure<'a> = &'a dyn Fn(&TextFont, &str, f32) -> Option<TextBox>;

impl MathBox {
    /// The zone broken at its last break that leaves at most `room` before
    /// it: (that part, the rest), or None when no break does. A display
    /// equation breaks at a relation when one fits at least m:wrapIndent into
    /// the line, however much more would fit: y=a+b+cd= breaks before its
    /// second =, y=a+b+c= (69.8pt) doesn't (probes `mp20_relpos`,
    /// `mp21_relmin`).
    pub(crate) fn fit(&self, room: f32) -> Option<(MathBox, MathBox)> {
        let fits = || self.breaks.iter().rev().filter(|b| b.0 <= room);
        let rel = fits().find(|b| self.display && b.2 && b.0 >= self.wrap_indent);
        let &(end, start, _) = rel.or_else(|| fits().next())?;
        Some(self.split((end, start)))
    }

    /// The parts before `end` and from `start` (a break point): the first
    /// keeps the items left of the break, the second the rest, moved to 0;
    /// each as tall as its own atoms (Word: a fraction on the first line
    /// doesn't raise the second, probe `mp16_brk`).
    fn split(&self, (end, start): (f32, f32)) -> (MathBox, MathBox) {
        let x = |it: &MathItem| match it {
            MathItem::Glyph { x, .. } | MathItem::Rule { x, .. } | MathItem::Text { x, .. } => *x,
        };
        let (head, tail): (Vec<_>, Vec<_>) = self.items.iter().cloned().partition(|it| x(it) < end);
        let i = self.atoms.partition_point(|a| a.0 < end);
        let head_atoms = self.atoms[..i].to_vec();
        let tail_atoms = self.atoms[i..]
            .iter()
            .map(|&(x, a, d)| (x - start, a, d))
            .collect();
        let part = |items, width, breaks, atoms: Vec<(f32, f32, f32)>| {
            let (ascent, descent) = extent(&atoms);
            MathBox {
                width,
                items,
                breaks,
                ascent,
                descent,
                atoms,
                display: self.display,
                wrap_indent: self.wrap_indent,
                leading: self.leading,
                line_ascent: self.line_ascent,
                line_descent: self.line_descent,
            }
        };
        let tail = tail.into_iter().map(|it| shift(it, -start, 0.0)).collect();
        let tail_breaks = self
            .breaks
            .iter()
            .filter(|b| b.0 > end)
            .map(|&(e, s, rel)| (e - start, s - start, rel))
            .collect();
        (
            part(head, end, Vec::new(), head_atoms),
            part(tail, self.width - start, tail_breaks, tail_atoms),
        )
    }
}

/// Lay out one zone at `size` points.
pub(crate) fn layout(
    root: &MathRow,
    display: bool,
    size: f32,
    font: &MathFont,
    settings: &MathSettings,
    text: TextMeasure,
) -> MathBox {
    let cx = Ctx {
        font,
        settings,
        base: size,
        display,
        text,
    };
    let style = if display { Style::D } else { Style::T };
    let st = St::new(style);
    let mut atoms = cx.atoms(root, st);
    let (xs, w) = cx.place(&mut atoms, st);
    let extents: Vec<_> = atoms
        .iter()
        .zip(&xs)
        .map(|(a, &x)| (x, a.asc, a.desc))
        .collect();
    let (ascent, descent) = extent(&extents);
    // Inline math breaks after an operator whatever m:brkBin says; a display
    // equation too wide for its line breaks before one unless m:brkBin is
    // after (probes `mp17_wrap`, `mp18_after`).
    let before = display && !settings.break_after;
    let (b, breaks) = assemble(atoms, &xs, before);
    let (line_ascent, line_descent) = font.line_metrics(size);
    MathBox {
        atoms: extents,
        display,
        wrap_indent: settings.wrap_indent,
        breaks,
        // A zone ending in an italic glyph keeps its italic correction: text
        // after x̂ starts 0.36pt further on (probe `mp22_acc`, four zones).
        width: w + b.ic,
        ascent,
        descent,
        items: b.items,
        leading: font.constant(Constants::math_leading, size),
        line_ascent,
        line_descent,
    }
}

/// TeX's styles: display, text, script, scriptscript.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Style {
    D,
    T,
    S,
    SS,
}

/// A style and whether it is cramped (superscripts sit lower).
#[derive(Debug, Clone, Copy)]
struct St {
    style: Style,
    cramped: bool,
}

impl St {
    fn new(style: Style) -> Self {
        Self {
            style,
            cramped: false,
        }
    }
    fn script_level(self) -> usize {
        match self.style {
            Style::D | Style::T => 0,
            Style::S => 1,
            Style::SS => 2,
        }
    }
    fn display(self) -> bool {
        self.style == Style::D
    }
    /// Superscript style.
    fn sup(self) -> Self {
        Self {
            style: match self.style {
                Style::D | Style::T => Style::S,
                _ => Style::SS,
            },
            cramped: self.cramped,
        }
    }
    /// Subscript style.
    fn sub(self) -> Self {
        self.sup().cramp()
    }
    /// Numerator style; the denominator is its cramped form. In display
    /// math Word never shrinks a fraction's parts (a nested a/b/c keeps 12pt
    /// letters at 12pt, text-style shifts); inline, they drop a script level.
    fn num(self, display_zone: bool) -> Self {
        Self {
            style: match self.style {
                Style::D => Style::T,
                Style::T if display_zone => Style::T,
                Style::T => Style::S,
                _ => Style::SS,
            },
            cramped: self.cramped,
        }
    }
    fn cramp(self) -> Self {
        Self {
            cramped: true,
            ..self
        }
    }
}

/// TeX's atom classes, for the space between neighbours.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Class {
    Ord,
    Op,
    Bin,
    Rel,
    Open,
    Close,
    Punct,
    /// A space typed in a math run: it is the spacing, nothing is added
    /// beside it (Word's ", " zone is a comma and a 0.22em space).
    Space,
    /// A stacked fraction: a thin space from its ordinary neighbours.
    Inner,
}

/// A box being built: like `MathBox`, plus what spacing needs to know of
/// its ends.
struct B {
    w: f32,
    asc: f32,
    desc: f32,
    items: Vec<MathItem>,
    first: Class,
    last: Class,
    /// Italic correction of the last glyph when it is italic.
    ic: f32,
    first_italic: bool,
    /// One glyph only (scripts attach to its italic correction).
    single_glyph: Option<u16>,
    /// A character of a run (or an m:nor run's text), not a structure.
    character: bool,
    /// The tallest delimiter inside: an enclosing one grows past it.
    fence: f32,
}

impl B {
    fn empty() -> Self {
        Self {
            w: 0.0,
            asc: 0.0,
            desc: 0.0,
            items: Vec::new(),
            first: Class::Ord,
            last: Class::Ord,
            ic: 0.0,
            first_italic: false,
            single_glyph: None,
            character: false,
            fence: 0.0,
        }
    }

    /// An empty box with `base`'s ends, for spacing it as `base`.
    fn ends_of(base: &B) -> Self {
        Self {
            first: base.first,
            last: base.last,
            ic: base.ic,
            first_italic: base.first_italic,
            ..Self::empty()
        }
    }

    fn ord(mut self) -> Self {
        self.first = Class::Ord;
        self.last = Class::Ord;
        self.single_glyph = None;
        self
    }

    /// Place `other` with its origin at (dx, dy) in this box.
    fn put(&mut self, other: B, dx: f32, dy: f32) {
        self.w = self.w.max(dx + other.w);
        self.asc = self.asc.max(other.asc + dy);
        self.desc = self.desc.max(other.desc - dy);
        self.fence = self.fence.max(other.fence);
        self.items
            .extend(other.items.into_iter().map(|it| shift(it, dx, dy)));
    }

    fn rule(&mut self, x: f32, y: f32, w: f32, h: f32) {
        self.w = self.w.max(x + w);
        self.asc = self.asc.max(y + h);
        self.desc = self.desc.max(-y);
        self.items.push(MathItem::Rule { x, y, w, h });
    }
}

fn shift(mut it: MathItem, dx: f32, dy: f32) -> MathItem {
    match &mut it {
        MathItem::Glyph { x, y, .. }
        | MathItem::Rule { x, y, .. }
        | MathItem::Text { x, y, .. } => {
            *x += dx;
            *y += dy;
        }
    }
    it
}

struct Ctx<'a> {
    font: &'a MathFont<'a>,
    settings: &'a MathSettings,
    base: f32,
    /// A display zone (`m:oMathPara`).
    display: bool,
    text: TextMeasure<'a>,
}

/// Floor to the half point, as Word sizes its script levels.
fn half_point_floor(v: f32) -> f32 {
    ((v * 2.0) + 1e-3).floor() / 2.0
}

type Constant<'a> = fn(&Constants<'a>) -> MathValue<'a>;

impl<'a> Ctx<'a> {
    fn size(&self, st: St) -> f32 {
        let pct = match st.script_level() {
            0 => return self.base,
            1 => self.font.script_percent(),
            _ => self.font.script_script_percent(),
        };
        half_point_floor(self.base * pct / 100.0)
    }

    /// A MATH constant in points at `st`'s size.
    fn k(&self, st: St, f: Constant<'a>) -> f32 {
        self.font.constant(f, self.size(st))
    }

    /// The display-style constant in display style, else the other.
    fn kd(&self, st: St, display: Constant<'a>, text: Constant<'a>) -> f32 {
        self.k(st, if st.display() { display } else { text })
    }

    fn glyph(&self, ch: char, st: St, color: Option<[u8; 3]>) -> B {
        let gid = self.font.glyph(ch, st.script_level());
        self.glyph_id(gid, ch, self.size(st), color)
    }

    fn glyph_id(&self, gid: u16, ch: char, size: f32, color: Option<[u8; 3]>) -> B {
        let (adv, asc, desc) = self.font.metrics(gid, size);
        B {
            w: adv,
            asc,
            desc,
            items: vec![MathItem::Glyph {
                gid,
                ch,
                x: 0.0,
                y: 0.0,
                size,
                color,
            }],
            single_glyph: Some(gid),
            ..B::empty()
        }
    }

    /// Lay out a row: atoms side by side with TeX's spacing. A run's
    /// characters are atoms of the row, so a sign at a run's start can be
    /// binary.
    fn row(&self, row: &MathRow, st: St) -> B {
        if row.is_empty() {
            return self.placeholder(st);
        }
        self.hbox(self.atoms(row, st), st)
    }

    /// Word keeps an empty argument's room: an invisible ⬚ at its size (an
    /// empty numerator, script, limit, radicand, matrix cell or delimiter;
    /// probes `mp23_nary`–`mp25_empty`).
    fn placeholder(&self, st: St) -> B {
        let gid = self.font.glyph('\u{2B1A}', st.script_level());
        let (w, asc, desc) = self.font.metrics(gid, self.size(st));
        B {
            w,
            asc,
            desc,
            ..B::empty()
        }
    }

    fn atoms(&self, row: &MathRow, st: St) -> Vec<B> {
        let mut atoms: Vec<B> = Vec::new();
        for node in row {
            match node {
                MathNode::Run(run) => atoms.extend(self.run_text(&run.text, run, st)),
                _ => atoms.push(self.node(node, st)),
            }
        }
        atoms
    }

    fn hbox(&self, mut atoms: Vec<B>, st: St) -> B {
        let (xs, _) = self.place(&mut atoms, st);
        if atoms.len() == 1 {
            return atoms.pop().unwrap();
        }
        assemble(atoms, &xs, false).0
    }

    /// Where each atom of a row goes, with TeX's spacing, and the row's width.
    fn place(&self, atoms: &mut [B], st: St) -> (Vec<f32>, f32) {
        // TeX rule: a binary operator with nothing to operate on is ordinary.
        for i in 0..atoms.len() {
            if atoms[i].first == Class::Bin {
                let prev = i.checked_sub(1).map(|p| atoms[p].last);
                let next = atoms.get(i + 1).map(|n| n.first);
                let unary = matches!(
                    prev,
                    None | Some(Class::Bin | Class::Op | Class::Rel | Class::Open | Class::Punct)
                ) || matches!(
                    next,
                    None | Some(Class::Rel | Class::Close | Class::Punct)
                );
                if unary {
                    atoms[i].first = Class::Ord;
                    atoms[i].last = Class::Ord;
                }
            }
        }
        let em = self.size(st);
        let mut xs = Vec::with_capacity(atoms.len());
        let mut w = 0.0f32;
        for (i, a) in atoms.iter().enumerate() {
            let mut x = w;
            if let Some(p) = i.checked_sub(1).map(|p| &atoms[p]) {
                x += space(p.last, a.first, st) * em / 18.0;
                // An italic glyph keeps its italic correction before an
                // upright atom unless that is an ordinary character: 𝑎1, 𝑏c
                // and 𝑣 if set tight, 𝑥+, 𝑏), 𝜋√ and 𝑥𝑦² apart (probes
                // `mp_spacing`, `mp17_wrap`, `mp14_nor`, case221).
                if !a.first_italic && !(a.character && a.first == Class::Ord) {
                    x += p.ic;
                }
            }
            xs.push(x);
            w = w.max(x + a.w);
        }
        (xs, w)
    }

    /// `text` set in `run`'s style: one atom per character.
    fn run_text(&self, text: &str, run: &MathRun, st: St) -> Vec<B> {
        let size = self.size(st);
        // Word sets m:nor text in the run's own font, upright whatever m:sty
        // says, at the run's size (scaled down in scripts), as one ordinary
        // atom: probe `mp14_nor`.
        if let Some(tf) = &run.normal {
            let size = size * tf.size / self.base;
            if let Some(tb) = (self.text)(tf, text, size) {
                return vec![B {
                    w: tb.w,
                    asc: tb.asc,
                    desc: tb.desc,
                    items: vec![MathItem::Text {
                        text: text.to_string(),
                        font: tb.font,
                        x: 0.0,
                        y: 0.0,
                        size,
                        color: run.color,
                    }],
                    character: true,
                    ..B::empty()
                }];
            }
        }
        text.chars()
            .map(|ch| {
                let (drawn, italic) = alphanumeric::styled(ch, run.style, run.script);
                let drawn = match drawn {
                    // Word draws a hyphen-minus in math as the minus sign.
                    '-' if run.normal.is_none() => '\u{2212}',
                    '\'' if run.normal.is_none() => '\u{2032}',
                    c => c,
                };
                let gid = self.font.glyph(drawn, st.script_level());
                let mut b = self.glyph_id(gid, drawn, size, run.color);
                if italic {
                    b.first_italic = true;
                    b.ic = self.font.italic_correction(gid, size);
                }
                let class = match drawn {
                    ' ' => Class::Space,
                    _ if run.normal.is_some() => Class::Ord,
                    c => class_of(c),
                };
                b.first = class;
                b.last = class;
                b.character = true;
                b
            })
            .collect()
    }

    fn node(&self, node: &MathNode, st: St) -> B {
        match node {
            MathNode::Run(run) => self.hbox(self.run_text(&run.text, run, st), st),
            MathNode::Frac { kind, num, den } => self.frac(*kind, num, den, st),
            MathNode::Scripts {
                base,
                sub,
                sup,
                pre,
            } => self.scripts(self.row(base, st), sub.as_ref(), sup.as_ref(), *pre, st),
            MathNode::Rad { degree, body } => self.radical(degree.as_ref(), body, st),
            MathNode::Nary {
                chr,
                under_over,
                sub,
                sup,
                grow,
                body,
            } => self.nary(
                *chr,
                *under_over,
                sub.as_ref(),
                sup.as_ref(),
                *grow,
                body,
                st,
            ),
            MathNode::Delim {
                open,
                close,
                sep,
                grow,
                items,
            } => self.delim(*open, *close, *sep, *grow, items, st),
            MathNode::Func { name, body } => {
                let mut name_b = self.row(name, st);
                name_b.first = Class::Op;
                name_b.last = Class::Op;
                self.hbox(vec![name_b, self.row(body, st)], st)
            }
            MathNode::Accent { chr, body } => self.accent(*chr, body, st),
            MathNode::Bar { top, body } => self.bar(*top, body, st),
            MathNode::GroupChr {
                chr,
                top,
                vert_top,
                body,
            } => self.group_chr(*chr, *top, top == vert_top, body, st),
            MathNode::Limit { upper, body, limit } => {
                let b = self.row(body, st);
                let lim = Some(self.row(limit, st.sup()));
                if *upper {
                    self.limits(b, None, lim, st)
                } else {
                    self.limits(b, lim, None, st)
                }
            }
            MathNode::Matrix { rows } => {
                let rows = rows
                    .iter()
                    .map(|r| r.iter().map(|c| self.row(c, st)).collect())
                    .collect();
                self.grid(rows, st, true)
            }
            MathNode::EqArr { rows } => self.eq_arr(rows, st),
            MathNode::Boxed(body) => self.row(body, st).ord(),
            MathNode::BorderBox { hide, body } => self.border_box(*hide, body, st),
            MathNode::Phantom {
                show,
                zero_width,
                zero_ascent,
                zero_descent,
                body,
            } => {
                let mut b = self.row(body, st).ord();
                if !show {
                    b.items.clear();
                }
                if *zero_width {
                    b.w = 0.0;
                }
                if *zero_ascent {
                    b.asc = 0.0;
                }
                if *zero_descent {
                    b.desc = 0.0;
                }
                b
            }
        }
    }

    fn frac(&self, kind: FracKind, num: &MathRow, den: &MathRow, st: St) -> B {
        let mut b = self.frac_box(kind, num, den, st);
        // A stacked fraction is TeX's inner atom: Word sets 𝑎½𝑏 with a thin
        // space either side (probes `mp_spacing`, `mp25_empty`).
        if matches!(kind, FracKind::Bar | FracKind::NoBar) {
            (b.first, b.last) = (Class::Inner, Class::Inner);
        }
        b
    }

    fn frac_box(&self, kind: FracKind, num: &MathRow, den: &MathRow, st: St) -> B {
        // m:smallFrac sets display fractions at text size.
        let st = if self.settings.small_frac && st.display() {
            St::new(Style::T)
        } else {
            st
        };
        if kind == FracKind::Linear {
            let slash = self.glyph('/', st, None).ord();
            let atoms = vec![self.row(num, st), slash, self.row(den, st)];
            return self.hbox(atoms, st).ord();
        }
        let parts = st.num(self.display && !self.settings.small_frac);
        let n = self.row(num, parts);
        let d = self.row(den, parts.cramp());
        if kind == FracKind::Skewed {
            return self.skewed(n, d, st);
        }
        let w = n.w.max(d.w);
        let (nx, dx) = ((w - n.w) / 2.0, (w - d.w) / 2.0);
        let mut b = B::empty();
        if kind == FracKind::NoBar {
            let mut up = self.kd(
                st,
                Constants::stack_top_display_style_shift_up,
                Constants::stack_top_shift_up,
            );
            let mut down = self.kd(
                st,
                Constants::stack_bottom_display_style_shift_down,
                Constants::stack_bottom_shift_down,
            );
            let gap = self.kd(
                st,
                Constants::stack_display_style_gap_min,
                Constants::stack_gap_min,
            );
            let actual = (up - n.desc) - (d.asc - down);
            if actual < gap {
                up += (gap - actual) / 2.0;
                down += (gap - actual) / 2.0;
            }
            b.put(n, nx, up);
            b.put(d, dx, -down);
            return b;
        }
        let axis = self.k(st, Constants::axis_height);
        let t = self.k(st, Constants::fraction_rule_thickness);
        let up = self
            .kd(
                st,
                Constants::fraction_numerator_display_style_shift_up,
                Constants::fraction_numerator_shift_up,
            )
            .max(
                axis + t / 2.0
                    + n.desc
                    + self.kd(
                        st,
                        Constants::fraction_num_display_style_gap_min,
                        Constants::fraction_numerator_gap_min,
                    ),
            );
        let down = self
            .kd(
                st,
                Constants::fraction_denominator_display_style_shift_down,
                Constants::fraction_denominator_shift_down,
            )
            .max(
                d.asc - (axis - t / 2.0)
                    + self.kd(
                        st,
                        Constants::fraction_denom_display_style_gap_min,
                        Constants::fraction_denominator_gap_min,
                    ),
            );
        b.put(n, nx, up);
        b.put(d, dx, -down);
        b.rule(0.0, axis - t / 2.0, w, t);
        b
    }

    /// `m:type="skw"`: numerator raised, denominator lowered, a slash between.
    fn skewed(&self, n: B, d: B, st: St) -> B {
        let gap = self.k(st, Constants::skewed_fraction_horizontal_gap);
        let slash = self.glyph('/', st, None);
        let up = self.k(st, Constants::superscript_shift_up);
        let down = self.k(st, Constants::subscript_shift_down);
        let (nw, sw) = (n.w, slash.w);
        let mut b = B::empty();
        b.put(n, 0.0, up);
        b.put(slash, nw + (gap - sw).max(0.0) / 2.0, 0.0);
        b.put(d, nw + gap.max(sw), -down);
        b
    }

    fn scripts(
        &self,
        base: B,
        sub: Option<&MathRow>,
        sup: Option<&MathRow>,
        pre: bool,
        st: St,
    ) -> B {
        let sub_b = sub.map(|r| self.row(r, st.sub()));
        let sup_b = sup.map(|r| self.row(r, st.sup()));
        if !pre {
            let (bw, ic) = (base.w, base.ic);
            // The superscript clears the base's italic correction.
            let drop = base.single_glyph.is_none();
            return self.attach(base, sub_b, sup_b, (bw, bw + ic), drop, st);
        }
        let drop = base.single_glyph.is_none();
        let (sub_shift, sup_shift) =
            self.script_shifts(&base, sub_b.as_ref(), sup_b.as_ref(), drop, st);
        let sw = sub_b.as_ref().map_or(0.0, |s| s.w);
        let pw = sup_b.as_ref().map_or(0.0, |s| s.w);
        let scripts_w = sw.max(pw);
        let mut b = B::ends_of(&base);
        if let Some(s) = sub_b {
            b.put(s, scripts_w - sw, -sub_shift);
        }
        if let Some(s) = sup_b {
            b.put(s, scripts_w - pw, sup_shift);
        }
        b.put(base, scripts_w, 0.0);
        b
    }

    /// `base` with a subscript at `sub_x` and a superscript at `sup_x`, then
    /// SpaceAfterScript; `drop`: see `script_shifts`.
    fn attach(
        &self,
        base: B,
        sub: Option<B>,
        sup: Option<B>,
        (sub_x, sup_x): (f32, f32),
        drop: bool,
        st: St,
    ) -> B {
        let (sub_shift, sup_shift) =
            self.script_shifts(&base, sub.as_ref(), sup.as_ref(), drop, st);
        let mut b = B::ends_of(&base);
        b.put(base, 0.0, 0.0);
        if let Some(s) = sub {
            b.put(s, sub_x, -sub_shift);
        }
        if let Some(s) = sup {
            b.put(s, sup_x, sup_shift);
        }
        b.w += self.k(st, Constants::space_after_script);
        b
    }

    /// Script shifts (sub down, sup up), the OpenType MATH rules. With
    /// `drop` the scripts also hang from the base's ink (a box's, or a display
    /// operator's); a lone glyph's don't (TeX's rule 18a).
    fn script_shifts(
        &self,
        base: &B,
        sub: Option<&B>,
        sup: Option<&B>,
        drop: bool,
        st: St,
    ) -> (f32, f32) {
        let (b_asc, b_desc) = if drop {
            (base.asc, base.desc)
        } else {
            (0.0, 0.0)
        };
        let mut sub_shift = sub.map_or(0.0, |s| {
            self.k(st, Constants::subscript_shift_down)
                .max(b_desc + self.k(st, Constants::subscript_baseline_drop_min))
                .max(s.asc - self.k(st, Constants::subscript_top_max))
        });
        let mut sup_shift = sup.map_or(0.0, |s| {
            let up = if st.cramped {
                self.k(st, Constants::superscript_shift_up_cramped)
            } else {
                self.k(st, Constants::superscript_shift_up)
            };
            up.max(b_asc - self.k(st, Constants::superscript_baseline_drop_max))
                .max(s.desc + self.k(st, Constants::superscript_bottom_min))
        });
        if let (Some(sb), Some(sp)) = (sub, sup) {
            let gap_min = self.k(st, Constants::sub_superscript_gap_min);
            let gap = (sup_shift - sp.desc) - (sb.asc - sub_shift);
            if gap < gap_min {
                sub_shift += gap_min - gap;
                let max_bottom = self.k(st, Constants::superscript_bottom_max_with_subscript);
                let delta = max_bottom - (sup_shift - sp.desc);
                if delta > 0.0 {
                    sup_shift += delta;
                    sub_shift -= delta;
                }
            }
        }
        (sub_shift, sup_shift)
    }

    fn radical(&self, degree: Option<&MathRow>, body: &MathRow, st: St) -> B {
        let body_b = self.row(body, st.cramp());
        let t = self.k(st, Constants::radical_rule_thickness);
        let gap = self.kd(
            st,
            Constants::radical_display_style_vertical_gap,
            Constants::radical_vertical_gap,
        );
        let need = body_b.asc + body_b.desc + gap + t;
        // A taller surd than needed shares its surplus above and below the
        // body inline; in display the bar keeps the gap and it all goes
        // below. Word lowers the surd at most RadicalVerticalGap (the
        // text-style one, in display too) below its designed place, so over
        // short contents the bar sits higher, and a surd that then no longer
        // reaches below the body gives way to the next size (√y takes the
        // taller surd): probes `mp3_rad`/`mp7_rad`, 10-72pt.
        let max_drop = self.k(st, Constants::radical_vertical_gap);
        let share = if st.display() { 0.0 } else { 0.5 };
        let mut target = need;
        let (surd, rule_top, surd_bottom) = loop {
            let surd = self.stretch('√', target, self.size(st), true);
            let surd_h = surd.asc + surd.desc;
            let mut rule_top = body_b.asc + gap + t + (surd_h - need).max(0.0) * share;
            // An assembled surd has no designed place to drop from.
            let variant = surd.single_glyph.is_some();
            if variant {
                rule_top = rule_top.max(surd.asc - max_drop);
            }
            let surd_bottom = surd_h - rule_top;
            if surd_bottom >= body_b.desc || surd_h < target || !variant {
                break (surd, rule_top, surd_bottom);
            }
            target = surd_h * 1.01;
        };
        // The surd's top is the overbar's top.
        let surd_dy = rule_top - surd.asc;
        let mut b = B::empty();
        let mut x = 0.0;
        if let Some(deg) = degree {
            let d = self.row(deg, St::new(Style::SS));
            let before = self.k(st, Constants::radical_kern_before_degree);
            let after = self.k(st, Constants::radical_kern_after_degree);
            let raise = self.font.radical_degree_bottom_raise_percent() / 100.0;
            let dy = raise * (rule_top + surd_bottom) - surd_bottom + d.desc;
            x = (before + d.w + after).max(0.0);
            b.put(d, before, dy);
        }
        let surd_w = surd.w;
        b.put(surd, x, surd_dy);
        x += surd_w;
        let bw = body_b.w;
        b.put(body_b, x, 0.0);
        b.rule(x, rule_top - t, bw, t);
        b.asc = b
            .asc
            .max(rule_top + self.k(st, Constants::radical_extra_ascender));
        b
    }

    #[allow(clippy::too_many_arguments)]
    fn nary(
        &self,
        chr: char,
        under_over: Option<bool>,
        sub: Option<&MathRow>,
        sup: Option<&MathRow>,
        grow: bool,
        body: &MathRow,
        st: St,
    ) -> B {
        let size = self.size(st);
        let body_b = self.row(body, st);
        // m:grow: the operator spans its operand.
        let grown = if grow { body_b.asc + body_b.desc } else { 0.0 };
        let op = if st.display() || grow {
            // Word sets every display operator one variant taller than the
            // font's DisplayOperatorMinHeight (1.22em) asks: probes of 20
            // operators all take the variant of 1.91-2.15em, never the ∫
            // family's 1.35em one, so the bar lies in (1.35, 1.91] em.
            let min_h = if st.display() {
                1.5 * self.font.display_operator_min_height(size)
            } else {
                0.0
            };
            self.stretch(chr, min_h.max(grown), size, true)
        } else {
            self.glyph(chr, st, None)
        };
        let op_ic = op
            .single_glyph
            .map_or(0.0, |g| self.font.italic_correction(g, size));
        // Large operators are centred on the math axis. A display operator's
        // limits drop from its ink like a box's scripts (Word's display ∫ at
        // 12pt: 13.7pt up, 11.3pt down); inline ones sit like a letter's
        // (inline ∑: 4.8 up, 2.8 down).
        let mut centred = B::empty();
        let dy = self.k(st, Constants::axis_height) - (op.asc - op.desc) / 2.0;
        centred.put(op, 0.0, dy);
        let is_integral = matches!(
            chr,
            '∫' | '∬' | '∭' | '∮' | '∯' | '∰' | '∱' | '∲' | '∳' | '⨌'
        );
        let default_under = if is_integral {
            self.settings.int_under_over
        } else {
            self.settings.nary_under_over
        };
        let lo = sub.map(|r| self.row(r, st.sub()));
        let hi = sup.map(|r| self.row(r, st.sup()));
        // Inline (text style) limits always go beside the operator.
        let mut beside = false;
        let mut op_b = if under_over.unwrap_or(default_under) && st.display() {
            self.limits(centred, lo, hi, st)
        } else if lo.is_some() || hi.is_some() {
            // The italic correction moves the lower limit left of the upper
            // one (OpenType MATH, large operators).
            let w = centred.w;
            beside = true;
            self.attach(centred, lo, hi, (w - op_ic, w), st.display(), st)
        } else {
            centred
        };
        op_b.first = Class::Op;
        op_b.last = Class::Op;
        // In display math the operand starts right where the limits beside
        // the operator end, with no thin space (Word's ∫₀¹f: f at the end of
        // the "1", probe `mp23_nary`); inline it keeps one.
        if beside && st.display() {
            let x = op_b.w;
            return assemble(vec![op_b, body_b], &[0.0, x], false).0;
        }
        self.hbox(vec![op_b, body_b], st)
    }

    /// Limits centred under and over `base` (`m:limLow`, `m:limUpp`, n-ary
    /// operators with undOvr limits).
    fn limits(&self, base: B, lo: Option<B>, hi: Option<B>, st: St) -> B {
        let w = base
            .w
            .max(lo.as_ref().map_or(0.0, |l| l.w))
            .max(hi.as_ref().map_or(0.0, |h| h.w));
        let (base_asc, base_desc) = (base.asc, base.desc);
        let mut b = B::empty();
        let mut centre = |part: B, dy: f32| {
            let x = (w - part.w) / 2.0;
            b.put(part, x, dy);
        };
        centre(base, 0.0);
        // Drawn lower limit first, so text extraction reads "∑ i=1 n".
        if let Some(l) = lo {
            let dy = (base_desc + self.k(st, Constants::lower_limit_gap_min) + l.asc)
                .max(base_desc + self.k(st, Constants::lower_limit_baseline_drop_min));
            centre(l, -dy);
        }
        if let Some(h) = hi {
            let dy = (base_asc + self.k(st, Constants::upper_limit_gap_min) + h.desc)
                .max(base_asc + self.k(st, Constants::upper_limit_baseline_rise_min));
            centre(h, dy);
        }
        b
    }

    #[allow(clippy::too_many_arguments)]
    fn delim(
        &self,
        open: Option<char>,
        close: Option<char>,
        sep: Option<char>,
        grow: bool,
        items: &[MathRow],
        st: St,
    ) -> B {
        let size = self.size(st);
        let axis = self.k(st, Constants::axis_height);
        let bodies: Vec<B> = items.iter().map(|r| self.row(r, st)).collect();
        let (asc, desc) = bodies
            .iter()
            .fold((0.0f32, 0.0f32), |(a, d), b| (a.max(b.asc), d.max(b.desc)));
        // Delimiters cover the content symmetrically about the axis (TeX's
        // δ) less a 3/8 em shortfall, and grow past any delimiter inside them;
        // even at base size they sit centred on the axis. m:grow off keeps the
        // base glyph. Fitted on 36 Word probes at 12pt (inline and display
        // x, y, b², x+y, fractions, nested fractions and delimiters): all but
        // the radicals (`current_focus/office-math.md`) agree for a shortfall
        // in [4.47, 4.58) pt.
        let inner = bodies.iter().map(|b| b.fence).fold(0.0f32, f32::max);
        let target = if grow {
            let cover = 2.0 * (asc - axis).max(desc + axis) - size * 3.0 / 8.0;
            // A hair over the inner delimiter takes the next variant.
            cover.max(inner + 0.01)
        } else {
            0.0
        };
        let fence = |ch: char, class: Class| {
            let g = self.stretch(ch, target, size, true);
            let dy = axis - (g.asc - g.desc) / 2.0;
            let mut b = B {
                first: class,
                last: class,
                fence: g.asc + g.desc,
                ..B::empty()
            };
            b.put(g, 0.0, dy);
            b
        };
        let mut parts: Vec<B> = Vec::new();
        parts.extend(open.map(|c| fence(c, Class::Open)));
        let n = bodies.len();
        for (i, body) in bodies.into_iter().enumerate() {
            parts.push(body.ord());
            if let Some(sep) = sep.filter(|_| i + 1 < n) {
                parts.push(fence(sep, Class::Punct));
            }
        }
        parts.extend(close.map(|c| fence(c, Class::Close)));
        let mut b = self.hbox(parts, st);
        b.first_italic = false;
        b.ic = 0.0;
        b
    }

    fn accent(&self, chr: char, body: &MathRow, st: St) -> B {
        let base = self.row(body, st.cramp());
        let size = self.size(st);
        let attach = |b: &B| {
            b.single_glyph
                .and_then(|g| self.font.top_accent_attachment(g, size))
                .unwrap_or(b.w / 2.0)
        };
        // Word takes the widest accent that fits within the base on both
        // sides of its attachment point (𝐴's apex leaves the narrowest hat),
        // flattened over bases taller than FlattenedAccentBaseHeight; an
        // extensible accent (⃗) stretches to that room when its assembly
        // fits it (probe `mp22_acc`: every glyph as Word's).
        let base_attach = attach(&base);
        let room = 2.0 * base_attach.min(base.w - base_attach);
        let flat = base.asc > self.k(st, Constants::flattened_accent_base_height);
        let found = self.font.accent(self.font.glyph(chr, 0), room, size, flat);
        let acc = self.stretched(found, chr, size, false);
        // Accents are drawn for x-height letters; a taller base lifts them.
        let dy = (base.asc - self.k(st, Constants::accent_base_height)).max(0.0);
        let dx = base_attach - attach(&acc);
        let mut b = B::ends_of(&base);
        let w = base.w;
        b.put(base, 0.0, 0.0);
        b.put(acc, dx, dy);
        // An accent wider than its base doesn't widen it.
        b.w = w;
        b
    }

    fn bar(&self, top: bool, body: &MathRow, st: St) -> B {
        let base = if top {
            self.row(body, st.cramp())
        } else {
            self.row(body, st)
        };
        let (w, asc, desc) = (base.w, base.asc, base.desc);
        let mut b = B::empty();
        b.put(base, 0.0, 0.0);
        if top {
            let gap = self.k(st, Constants::overbar_vertical_gap);
            let t = self.k(st, Constants::overbar_rule_thickness);
            b.rule(0.0, asc + gap, w, t);
            b.asc = b
                .asc
                .max(asc + gap + t + self.k(st, Constants::overbar_extra_ascender));
        } else {
            let gap = self.k(st, Constants::underbar_vertical_gap);
            let t = self.k(st, Constants::underbar_rule_thickness);
            b.rule(0.0, -(desc + gap + t), w, t);
            b.desc = b
                .desc
                .max(desc + gap + t + self.k(st, Constants::underbar_extra_descender));
        }
        b
    }

    /// A brace or arrow over or under its content. Word keeps the character
    /// where its design puts it (a ⏟ hangs below the baseline) unless the
    /// content comes within the stretch-stack gap of it: x+y+z's descenders
    /// push the brace 1.44pt down, a+b+c's leave it. With `on_chr` (m:vertJc
    /// on the character's side) the character carries the baseline and the
    /// content is its script-size limit: Word sets a pos="top" brace's content
    /// at script size beneath it, m:vertJc defaulting to top. That limit sits
    /// a stretch stack's shift away (6.91pt down, 9.38 up at 12pt, whatever
    /// the character), further only to clear the character's ink by the
    /// stretch-stack gap (an a/b under an arrow). A bracket pair takes the
    /// form its side calls for: a ⏞ asked to go below draws as ⏟ (probes
    /// `mp9_gc`, `mp10_gc`).
    fn group_chr(&self, chr: char, top: bool, on_chr: bool, body: &MathRow, st: St) -> B {
        let chr = match (chr, top) {
            ('⏟', true) => '⏞',
            ('⏞', false) => '⏟',
            ('⏝', true) => '⏜',
            ('⏜', false) => '⏝',
            ('⎵', true) => '⎴',
            ('⎴', false) => '⎵',
            _ => chr,
        };
        let base = self.row(body, if on_chr { st.sup() } else { st });
        let brace = self.stretch(chr, base.w, self.size(st), false);
        // How far the character sits above (+) or below (-) its content:
        // the stretch-stack gap from the ink, at least the stretch-stack
        // shift when the content is the character's limit.
        let k = |c| self.k(st, c);
        let up = if top {
            let floor = if on_chr {
                k(Constants::stretch_stack_bottom_shift_down)
            } else {
                0.0
            };
            (base.asc + k(Constants::stretch_stack_gap_below_min) + brace.desc).max(floor)
        } else {
            let floor = if on_chr {
                k(Constants::stretch_stack_top_shift_up)
            } else {
                0.0
            };
            -(base.desc + k(Constants::stretch_stack_gap_above_min) + brace.asc).max(floor)
        };
        // The baseline stays with the character in limit mode, else the content.
        let (fixed, moved, dy) = if on_chr {
            (brace, base, -up)
        } else {
            (base, brace, up)
        };
        let w = fixed.w.max(moved.w);
        let mut b = B::ends_of(&fixed);
        let (fx, mx) = ((w - fixed.w) / 2.0, (w - moved.w) / 2.0);
        b.put(fixed, fx, 0.0);
        b.put(moved, mx, dy);
        b
    }

    /// An equation array. A row's `&`s split it into segments: every odd
    /// one is an alignment point, every even one a place for padding. Each
    /// segment keeps its place in the row laid out whole (a space between
    /// atoms goes with the segment before it), columns take the widest
    /// segment, and the segment before an alignment point is set flush right
    /// to it, the one after flush left. Rows without any `&` are centred
    /// (probe `mp11_eqarr`: Word's rows align to 0.25pt).
    fn eq_arr(&self, rows: &[MathRow], st: St) -> B {
        let mut laid = Vec::with_capacity(rows.len());
        for row in rows {
            let (mut atoms, mut marks) = (Vec::new(), Vec::new());
            for node in row {
                match node {
                    MathNode::Run(run) => {
                        for (i, text) in run.text.split('&').enumerate() {
                            if i > 0 {
                                marks.push(atoms.len());
                            }
                            atoms.extend(self.run_text(text, run, st));
                        }
                    }
                    _ => atoms.push(self.node(node, st)),
                }
            }
            let (xs, w) = self.place(&mut atoms, st);
            let mut bounds = vec![0.0];
            bounds.extend(marks.iter().map(|&m| xs.get(m).copied().unwrap_or(w)));
            bounds.push(w);
            laid.push((atoms, xs, marks, bounds));
        }
        if laid.iter().all(|(_, _, marks, _)| marks.is_empty()) {
            let rows = laid
                .into_iter()
                .map(|(atoms, xs, _, _)| vec![assemble(atoms, &xs, false).0])
                .collect();
            return self.grid(rows, st, false);
        }
        let ncols = laid
            .iter()
            .map(|(_, _, _, bounds)| bounds.len() - 1)
            .max()
            .unwrap_or(0);
        let mut col_w = vec![0.0f32; ncols];
        for (_, _, _, bounds) in &laid {
            for (i, seg) in bounds.windows(2).enumerate() {
                col_w[i] = col_w[i].max(seg[1] - seg[0]);
            }
        }
        let col_x: Vec<f32> = col_w
            .iter()
            .scan(0.0, |x, w| Some(std::mem::replace(x, *x + w)))
            .collect();
        let total: f32 = col_w.iter().sum();
        let amp = self.glyph('&', st, None);
        let rows = laid
            .into_iter()
            .map(|(atoms, xs, marks, bounds)| {
                // Segment s moves to its column, flush right before an
                // alignment point (even s), flush left after it.
                let dx: Vec<f32> = bounds
                    .windows(2)
                    .enumerate()
                    .map(|(s, seg)| {
                        let flush = if s % 2 == 0 {
                            col_w[s] - (seg[1] - seg[0])
                        } else {
                            0.0
                        };
                        col_x[s] + flush - seg[0]
                    })
                    .collect();
                let xs: Vec<f32> = xs
                    .iter()
                    .enumerate()
                    .map(|(j, x)| x + dx[marks.partition_point(|&m| m <= j)])
                    .collect();
                let (mut b, _) = assemble(atoms, &xs, false);
                b.w = total;
                // The `&`s Word measures as glyphs: an "x&=y" row stands as
                // tall as an ampersand (20.0pt under a display a/b row, 17.5
                // without the &).
                if !marks.is_empty() {
                    b.asc = b.asc.max(amp.asc);
                    b.desc = b.desc.max(amp.desc);
                }
                vec![b]
            })
            .collect();
        self.grid(rows, st, false)
    }

    /// Matrix cells (centred, columns an em apart) or equation-array rows,
    /// stacked a math line apart and centred on the math axis.
    fn grid(&self, rows: Vec<Vec<B>>, st: St, matrix: bool) -> B {
        let size = self.size(st);
        let (line_asc, line_desc) = self.font.line_metrics(size);
        let ncols = rows.iter().map(Vec::len).max().unwrap_or(0);
        let mut col_w = vec![0.0f32; ncols];
        for r in &rows {
            for (i, c) in r.iter().enumerate() {
                col_w[i] = col_w[i].max(c.w);
            }
        }
        let col_gap = if matrix { size } else { 0.0 };
        let total_w = col_w.iter().sum::<f32>() + col_gap * ncols.saturating_sub(1) as f32;
        // Rows step like lines of the math font (a math line apart for plain
        // contents: Word 14.16pt at 12pt), taller ink pushing them apart; the
        // grid is measured and centred on the axis by its ink. Equation-array
        // rows keep a math line apart unless their ink comes within
        // StackDisplayStyleGapMin (inline too): 17.5pt under a display a/b
        // row, 15.8 under an inline one (probe `mp13_eqarr`).
        let stack_gap = self.k(st, Constants::stack_display_style_gap_min);
        let mut rows_b = B::empty();
        let (mut y, mut prev_desc) = (0.0f32, None::<f32>);
        for r in rows {
            let asc = r.iter().map(|c| c.asc).fold(0.0, f32::max);
            let desc = r.iter().map(|c| c.desc).fold(0.0, f32::max);
            if let Some(pd) = prev_desc {
                y -= if matrix {
                    pd.max(line_desc) + asc.max(line_asc)
                } else {
                    (line_asc + line_desc).max(pd + asc + stack_gap)
                };
            }
            prev_desc = Some(desc);
            let mut x = 0.0;
            for (i, c) in r.into_iter().enumerate() {
                let cx = if matrix {
                    x + (col_w[i] - c.w) / 2.0
                } else {
                    (total_w - c.w) / 2.0
                };
                rows_b.put(c, cx, y);
                x += col_w[i] + col_gap;
            }
        }
        let dy = self.k(st, Constants::axis_height) - (rows_b.asc - rows_b.desc) / 2.0;
        let mut b = B::empty();
        b.put(rows_b, 0.0, dy);
        b.w = total_w;
        b
    }

    fn border_box(&self, hide: [bool; 4], body: &MathRow, st: St) -> B {
        let base = self.row(body, st);
        // An overbar's gap and rule all round: Word's border sits 1.94pt off
        // 12pt content with a 0.72pt rule (OverbarVerticalGap 2.02).
        let t = self.k(st, Constants::overbar_rule_thickness);
        let pad = self.k(st, Constants::overbar_vertical_gap) + t;
        let (w, asc, desc) = (base.w + 2.0 * pad, base.asc + pad, base.desc + pad);
        let mut b = B::empty();
        b.put(base, pad, 0.0);
        let [top, bot, left, right] = hide;
        if !top {
            b.rule(0.0, asc - t, w, t);
        }
        if !bot {
            b.rule(0.0, -desc, w, t);
        }
        if !left {
            b.rule(0.0, -desc, t, asc + desc);
        }
        if !right {
            b.rule(w - t, -desc, t, asc + desc);
        }
        b
    }

    /// The glyph for `ch` at least `target` tall (or wide): the base glyph,
    /// a larger variant, or an assembly of parts.
    fn stretch(&self, ch: char, target: f32, size: f32, vertical: bool) -> B {
        let gid = self.font.glyph(ch, 0);
        self.stretched(
            self.font.stretch(gid, target, size, vertical),
            ch,
            size,
            vertical,
        )
    }

    /// A stretched glyph or assembly as a box.
    fn stretched(&self, s: font::Stretched, ch: char, size: f32, vertical: bool) -> B {
        match s {
            font::Stretched::Glyph(g) => self.glyph_id(g, ch, size, None),
            font::Stretched::Assembly(parts) => {
                // The parts' own extents, which needn't reach the baseline (a
                // horizontal brace hangs below it).
                let mut b = B {
                    asc: f32::NEG_INFINITY,
                    desc: f32::NEG_INFINITY,
                    ..B::empty()
                };
                for (g, offset) in parts {
                    let (dx, dy) = if vertical {
                        (0.0, offset)
                    } else {
                        (offset, 0.0)
                    };
                    b.put(self.glyph_id(g, ch, size, None), dx, dy);
                }
                b
            }
        }
    }
}

/// The tallest ascent and deepest descent of a row's atoms.
fn extent(atoms: &[(f32, f32, f32)]) -> (f32, f32) {
    atoms.iter().fold((0.0f32, 0.0f32), |(a, d), &(_, aa, dd)| {
        (a.max(aa), d.max(dd))
    })
}

/// A row of atoms at `xs`, with where it may break: after each binary or
/// relation operator (`before`: before each one but the first atom), as (end
/// of the line's part, start of the next, at a relation).
fn assemble(atoms: Vec<B>, xs: &[f32], before: bool) -> (B, Vec<(f32, f32, bool)>) {
    let mut breaks = Vec::new();
    let mut out = atoms.first().map_or_else(B::empty, B::ends_of);
    let mut last = (Class::Ord, 0.0);
    for (i, (a, &x)) in atoms.into_iter().zip(xs).enumerate() {
        let c = if before { a.first } else { last.0 };
        if i > 0 && matches!(c, Class::Bin | Class::Rel) {
            breaks.push((out.w, x, c == Class::Rel));
        }
        last = (a.last, a.ic);
        out.put(a, x, 0.0);
    }
    (out.last, out.ic) = last;
    (out, breaks)
}

/// TeX's inter-atom space in eighteenths of an em; the medium and thick
/// spaces (and most thin ones) vanish in script styles.
fn space(left: Class, right: Class, st: St) -> f32 {
    use Class::*;
    let script = st.script_level() > 0;
    let (n, script_too) = match (left, right) {
        (Space, _) | (_, Space) => (0.0, true),
        (Ord, Op) | (Op, Ord) | (Op, Op) | (Close, Op) => (3.0, true),
        (Ord | Op | Close | Inner, Bin) | (Bin, Ord | Op | Open | Inner) => (4.0, false),
        (Ord | Op | Close | Inner, Rel) | (Rel, Ord | Op | Open | Inner) => (5.0, false),
        (Punct, Ord | Op | Rel | Open | Close | Punct | Inner) => (3.0, false),
        (Ord | Op | Close | Inner, Inner) | (Inner, Ord | Op | Open | Punct) => (3.0, false),
        _ => (0.0, true),
    };
    if script && !script_too { 0.0 } else { n }
}

/// The class of a character in a math run.
fn class_of(ch: char) -> Class {
    match ch {
        '+' | '\u{2212}' | '±' | '∓' | '×' | '÷' | '⋅' | '∙' | '·' | '∗' | '∘' | '∩' | '∪'
        | '⊕' | '⊖' | '⊗' | '⊘' | '⊙' | '∧' | '∨' | '∖' | '⋆' | '⋄' | '⊎' | '⊓' | '⊔' | '≀' => {
            Class::Bin
        }
        '='
        | '<'
        | '>'
        | '≤'
        | '≥'
        | '≠'
        | '≈'
        | '≡'
        | '≅'
        | '∼'
        | '≃'
        | '≪'
        | '≫'
        | '∝'
        | '∈'
        | '∉'
        | '∋'
        | '⊂'
        | '⊃'
        | '⊆'
        | '⊇'
        | '⊊'
        | '⊋'
        | '∣'
        | '∥'
        | '⊥'
        | '≺'
        | '≻'
        | '≼'
        | '≽'
        | '≔'
        | ':'
        | '≐'
        | '≜'
        | '≦'
        | '≧'
        | '≲'
        | '≳'
        | '⊢'
        | '⊨'
        | '∶'
        | '≍'
        | '≑'
        | '\u{2192}'..='\u{21FF}'
        | '\u{2190}'
        | '\u{2191}'
        | '\u{27F0}'..='\u{27FF}' => Class::Rel,
        '(' | '[' | '{' | '⟨' | '⌈' | '⌊' | '⟦' | '〈' => Class::Open,
        ')' | ']' | '}' | '⟩' | '⌉' | '⌋' | '⟧' | '〉' | '!' => Class::Close,
        ',' | ';' => Class::Punct,
        '∑' | '∏' | '∐' | '∫' | '∬' | '∭' | '∮' | '∯' | '∰' | '⋂' | '⋃' | '⋀' | '⋁' | '⨀' | '⨁'
        | '⨂' | '⨄' | '⨆' => Class::Op,
        _ => Class::Ord,
    }
}

#[cfg(test)]
mod tests;
