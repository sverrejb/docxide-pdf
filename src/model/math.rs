//! Office Math (OMML, ISO/IEC 29500-1 §22.1) as a tree, laid out by `crate::math`.

/// One `m:oMath` zone: inline in a paragraph, or one equation of an
/// `m:oMathPara` (display math). Its run's size is the zone's.
#[derive(Debug)]
pub struct MathZone {
    pub root: MathRow,
    pub display: bool,
}

pub type MathRow = Vec<MathNode>;

impl MathZone {
    /// Every run in the zone, nested ones included.
    pub fn runs(&self) -> Vec<&MathRun> {
        let mut out = Vec::new();
        let mut stack: Vec<&MathRow> = vec![&self.root];
        while let Some(row) = stack.pop() {
            for node in row {
                match node {
                    MathNode::Run(r) => out.push(r),
                    _ => stack.extend(node.rows()),
                }
            }
        }
        out
    }
}

impl MathNode {
    /// The rows directly inside this node.
    fn rows(&self) -> Vec<&MathRow> {
        use MathNode::*;
        match self {
            Run(_) => Vec::new(),
            Frac { num, den, .. } => vec![num, den],
            Scripts { base, sub, sup, .. } => [Some(base), sub.as_ref(), sup.as_ref()]
                .into_iter()
                .flatten()
                .collect(),
            Rad { degree, body } => degree.iter().chain([body]).collect(),
            Nary { sub, sup, body, .. } => sub.iter().chain(sup).chain([body]).collect(),
            Delim { items, .. } => items.iter().collect(),
            Func { name, body } => vec![name, body],
            Limit { body, limit, .. } => vec![body, limit],
            Matrix { rows } => rows.iter().flatten().collect(),
            EqArr { rows } => rows.iter().collect(),
            Accent { body, .. }
            | Bar { body, .. }
            | GroupChr { body, .. }
            | Boxed(body)
            | BorderBox { body, .. }
            | Phantom { body, .. } => vec![body],
        }
    }
}

#[derive(Debug, Clone)]
pub enum MathNode {
    Run(MathRun),
    Frac {
        kind: FracKind,
        num: MathRow,
        den: MathRow,
    },
    /// `m:sSub`, `m:sSup`, `m:sSubSup` and (with `pre`) `m:sPre`.
    Scripts {
        base: MathRow,
        sub: Option<MathRow>,
        sup: Option<MathRow>,
        pre: bool,
    },
    Rad {
        degree: Option<MathRow>,
        body: MathRow,
    },
    Nary {
        chr: char,
        /// `m:limLoc`; None takes the document default (`m:intLim`/`m:naryLim`).
        under_over: Option<bool>,
        sub: Option<MathRow>,
        sup: Option<MathRow>,
        grow: bool,
        body: MathRow,
    },
    Delim {
        open: Option<char>,
        close: Option<char>,
        sep: Option<char>,
        grow: bool,
        items: Vec<MathRow>,
    },
    Func {
        name: MathRow,
        body: MathRow,
    },
    Accent {
        chr: char,
        body: MathRow,
    },
    Bar {
        top: bool,
        body: MathRow,
    },
    GroupChr {
        chr: char,
        top: bool,
        /// `m:vertJc` is top (Word's default when absent).
        vert_top: bool,
        body: MathRow,
    },
    /// `m:limLow` (`upper` false) and `m:limUpp`.
    Limit {
        upper: bool,
        body: MathRow,
        limit: MathRow,
    },
    Matrix {
        rows: Vec<Vec<MathRow>>,
    },
    EqArr {
        rows: Vec<MathRow>,
    },
    Boxed(MathRow),
    BorderBox {
        /// Hidden sides: top, bottom, left, right.
        hide: [bool; 4],
        body: MathRow,
    },
    Phantom {
        show: bool,
        zero_width: bool,
        zero_ascent: bool,
        zero_descent: bool,
        body: MathRow,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum FracKind {
    Bar,
    Skewed,
    Linear,
    NoBar,
}

#[derive(Debug, Clone)]
pub struct MathRun {
    pub text: String,
    /// `m:sty`; None: Word's default, italic letters and upright everything else.
    pub style: Option<MathStyle>,
    pub script: MathScript,
    /// `m:nor`: ordinary text in the run's own font, no math italic and no
    /// operator spacing.
    pub normal: Option<TextFont>,
    pub color: Option<[u8; 3]>,
    /// Struck through: deleted math in the markup view.
    pub strike: bool,
}

/// The font an `m:nor` run is drawn in: its `w:rPr` resolved like a text
/// run's.
#[derive(Debug, Clone, PartialEq)]
pub struct TextFont {
    pub name: String,
    pub bold: bool,
    pub italic: bool,
    pub size: f32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MathStyle {
    Plain,
    Bold,
    Italic,
    BoldItalic,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum MathScript {
    #[default]
    Roman,
    Script,
    Fraktur,
    DoubleStruck,
    SansSerif,
    Monospace,
}

/// Document-level math properties (`m:mathPr` in settings.xml, §22.1.2.62).
#[derive(Debug, Clone)]
pub struct MathSettings {
    pub font: String,
    /// `m:smallFrac`: fractions at script size even in display math.
    pub small_frac: bool,
    /// Limits of an integral go under/over it (`m:intLim` undOvr); Word's
    /// default puts them beside it.
    pub int_under_over: bool,
    /// Limits of other n-ary operators go under/over (`m:naryLim`, default undOvr).
    pub nary_under_over: bool,
    /// `m:wrapIndent` in points: how far a display equation's continuation
    /// lines are indented (default 1440 twips).
    pub wrap_indent: f32,
    /// `m:brkBin` after: a display equation breaks after an operator, not
    /// before it.
    pub break_after: bool,
}

impl Default for MathSettings {
    fn default() -> Self {
        Self {
            font: "Cambria Math".to_string(),
            small_frac: false,
            int_under_over: false,
            nary_under_over: true,
            wrap_indent: 72.0,
            break_after: false,
        }
    }
}
