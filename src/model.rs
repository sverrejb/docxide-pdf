#[allow(dead_code)]
pub struct Document {
    pub page_width: f32,
    pub page_height: f32,
    pub margin_top: f32,
    pub margin_bottom: f32,
    pub margin_left: f32,
    pub margin_right: f32,
    pub line_pitch: f32,
    pub paragraphs: Vec<Paragraph>,
}

pub struct Paragraph {
    pub runs: Vec<Run>,
    pub space_before: f32,
    pub space_after: f32,
}

#[allow(dead_code)]
pub struct Run {
    pub text: String,
    pub font_size: f32,
    pub font_name: String,
    pub bold: bool,
    pub italic: bool,
}
