use pdf_writer::Content;

/// Set fill color from an RGB byte array.
pub(super) fn fill_rgb(content: &mut Content, [r, g, b]: [u8; 3]) {
    content.set_fill_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
}

/// Set stroke color from an RGB byte array.
pub(super) fn stroke_rgb(content: &mut Content, [r, g, b]: [u8; 3]) {
    content.set_stroke_rgb(r as f32 / 255.0, g as f32 / 255.0, b as f32 / 255.0);
}

/// Set fill color from an optional RGB byte array, defaulting to black.
pub(super) fn fill_color_or_black(content: &mut Content, color: Option<[u8; 3]>) {
    if let Some(c) = color {
        fill_rgb(content, c);
    } else {
        content.set_fill_gray(0.0);
    }
}

/// Set stroke color from an optional RGB byte array, defaulting to black.
pub(super) fn stroke_color_or_black(content: &mut Content, color: Option<[u8; 3]>) {
    if let Some(c) = color {
        stroke_rgb(content, c);
    } else {
        content.set_stroke_gray(0.0);
    }
}
