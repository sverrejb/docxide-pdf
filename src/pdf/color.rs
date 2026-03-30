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

/// Draw an image drop shadow as a filled rect expanded by the blur radius.
/// `x`, `y_bottom` are the image's bottom-left corner in PDF coords.
/// `offset_y` is positive-down (screen); PDF y-axis is up so we subtract.
pub(super) fn draw_image_shadow(
    content: &mut Content,
    shadow: &crate::model::ImageShadow,
    x: f32,
    y_bottom: f32,
    width: f32,
    height: f32,
) {
    let expand = shadow.blur_radius * 0.5;
    content.save_state();
    fill_color_or_black(content, Some(shadow.color));
    content.rect(
        x + shadow.offset_x - expand,
        y_bottom - shadow.offset_y - expand,
        width + expand * 2.0,
        height + expand * 2.0,
    );
    content.fill_nonzero();
    content.restore_state();
}
