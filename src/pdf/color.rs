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

/// Draw an image drop shadow with soft edges using layered semi-transparent
/// rectangles. Layers expand in the shadow-cast direction with a quadratic
/// distribution that concentrates opacity near the image edge.
///
/// When `alpha_states` is provided, uses proper PDF transparency (ExtGState).
/// When `None`, falls back to a single pre-blended rectangle.
pub(super) fn draw_image_shadow(
    content: &mut Content,
    shadow: &crate::model::ImageShadow,
    x: f32,
    y_bottom: f32,
    width: f32,
    height: f32,
    alpha_states: Option<&mut std::collections::HashSet<u8>>,
) {
    use pdf_writer::Name;

    let sx = x + shadow.offset_x;
    let sy = y_bottom - shadow.offset_y;

    let Some(alpha_states) = alpha_states else {
        let a = shadow.alpha;
        let blended = [
            (a * shadow.color[0] as f32 + (1.0 - a) * 255.0) as u8,
            (a * shadow.color[1] as f32 + (1.0 - a) * 255.0) as u8,
            (a * shadow.color[2] as f32 + (1.0 - a) * 255.0) as u8,
        ];
        content.save_state();
        fill_color_or_black(content, Some(blended));
        content.rect(sx, sy, width, height);
        content.fill_nonzero();
        content.restore_state();
        return;
    };

    const NUM_LAYERS: usize = 10;
    let blur = shadow.blur_radius * 0.75;

    let per_alpha = 1.0 - (1.0 - shadow.alpha).powf(1.0 / NUM_LAYERS as f32);
    let pct = (per_alpha * 100.0).round().max(1.0).min(100.0) as u8;
    alpha_states.insert(pct);

    // Expand primarily in the shadow-cast direction (the image covers the other sides)
    let exp_right = if shadow.offset_x >= 0.0 { 1.0 } else { 0.0 };
    let exp_left = if shadow.offset_x <= 0.0 { 1.0 } else { 0.0 };
    // PDF y-up: offset_y>0 means screen-down = expand in -y
    let exp_down = if shadow.offset_y >= 0.0 { 1.0 } else { 0.0 };
    let exp_up = if shadow.offset_y <= 0.0 { 1.0 } else { 0.0 };

    content.save_state();
    fill_rgb(content, shadow.color);
    content.set_parameters(Name(format!("GSa{pct}").as_bytes()));

    // Quadratic distribution: layers are denser near the image (inner edge)
    // producing a more concentrated shadow body with gentle outer fade
    for i in (0..NUM_LAYERS).rev() {
        let frac = (i + 1) as f32 / NUM_LAYERS as f32;
        let t = blur * frac * frac;
        let lx = sx - t * exp_left;
        let ly = sy - t * exp_down;
        let lw = width + t * (exp_left + exp_right);
        let lh = height + t * (exp_up + exp_down);
        content.rect(lx, ly, lw, lh);
        content.fill_nonzero();
    }

    alpha_states.insert(100);
    content.set_parameters(Name(b"GSa100"));
    content.restore_state();
}
