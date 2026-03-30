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

/// Draw an image drop shadow using layered semi-transparent rectangles to
/// approximate gaussian blur. Each layer is slightly larger and overlaps,
/// so the center accumulates to the target alpha while edges fade out.
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
        // Fallback: single pre-blended rectangle
        let a = shadow.alpha;
        let blended = [
            (a * shadow.color[0] as f32 + (1.0 - a) * 255.0) as u8,
            (a * shadow.color[1] as f32 + (1.0 - a) * 255.0) as u8,
            (a * shadow.color[2] as f32 + (1.0 - a) * 255.0) as u8,
        ];
        let expand = shadow.blur_radius * 0.5;
        content.save_state();
        fill_color_or_black(content, Some(blended));
        content.rect(sx - expand, sy - expand, width + expand * 2.0, height + expand * 2.0);
        content.fill_nonzero();
        content.restore_state();
        return;
    };

    const NUM_LAYERS: usize = 10;

    // Solve for per-layer alpha so stacking NUM_LAYERS gives target alpha:
    // 1 - (1 - per_alpha)^N = target  =>  per_alpha = 1 - (1-target)^(1/N)
    let per_alpha = 1.0 - (1.0 - shadow.alpha).powf(1.0 / NUM_LAYERS as f32);
    let pct = (per_alpha * 100.0).round().max(1.0).min(100.0) as u8;
    alpha_states.insert(pct);

    let gs_name_str = format!("GSa{pct}");

    content.save_state();
    fill_rgb(content, shadow.color);
    content.set_parameters(Name(gs_name_str.as_bytes()));

    // Draw layers from outermost (largest expansion) to innermost (no expansion).
    // Each is a full rectangle so inner areas get covered by all layers.
    for i in (0..NUM_LAYERS).rev() {
        let t = (i + 1) as f32 / NUM_LAYERS as f32;
        let expand = shadow.blur_radius * t;
        content.rect(
            sx - expand,
            sy - expand,
            width + expand * 2.0,
            height + expand * 2.0,
        );
        content.fill_nonzero();
    }

    // Restore full opacity and previous state
    alpha_states.insert(100);
    content.set_parameters(Name(b"GSa100"));
    content.restore_state();
}
