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

/// Pixels per point for shadow mask rasterization. Blur is low-frequency,
/// so half-resolution is sufficient; PDF `interpolate(true)` smooths the rest.
pub(super) const SHADOW_SCALE: f32 = 0.5;

/// Generate a grayscale blur mask for a drop shadow.
/// Returns `(pixel_data, pixel_width, pixel_height)`.
/// The interior rectangle is opaque at `alpha`; the exterior fades via
/// 3-pass box blur (approximates Gaussian).
pub(super) fn generate_shadow_mask(
    width_pt: f32,
    height_pt: f32,
    blur_radius: f32,
    alpha: f32,
) -> (Vec<u8>, u32, u32) {
    let total_w = width_pt + 2.0 * blur_radius;
    let total_h = height_pt + 2.0 * blur_radius;
    let px_w = (total_w * SHADOW_SCALE).ceil().max(1.0) as u32;
    let px_h = (total_h * SHADOW_SCALE).ceil().max(1.0) as u32;

    let mut buf = vec![0u8; (px_w * px_h) as usize];

    // Fill interior rectangle (the shadow body before blur)
    let pad_x = (blur_radius * SHADOW_SCALE).round() as u32;
    let pad_y = (blur_radius * SHADOW_SCALE).round() as u32;
    let inner_val = (alpha * 255.0).round().min(255.0) as u8;
    let inner_w = px_w.saturating_sub(2 * pad_x);
    let inner_h = px_h.saturating_sub(2 * pad_y);
    for row in 0..inner_h {
        let y = pad_y + row;
        if y >= px_h {
            break;
        }
        for col in 0..inner_w {
            let x = pad_x + col;
            if x >= px_w {
                break;
            }
            buf[(y * px_w + x) as usize] = inner_val;
        }
    }

    // 3-pass box blur approximates Gaussian blur
    let box_r = ((blur_radius * SHADOW_SCALE) / 3.0).ceil().max(0.0) as usize;
    if box_r > 0 {
        let w = px_w as usize;
        let h = px_h as usize;
        let mut tmp = vec![0u8; w * h];
        for _pass in 0..3 {
            // Horizontal pass: buf -> tmp
            for y in 0..h {
                let mut acc: u32 = 0;
                // Seed accumulator with first (box_r+1) pixels
                for x in 0..=box_r.min(w - 1) {
                    acc += buf[y * w + x] as u32;
                }
                for x in 0..w {
                    // Add entering pixel on the right
                    let right = x + box_r;
                    if right < w && x > 0 {
                        acc += buf[y * w + right] as u32;
                    }
                    // Compute divisor: number of pixels actually in the window
                    let left_edge = if x > box_r { x - box_r } else { 0 };
                    let right_edge = right.min(w - 1);
                    let count = (right_edge - left_edge + 1) as u32;
                    tmp[y * w + x] = (acc / count).min(255) as u8;
                    // Remove leaving pixel on the left
                    if x >= box_r {
                        acc -= buf[y * w + (x - box_r)] as u32;
                    }
                }
            }
            // Vertical pass: tmp -> buf
            for x in 0..w {
                let mut acc: u32 = 0;
                for y in 0..=box_r.min(h - 1) {
                    acc += tmp[y * w + x] as u32;
                }
                for y in 0..h {
                    let bottom = y + box_r;
                    if bottom < h && y > 0 {
                        acc += tmp[bottom * w + x] as u32;
                    }
                    let top_edge = if y > box_r { y - box_r } else { 0 };
                    let bottom_edge = bottom.min(h - 1);
                    let count = (bottom_edge - top_edge + 1) as u32;
                    buf[y * w + x] = (acc / count).min(255) as u8;
                    if y >= box_r {
                        acc -= tmp[(y - box_r) * w + x] as u32;
                    }
                }
            }
        }
    }

    (buf, px_w, px_h)
}

/// Draw an image drop shadow by referencing a pre-embedded shadow XObject.
///
/// When `shadow_xobj_name` is provided, draws the rasterized blur-mask shadow
/// image at the correct offset position. When `None`, falls back to a single
/// pre-blended rectangle (for zero-blur or when embedding was skipped).
pub(super) fn draw_image_shadow(
    content: &mut Content,
    shadow: &crate::model::ImageShadow,
    x: f32,
    y_bottom: f32,
    width: f32,
    height: f32,
    shadow_xobj_name: Option<&str>,
) {
    use pdf_writer::Name;

    if let Some(name) = shadow_xobj_name {
        let blur = shadow.blur_radius;
        // Shadow image covers (width + 2*blur) x (height + 2*blur), centered
        // on the shadow offset position
        let sx = x + shadow.offset_x - blur;
        // PDF y-up: offset_y is in screen coords (positive = down)
        let sy = y_bottom - shadow.offset_y - blur;
        let sw = width + 2.0 * blur;
        let sh = height + 2.0 * blur;

        content.save_state();
        content.transform([sw, 0.0, 0.0, sh, sx, sy]);
        content.x_object(Name(name.as_bytes()));
        content.restore_state();
        return;
    }

    // Fallback: single pre-blended rectangle (no blur mask available)
    let sx = x + shadow.offset_x;
    let sy = y_bottom - shadow.offset_y;
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
}
