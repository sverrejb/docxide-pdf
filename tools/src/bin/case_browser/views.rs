//! The central image views: single, side-by-side, overlay, delta and compare.

use crate::App;
use crate::annotations::{
    Annotation, ImageClick, ImageSource, draw_annotation_markers, pending_from_click,
};
use crate::cases::file_age;
use eframe::egui;
use std::path::{Path, PathBuf};

/// The full [0,0]–[1,1] UV rectangle for drawing a texture untransformed.
/// (egui's `Rect::EVERYTHING` is the *infinite* rect, not this unit rect.)
pub const FULL_UV: egui::Rect =
    egui::Rect::from_min_max(egui::pos2(0.0, 0.0), egui::pos2(1.0, 1.0));

#[derive(Clone, Copy)]
pub enum ViewMode {
    SideBySide,
    Reference,
    Generated,
    Overlay,
    Delta,
    Compare,
}

/// Keyboard shortcuts that select a view mode. Several modes have two keys
/// (a number and a mnemonic letter), so duplicate targets are expected.
pub const VIEW_MODE_KEYS: &[(egui::Key, ViewMode)] = &[
    (egui::Key::Num1, ViewMode::SideBySide),
    (egui::Key::Num2, ViewMode::Reference),
    (egui::Key::R, ViewMode::Reference),
    (egui::Key::Num3, ViewMode::Generated),
    (egui::Key::G, ViewMode::Generated),
    (egui::Key::Num4, ViewMode::Overlay),
    (egui::Key::O, ViewMode::Overlay),
    (egui::Key::Num5, ViewMode::Delta),
    (egui::Key::D, ViewMode::Delta),
    (egui::Key::Num6, ViewMode::Compare),
    (egui::Key::C, ViewMode::Compare),
];

pub fn fit_size(tex: &egui::TextureHandle, max_w: f32, max_h: f32) -> egui::Vec2 {
    let aspect = tex.size()[0] as f32 / tex.size()[1] as f32;
    let w = max_h * aspect;
    if w > max_w {
        egui::vec2(max_w, max_w / aspect)
    } else {
        egui::vec2(w, max_h)
    }
}

pub fn draw_grid_overlay(ctx: &egui::Context, rect: egui::Rect, spacing: f32) {
    let mut painter = ctx.layer_painter(egui::LayerId::new(
        egui::Order::Foreground,
        egui::Id::new("grid_overlay"),
    ));
    painter.set_clip_rect(ctx.screen_rect());
    let color = egui::Color32::from_gray(140);
    let mut y = rect.top() + spacing;
    let mut thick = true;
    while y < rect.bottom() {
        let width = if thick { 1.0 } else { 0.5 };
        painter.line_segment(
            [egui::pos2(rect.left(), y), egui::pos2(rect.right(), y)],
            egui::Stroke::new(width, color),
        );
        y += spacing;
        thick = !thick;
    }
}

pub fn show_image_clickable(
    app: &mut App,
    ctx: &egui::Context,
    ui: &mut egui::Ui,
    path: &PathBuf,
    max_w: f32,
    max_h: f32,
    source: ImageSource,
    page: usize,
) -> Option<ImageClick> {
    let mut click = None;
    if let Some(tex) = app.load_texture(ctx, path) {
        let size = fit_size(&tex, max_w, max_h);
        let tex_size = [tex.size()[0], tex.size()[1]];
        let (rect, resp) = ui.allocate_exact_size(size, egui::Sense::click());
        ui.painter()
            .image(tex.id(), rect, FULL_UV, egui::Color32::WHITE);
        if app.show_grid {
            draw_grid_overlay(ctx, rect, app.grid_spacing);
        }
        if app.show_annotations {
            let case_name = &app.cases[app.current_case].name;
            draw_annotation_markers(ui, rect, tex_size, &app.annotations.annotations, case_name, page, source);
        }
        if resp.clicked() {
            if let Some(pos) = resp.interact_pointer_pos() {
                click = Some(ImageClick {
                    screen_pos: pos,
                    image_rect: rect,
                    tex_size,
                });
            }
        }
    } else {
        ui.label("(not found)");
    }
    click
}

pub fn show_single(
    app: &mut App,
    ctx: &egui::Context,
    ui: &mut egui::Ui,
    path: &PathBuf,
    label: &str,
    source: ImageSource,
    page: usize,
) {
    let age = file_age(path).unwrap_or_default();
    ui.horizontal(|ui| {
        ui.label(label);
        ui.label(egui::RichText::new(age).weak());
    });
    let available = ui.available_size();
    let click = egui::ScrollArea::both()
        .show(ui, |ui| {
            show_image_clickable(app, ctx, ui, path, available.x, available.y - 20.0, source, page)
        })
        .inner;
    if let Some(click) = click {
        app.pending_annotation = Some(pending_from_click(&click, page, source));
    }
}

pub fn show_side_by_side(
    app: &mut App,
    ctx: &egui::Context,
    ui: &mut egui::Ui,
    ref_path: &PathBuf,
    gen_path: &PathBuf,
    page: usize,
) {
    let ref_age = file_age(ref_path).unwrap_or_default();
    let gen_age = file_age(gen_path).unwrap_or_default();

    let available = ui.available_size();
    let half_w = (available.x - 10.0) / 2.0;
    let max_h = available.y - 30.0;

    let mut ref_click = None;
    let mut gen_click = None;

    ui.horizontal_top(|ui| {
        ui.vertical(|ui| {
            ui.horizontal(|ui| {
                ui.label("Reference");
                ui.label(egui::RichText::new(&ref_age).weak());
            });
            ref_click = show_image_clickable(
                app, ctx, ui, ref_path, half_w, max_h, ImageSource::Reference, page,
            );
        });

        ui.separator();

        ui.vertical(|ui| {
            ui.horizontal(|ui| {
                ui.label("Generated");
                ui.label(egui::RichText::new(&gen_age).weak());
            });
            gen_click = show_image_clickable(
                app, ctx, ui, gen_path, half_w, max_h, ImageSource::Generated, page,
            );
        });
    });

    let click_info = ref_click
        .map(|c| (c, ImageSource::Reference))
        .or_else(|| gen_click.map(|c| (c, ImageSource::Generated)));
    if let Some((click, source)) = click_info {
        app.pending_annotation = Some(pending_from_click(&click, page, source));
    }
}

pub struct OverlayResult {
    pub texture: egui::TextureHandle,
    pub change_rects: Vec<[u32; 4]>,
}

pub fn build_overlay_texture(
    ctx: &egui::Context,
    ref_path: &Path,
    gen_path: &Path,
    key: (usize, usize),
) -> Option<OverlayResult> {
    let ref_img = image::open(ref_path).ok()?.to_rgba8();
    let gen_img = image::open(gen_path).ok()?.to_rgba8();
    let w = ref_img.width().min(gen_img.width());
    let h = ref_img.height().min(gen_img.height());

    let cell = 32u32;
    let gw = (w + cell - 1) / cell;
    let gh = (h + cell - 1) / cell;
    // Count changed pixels per cell (not a bool) so isolated antialiasing flips
    // near the luma-200 ink threshold don't trip a whole cell — see find_change_rects.
    let mut grid = vec![0u32; (gw * gh) as usize];

    let mut pixels = vec![0u8; (w * h * 4) as usize];
    for y in 0..h as usize {
        for x in 0..w as usize {
            let rp = ref_img.get_pixel(x as u32, y as u32).0;
            let gp = gen_img.get_pixel(x as u32, y as u32).0;
            let ref_ink = luma(rp[0], rp[1], rp[2]) < 200;
            let gen_ink = luma(gp[0], gp[1], gp[2]) < 200;
            let c = match (ref_ink, gen_ink) {
                (true, true) => [80, 80, 80, 255],
                (true, false) => [0, 80, 220, 255],
                (false, true) => [220, 40, 40, 255],
                (false, false) => [255, 255, 255, 255],
            };
            if ref_ink != gen_ink {
                let gx = x as u32 / cell;
                let gy = y as u32 / cell;
                grid[(gy * gw + gx) as usize] += 1;
            }
            let i = (y * w as usize + x) * 4;
            pixels[i..i + 4].copy_from_slice(&c);
        }
    }

    let change_rects = find_change_rects(&grid, gw, gh, cell, w, h);

    let image = egui::ColorImage::from_rgba_unmultiplied([w as usize, h as usize], &pixels);
    let name = format!("overlay_{}_{}", key.0, key.1);
    Some(OverlayResult {
        texture: ctx.load_texture(name, image, egui::TextureOptions::LINEAR),
        change_rects,
    })
}

/// Find bounding rectangles around clusters of changed pixels using a grid-based
/// connected-component approach. Returns `[x, y, width, height]` in pixel coords.
///
/// `grid` holds the count of differing pixels per 32×32 cell. Two thresholds keep
/// the boxes off invisible noise: a cell only counts as changed with `CELL_MIN`+
/// differing pixels (one stray antialiased edge pixel isn't a change), and a whole
/// cluster is dropped unless it holds `CLUSTER_MIN`+ differing pixels total.
pub fn find_change_rects(
    grid: &[u32],
    gw: u32,
    gh: u32,
    cell: u32,
    img_w: u32,
    img_h: u32,
) -> Vec<[u32; 4]> {
    const CELL_MIN: u32 = 8;
    const CLUSTER_MIN: u32 = 40;
    let mut visited = vec![false; grid.len()];
    let mut rects = Vec::new();

    for gy in 0..gh {
        for gx in 0..gw {
            let idx = (gy * gw + gx) as usize;
            if grid[idx] < CELL_MIN || visited[idx] {
                continue;
            }
            // Flood fill (8-connected) to find connected component
            let mut min_x = gx;
            let mut min_y = gy;
            let mut max_x = gx;
            let mut max_y = gy;
            let mut total = 0u32;
            let mut stack = vec![(gx, gy)];
            while let Some((cx, cy)) = stack.pop() {
                let ci = (cy * gw + cx) as usize;
                if visited[ci] || grid[ci] < CELL_MIN {
                    continue;
                }
                visited[ci] = true;
                total += grid[ci];
                min_x = min_x.min(cx);
                min_y = min_y.min(cy);
                max_x = max_x.max(cx);
                max_y = max_y.max(cy);
                // 8-connected neighbors
                for dy in -1i32..=1 {
                    for dx in -1i32..=1 {
                        if dx == 0 && dy == 0 {
                            continue;
                        }
                        let nx = cx as i32 + dx;
                        let ny = cy as i32 + dy;
                        if nx >= 0 && (nx as u32) < gw && ny >= 0 && (ny as u32) < gh {
                            stack.push((nx as u32, ny as u32));
                        }
                    }
                }
            }
            if total < CLUSTER_MIN {
                continue;
            }
            let pad = 8u32;
            let x0 = (min_x * cell).saturating_sub(pad);
            let y0 = (min_y * cell).saturating_sub(pad);
            let x1 = ((max_x + 1) * cell + pad).min(img_w);
            let y1 = ((max_y + 1) * cell + pad).min(img_h);
            rects.push([x0, y0, x1 - x0, y1 - y0]);
        }
    }

    rects
}

pub fn luma(r: u8, g: u8, b: u8) -> u8 {
    ((r as u16 * 77 + g as u16 * 150 + b as u16 * 29) >> 8) as u8
}

pub fn show_overlay(app: &mut App, ctx: &egui::Context, ui: &mut egui::Ui, page: usize) {
    let key = (app.current_case, page);
    if !app.overlay_cache.contains_key(&key) {
        let ref_path = app.page_path("reference", page);
        let gen_path = app.page_path("generated", page);
        let result = build_overlay_texture(ctx, &ref_path, &gen_path, key);
        app.overlay_cache.insert(key, result);
    }

    ui.horizontal(|ui| {
        ui.label("Overlay");
        ui.colored_label(egui::Color32::from_rgb(0, 80, 220), "Blue=ref only");
        ui.colored_label(egui::Color32::from_rgb(220, 40, 40), "Red=gen only");
        ui.label("Gray=both");
    });

    let available = ui.available_size();
    if let Some(Some(result)) = app.overlay_cache.get(&key) {
        let tex = result.texture.clone();
        let tex_size = [tex.size()[0], tex.size()[1]];
        let show_grid = app.show_grid;
        let grid_spacing = app.grid_spacing;
        let show_annotations = app.show_annotations;
        let annotations_snapshot: Vec<Annotation> = app.annotations.annotations.clone();
        let case_name = app.cases[app.current_case].name.clone();
        let click = egui::ScrollArea::both()
            .show(ui, |ui| {
                let size = fit_size(&tex, available.x, available.y - 20.0);
                let (rect, resp) = ui.allocate_exact_size(size, egui::Sense::click());
                ui.painter().image(
                    tex.id(),
                    rect,
                    FULL_UV,
                    egui::Color32::WHITE,
                );
                if show_grid {
                    draw_grid_overlay(ui.ctx(), rect, grid_spacing);
                }
                if show_annotations {
                    draw_annotation_markers(
                        ui, rect, tex_size, &annotations_snapshot, &case_name, page, ImageSource::Reference,
                    );
                    draw_annotation_markers(
                        ui, rect, tex_size, &annotations_snapshot, &case_name, page, ImageSource::Generated,
                    );
                }
                if resp.clicked() {
                    resp.interact_pointer_pos().map(|pos| ImageClick {
                        screen_pos: pos,
                        image_rect: rect,
                        tex_size,
                    })
                } else {
                    None
                }
            })
            .inner;
        if let Some(click) = click {
            app.pending_annotation =
                Some(pending_from_click(&click, page, ImageSource::Generated));
        }
    } else {
        ui.label("Could not load reference and/or generated images");
    }
}

pub fn show_compare(app: &mut App, ctx: &egui::Context, ui: &mut egui::Ui, page: usize) {
    let ref_path = app.page_path("reference", page);
    let gen_path = app.page_path("generated", page);
    let lo_path = app.page_path("libreoffice", page);

    let ref_age = file_age(&ref_path).unwrap_or_default();
    let gen_age = file_age(&gen_path).unwrap_or_default();
    let lo_age = file_age(&lo_path).unwrap_or_default();

    let available = ui.available_size();
    let third_w = (available.x - 20.0) / 3.0;
    let max_h = available.y - 30.0;

    let mut ref_click = None;
    let mut gen_click = None;

    ui.horizontal_top(|ui| {
        ui.vertical(|ui| {
            ui.horizontal(|ui| {
                ui.label("Reference (Word)");
                ui.label(egui::RichText::new(&ref_age).weak());
            });
            ref_click = show_image_clickable(
                app, ctx, ui, &ref_path, third_w, max_h, ImageSource::Reference, page,
            );
        });

        ui.separator();

        ui.vertical(|ui| {
            ui.horizontal(|ui| {
                ui.label(egui::RichText::new("docxide-pdf").strong());
                ui.label(egui::RichText::new(&gen_age).weak());
            });
            gen_click = show_image_clickable(
                app, ctx, ui, &gen_path, third_w, max_h, ImageSource::Generated, page,
            );
        });

        ui.separator();

        ui.vertical(|ui| {
            ui.horizontal(|ui| {
                ui.label("LibreOffice");
                if lo_path.exists() {
                    ui.label(egui::RichText::new(&lo_age).weak());
                }
            });
            show_image_static(app, ctx, ui, &lo_path, third_w, max_h);
        });
    });

    let click_info = ref_click
        .map(|c| (c, ImageSource::Reference))
        .or_else(|| gen_click.map(|c| (c, ImageSource::Generated)));
    if let Some((click, source)) = click_info {
        app.pending_annotation = Some(pending_from_click(&click, page, source));
    }
}

pub fn show_image_static(
    app: &mut App,
    ctx: &egui::Context,
    ui: &mut egui::Ui,
    path: &PathBuf,
    max_w: f32,
    max_h: f32,
) {
    if !path.exists() {
        ui.allocate_ui(egui::vec2(max_w, max_h), |ui| {
            ui.centered_and_justified(|ui| {
                ui.weak("Not rendered.\nRun: cargo test --test suite libreoffice_comparison");
            });
        });
        return;
    }
    if let Some(tex) = app.load_texture(ctx, path) {
        let size = fit_size(&tex, max_w, max_h);
        let (rect, _) = ui.allocate_exact_size(size, egui::Sense::hover());
        ui.painter().image(
            tex.id(),
            rect,
            FULL_UV,
            egui::Color32::WHITE,
        );
        if app.show_grid {
            draw_grid_overlay(ctx, rect, app.grid_spacing);
        }
    } else {
        ui.label("(failed to load)");
    }
}

pub fn show_delta(app: &mut App, ctx: &egui::Context, ui: &mut egui::Ui, page: usize) {
    let ack_path = app.page_path("acknowledged", page);
    if !ack_path.exists() {
        ui.centered_and_justified(|ui| {
            ui.label("No acknowledged snapshot — acknowledge a visual change first");
        });
        return;
    }

    let key = (app.current_case, page);
    if !app.delta_cache.contains_key(&key) {
        let gen_path = app.page_path("generated", page);
        let result = build_overlay_texture(ctx, &ack_path, &gen_path, key);
        app.delta_cache.insert(key, result);
    }

    ui.horizontal(|ui| {
        ui.label("Delta (acknowledged vs generated)");
        ui.colored_label(egui::Color32::from_rgb(0, 80, 220), "Blue=removed");
        ui.colored_label(egui::Color32::from_rgb(220, 40, 40), "Red=added");
        ui.label("Gray=unchanged");
    });

    let available = ui.available_size();
    if let Some(Some(result)) = app.delta_cache.get(&key) {
        let tex = result.texture.clone();
        let tex_size = [tex.size()[0], tex.size()[1]];
        let change_rects = result.change_rects.clone();
        let show_grid = app.show_grid;
        let grid_spacing = app.grid_spacing;
        let click = egui::ScrollArea::both()
            .show(ui, |ui| {
                let size = fit_size(&tex, available.x, available.y - 20.0);
                let (rect, resp) = ui.allocate_exact_size(size, egui::Sense::click());
                ui.painter().image(
                    tex.id(),
                    rect,
                    FULL_UV,
                    egui::Color32::WHITE,
                );
                if show_grid {
                    draw_grid_overlay(ui.ctx(), rect, grid_spacing);
                }
                // Draw red outlines around change regions
                let scale_x = rect.width() / tex_size[0] as f32;
                let scale_y = rect.height() / tex_size[1] as f32;
                for &[rx, ry, rw, rh] in &change_rects {
                    let outline = egui::Rect::from_min_size(
                        egui::pos2(
                            rect.min.x + rx as f32 * scale_x,
                            rect.min.y + ry as f32 * scale_y,
                        ),
                        egui::vec2(rw as f32 * scale_x, rh as f32 * scale_y),
                    );
                    ui.painter().rect_stroke(
                        outline,
                        0.0,
                        egui::Stroke::new(2.0, egui::Color32::from_rgb(255, 0, 0)),
                        egui::StrokeKind::Outside,
                    );
                }
                if resp.clicked() {
                    resp.interact_pointer_pos().map(|pos| ImageClick {
                        screen_pos: pos,
                        image_rect: rect,
                        tex_size,
                    })
                } else {
                    None
                }
            })
            .inner;
        if let Some(click) = click {
            app.pending_annotation =
                Some(pending_from_click(&click, page, ImageSource::Generated));
        }
    } else {
        ui.label("Could not load acknowledged and/or generated images");
    }
}

pub fn central(app: &mut App, ctx: &egui::Context, ui: &mut egui::Ui) {
    let page = app.current_page;
    match app.view_mode {
        ViewMode::SideBySide => {
            let ref_path = app.page_path("reference", page);
            let gen_path = app.page_path("generated", page);
            show_side_by_side(app, ctx, ui, &ref_path, &gen_path, page);
        }
        ViewMode::Reference => {
            let path = app.page_path("reference", page);
            show_single(app, ctx, ui, &path, "Reference", ImageSource::Reference, page);
        }
        ViewMode::Generated => {
            let path = app.page_path("generated", page);
            show_single(app, ctx, ui, &path, "Generated", ImageSource::Generated, page);
        }
        ViewMode::Overlay => {
            show_overlay(app, ctx, ui, page);
        }
        ViewMode::Delta => {
            show_delta(app, ctx, ui, page);
        }
        ViewMode::Compare => {
            show_compare(app, ctx, ui, page);
        }
    }
}
