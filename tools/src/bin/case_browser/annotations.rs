//! Click-to-annotate notes: the JSON store, screen <-> PDF-point conversion and markers.

use eframe::egui;
use serde::{Deserialize, Serialize};
use std::path::{Path, PathBuf};

pub const RENDER_DPI: f32 = 150.0;
pub const PIXELS_PER_POINT: f32 = RENDER_DPI / 72.0;

// --- Annotation types ---

#[derive(Clone, Copy, PartialEq, Serialize, Deserialize)]
pub enum ImageSource {
    Reference,
    Generated,
}

#[derive(Clone, Serialize, Deserialize)]
pub struct Annotation {
    pub id: u64,
    pub case: String,
    pub page: usize,
    pub source: ImageSource,
    pub x_pt: f32,
    pub y_pt: f32,
    pub note: String,
    #[serde(default)]
    pub fixed: bool,
    // Cleared annotations stay in the JSON (kept as a record) but are hidden
    // from the UI. Clearing replaces deletion as the "resolved" signal.
    #[serde(default)]
    pub cleared: bool,
}

#[derive(Default, Serialize, Deserialize)]
pub struct AnnotationStore {
    pub next_id: u64,
    pub annotations: Vec<Annotation>,
}

pub fn annotations_path(output_dir: &Path) -> PathBuf {
    output_dir.join("annotations.json")
}

pub fn load_annotations(output_dir: &Path) -> AnnotationStore {
    let path = annotations_path(output_dir);
    std::fs::read_to_string(&path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

pub fn save_annotations(output_dir: &Path, store: &AnnotationStore) {
    let path = annotations_path(output_dir);
    if let Ok(json) = serde_json::to_string_pretty(store) {
        let _ = std::fs::write(path, json);
    }
}

/// Marker/label color for an annotation by its image source (orange for
/// reference, green for generated). Fixed annotations are greyed out by the
/// caller instead.
pub fn source_color(source: ImageSource) -> egui::Color32 {
    match source {
        ImageSource::Reference => egui::Color32::from_rgb(255, 165, 0),
        ImageSource::Generated => egui::Color32::from_rgb(0, 180, 80),
    }
}

pub struct PendingAnnotation {
    pub page: usize,
    pub source: ImageSource,
    pub x_pt: f32,
    pub y_pt: f32,
    pub text_buf: String,
}

pub struct ImageClick {
    pub screen_pos: egui::Pos2,
    pub image_rect: egui::Rect,
    pub tex_size: [usize; 2],
}

pub fn screen_to_pdf_pts(click: &ImageClick) -> (f32, f32) {
    let rel_x = click.screen_pos.x - click.image_rect.left();
    let rel_y = click.screen_pos.y - click.image_rect.top();
    let scale = click.tex_size[0] as f32 / click.image_rect.width();
    let x_pt = (rel_x * scale) / PIXELS_PER_POINT;
    let page_height_pt = click.tex_size[1] as f32 / PIXELS_PER_POINT;
    let y_from_top = (rel_y * scale) / PIXELS_PER_POINT;
    let y_pt = page_height_pt - y_from_top;
    (x_pt, y_pt)
}

/// Build a pending annotation from an image click, converting screen coords to
/// PDF points. Shared by every view that lets you click to drop a note.
pub fn pending_from_click(
    click: &ImageClick,
    page: usize,
    source: ImageSource,
) -> PendingAnnotation {
    let (x_pt, y_pt) = screen_to_pdf_pts(click);
    PendingAnnotation {
        page,
        source,
        x_pt,
        y_pt,
        text_buf: String::new(),
    }
}

pub fn pdf_pts_to_screen(
    x_pt: f32,
    y_pt: f32,
    image_rect: &egui::Rect,
    tex_size: [usize; 2],
) -> egui::Pos2 {
    let scale = tex_size[0] as f32 / image_rect.width();
    let page_height_pt = tex_size[1] as f32 / PIXELS_PER_POINT;
    let y_from_top = page_height_pt - y_pt;
    let sx = image_rect.left() + (x_pt * PIXELS_PER_POINT) / scale;
    let sy = image_rect.top() + (y_from_top * PIXELS_PER_POINT) / scale;
    egui::pos2(sx, sy)
}

pub fn draw_annotation_markers(
    ui: &egui::Ui,
    image_rect: egui::Rect,
    tex_size: [usize; 2],
    annotations: &[Annotation],
    case_name: &str,
    page: usize,
    source: ImageSource,
) {
    let painter = ui.painter();
    let filtered: Vec<&Annotation> = annotations
        .iter()
        .filter(|a| a.case == case_name && a.page == page && a.source == source && !a.cleared)
        .collect();
    for (i, ann) in filtered.iter().enumerate() {
        let center = pdf_pts_to_screen(ann.x_pt, ann.y_pt, &image_rect, tex_size);
        if !image_rect.contains(center) {
            continue;
        }
        let color = if ann.fixed {
            egui::Color32::from_gray(160)
        } else {
            source_color(ann.source)
        };
        painter.circle_filled(center, 10.0, color);
        painter.text(
            center,
            egui::Align2::CENTER_CENTER,
            format!("{}", i + 1),
            egui::FontId::proportional(11.0),
            egui::Color32::WHITE,
        );
        let hover_rect = egui::Rect::from_center_size(center, egui::vec2(20.0, 20.0));
        if ui.rect_contains_pointer(hover_rect) {
            egui::show_tooltip(ui.ctx(), ui.layer_id(), ui.id().with(ann.id), |ui| {
                ui.label(&ann.note);
            });
        }
    }
}
