use eframe::egui;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::time::SystemTime;

const RENDER_DPI: f32 = 150.0;
const PIXELS_PER_POINT: f32 = RENDER_DPI / 72.0;

fn main() -> eframe::Result<()> {
    let output_dir = find_output_dir();
    let cases = discover_cases(&output_dir);

    if cases.is_empty() {
        eprintln!("No cases found in {}", output_dir.display());
        std::process::exit(1);
    }

    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_inner_size([1400.0, 900.0])
            .with_title("Case Browser"),
        ..Default::default()
    };

    eframe::run_native(
        "Case Browser",
        options,
        Box::new(move |cc| {
            egui_extras::install_image_loaders(&cc.egui_ctx);
            Ok(Box::new(App::new(cases, output_dir)))
        }),
    )
}

fn find_output_dir() -> PathBuf {
    // Try relative to CWD first, then walk up
    let candidates = [
        PathBuf::from("tests/output"),
        PathBuf::from("../tests/output"),
    ];
    for c in &candidates {
        if c.is_dir() {
            return c.clone();
        }
    }
    eprintln!("Could not find tests/output directory. Run from the project root.");
    std::process::exit(1);
}

#[derive(Clone)]
struct CaseInfo {
    name: String,
    dir: PathBuf,
    page_count: usize,
}

fn discover_cases(output_dir: &Path) -> Vec<CaseInfo> {
    let mut cases = Vec::new();

    // Scan cases/, scraped/, samples/ subdirs
    for subdir in &["cases", "scraped", "samples"] {
        let dir = output_dir.join(subdir);
        if !dir.is_dir() {
            continue;
        }
        let Ok(entries) = std::fs::read_dir(&dir) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if !path.is_dir() {
                continue;
            }
            let ref_dir = path.join("reference");
            let gen_dir = path.join("generated");
            if !ref_dir.is_dir() && !gen_dir.is_dir() {
                continue;
            }
            let page_count = count_pages(&path);
            if page_count == 0 {
                continue;
            }
            let name = if *subdir == "cases" {
                entry.file_name().to_string_lossy().into_owned()
            } else {
                format!("{}/{}", subdir, entry.file_name().to_string_lossy())
            };
            cases.push(CaseInfo {
                name,
                dir: path,
                page_count,
            });
        }
    }

    cases.sort_by(|a, b| natural_sort_key(&a.name).cmp(&natural_sort_key(&b.name)));
    cases
}

fn natural_sort_key(s: &str) -> Vec<NatPart> {
    let mut parts = Vec::new();
    let mut chars = s.chars().peekable();
    while chars.peek().is_some() {
        if chars.peek().unwrap().is_ascii_digit() {
            let mut num = String::new();
            while chars.peek().is_some() && chars.peek().unwrap().is_ascii_digit() {
                num.push(chars.next().unwrap());
            }
            parts.push(NatPart::Num(num.parse().unwrap_or(0)));
        } else {
            let mut text = String::new();
            while chars.peek().is_some() && !chars.peek().unwrap().is_ascii_digit() {
                text.push(chars.next().unwrap());
            }
            parts.push(NatPart::Str(text));
        }
    }
    parts
}

#[derive(Clone, PartialEq, Eq, PartialOrd, Ord)]
enum NatPart {
    Str(String),
    Num(u64),
}

// --- Annotation types ---

#[derive(Clone, Copy, PartialEq, Serialize, Deserialize)]
enum ImageSource {
    Reference,
    Generated,
}

#[derive(Clone, Serialize, Deserialize)]
struct Annotation {
    id: u64,
    case: String,
    page: usize,
    source: ImageSource,
    x_pt: f32,
    y_pt: f32,
    note: String,
    #[serde(default)]
    fixed: bool,
}

#[derive(Default, Serialize, Deserialize)]
struct AnnotationStore {
    next_id: u64,
    annotations: Vec<Annotation>,
}

fn annotations_path(output_dir: &Path) -> PathBuf {
    output_dir.join("annotations.json")
}

fn load_annotations(output_dir: &Path) -> AnnotationStore {
    let path = annotations_path(output_dir);
    std::fs::read_to_string(&path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

fn save_annotations(output_dir: &Path, store: &AnnotationStore) {
    let path = annotations_path(output_dir);
    if let Ok(json) = serde_json::to_string_pretty(store) {
        let _ = std::fs::write(path, json);
    }
}

// --- Baseline scores ---

#[derive(Default, Deserialize)]
struct BaselineScores {
    jaccard: Option<f64>,
    ssim: Option<f64>,
}

fn baseline_key(name: &str) -> String {
    let (group, case_name) = if let Some(idx) = name.find('/') {
        (&name[..idx], &name[idx + 1..])
    } else {
        ("cases", name)
    };
    let short = if case_name.len() > 16 {
        format!("{}..", &case_name[..16])
    } else {
        case_name.to_string()
    };
    format!("{}/{}", group, short)
}

fn load_baselines(output_dir: &Path) -> HashMap<String, BaselineScores> {
    let path = output_dir
        .parent()
        .unwrap_or(output_dir)
        .join("baselines.json");
    std::fs::read_to_string(&path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

fn score_color(value: f64, threshold: f64) -> egui::Color32 {
    if value >= threshold {
        egui::Color32::from_rgb(0, 180, 0)
    } else {
        egui::Color32::from_rgb(220, 40, 40)
    }
}

struct PendingAnnotation {
    page: usize,
    source: ImageSource,
    x_pt: f32,
    y_pt: f32,
    text_buf: String,
}

struct ImageClick {
    screen_pos: egui::Pos2,
    image_rect: egui::Rect,
    tex_size: [usize; 2],
}

fn screen_to_pdf_pts(click: &ImageClick) -> (f32, f32) {
    let rel_x = click.screen_pos.x - click.image_rect.left();
    let rel_y = click.screen_pos.y - click.image_rect.top();
    let scale = click.tex_size[0] as f32 / click.image_rect.width();
    let x_pt = (rel_x * scale) / PIXELS_PER_POINT;
    let page_height_pt = click.tex_size[1] as f32 / PIXELS_PER_POINT;
    let y_from_top = (rel_y * scale) / PIXELS_PER_POINT;
    let y_pt = page_height_pt - y_from_top;
    (x_pt, y_pt)
}

fn pdf_pts_to_screen(
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

fn count_pages(case_dir: &Path) -> usize {
    // Count from whichever has more pages (reference or generated)
    let r = count_pngs(&case_dir.join("reference"));
    let g = count_pngs(&case_dir.join("generated"));
    r.max(g)
}

fn count_pngs(dir: &Path) -> usize {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return 0;
    };
    entries
        .flatten()
        .filter(|e| {
            e.path()
                .extension()
                .is_some_and(|ext| ext.eq_ignore_ascii_case("png"))
        })
        .count()
}

enum ViewMode {
    SideBySide,
    Reference,
    Generated,
    Overlay,
}

struct App {
    cases: Vec<CaseInfo>,
    output_dir: PathBuf,
    current_case: usize,
    current_page: usize,
    view_mode: ViewMode,
    texture_cache: HashMap<PathBuf, Option<egui::TextureHandle>>,
    overlay_cache: HashMap<(usize, usize), Option<egui::TextureHandle>>,
    scroll_to_current: bool,
    show_grid: bool,
    grid_spacing: f32,
    refresh_flash: f32,
    annotations: AnnotationStore,
    pending_annotation: Option<PendingAnnotation>,
    show_annotations: bool,
    show_notes_panel: bool,
    baselines: HashMap<String, BaselineScores>,
}

impl App {
    fn new(cases: Vec<CaseInfo>, output_dir: PathBuf) -> Self {
        let annotations = load_annotations(&output_dir);
        let baselines = load_baselines(&output_dir);
        Self {
            cases,
            output_dir,
            current_case: 0,
            current_page: 0,
            view_mode: ViewMode::SideBySide,
            texture_cache: HashMap::new(),
            overlay_cache: HashMap::new(),
            scroll_to_current: true,
            show_grid: false,
            grid_spacing: 18.0,
            refresh_flash: 0.0,
            annotations,
            pending_annotation: None,
            show_annotations: true,
            show_notes_panel: true,
            baselines,
        }
    }

    fn current(&self) -> &CaseInfo {
        &self.cases[self.current_case]
    }

    fn set_case(&mut self, idx: usize) {
        if idx != self.current_case {
            self.current_case = idx;
            self.current_page = 0;
            self.scroll_to_current = true;
            self.texture_cache.clear();
            self.overlay_cache.clear();
            self.pending_annotation = None;
        }
    }

    fn save_annotations(&self) {
        save_annotations(&self.output_dir, &self.annotations);
    }

    fn refresh(&mut self) {
        self.texture_cache.clear();
        self.overlay_cache.clear();
        self.refresh_flash = 1.0;
    }

    fn page_path(&self, subdir: &str, page: usize) -> PathBuf {
        self.current()
            .dir
            .join(subdir)
            .join(format!("page_{:03}.png", page + 1))
    }

    fn load_texture(
        &mut self,
        ctx: &egui::Context,
        path: &PathBuf,
    ) -> Option<egui::TextureHandle> {
        if let Some(cached) = self.texture_cache.get(path) {
            return cached.clone();
        }

        let result = if path.exists() {
            match image::open(path) {
                Ok(img) => {
                    let rgba = img.to_rgba8();
                    let size = [rgba.width() as usize, rgba.height() as usize];
                    let pixels = rgba.into_raw();
                    let image = egui::ColorImage::from_rgba_unmultiplied(size, &pixels);
                    let name = path.to_string_lossy().to_string();
                    Some(ctx.load_texture(name, image, egui::TextureOptions::LINEAR))
                }
                Err(_) => None,
            }
        } else {
            None
        };

        self.texture_cache.insert(path.clone(), result.clone());
        result
    }
}

impl eframe::App for App {
    fn update(&mut self, ctx: &egui::Context, _frame: &mut eframe::Frame) {
        // Handle keyboard input (skip when text input has focus)
        let text_has_focus = self.pending_annotation.is_some()
            && ctx.memory(|m| m.focused().is_some());
        ctx.input(|i| {
            if i.key_pressed(egui::Key::Escape) {
                self.pending_annotation = None;
            }
            if text_has_focus {
                return;
            }
            if i.key_pressed(egui::Key::ArrowRight) {
                let next = (self.current_case + 1).min(self.cases.len() - 1);
                self.set_case(next);
            }
            if i.key_pressed(egui::Key::ArrowLeft) {
                let next = self.current_case.saturating_sub(1);
                self.set_case(next);
            }
            if i.key_pressed(egui::Key::ArrowDown) {
                self.current_page = (self.current_page + 1).min(self.current().page_count - 1);
            }
            if i.key_pressed(egui::Key::ArrowUp) {
                self.current_page = self.current_page.saturating_sub(1);
            }
            if i.key_pressed(egui::Key::Num1) || i.key_pressed(egui::Key::S) {
                self.view_mode = ViewMode::SideBySide;
            }
            if i.key_pressed(egui::Key::Num2) || i.key_pressed(egui::Key::R) {
                self.view_mode = ViewMode::Reference;
            }
            if i.key_pressed(egui::Key::Num3) || i.key_pressed(egui::Key::G) {
                self.view_mode = ViewMode::Generated;
            }
            if i.key_pressed(egui::Key::Num4) || i.key_pressed(egui::Key::O) {
                self.view_mode = ViewMode::Overlay;
            }
            if i.key_pressed(egui::Key::F5) || i.key_pressed(egui::Key::F) {
                self.refresh();
            }
            if i.key_pressed(egui::Key::L) {
                self.show_grid = !self.show_grid;
            }
            if i.key_pressed(egui::Key::Equals) || i.key_pressed(egui::Key::Plus) {
                self.grid_spacing = (self.grid_spacing + 2.0).min(100.0);
            }
            if i.key_pressed(egui::Key::Minus) {
                self.grid_spacing = (self.grid_spacing - 2.0).max(4.0);
            }
            if i.key_pressed(egui::Key::A) {
                self.show_annotations = !self.show_annotations;
            }
            if i.key_pressed(egui::Key::N) {
                self.show_notes_panel = !self.show_notes_panel;
            }
        });

        // Right panel: case list
        let case_labels: Vec<(String, usize)> = self
            .cases
            .iter()
            .enumerate()
            .map(|(i, c)| {
                let key = baseline_key(&c.name);
                let label = if let Some(b) = self.baselines.get(&key) {
                    let j = b.jaccard.map(|v| format!("{:.0}", v * 100.0)).unwrap_or_else(|| "-".into());
                    let s = b.ssim.map(|v| format!("{:.0}", v * 100.0)).unwrap_or_else(|| "-".into());
                    format!("{} ({}p) {}/{}", c.name, c.page_count, j, s)
                } else {
                    format!("{} ({}p)", c.name, c.page_count)
                };
                (label, i)
            })
            .collect();
        let cur = self.current_case;
        let scroll = self.scroll_to_current;

        let mut clicked_case = None;
        let mut did_scroll = false;

        egui::SidePanel::right("case_list")
            .default_width(260.0)
            .show(ctx, |ui| {
                ui.heading("Cases");
                ui.separator();
                egui::ScrollArea::vertical().show(ui, |ui| {
                    for (label, i) in &case_labels {
                        let selected = *i == cur;
                        let resp = ui.selectable_label(selected, label);
                        if resp.clicked() {
                            clicked_case = Some(*i);
                        }
                        if selected && scroll {
                            resp.scroll_to_me(Some(egui::Align::Center));
                            did_scroll = true;
                        }
                    }
                });
            });

        if let Some(idx) = clicked_case {
            self.set_case(idx);
        }
        if did_scroll {
            self.scroll_to_current = false;
        }

        // Top bar: case name and view mode
        egui::TopBottomPanel::top("top_bar").show(ctx, |ui| {
            ui.horizontal(|ui| {
                let case = &self.cases[self.current_case];
                ui.strong(&case.name);
                ui.separator();
                let mode_label = match self.view_mode {
                    ViewMode::SideBySide => "[S]ide-by-side",
                    ViewMode::Reference => "[R]eference",
                    ViewMode::Generated => "[G]enerated",
                    ViewMode::Overlay => "[O]verlay",
                };
                ui.label(format!("View: {}", mode_label));
                let key = baseline_key(&case.name);
                if let Some(b) = self.baselines.get(&key) {
                    ui.separator();
                    if let Some(j) = b.jaccard {
                        ui.colored_label(
                            score_color(j, 0.20),
                            format!("Jaccard: {:.1}%", j * 100.0),
                        );
                    }
                    if let Some(s) = b.ssim {
                        ui.colored_label(
                            score_color(s, 0.75),
                            format!("SSIM: {:.1}%", s * 100.0),
                        );
                    }
                }
                ui.separator();
                if self.show_grid {
                    ui.label(format!("Grid: {:.0}px (+/-)", self.grid_spacing));
                }
                ui.separator();
                ui.label("[L]ines  [F]refresh  [A]nnot  [N]otes");
                if self.refresh_flash > 0.0 {
                    let alpha = (self.refresh_flash * 255.0) as u8;
                    ui.label(
                        egui::RichText::new("Refreshed")
                            .color(egui::Color32::from_rgba_unmultiplied(0, 180, 0, alpha)),
                    );
                }
            });
        });

        if self.refresh_flash > 0.0 {
            self.refresh_flash = (self.refresh_flash - 0.05).max(0.0);
            ctx.request_repaint();
        }

        // Bottom bar: page number
        egui::TopBottomPanel::bottom("bottom_bar").show(ctx, |ui| {
            ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                let case = &self.cases[self.current_case];
                ui.label(
                    egui::RichText::new(format!(
                        "Page {}/{}",
                        self.current_page + 1,
                        case.page_count
                    ))
                    .size(18.0),
                );
            });
        });

        // Notes panel (left side)
        if self.show_notes_panel {
            let page = self.current_page;
            let case_name = self.cases[self.current_case].name.clone();
            let page_ann_ids: Vec<u64> = self
                .annotations
                .annotations
                .iter()
                .filter(|a| a.case == case_name && a.page == page)
                .map(|a| a.id)
                .collect();
            let mut delete_id = None;
            let mut toggle_fixed_id = None;
            egui::SidePanel::left("notes_panel")
                .default_width(220.0)
                .show(ctx, |ui| {
                    ui.heading("Notes");
                    ui.separator();
                    if page_ann_ids.is_empty() {
                        ui.weak("Click an image to add a note");
                    }
                    egui::ScrollArea::vertical().show(ui, |ui| {
                        for (display_idx, ann_id) in page_ann_ids.iter().enumerate() {
                            let ann = self.annotations.annotations.iter().find(|a| a.id == *ann_id).unwrap();
                            let color = if ann.fixed {
                                egui::Color32::from_gray(140)
                            } else {
                                match ann.source {
                                    ImageSource::Reference => egui::Color32::from_rgb(255, 165, 0),
                                    ImageSource::Generated => egui::Color32::from_rgb(0, 180, 80),
                                }
                            };
                            let src = match ann.source {
                                ImageSource::Reference => "Ref",
                                ImageSource::Generated => "Gen",
                            };
                            ui.horizontal(|ui| {
                                ui.colored_label(
                                    color,
                                    format!("#{}", display_idx + 1),
                                );
                                ui.label(format!(
                                    "{} ({:.0}, {:.0})",
                                    src, ann.x_pt, ann.y_pt
                                ));
                                if ui
                                    .small_button("X")
                                    .on_hover_text("Delete this note")
                                    .clicked()
                                {
                                    delete_id = Some(*ann_id);
                                }
                            });
                            if ann.fixed {
                                ui.label(egui::RichText::new(&ann.note).strikethrough());
                            } else {
                                ui.label(&ann.note);
                            }
                            let mut fixed = ann.fixed;
                            if ui.checkbox(&mut fixed, "Fixed").changed() {
                                toggle_fixed_id = Some(*ann_id);
                            }
                            ui.separator();
                        }
                    });
                });
            if let Some(id) = delete_id {
                self.annotations.annotations.retain(|a| a.id != id);
                self.save_annotations();
            }
            if let Some(id) = toggle_fixed_id {
                if let Some(ann) = self.annotations.annotations.iter_mut().find(|a| a.id == id) {
                    ann.fixed = !ann.fixed;
                }
                self.save_annotations();
            }
        }

        // Annotation input popup
        let mut commit_annotation = false;
        let mut cancel_annotation = false;
        if self.pending_annotation.is_some() {
            egui::Window::new("Add Note")
                .collapsible(false)
                .resizable(false)
                .show(ctx, |ui| {
                    let pending = self.pending_annotation.as_mut().unwrap();
                    let src = match pending.source {
                        ImageSource::Reference => "Reference",
                        ImageSource::Generated => "Generated",
                    };
                    ui.label(format!(
                        "{} page {} at ({:.0}, {:.0}) pt",
                        src,
                        pending.page + 1,
                        pending.x_pt,
                        pending.y_pt
                    ));
                    let resp = ui.text_edit_singleline(&mut pending.text_buf);
                    resp.request_focus();
                    let cmd_enter = ui.input(|i| {
                        i.key_pressed(egui::Key::Enter) && i.modifiers.command
                    });
                    if cmd_enter
                        || (resp.lost_focus()
                            && ui.input(|i| i.key_pressed(egui::Key::Enter)))
                    {
                        commit_annotation = true;
                    }
                    ui.horizontal(|ui| {
                        if ui.button("Save").clicked() {
                            commit_annotation = true;
                        }
                        if ui.button("Cancel").clicked() {
                            cancel_annotation = true;
                        }
                    });
                });
        }
        if commit_annotation {
            if let Some(pending) = self.pending_annotation.take() {
                if !pending.text_buf.trim().is_empty() {
                    let id = self.annotations.next_id;
                    self.annotations.next_id += 1;
                    self.annotations.annotations.push(Annotation {
                        id,
                        case: self.cases[self.current_case].name.clone(),
                        page: pending.page,
                        source: pending.source,
                        x_pt: pending.x_pt,
                        y_pt: pending.y_pt,
                        note: pending.text_buf.trim().to_string(),
                        fixed: false,
                    });
                    self.save_annotations();
                }
            }
        }
        if cancel_annotation {
            self.pending_annotation = None;
        }

        // Central panel: images
        egui::CentralPanel::default().show(ctx, |ui| {
            let page = self.current_page;
            match self.view_mode {
                ViewMode::SideBySide => {
                    let ref_path = self.page_path("reference", page);
                    let gen_path = self.page_path("generated", page);
                    show_side_by_side(self, ctx, ui, &ref_path, &gen_path, page);
                }
                ViewMode::Reference => {
                    let path = self.page_path("reference", page);
                    show_single(self, ctx, ui, &path, "Reference", ImageSource::Reference, page);
                }
                ViewMode::Generated => {
                    let path = self.page_path("generated", page);
                    show_single(self, ctx, ui, &path, "Generated", ImageSource::Generated, page);
                }
                ViewMode::Overlay => {
                    show_overlay(self, ctx, ui, page);
                }
            }
        });
    }
}

fn fit_size(tex: &egui::TextureHandle, max_w: f32, max_h: f32) -> egui::Vec2 {
    let aspect = tex.size()[0] as f32 / tex.size()[1] as f32;
    let w = max_h * aspect;
    if w > max_w {
        egui::vec2(max_w, max_w / aspect)
    } else {
        egui::vec2(w, max_h)
    }
}

fn draw_grid_overlay(ctx: &egui::Context, rect: egui::Rect, spacing: f32) {
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

fn show_image_clickable(
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
            .image(tex.id(), rect, egui::Rect::from_min_max(egui::pos2(0.0, 0.0), egui::pos2(1.0, 1.0)), egui::Color32::WHITE);
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

fn draw_annotation_markers(
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
        .filter(|a| a.case == case_name && a.page == page && a.source == source)
        .collect();
    for (i, ann) in filtered.iter().enumerate() {
        let center = pdf_pts_to_screen(ann.x_pt, ann.y_pt, &image_rect, tex_size);
        if !image_rect.contains(center) {
            continue;
        }
        let color = if ann.fixed {
            egui::Color32::from_gray(160)
        } else {
            match ann.source {
                ImageSource::Reference => egui::Color32::from_rgb(255, 165, 0),
                ImageSource::Generated => egui::Color32::from_rgb(0, 180, 80),
            }
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

fn show_single(
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
        let (x_pt, y_pt) = screen_to_pdf_pts(&click);
        app.pending_annotation = Some(PendingAnnotation {
            page,
            source,
            x_pt,
            y_pt,
            text_buf: String::new(),
        });
    }
}

fn show_side_by_side(
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
        let (x_pt, y_pt) = screen_to_pdf_pts(&click);
        app.pending_annotation = Some(PendingAnnotation {
            page,
            source,
            x_pt,
            y_pt,
            text_buf: String::new(),
        });
    }
}

fn build_overlay_texture(
    ctx: &egui::Context,
    ref_path: &Path,
    gen_path: &Path,
    key: (usize, usize),
) -> Option<egui::TextureHandle> {
    let ref_img = image::open(ref_path).ok()?.to_rgba8();
    let gen_img = image::open(gen_path).ok()?.to_rgba8();
    let w = ref_img.width().min(gen_img.width());
    let h = ref_img.height().min(gen_img.height());

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
            let i = (y * w as usize + x) * 4;
            pixels[i..i + 4].copy_from_slice(&c);
        }
    }

    let image = egui::ColorImage::from_rgba_unmultiplied([w as usize, h as usize], &pixels);
    let name = format!("overlay_{}_{}", key.0, key.1);
    Some(ctx.load_texture(name, image, egui::TextureOptions::LINEAR))
}

fn luma(r: u8, g: u8, b: u8) -> u8 {
    ((r as u16 * 77 + g as u16 * 150 + b as u16 * 29) >> 8) as u8
}

fn file_age(path: &Path) -> Option<String> {
    let modified = std::fs::metadata(path).ok()?.modified().ok()?;
    let elapsed = SystemTime::now().duration_since(modified).ok()?;
    let total_mins = elapsed.as_secs() / 60;
    let hours = total_mins / 60;
    let mins = total_mins % 60;
    if hours > 0 {
        Some(format!("{}h {}m ago", hours, mins))
    } else {
        Some(format!("{}m ago", mins))
    }
}

fn show_overlay(app: &mut App, ctx: &egui::Context, ui: &mut egui::Ui, page: usize) {
    let key = (app.current_case, page);
    if !app.overlay_cache.contains_key(&key) {
        let ref_path = app.page_path("reference", page);
        let gen_path = app.page_path("generated", page);
        let tex = build_overlay_texture(ctx, &ref_path, &gen_path, key);
        app.overlay_cache.insert(key, tex);
    }

    ui.horizontal(|ui| {
        ui.label("Overlay");
        ui.colored_label(egui::Color32::from_rgb(0, 80, 220), "Blue=ref only");
        ui.colored_label(egui::Color32::from_rgb(220, 40, 40), "Red=gen only");
        ui.label("Gray=both");
    });

    let available = ui.available_size();
    if let Some(Some(tex)) = app.overlay_cache.get(&key) {
        let tex = tex.clone();
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
                    egui::Rect::from_min_max(egui::pos2(0.0, 0.0), egui::pos2(1.0, 1.0)),
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
            let (x_pt, y_pt) = screen_to_pdf_pts(&click);
            app.pending_annotation = Some(PendingAnnotation {
                page,
                source: ImageSource::Generated,
                x_pt,
                y_pt,
                text_buf: String::new(),
            });
        }
    } else {
        ui.label("Could not load reference and/or generated images");
    }
}
