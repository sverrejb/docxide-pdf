use eframe::egui;
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
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

    // Scan all subdirs (cases/, scraped/, samples/, fonts/, new/, etc.)
    let Ok(top_entries) = std::fs::read_dir(output_dir) else {
        return cases;
    };
    for top_entry in top_entries.flatten() {
        let subdir_path = top_entry.path();
        if !subdir_path.is_dir() {
            continue;
        }
        let subdir = top_entry.file_name().to_string_lossy().into_owned();
        let Ok(entries) = std::fs::read_dir(&subdir_path) else {
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
            let name = if subdir == "cases" {
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

fn load_hashes(path: &Path) -> HashMap<String, Vec<String>> {
    std::fs::read_to_string(path)
        .ok()
        .and_then(|s| serde_json::from_str(&s).ok())
        .unwrap_or_default()
}

fn compute_changed_cases(
    committed: &HashMap<String, Vec<String>>,
    latest: &HashMap<String, Vec<String>>,
) -> HashSet<String> {
    latest
        .iter()
        .filter(|(name, hashes)| committed.get(*name) != Some(hashes))
        .map(|(name, _)| name.clone())
        .collect()
}

/// Copy generated page PNGs to acknowledged/ as a snapshot of the current output.
fn snapshot_acknowledged(case_dir: &Path) {
    let gen_dir = case_dir.join("generated");
    let ack_dir = case_dir.join("acknowledged");
    let Ok(entries) = std::fs::read_dir(&gen_dir) else {
        return;
    };
    let _ = std::fs::create_dir_all(&ack_dir);
    // Remove old acknowledged files first
    if let Ok(old) = std::fs::read_dir(&ack_dir) {
        for e in old.flatten() {
            let _ = std::fs::remove_file(e.path());
        }
    }
    for entry in entries.flatten() {
        let path = entry.path();
        if path.extension().is_some_and(|e| e.eq_ignore_ascii_case("png")) {
            let dest = ack_dir.join(path.file_name().unwrap());
            let _ = std::fs::copy(&path, &dest);
        }
    }
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
    Delta,
}

struct App {
    cases: Vec<CaseInfo>,
    output_dir: PathBuf,
    current_case: usize,
    current_page: usize,
    view_mode: ViewMode,
    texture_cache: HashMap<PathBuf, Option<egui::TextureHandle>>,
    overlay_cache: HashMap<(usize, usize), Option<OverlayResult>>,
    delta_cache: HashMap<(usize, usize), Option<OverlayResult>>,
    scroll_to_current: bool,
    show_grid: bool,
    grid_spacing: f32,
    refresh_flash: f32,
    annotations: AnnotationStore,
    pending_annotation: Option<PendingAnnotation>,
    show_annotations: bool,
    show_notes_panel: bool,
    baselines: HashMap<String, BaselineScores>,
    visual_hashes: HashMap<String, Vec<String>>,
    latest_hashes: HashMap<String, Vec<String>>,
    changed_cases: HashSet<String>,
    collapsed_groups: HashSet<String>,
    search_query: String,
    search_active: bool,
}

impl App {
    fn new(cases: Vec<CaseInfo>, output_dir: PathBuf) -> Self {
        let annotations = load_annotations(&output_dir);
        let baselines = load_baselines(&output_dir);
        let tests_dir = output_dir.parent().unwrap_or(&output_dir);
        let visual_hashes = load_hashes(&tests_dir.join("visual_hashes.json"));
        let latest_hashes = load_hashes(&output_dir.join("latest_hashes.json"));
        let changed_cases = compute_changed_cases(&visual_hashes, &latest_hashes);
        Self {
            cases,
            output_dir,
            current_case: 0,
            current_page: 0,
            view_mode: ViewMode::SideBySide,
            texture_cache: HashMap::new(),
            overlay_cache: HashMap::new(),
            delta_cache: HashMap::new(),
            scroll_to_current: true,
            show_grid: false,
            grid_spacing: 18.0,
            refresh_flash: 0.0,
            annotations,
            pending_annotation: None,
            show_annotations: true,
            show_notes_panel: true,
            baselines,
            visual_hashes,
            latest_hashes,
            changed_cases,
            collapsed_groups: HashSet::new(),
            search_query: String::new(),
            search_active: false,
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
            self.delta_cache.clear();
            self.pending_annotation = None;
        }
    }

    fn save_annotations(&self) {
        save_annotations(&self.output_dir, &self.annotations);
    }

    fn refresh(&mut self) {
        let old_name = self.cases[self.current_case].name.clone();
        self.cases = discover_cases(&self.output_dir);
        self.baselines = load_baselines(&self.output_dir);
        self.annotations = load_annotations(&self.output_dir);
        let tests_dir = self.output_dir.parent().unwrap_or(&self.output_dir);
        self.visual_hashes = load_hashes(&tests_dir.join("visual_hashes.json"));
        self.latest_hashes = load_hashes(&self.output_dir.join("latest_hashes.json"));
        self.changed_cases = compute_changed_cases(&self.visual_hashes, &self.latest_hashes);
        // Try to stay on the same case after refresh
        self.current_case = self
            .cases
            .iter()
            .position(|c| c.name == old_name)
            .unwrap_or(0);
        // Clamp page in case page count changed
        if !self.cases.is_empty() {
            self.current_page = self
                .current_page
                .min(self.cases[self.current_case].page_count.saturating_sub(1));
        }
        self.texture_cache.clear();
        self.overlay_cache.clear();
        self.delta_cache.clear();
        self.refresh_flash = 1.0;
        self.scroll_to_current = true;
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
        let text_has_focus = (self.pending_annotation.is_some() || self.search_active)
            && ctx.memory(|m| m.focused().is_some());
        ctx.input(|i| {
            if i.key_pressed(egui::Key::Escape) {
                if self.search_active {
                    self.search_active = false;
                    self.search_query.clear();
                } else {
                    self.pending_annotation = None;
                }
            }
            if self.search_active && i.key_pressed(egui::Key::Enter) {
                let q = self.search_query.to_lowercase().replace(' ', "_");
                if !q.is_empty() {
                    if let Some(idx) = self.cases.iter().position(|c| {
                        c.name.to_lowercase().replace(' ', "_").contains(&q)
                    }) {
                        self.set_case(idx);
                    }
                }
                self.search_active = false;
                self.search_query.clear();
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
            if i.key_pressed(egui::Key::Num1) {
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
            if i.key_pressed(egui::Key::Num5) || i.key_pressed(egui::Key::D) {
                self.view_mode = ViewMode::Delta;
            }
            if i.key_pressed(egui::Key::S) {
                self.search_active = true;
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
        // Build per-case annotation status: (total, fixed)
        let mut annotation_counts: HashMap<&str, (usize, usize)> = HashMap::new();
        for a in &self.annotations.annotations {
            let entry = annotation_counts.entry(a.case.as_str()).or_insert((0, 0));
            entry.0 += 1;
            if a.fixed {
                entry.1 += 1;
            }
        }
        // Build case labels with group info
        struct CaseLabel {
            group: String,
            label: String,
            index: usize,
            color: Option<egui::Color32>,
            hash_changed: bool,
        }
        let case_labels: Vec<CaseLabel> = self
            .cases
            .iter()
            .enumerate()
            .map(|(i, c)| {
                let (group, display_name) = if let Some(pos) = c.name.find('/') {
                    (c.name[..pos].to_string(), c.name[pos + 1..].to_string())
                } else {
                    ("cases".to_string(), c.name.clone())
                };
                let key = baseline_key(&c.name);
                let color = match annotation_counts.get(c.name.as_str()) {
                    Some(&(total, fixed)) if fixed == total => Some(egui::Color32::from_rgb(0, 180, 0)),
                    Some(&(_, fixed)) if fixed > 0 => Some(egui::Color32::from_rgb(220, 180, 0)),
                    Some(_) => Some(egui::Color32::from_rgb(220, 40, 40)),
                    None => None,
                };
                let hash_changed = self.changed_cases.contains(&key);
                let pad = if color.is_some() { "     " } else { "" };
                let prefix = if hash_changed { "\u{26A0} " } else { "" };
                let label = if let Some(b) = self.baselines.get(&key) {
                    let j = b.jaccard.map(|v| format!("{:.0}", v * 100.0)).unwrap_or_else(|| "-".into());
                    let s = b.ssim.map(|v| format!("{:.0}", v * 100.0)).unwrap_or_else(|| "-".into());
                    format!("{}{} ({}p) {}/{}{}", prefix, display_name, c.page_count, j, s, pad)
                } else {
                    format!("{}{} ({}p){}", prefix, display_name, c.page_count, pad)
                };
                CaseLabel { group, label, index: i, color, hash_changed }
            })
            .collect();

        // Group labels by folder, preserving order
        let mut groups: Vec<(String, Vec<&CaseLabel>)> = Vec::new();
        for cl in &case_labels {
            if groups.last().is_some_and(|(g, _)| g == &cl.group) {
                groups.last_mut().unwrap().1.push(cl);
            } else {
                groups.push((cl.group.clone(), vec![cl]));
            }
        }

        let cur = self.current_case;
        let scroll = self.scroll_to_current;
        // Expand the group containing the current case when scrolling to it
        if scroll {
            let cur_group = case_labels.get(cur).map(|cl| cl.group.as_str());
            if let Some(g) = cur_group {
                self.collapsed_groups.remove(g);
            }
        }

        let mut clicked_case = None;
        let mut did_scroll = false;

        // Normalize query for substring matching (lowercase, spaces→underscores)
        let query_norm: String = self.search_query.to_lowercase().replace(' ', "_");
        let search_active = self.search_active;

        egui::SidePanel::right("case_list")
            .default_width(260.0)
            .show(ctx, |ui| {
                ui.heading("Cases");
                if search_active {
                    let resp = ui.text_edit_singleline(&mut self.search_query);
                    if self.search_active {
                        resp.request_focus();
                    }
                } else {
                    ui.separator();
                }
                egui::ScrollArea::vertical().show(ui, |ui| {
                    for (group_name, items) in &groups {
                        // Filter items by search query
                        let filtered: Vec<&&CaseLabel> = if query_norm.is_empty() {
                            items.iter().collect()
                        } else {
                            items
                                .iter()
                                .filter(|cl| {
                                    let name_norm = cl.label.to_lowercase().replace(' ', "_");
                                    name_norm.contains(&query_norm)
                                })
                                .collect()
                        };
                        if filtered.is_empty() {
                            continue;
                        }
                        let collapsed = self.collapsed_groups.contains(group_name) && query_norm.is_empty();
                        let header = format!(
                            "{} {} ({})",
                            if collapsed { "\u{25B6}" } else { "\u{25BC}" },
                            group_name,
                            filtered.len()
                        );
                        let header_resp = ui.selectable_label(false,
                            egui::RichText::new(&header).strong()
                        );
                        if header_resp.clicked() && query_norm.is_empty() {
                            if self.collapsed_groups.contains(group_name) {
                                self.collapsed_groups.remove(group_name);
                            } else {
                                self.collapsed_groups.insert(group_name.clone());
                            }
                        }
                        if !collapsed {
                            for cl in &filtered {
                                let selected = cl.index == cur;
                                let label_text = if cl.hash_changed {
                                    egui::RichText::new(&cl.label)
                                        .color(egui::Color32::from_rgb(220, 40, 40))
                                } else {
                                    egui::RichText::new(&cl.label)
                                };
                                let resp = ui.selectable_label(selected, label_text);
                                if let Some(c) = cl.color {
                                    let center = egui::pos2(resp.rect.right() - 6.0, resp.rect.center().y);
                                    ui.painter().circle_filled(center, 4.0, c);
                                }
                                if resp.clicked() {
                                    clicked_case = Some(cl.index);
                                }
                                if selected && scroll {
                                    resp.scroll_to_me(Some(egui::Align::Center));
                                    did_scroll = true;
                                }
                            }
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
        let mut acknowledge_current = false;
        egui::TopBottomPanel::top("top_bar").show(ctx, |ui| {
            ui.horizontal(|ui| {
                let case = &self.cases[self.current_case];
                ui.strong(&case.name);
                ui.separator();
                let mode_label = match self.view_mode {
                    ViewMode::SideBySide => "[1] Side-by-side",
                    ViewMode::Reference => "[R]eference",
                    ViewMode::Generated => "[G]enerated",
                    ViewMode::Overlay => "[O]verlay",
                    ViewMode::Delta => "[D]elta",
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
                if self.changed_cases.contains(&key) {
                    ui.separator();
                    ui.colored_label(
                        egui::Color32::from_rgb(220, 40, 40),
                        "Visual change",
                    );
                    if ui.button("Acknowledge").clicked() {
                        acknowledge_current = true;
                    }
                }
                ui.separator();
                if self.show_grid {
                    ui.label(format!("Grid: {:.0}px (+/-)", self.grid_spacing));
                }
                ui.separator();
                ui.label("[S]earch  [D]elta  [L]ines  [F]refresh  [A]nnot  [N]otes");
                if self.refresh_flash > 0.0 {
                    let alpha = (self.refresh_flash * 255.0) as u8;
                    ui.label(
                        egui::RichText::new("Refreshed")
                            .color(egui::Color32::from_rgba_unmultiplied(0, 180, 0, alpha)),
                    );
                }
            });
        });

        if acknowledge_current {
            let key = baseline_key(&self.cases[self.current_case].name);
            if let Some(hashes) = self.latest_hashes.get(&key).cloned() {
                self.visual_hashes.insert(key.clone(), hashes);
                self.changed_cases.remove(&key);
                // Write updated visual_hashes.json
                let tests_dir = self.output_dir.parent().unwrap_or(&self.output_dir);
                let path = tests_dir.join("visual_hashes.json");
                let sorted: std::collections::BTreeMap<_, _> =
                    self.visual_hashes.iter().collect();
                if let Ok(json) = serde_json::to_string_pretty(&sorted) {
                    let _ = std::fs::write(path, json + "\n");
                }
                // Snapshot current generated PNGs as the acknowledged baseline
                snapshot_acknowledged(&self.cases[self.current_case].dir);
                self.delta_cache.clear();
            }
        }

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
                ViewMode::Delta => {
                    show_delta(self, ctx, ui, page);
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

struct OverlayResult {
    texture: egui::TextureHandle,
    change_rects: Vec<[u32; 4]>,
}

fn build_overlay_texture(
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
    let mut grid = vec![false; (gw * gh) as usize];

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
                grid[(gy * gw + gx) as usize] = true;
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
fn find_change_rects(
    grid: &[bool],
    gw: u32,
    gh: u32,
    cell: u32,
    img_w: u32,
    img_h: u32,
) -> Vec<[u32; 4]> {
    let mut visited = vec![false; grid.len()];
    let mut rects = Vec::new();

    for gy in 0..gh {
        for gx in 0..gw {
            let idx = (gy * gw + gx) as usize;
            if !grid[idx] || visited[idx] {
                continue;
            }
            // Flood fill (8-connected) to find connected component
            let mut min_x = gx;
            let mut min_y = gy;
            let mut max_x = gx;
            let mut max_y = gy;
            let mut stack = vec![(gx, gy)];
            while let Some((cx, cy)) = stack.pop() {
                let ci = (cy * gw + cx) as usize;
                if visited[ci] || !grid[ci] {
                    continue;
                }
                visited[ci] = true;
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

fn show_delta(app: &mut App, ctx: &egui::Context, ui: &mut egui::Ui, page: usize) {
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
                    egui::Rect::from_min_max(egui::pos2(0.0, 0.0), egui::pos2(1.0, 1.0)),
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
        ui.label("Could not load acknowledged and/or generated images");
    }
}
