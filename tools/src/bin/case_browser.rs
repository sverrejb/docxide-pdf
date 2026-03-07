use eframe::egui;
use std::collections::HashMap;
use std::path::{Path, PathBuf};
use std::time::SystemTime;

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
    current_case: usize,
    current_page: usize,
    view_mode: ViewMode,
    texture_cache: HashMap<PathBuf, Option<egui::TextureHandle>>,
    overlay_cache: HashMap<(usize, usize), Option<egui::TextureHandle>>,
    scroll_to_current: bool,
    show_grid: bool,
    grid_spacing: f32,
    refresh_flash: f32,
}

impl App {
    fn new(cases: Vec<CaseInfo>, _output_dir: PathBuf) -> Self {
        Self {
            cases,
            current_case: 0,
            current_page: 0,
            view_mode: ViewMode::SideBySide,
            texture_cache: HashMap::new(),
            overlay_cache: HashMap::new(),
            scroll_to_current: true,
            show_grid: false,
            grid_spacing: 18.0,
            refresh_flash: 0.0,
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
        }
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
        // Handle keyboard input
        ctx.input(|i| {
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
        });

        // Right panel: case list
        let case_labels: Vec<(String, usize)> = self
            .cases
            .iter()
            .enumerate()
            .map(|(i, c)| (format!("{} ({}p)", c.name, c.page_count), i))
            .collect();
        let cur = self.current_case;
        let scroll = self.scroll_to_current;

        let mut clicked_case = None;
        let mut did_scroll = false;

        egui::SidePanel::right("case_list")
            .default_width(200.0)
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
                ui.separator();
                if self.show_grid {
                    ui.label(format!("Grid: {:.0}px (+/-)", self.grid_spacing));
                }
                ui.separator();
                ui.label("[L]ines overlay  [F]refresh");
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

        // Central panel: images
        egui::CentralPanel::default().show(ctx, |ui| {
            let page = self.current_page;
            match self.view_mode {
                ViewMode::SideBySide => {
                    let ref_path = self.page_path("reference", page);
                    let gen_path = self.page_path("generated", page);
                    show_side_by_side(self, ctx, ui, &ref_path, &gen_path);
                }
                ViewMode::Reference => {
                    let path = self.page_path("reference", page);
                    show_single(self, ctx, ui, &path, "Reference");
                }
                ViewMode::Generated => {
                    let path = self.page_path("generated", page);
                    show_single(self, ctx, ui, &path, "Generated");
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

fn show_image(
    app: &mut App,
    ctx: &egui::Context,
    ui: &mut egui::Ui,
    path: &PathBuf,
    max_w: f32,
    max_h: f32,
) {
    if let Some(tex) = app.load_texture(ctx, path) {
        let size = fit_size(&tex, max_w, max_h);
        let resp = ui.image(egui::load::SizedTexture::new(tex.id(), size));
        if app.show_grid {
            draw_grid_overlay(ctx, resp.rect, app.grid_spacing);
        }
    } else {
        ui.label("(not found)");
    }
}

fn show_single(app: &mut App, ctx: &egui::Context, ui: &mut egui::Ui, path: &PathBuf, label: &str) {
    let age = file_age(path).unwrap_or_default();
    ui.horizontal(|ui| {
        ui.label(label);
        ui.label(egui::RichText::new(age).weak());
    });
    let available = ui.available_size();
    egui::ScrollArea::both().show(ui, |ui| {
        show_image(app, ctx, ui, path, available.x, available.y - 20.0);
    });
}

fn show_side_by_side(
    app: &mut App,
    ctx: &egui::Context,
    ui: &mut egui::Ui,
    ref_path: &PathBuf,
    gen_path: &PathBuf,
) {
    let ref_tex = app.load_texture(ctx, ref_path);
    let gen_tex = app.load_texture(ctx, gen_path);
    let show_grid = app.show_grid;
    let grid_spacing = app.grid_spacing;
    let ref_age = file_age(ref_path).unwrap_or_default();
    let gen_age = file_age(gen_path).unwrap_or_default();

    let available = ui.available_size();
    let half_w = (available.x - 10.0) / 2.0;
    let max_h = available.y - 30.0;

    ui.horizontal_top(|ui| {
        ui.vertical(|ui| {
            ui.horizontal(|ui| {
                ui.label("Reference");
                ui.label(egui::RichText::new(&ref_age).weak());
            });
            if let Some(tex) = &ref_tex {
                let size = fit_size(tex, half_w, max_h);
                let resp = ui.image(egui::load::SizedTexture::new(tex.id(), size));
                if show_grid {
                    draw_grid_overlay(ui.ctx(), resp.rect, grid_spacing);
                }
            } else {
                ui.label("(not found)");
            }
        });

        ui.separator();

        ui.vertical(|ui| {
            ui.horizontal(|ui| {
                ui.label("Generated");
                ui.label(egui::RichText::new(&gen_age).weak());
            });
            if let Some(tex) = &gen_tex {
                let size = fit_size(tex, half_w, max_h);
                let resp = ui.image(egui::load::SizedTexture::new(tex.id(), size));
                if show_grid {
                    draw_grid_overlay(ui.ctx(), resp.rect, grid_spacing);
                }
            } else {
                ui.label("(not found)");
            }
        });
    });
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
        let show_grid = app.show_grid;
        let grid_spacing = app.grid_spacing;
        egui::ScrollArea::both().show(ui, |ui| {
            let size = fit_size(&tex, available.x, available.y - 20.0);
            let resp = ui.image(egui::load::SizedTexture::new(tex.id(), size));
            if show_grid {
                draw_grid_overlay(ui.ctx(), resp.rect, grid_spacing);
            }
        });
    } else {
        ui.label("Could not load reference and/or generated images");
    }
}
