//! Case browser: Word reference vs our output for every fixture in tests/output,
//! with click-to-annotate notes, one-case re-generation and headless Claude fixes.

mod annotations;
mod cases;
mod jobs;
mod panels;
mod views;

use annotations::{AnnotationStore, PendingAnnotation, load_annotations, save_annotations};
use cases::{
    BaselineScores, CaseInfo, compute_changed_cases, discover_cases, find_output_dir,
    load_baselines, load_hashes,
};
use eframe::egui;
use jobs::{FixOutcome, Job, JobKind};
use std::collections::{HashMap, HashSet};
use std::path::{Path, PathBuf};
use std::time::Duration;
use views::{OverlayResult, ViewMode};

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

pub struct App {
    pub cases: Vec<CaseInfo>,
    pub output_dir: PathBuf,
    pub current_case: usize,
    pub current_page: usize,
    pub view_mode: ViewMode,
    pub texture_cache: HashMap<PathBuf, Option<egui::TextureHandle>>,
    pub overlay_cache: HashMap<(usize, usize), Option<OverlayResult>>,
    pub delta_cache: HashMap<(usize, usize), Option<OverlayResult>>,
    pub scroll_to_current: bool,
    pub show_grid: bool,
    pub grid_spacing: f32,
    pub refresh_flash: f32,
    pub annotations: AnnotationStore,
    pub pending_annotation: Option<PendingAnnotation>,
    pub show_annotations: bool,
    pub show_notes_panel: bool,
    pub baselines: HashMap<String, BaselineScores>,
    pub visual_hashes: HashMap<String, Vec<String>>,
    pub latest_hashes: HashMap<String, Vec<String>>,
    pub changed_cases: HashSet<String>,
    pub collapsed_groups: HashSet<String>,
    pub search_query: String,
    pub search_active: bool,
    pub filter_annotated: bool,
    pub filter_changed: bool,
    /// Running background processes: at most one re-generate, any number of Claude fixes.
    pub jobs: Vec<Job>,
    /// Outcome of fixes finished this session, by annotation id.
    pub fix_results: HashMap<u64, FixOutcome>,
}

impl App {
    pub fn new(cases: Vec<CaseInfo>, output_dir: PathBuf) -> Self {
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
            filter_annotated: false,
            filter_changed: false,
            jobs: Vec::new(),
            fix_results: HashMap::new(),
        }
    }

    /// cargo and claude must run from the project root: output_dir is <root>/tests/output.
    fn project_root(&self) -> PathBuf {
        self.output_dir
            .canonicalize()
            .ok()
            .and_then(|p| p.parent().and_then(Path::parent).map(Path::to_path_buf))
            .unwrap_or_else(|| PathBuf::from("."))
    }

    /// Re-run the visual_comparison test rig for the current case only, in the
    /// background. Refreshes scores, hashes and screenshots when it finishes.
    pub fn start_regen(&mut self) {
        if self.jobs.iter().any(|j| j.fixing().is_none()) {
            return;
        }
        match jobs::regen(&self.project_root(), self.current()) {
            Ok(job) => self.jobs.push(job),
            Err(e) => eprintln!("re-generate failed to start: {e}"),
        }
    }

    pub fn start_fix(&mut self, ann_id: u64) {
        if self.jobs.iter().any(|j| j.fixing() == Some(ann_id)) {
            return;
        }
        let Some(ann) = self.annotations.annotations.iter().find(|a| a.id == ann_id) else {
            return;
        };
        let Some(case) = self.cases.iter().find(|c| c.name == ann.case) else {
            return;
        };
        let log = jobs::fix_log_path(&self.output_dir, ann_id);
        match jobs::fix(&self.project_root(), &log, case, ann) {
            Ok(job) => {
                self.fix_results.remove(&ann_id);
                self.jobs.push(job);
            }
            Err(e) => eprintln!("fix failed to start: {e}"),
        }
    }

    pub fn stop_fix(&mut self, ann_id: u64) {
        if let Some(i) = self.jobs.iter().position(|j| j.fixing() == Some(ann_id)) {
            self.jobs.remove(i).stop();
            self.fix_results.insert(ann_id, FixOutcome::Stopped);
        }
    }

    /// Reap finished jobs. A re-generate rewrote tests/output, so reload it. A fix
    /// ran in its own worktree and reaches the main tree only through the `fixed`
    /// flag in that worktree's copy of annotations.json, mirrored here.
    fn poll_jobs(&mut self, ctx: &egui::Context) {
        let mut finished = Vec::new();
        let mut i = 0;
        while i < self.jobs.len() {
            match self.jobs[i].child.try_wait() {
                Ok(None) => i += 1,
                Ok(Some(status)) => finished.push((self.jobs.remove(i).kind, status.success())),
                Err(_) => finished.push((self.jobs.remove(i).kind, false)),
            }
        }
        if !self.jobs.is_empty() {
            ctx.request_repaint_after(Duration::from_millis(500));
        }
        for (kind, succeeded) in finished {
            match kind {
                JobKind::Regen => self.refresh(),
                JobKind::Fix(id) => {
                    let wt_output =
                        jobs::worktree_path(&self.project_root(), id).join("tests/output");
                    let marked = load_annotations(&wt_output)
                        .annotations
                        .iter()
                        .any(|a| a.id == id && a.fixed);
                    if marked
                        && let Some(a) =
                            self.annotations.annotations.iter_mut().find(|a| a.id == id)
                    {
                        a.fixed = true;
                        self.save_annotations();
                    }
                    let outcome = if succeeded {
                        FixOutcome::Done
                    } else {
                        FixOutcome::Failed
                    };
                    self.fix_results.insert(id, outcome);
                }
            }
        }
    }

    pub fn current(&self) -> &CaseInfo {
        &self.cases[self.current_case]
    }

    pub fn set_case(&mut self, idx: usize) {
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

    pub fn save_annotations(&self) {
        save_annotations(&self.output_dir, &self.annotations);
    }

    pub fn refresh(&mut self) {
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

    pub fn page_path(&self, subdir: &str, page: usize) -> PathBuf {
        self.current()
            .dir
            .join(subdir)
            .join(format!("page_{:03}.png", page + 1))
    }

    pub fn load_texture(
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
        panels::handle_keys(self, ctx);
        self.poll_jobs(ctx);
        panels::case_list(self, ctx);
        panels::top_bar(self, ctx);
        if self.refresh_flash > 0.0 {
            self.refresh_flash = (self.refresh_flash - 0.05).max(0.0);
            ctx.request_repaint();
        }
        panels::bottom_bar(self, ctx);
        panels::notes_panel(self, ctx);
        panels::add_note_window(self, ctx);
        egui::CentralPanel::default().show(ctx, |ui| views::central(self, ctx, ui));
    }
}
