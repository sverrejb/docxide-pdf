//! Chrome around the image views: keyboard shortcuts, case list, top/bottom bars, notes.

use crate::App;
use crate::annotations::{Annotation, ImageSource, source_color};
use crate::cases::{baseline_key, normalize_name, snapshot_acknowledged};
use crate::jobs::{self, FixOutcome};
use crate::views::{VIEW_MODE_KEYS, ViewMode};
use eframe::egui;
use std::collections::HashMap;

pub fn score_color(value: f64, threshold: f64) -> egui::Color32 {
    if value >= threshold {
        egui::Color32::from_rgb(0, 180, 0)
    } else {
        egui::Color32::from_rgb(220, 40, 40)
    }
}

/// Paint a group disclosure triangle as a filled shape.
///
/// Drawn rather than rendered from a font glyph: the right-pointing triangle
/// `\u{25B6}` has an emoji presentation and is absent from egui's monochrome
/// text font, so it showed as a tofu box while the down triangle `\u{25BC}`
/// rendered fine. A painted polygon sidesteps font glyph coverage entirely.
pub fn paint_disclosure(
    painter: &egui::Painter,
    center: egui::Pos2,
    collapsed: bool,
    color: egui::Color32,
) {
    let r = 4.0;
    let points = if collapsed {
        vec![
            egui::pos2(center.x - r * 0.6, center.y - r),
            egui::pos2(center.x - r * 0.6, center.y + r),
            egui::pos2(center.x + r * 0.8, center.y),
        ]
    } else {
        vec![
            egui::pos2(center.x - r, center.y - r * 0.6),
            egui::pos2(center.x + r, center.y - r * 0.6),
            egui::pos2(center.x, center.y + r * 0.8),
        ]
    };
    painter.add(egui::Shape::convex_polygon(
        points,
        color,
        egui::Stroke::NONE,
    ));
}

pub fn handle_keys(app: &mut App, ctx: &egui::Context) {
    // Both text fields request focus every frame; with both open they steal it from
    // each other and neither ever sees a keystroke, so a new note closes the search.
    if app.pending_annotation.is_some() && app.search_active {
        app.search_active = false;
        app.search_query.clear();
    }
    // Handle keyboard input (skip when text input has focus)
    let text_has_focus = (app.pending_annotation.is_some() || app.search_active)
        && ctx.memory(|m| m.focused().is_some());
    ctx.input(|i| {
        if i.key_pressed(egui::Key::Escape) {
            if app.search_active {
                app.search_active = false;
                app.search_query.clear();
            } else {
                app.pending_annotation = None;
            }
        }
        if app.search_active && i.key_pressed(egui::Key::Enter) {
            let q = normalize_name(&app.search_query);
            if !q.is_empty()
                && let Some(idx) = app
                    .cases
                    .iter()
                    .position(|c| normalize_name(&c.name).contains(&q))
            {
                app.set_case(idx);
            }
            app.search_active = false;
            app.search_query.clear();
        }
        if text_has_focus {
            return;
        }
        if i.key_pressed(egui::Key::ArrowRight) {
            let next = (app.current_case + 1).min(app.cases.len() - 1);
            app.set_case(next);
        }
        if i.key_pressed(egui::Key::ArrowLeft) {
            let next = app.current_case.saturating_sub(1);
            app.set_case(next);
        }
        if i.key_pressed(egui::Key::ArrowDown) {
            app.current_page = (app.current_page + 1).min(app.current().page_count - 1);
        }
        if i.key_pressed(egui::Key::ArrowUp) {
            app.current_page = app.current_page.saturating_sub(1);
        }
        for &(key, mode) in VIEW_MODE_KEYS {
            if i.key_pressed(key) {
                app.view_mode = mode;
            }
        }
        if i.key_pressed(egui::Key::S) {
            app.search_active = true;
        }
        if i.key_pressed(egui::Key::F5) || i.key_pressed(egui::Key::F) {
            app.refresh();
        }
        if i.key_pressed(egui::Key::L) {
            app.show_grid = !app.show_grid;
        }
        if i.key_pressed(egui::Key::Equals) || i.key_pressed(egui::Key::Plus) {
            app.grid_spacing = (app.grid_spacing + 2.0).min(100.0);
        }
        if i.key_pressed(egui::Key::Minus) {
            app.grid_spacing = (app.grid_spacing - 2.0).max(4.0);
        }
        if i.key_pressed(egui::Key::A) {
            app.show_annotations = !app.show_annotations;
        }
        if i.key_pressed(egui::Key::N) {
            app.show_notes_panel = !app.show_notes_panel;
        }
    });
}

pub fn case_list(app: &mut App, ctx: &egui::Context) {
    // Right panel: case list
    // Build per-case annotation status: (total, fixed)
    let mut annotation_counts: HashMap<&str, (usize, usize)> = HashMap::new();
    for a in &app.annotations.annotations {
        if a.cleared {
            continue;
        }
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
        has_annotations: bool,
    }
    let case_labels: Vec<CaseLabel> = app
        .cases
        .iter()
        .enumerate()
        .map(|(i, c)| {
            let (group, display_name) = c.group_and_name();
            let group = group.to_string();
            let key = baseline_key(&c.name);
            let color = match annotation_counts.get(c.name.as_str()) {
                Some(&(total, fixed)) if fixed == total => Some(egui::Color32::from_rgb(0, 180, 0)),
                Some(&(_, fixed)) if fixed > 0 => Some(egui::Color32::from_rgb(220, 180, 0)),
                Some(_) => Some(egui::Color32::from_rgb(220, 40, 40)),
                None => None,
            };
            let hash_changed = app.changed_cases.contains(&key);
            let has_annotations = annotation_counts.contains_key(c.name.as_str());
            let pad = if color.is_some() { "     " } else { "" };
            let prefix = if hash_changed { "\u{26A0} " } else { "" };
            let label = if let Some(b) = app.baselines.get(&key) {
                let j = b
                    .jaccard
                    .map(|v| format!("{:.0}", v * 100.0))
                    .unwrap_or_else(|| "-".into());
                let s = b
                    .ssim
                    .map(|v| format!("{:.0}", v * 100.0))
                    .unwrap_or_else(|| "-".into());
                format!(
                    "{}{} ({}p) {}/{}{}",
                    prefix, display_name, c.page_count, j, s, pad
                )
            } else {
                format!("{}{} ({}p){}", prefix, display_name, c.page_count, pad)
            };
            CaseLabel {
                group,
                label,
                index: i,
                color,
                hash_changed,
                has_annotations,
            }
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

    let cur = app.current_case;
    let scroll = app.scroll_to_current;
    // Expand the group containing the current case when scrolling to it
    if scroll {
        let cur_group = case_labels.get(cur).map(|cl| cl.group.as_str());
        if let Some(g) = cur_group {
            app.collapsed_groups.remove(g);
        }
    }

    let mut clicked_case = None;
    let mut did_scroll = false;

    // Normalize query for substring matching (lowercase, spaces→underscores)
    let query_norm: String = normalize_name(&app.search_query);
    let search_active = app.search_active;

    egui::SidePanel::right("case_list")
        .default_width(260.0)
        .show(ctx, |ui| {
            ui.heading("Cases");
            if search_active {
                let resp = ui.text_edit_singleline(&mut app.search_query);
                if app.search_active {
                    resp.request_focus();
                }
            }
            ui.horizontal(|ui| {
                ui.checkbox(&mut app.filter_annotated, "Annotated")
                    .on_hover_text("Only show cases that have annotations");
                ui.checkbox(&mut app.filter_changed, "Changed")
                    .on_hover_text("Only show cases with unacknowledged visual changes");
            });
            ui.separator();

            // Read filter state after the checkboxes so toggles apply this frame.
            let filter_annotated = app.filter_annotated;
            let filter_changed = app.filter_changed;
            // Any active filter force-expands groups (mirrors search behaviour).
            let any_filter = !query_norm.is_empty() || filter_annotated || filter_changed;

            egui::ScrollArea::vertical().show(ui, |ui| {
                for (group_name, items) in &groups {
                    let filtered: Vec<&&CaseLabel> = items
                        .iter()
                        .filter(|cl| {
                            if filter_annotated && !cl.has_annotations {
                                return false;
                            }
                            if filter_changed && !cl.hash_changed {
                                return false;
                            }
                            if !query_norm.is_empty()
                                && !normalize_name(&cl.label).contains(&query_norm)
                            {
                                return false;
                            }
                            true
                        })
                        .collect();
                    if filtered.is_empty() {
                        continue;
                    }
                    let collapsed = app.collapsed_groups.contains(group_name) && !any_filter;
                    // Leading spaces reserve room for the painted disclosure triangle.
                    let header = format!("   {} ({})", group_name, filtered.len());
                    let header_resp =
                        ui.selectable_label(false, egui::RichText::new(&header).strong());
                    let tri_center =
                        egui::pos2(header_resp.rect.left() + 7.0, header_resp.rect.center().y);
                    paint_disclosure(
                        ui.painter(),
                        tri_center,
                        collapsed,
                        ui.visuals().strong_text_color(),
                    );
                    if header_resp.clicked() && !any_filter {
                        if app.collapsed_groups.contains(group_name) {
                            app.collapsed_groups.remove(group_name);
                        } else {
                            app.collapsed_groups.insert(group_name.clone());
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
                                let center =
                                    egui::pos2(resp.rect.right() - 6.0, resp.rect.center().y);
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
        app.set_case(idx);
    }
    if did_scroll {
        app.scroll_to_current = false;
    }
}

pub fn top_bar(app: &mut App, ctx: &egui::Context) {
    // Top bar: case name and view mode
    let mut acknowledge_current = false;
    let mut start_regen = false;
    let mut stop_fix = None;
    egui::TopBottomPanel::top("top_bar").show(ctx, |ui| {
        ui.horizontal(|ui| {
            let case = &app.cases[app.current_case];
            ui.strong(&case.name);
            ui.separator();
            let regen_running = app.jobs.iter().any(|j| j.fixing().is_none());
            if !regen_running && ui.button("Re-generate").clicked() {
                start_regen = true;
            }
            if !app.jobs.is_empty() {
                ui.spinner();
                for job in &app.jobs {
                    let what = match job.fixing() {
                        Some(id) => format!("fixing note {id}"),
                        None => "regenerating".to_string(),
                    };
                    ui.colored_label(
                        egui::Color32::from_rgb(220, 180, 0),
                        format!("{what} {}", job.elapsed()),
                    );
                    if let Some(id) = job.fixing()
                        && ui.small_button("stop").clicked()
                    {
                        stop_fix = Some(id);
                    }
                }
            }
            ui.separator();
            let mode_label = match app.view_mode {
                ViewMode::SideBySide => "[1] Side-by-side",
                ViewMode::Reference => "[R]eference",
                ViewMode::Generated => "[G]enerated",
                ViewMode::Overlay => "[O]verlay",
                ViewMode::Delta => "[D]elta",
                ViewMode::Compare => "[C]ompare (Ref | Ours | LO)",
            };
            ui.label(format!("View: {}", mode_label));
            let key = baseline_key(&case.name);
            if let Some(b) = app.baselines.get(&key) {
                ui.separator();
                if let Some(j) = b.jaccard {
                    ui.colored_label(score_color(j, 0.20), format!("Jaccard: {:.1}%", j * 100.0));
                }
                if let Some(s) = b.ssim {
                    ui.colored_label(score_color(s, 0.75), format!("SSIM: {:.1}%", s * 100.0));
                }
            }
            if app.changed_cases.contains(&key) {
                ui.separator();
                ui.colored_label(egui::Color32::from_rgb(220, 40, 40), "Visual change");
                if ui.button("Acknowledge").clicked() {
                    acknowledge_current = true;
                }
            }
            ui.separator();
            if app.show_grid {
                ui.label(format!("Grid: {:.0}px (+/-)", app.grid_spacing));
            }
            ui.separator();
            ui.label("[S]earch  [D]elta  [C]ompare  [L]ines  [F]refresh  [A]nnot  [N]otes");
            if app.refresh_flash > 0.0 {
                let alpha = (app.refresh_flash * 255.0) as u8;
                ui.label(
                    egui::RichText::new("Refreshed")
                        .color(egui::Color32::from_rgba_unmultiplied(0, 180, 0, alpha)),
                );
            }
        });
    });

    if start_regen {
        app.start_regen();
    }
    if let Some(id) = stop_fix {
        app.stop_fix(id);
    }

    if acknowledge_current {
        let key = baseline_key(&app.cases[app.current_case].name);
        if let Some(hashes) = app.latest_hashes.get(&key).cloned() {
            app.visual_hashes.insert(key.clone(), hashes);
            app.changed_cases.remove(&key);
            // Write updated visual_hashes.json
            let tests_dir = app.output_dir.parent().unwrap_or(&app.output_dir);
            let path = tests_dir.join("visual_hashes.json");
            let sorted: std::collections::BTreeMap<_, _> = app.visual_hashes.iter().collect();
            if let Ok(json) = serde_json::to_string_pretty(&sorted) {
                let _ = std::fs::write(path, json + "\n");
            }
            // Snapshot current generated PNGs as the acknowledged baseline
            snapshot_acknowledged(&app.cases[app.current_case].dir);
            app.delta_cache.clear();
        }
    }
}

pub fn bottom_bar(app: &mut App, ctx: &egui::Context) {
    // Bottom bar: page number
    egui::TopBottomPanel::bottom("bottom_bar").show(ctx, |ui| {
        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
            let case = &app.cases[app.current_case];
            ui.label(
                egui::RichText::new(format!("Page {}/{}", app.current_page + 1, case.page_count))
                    .size(18.0),
            );
        });
    });
}

pub fn notes_panel(app: &mut App, ctx: &egui::Context) {
    // Notes panel (left side)
    if app.show_notes_panel {
        let page = app.current_page;
        let case_name = app.cases[app.current_case].name.clone();
        let page_ann_ids: Vec<u64> = app
            .annotations
            .annotations
            .iter()
            .filter(|a| a.case == case_name && a.page == page && !a.cleared)
            .map(|a| a.id)
            .collect();
        let mut clear_id = None;
        let mut toggle_fixed_id = None;
        let mut start_fix_id = None;
        let mut stop_fix_id = None;
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
                        let ann = app.annotations.annotations.iter().find(|a| a.id == *ann_id).unwrap();
                        let color = if ann.fixed {
                            egui::Color32::from_gray(140)
                        } else {
                            source_color(ann.source)
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
                            ui.weak(format!("id {}", ann.id));
                            if ui
                                .small_button("X")
                                .on_hover_text("Clear this note (hidden, kept in file)")
                                .clicked()
                            {
                                clear_id = Some(*ann_id);
                            }
                            let fixing = app.jobs.iter().find(|j| j.fixing() == Some(*ann_id));
                            if let Some(job) = fixing {
                                ui.spinner();
                                ui.weak(job.elapsed());
                                if ui.small_button("Stop").clicked() {
                                    stop_fix_id = Some(*ann_id);
                                }
                            } else if !ann.fixed
                                && ui
                                    .small_button("Fix")
                                    .on_hover_text(
                                        "Hand this note to a headless `claude -p` run in its own worktree",
                                    )
                                    .clicked()
                            {
                                start_fix_id = Some(*ann_id);
                            }
                        });
                        if let Some(outcome) = app.fix_results.get(ann_id) {
                            let (color, text) = match outcome {
                                FixOutcome::Done => (
                                    egui::Color32::from_rgb(0, 180, 0),
                                    format!("Claude finished on {}", jobs::fix_branch(*ann_id)),
                                ),
                                FixOutcome::Failed => {
                                    (egui::Color32::from_rgb(220, 40, 40), "Claude failed".to_string())
                                }
                                FixOutcome::Stopped => {
                                    (egui::Color32::from_gray(140), "Stopped".to_string())
                                }
                            };
                            ui.colored_label(color, text).on_hover_text(format!(
                                "log: {}",
                                jobs::fix_log_path(&app.output_dir, *ann_id).display()
                            ));
                        }
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
        if let Some(id) = clear_id {
            if let Some(ann) = app.annotations.annotations.iter_mut().find(|a| a.id == id) {
                ann.cleared = true;
            }
            app.save_annotations();
        }
        if let Some(id) = toggle_fixed_id {
            if let Some(ann) = app.annotations.annotations.iter_mut().find(|a| a.id == id) {
                ann.fixed = !ann.fixed;
            }
            app.save_annotations();
        }
        if let Some(id) = start_fix_id {
            app.start_fix(id);
        }
        if let Some(id) = stop_fix_id {
            app.stop_fix(id);
        }
    }
}

pub fn add_note_window(app: &mut App, ctx: &egui::Context) {
    // Annotation input popup
    let mut commit_annotation = false;
    let mut cancel_annotation = false;
    if app.pending_annotation.is_some() {
        egui::Window::new("Add Note")
            .collapsible(false)
            .resizable(false)
            .show(ctx, |ui| {
                let pending = app.pending_annotation.as_mut().unwrap();
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
                let cmd_enter =
                    ui.input(|i| i.key_pressed(egui::Key::Enter) && i.modifiers.command);
                if cmd_enter || (resp.lost_focus() && ui.input(|i| i.key_pressed(egui::Key::Enter)))
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
        if let Some(pending) = app.pending_annotation.take() {
            if !pending.text_buf.trim().is_empty() {
                let id = app.annotations.next_id;
                app.annotations.next_id += 1;
                app.annotations.annotations.push(Annotation {
                    id,
                    case: app.cases[app.current_case].name.clone(),
                    page: pending.page,
                    source: pending.source,
                    x_pt: pending.x_pt,
                    y_pt: pending.y_pt,
                    note: pending.text_buf.trim().to_string(),
                    fixed: false,
                    cleared: false,
                });
                app.save_annotations();
            }
        }
    }
    if cancel_annotation {
        app.pending_annotation = None;
    }
}
