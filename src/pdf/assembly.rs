use std::collections::{HashMap, HashSet};

use pdf_writer::{Content, Filter, Name, Pdf, Rect, Ref, Str, TextStr};

use crate::fonts::FontEntry;
use crate::model::Document;

use super::comments::render_comment_pane;
use super::layout::LinkAnnotation;
use super::GradientSpec;

pub(crate) struct HeadingEntry {
    pub(super) title: String,
    pub(super) level: u8,
    pub(super) page_idx: usize,
    pub(super) y_position: f32,
}

fn srgb_to_linear(s: f32) -> f32 {
    if s <= 0.04045 {
        s / 12.92
    } else {
        ((s + 0.055) / 1.055).powf(2.4)
    }
}

fn srgb_to_linear_rgb(c: [u8; 3]) -> [f32; 3] {
    [
        srgb_to_linear(c[0] as f32 / 255.0),
        srgb_to_linear(c[1] as f32 / 255.0),
        srgb_to_linear(c[2] as f32 / 255.0),
    ]
}

#[allow(clippy::too_many_arguments)]
pub(super) fn assemble_pdf_pages(
    pdf: &mut Pdf,
    alloc: &mut impl FnMut() -> Ref,
    catalog_id: Ref,
    pages_id: Ref,
    all_contents: Vec<Content>,
    all_hf_contents: &mut Vec<Option<Content>>,
    all_page_links: &[Vec<LinkAnnotation>],
    all_page_comment_anchors: &[Vec<(u32, f32, f32)>],
    all_page_alpha_states: &[HashSet<u8>],
    all_page_gradient_specs: &[Vec<GradientSpec>],
    page_section_indices: &[(usize, bool, usize)],
    seen_fonts: &HashMap<String, FontEntry>,
    font_order: &[String],
    image_xobjects: &[(String, Ref)],
    doc: &Document,
    bookmark_positions: &HashMap<String, (usize, f32)>,
    heading_entries: &[HeadingEntry],
) {
    let n = all_contents.len();
    let page_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();
    let content_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();

    let has_any_comments = !doc.comments.is_empty()
        && all_page_comment_anchors.iter().any(|p| !p.is_empty());

    let page_annot_refs: Vec<Vec<Ref>> = all_page_links
        .iter()
        .map(|links| {
            links
                .iter()
                .filter_map(|link| {
                    let annot_ref = alloc();
                    let mut annot = pdf.annotation(annot_ref);
                    annot
                        .subtype(pdf_writer::types::AnnotationType::Link)
                        .rect(link.rect)
                        .border(0.0, 0.0, 0.0, None);
                    if let Some(anchor) = link.url.strip_prefix('#') {
                        if let Some(&(dest_page_idx, dest_y)) = bookmark_positions.get(anchor) {
                            debug_assert!(
                                dest_page_idx < page_ids.len(),
                                "bookmark '{anchor}' points to page {dest_page_idx} but only {} pages exist",
                                page_ids.len(),
                            );
                            let safe_idx = dest_page_idx.min(page_ids.len().saturating_sub(1));
                            annot
                                .action()
                                .action_type(pdf_writer::types::ActionType::GoTo)
                                .destination()
                                .page(page_ids[safe_idx])
                                .xyz(0.0, dest_y, None);
                            Some(annot_ref)
                        } else {
                            None
                        }
                    } else {
                        annot
                            .action()
                            .action_type(pdf_writer::types::ActionType::Uri)
                            .uri(Str(link.url.as_bytes()));
                        Some(annot_ref)
                    }
                })
                .collect()
        })
        .collect();

    let all_alpha_values: HashSet<u8> = all_page_alpha_states
        .iter()
        .flat_map(|s| s.iter().copied())
        .collect();
    let alpha_gs_refs: HashMap<u8, Ref> = all_alpha_values
        .iter()
        .map(|&pct| {
            let gs_ref = alloc();
            pdf.ext_graphics(gs_ref)
                .non_stroking_alpha(pct as f32 / 100.0);
            (pct, gs_ref)
        })
        .collect();

    let all_page_pattern_refs: Vec<Vec<(String, Ref)>> = all_page_gradient_specs
        .iter()
        .map(|specs| {
            specs
                .iter()
                .map(|spec| {
                    let func_ref = if spec.stops.len() <= 2 {
                        let (c0, c1) = if spec.stops.len() >= 2 {
                            (spec.stops[0].0, spec.stops[spec.stops.len() - 1].0)
                        } else {
                            (spec.stops[0].0, spec.stops[0].0)
                        };
                        let fref = alloc();
                        pdf.exponential_function(fref)
                            .domain([0.0, 1.0])
                            .c0(srgb_to_linear_rgb(c0))
                            .c1(srgb_to_linear_rgb(c1))
                            .n(1.0);
                        fref
                    } else {
                        let sub_refs: Vec<Ref> = spec
                            .stops
                            .windows(2)
                            .map(|pair| {
                                let fref = alloc();
                                pdf.exponential_function(fref)
                                    .domain([0.0, 1.0])
                                    .c0(srgb_to_linear_rgb(pair[0].0))
                                    .c1(srgb_to_linear_rgb(pair[1].0))
                                    .n(1.0);
                                fref
                            })
                            .collect();

                        let bounds: Vec<f32> = spec.stops[1..spec.stops.len() - 1]
                            .iter()
                            .map(|s| s.1)
                            .collect();
                        let encode: Vec<f32> =
                            sub_refs.iter().flat_map(|_| [0.0, 1.0]).collect();

                        let stitch_ref = alloc();
                        pdf.stitching_function(stitch_ref)
                            .domain([0.0, 1.0])
                            .functions(sub_refs)
                            .bounds(bounds)
                            .encode(encode);
                        stitch_ref
                    };

                    let ang_rad = spec.angle_deg.to_radians();
                    let (sin_a, cos_a) = ang_rad.sin_cos();
                    let cx = spec.x + spec.w / 2.0;
                    let cy = spec.y + spec.h / 2.0;
                    let half_len = ((spec.w / 2.0 * cos_a).powi(2)
                        + (spec.h / 2.0 * sin_a).powi(2))
                    .sqrt();
                    let x0 = cx - half_len * cos_a;
                    let y0 = cy + half_len * sin_a;
                    let x1 = cx + half_len * cos_a;
                    let y1 = cy - half_len * sin_a;

                    let pat_ref = alloc();
                    let mut pattern = pdf.shading_pattern(pat_ref);
                    let mut shading = pattern.function_shading();
                    shading
                        .shading_type(pdf_writer::types::FunctionShadingType::Axial)
                        .color_space()
                        .cal_rgb(
                            [0.9505, 1.0, 1.0890],
                            None,
                            None,
                            Some([
                                0.4124, 0.2126, 0.0193, 0.3576, 0.7152, 0.1192, 0.1805, 0.0722,
                                0.9505,
                            ]),
                        );
                    shading
                        .function(func_ref)
                        .coords([x0, y0, x1, y1])
                        .extend([true, true]);

                    (spec.pattern_name.clone(), pat_ref)
                })
                .collect()
        })
        .collect();

    for (i, mut c) in all_contents.into_iter().enumerate() {
        if has_any_comments {
            let (.., si) = page_section_indices[i];
            let sp = &doc.sections[si].properties;
            let anchors = &all_page_comment_anchors[i];
            render_comment_pane(
                &mut c,
                &doc.comments,
                anchors,
                sp.page_width,
                sp.page_height,
                seen_fonts,
            );
        }
        let body_raw = c.finish();
        if let Some(hf) = all_hf_contents[i].take() {
            let hf_raw = hf.finish();
            let mut combined = Vec::with_capacity(hf_raw.len() + 1 + body_raw.len());
            combined.extend_from_slice(hf_raw.as_slice());
            combined.push(b'\n');
            combined.extend_from_slice(body_raw.as_slice());
            let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&combined, 6);
            pdf.stream(content_ids[i], &compressed)
                .filter(Filter::FlateDecode);
        } else {
            let compressed = miniz_oxide::deflate::compress_to_vec_zlib(body_raw.as_slice(), 6);
            pdf.stream(content_ids[i], &compressed)
                .filter(Filter::FlateDecode);
        }
    }

    // Build PDF outline (bookmarks panel) from heading entries
    let outline_id = if !heading_entries.is_empty() {
        let oid = alloc();
        let item_refs: Vec<Ref> = heading_entries.iter().map(|_| alloc()).collect();

        // Build parent/first-child/last-child/prev/next relationships using a stack.
        // Each item's parent is the nearest preceding item with a smaller level, or the root.
        // children_of[i] collects indices of direct children of item i.
        // children_of_root collects top-level items.
        let mut parent_idx: Vec<Option<usize>> = vec![None; heading_entries.len()];
        let mut stack: Vec<usize> = Vec::new(); // indices of ancestors
        for (i, entry) in heading_entries.iter().enumerate() {
            while let Some(&top) = stack.last() {
                if heading_entries[top].level < entry.level {
                    break;
                }
                stack.pop();
            }
            parent_idx[i] = stack.last().copied();
            stack.push(i);
        }

        // Group children by parent
        let mut children_of: Vec<Vec<usize>> = vec![Vec::new(); heading_entries.len()];
        let mut root_children: Vec<usize> = Vec::new();
        for (i, parent) in parent_idx.iter().enumerate() {
            match parent {
                Some(p) => children_of[*p].push(i),
                None => root_children.push(i),
            }
        }

        // Count all visible descendants (open by default)
        fn count_descendants(idx: usize, children_of: &[Vec<usize>]) -> i32 {
            let mut total = children_of[idx].len() as i32;
            for &child in &children_of[idx] {
                total += count_descendants(child, children_of);
            }
            total
        }

        let total_visible: i32 = heading_entries.len() as i32;

        // Write outline items
        for (i, entry) in heading_entries.iter().enumerate() {
            let parent_ref = match parent_idx[i] {
                Some(p) => item_refs[p],
                None => oid,
            };
            let siblings = match parent_idx[i] {
                Some(p) => &children_of[p],
                None => &root_children,
            };
            let pos_in_siblings = siblings.iter().position(|&s| s == i).unwrap();

            let mut item = pdf.outline_item(item_refs[i]);
            item.title(TextStr(&entry.title))
                .parent(parent_ref);

            if pos_in_siblings > 0 {
                item.prev(item_refs[siblings[pos_in_siblings - 1]]);
            }
            if pos_in_siblings + 1 < siblings.len() {
                item.next(item_refs[siblings[pos_in_siblings + 1]]);
            }

            if !children_of[i].is_empty() {
                item.first(item_refs[*children_of[i].first().unwrap()])
                    .last(item_refs[*children_of[i].last().unwrap()])
                    .count(count_descendants(i, &children_of));
            }

            item.dest().page(page_ids[entry.page_idx]).xyz(0.0, entry.y_position, None);
        }

        // Write outline root
        pdf.outline(oid)
            .first(item_refs[root_children[0]])
            .last(item_refs[*root_children.last().unwrap()])
            .count(total_visible);

        Some(oid)
    } else {
        None
    };

    {
        let mut catalog = pdf.catalog(catalog_id);
        catalog.pages(pages_id);
        if let Some(oid) = outline_id {
            catalog.outlines(oid)
                .page_mode(pdf_writer::types::PageMode::UseOutlines);
        }
    }

    if doc.title.is_some() || doc.author.is_some() || doc.subject.is_some() || doc.keywords.is_some() {
        let info_id = alloc();
        let mut info = pdf.document_info(info_id);
        if let Some(ref t) = doc.title {
            info.title(TextStr(t));
        }
        if let Some(ref a) = doc.author {
            info.author(TextStr(a));
        }
        if let Some(ref s) = doc.subject {
            info.subject(TextStr(s));
        }
        if let Some(ref k) = doc.keywords {
            info.keywords(TextStr(k));
        }
        info.producer(TextStr("docxside-pdf"));
    }

    pdf.pages(pages_id)
        .kids(page_ids.iter().copied())
        .count(n as i32);

    let font_pairs: Vec<(String, Ref)> = font_order
        .iter()
        .map(|name| (seen_fonts[name].pdf_name.clone(), seen_fonts[name].font_ref))
        .collect();

    for i in 0..n {
        let (.., si) = page_section_indices[i];
        let sp = &doc.sections[si].properties;
        let mut page = pdf.page(page_ids[i]);
        page.media_box(Rect::new(0.0, 0.0, sp.page_width, sp.page_height))
            .parent(pages_id)
            .contents(content_ids[i]);
        if !page_annot_refs[i].is_empty() {
            page.annotations(page_annot_refs[i].iter().copied());
        }
        {
            let mut resources = page.resources();
            {
                let mut fonts = resources.fonts();
                for (name, font_ref) in &font_pairs {
                    fonts.pair(Name(name.as_bytes()), *font_ref);
                }
            }
            if !image_xobjects.is_empty() {
                let mut xobjects = resources.x_objects();
                for (name, xobj_ref) in image_xobjects {
                    xobjects.pair(Name(name.as_bytes()), *xobj_ref);
                }
            }
            if let Some(alpha_set) = all_page_alpha_states.get(i).filter(|s| !s.is_empty()) {
                let mut gs_dict = resources.ext_g_states();
                for &pct in alpha_set {
                    let gs_name = format!("GSa{pct}");
                    let gs_ref = alpha_gs_refs[&pct];
                    gs_dict.pair(Name(gs_name.as_bytes()), gs_ref);
                }
            }
            if let Some(pat_refs) = all_page_pattern_refs.get(i).filter(|p| !p.is_empty()) {
                let mut patterns = resources.patterns();
                for (name, pat_ref) in pat_refs {
                    patterns.pair(Name(name.as_bytes()), *pat_ref);
                }
            }
        }
    }
}
