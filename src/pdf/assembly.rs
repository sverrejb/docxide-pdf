use std::collections::{BTreeSet, HashMap, HashSet};

use pdf_writer::{Content, Filter, Name, Pdf, Rect, Ref, Str, TextStr};

use crate::fonts::FontEntry;
use crate::model::{Document, PageBorderDisplay, PageBorders, ParagraphBorder, SectionProperties};

use super::GradientSpec;
use super::comments::{page_zoom, render_comment_pane};
use super::layout::LinkAnnotation;

pub(crate) struct HeadingEntry {
    pub(super) title: String,
    pub(super) level: u8,
    pub(super) page_idx: usize,
    pub(super) y_position: f32,
}

/// Draw a `w:pgBorders` box for one page into its own content stream (page
/// coordinates, unscaled). Each edge is stroked independently so asymmetric
/// borders work; horizontal edges overrun by the adjacent vertical edge's
/// half-width to close the butt-capped corners.
fn render_page_borders(pb: &PageBorders, sp: &SectionProperties) -> Content {
    let mut content = Content::new();
    // For offsetFrom="page", @space is the gap from the page edge to the border;
    // the stroke is centered on (space + width/2) inward. For "text", @space is
    // measured outward from the text margins.
    let edge_pos = |b: &ParagraphBorder, margin: f32| {
        let inset = b.space_pt + b.extent() / 2.0;
        if pb.offset_from_page {
            inset
        } else {
            margin - inset
        }
    };
    let left_x = pb.left.as_ref().map(|b| edge_pos(b, sp.margin_left));
    let right_x = pb
        .right
        .as_ref()
        .map(|b| sp.page_width - edge_pos(b, sp.margin_right));
    let top_y = pb
        .top
        .as_ref()
        .map(|b| sp.page_height - edge_pos(b, sp.margin_top));
    let bottom_y = pb.bottom.as_ref().map(|b| edge_pos(b, sp.margin_bottom));

    // Fall back to the opposite/adjacent edge's coordinate so a line still spans
    // the full box when only some edges are present.
    let lx = left_x.unwrap_or(0.0);
    let rx = right_x.unwrap_or(sp.page_width);
    let ty = top_y.unwrap_or(sp.page_height);
    let by = bottom_y.unwrap_or(0.0);
    let l_ext = pb.left.as_ref().map(|b| b.extent() / 2.0).unwrap_or(0.0);
    let r_ext = pb.right.as_ref().map(|b| b.extent() / 2.0).unwrap_or(0.0);

    let stroke = super::color::stroke_border;
    if let (Some(b), Some(y)) = (&pb.top, top_y) {
        stroke(&mut content, b, (lx - l_ext, y), (rx + r_ext, y), -1.0);
    }
    if let (Some(b), Some(y)) = (&pb.bottom, bottom_y) {
        stroke(&mut content, b, (lx - l_ext, y), (rx + r_ext, y), 1.0);
    }
    if let (Some(b), Some(x)) = (&pb.left, left_x) {
        stroke(&mut content, b, (x, ty), (x, by), 1.0);
    }
    if let (Some(b), Some(x)) = (&pb.right, right_x) {
        stroke(&mut content, b, (x, ty), (x, by), -1.0);
    }
    content
}

/// XMP packet for the catalog `/Metadata` stream; PDF/UA wants dc:title here,
/// not only in the Info dictionary.
fn xmp_packet(doc: &Document, lang: &str, claims_ua: bool) -> String {
    let esc = |s: &str| {
        s.replace('&', "&amp;")
            .replace('<', "&lt;")
            .replace('>', "&gt;")
            .replace('"', "&quot;")
    };
    let mut props = format!(
        "<pdf:Producer>docxide-pdf</pdf:Producer>\
         <dc:language><rdf:Bag><rdf:li>{}</rdf:li></rdf:Bag></dc:language>",
        esc(lang)
    );
    if let Some(t) = &doc.title {
        props += &format!(
            "<dc:title><rdf:Alt><rdf:li xml:lang=\"x-default\">{}</rdf:li></rdf:Alt></dc:title>",
            esc(t)
        );
    }
    if let Some(a) = &doc.author {
        props += &format!(
            "<dc:creator><rdf:Seq><rdf:li>{}</rdf:li></rdf:Seq></dc:creator>",
            esc(a)
        );
    }
    if claims_ua {
        props += "<pdfuaid:part>1</pdfuaid:part>";
    }
    format!(
        "<?xpacket begin=\"\u{feff}\" id=\"W5M0MpCehiHzreSzNTczkc9d\"?>\
         <x:xmpmeta xmlns:x=\"adobe:ns:meta/\">\
         <rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">\
         <rdf:Description rdf:about=\"\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" \
         xmlns:pdf=\"http://ns.adobe.com/pdf/1.3/\" \
         xmlns:pdfuaid=\"http://www.aiim.org/pdfua/ns/id/\">{props}</rdf:Description>\
         </rdf:RDF></x:xmpmeta><?xpacket end=\"w\"?>"
    )
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
    // Per-page §17.6.23 vertical-align offset: shift the body block down by this
    // many points (0 = top-aligned, the default).
    valign_offsets: Vec<f32>,
    all_contents: Vec<Content>,
    all_deferred_shapes: Vec<Vec<(u32, Content)>>,
    all_hf_contents: &mut [Option<Content>],
    all_page_links: &[Vec<LinkAnnotation>],
    all_page_comment_anchors: &[Vec<(u32, f32, f32, f32)>],
    all_page_alpha_states: &[HashSet<u8>],
    all_page_gradient_specs: &[Vec<GradientSpec>],
    page_section_indices: &[(usize, bool, usize)],
    seen_fonts: &HashMap<String, FontEntry>,
    font_order: &[String],
    image_xobjects: &[(String, Ref)],
    doc: &Document,
    bookmark_positions: &HashMap<String, (usize, f32)>,
    heading_entries: &[HeadingEntry],
    tags: &super::tagging::Tags,
) {
    let n = all_contents.len();
    let page_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();
    let content_ids: Vec<Ref> = (0..n).map(|_| alloc()).collect();

    let has_any_comments =
        !doc.comments.is_empty() && all_page_comment_anchors.iter().any(|p| !p.is_empty());
    let comment_keys = has_any_comments.then(|| super::fonts::comment_pane_keys(doc));

    // (Link element, page, annotation) for the structure tree's OBJR kids; the
    // annotations' /StructParent keys follow the pages' keys.
    let mut tagged_annots: Vec<(usize, usize, Ref)> = Vec::new();
    let page_annot_refs: Vec<Vec<Ref>> = all_page_links
        .iter()
        .enumerate()
        .map(|(page, links)| {
            links
                .iter()
                .filter_map(|link| {
                    // An internal link to a bookmark that doesn't exist gets no annotation.
                    let dest = match link.url.strip_prefix('#') {
                        Some(anchor) => Some(*bookmark_positions.get(anchor)?),
                        None => None,
                    };
                    let annot_ref = alloc();
                    let mut annot = pdf.annotation(annot_ref);
                    annot
                        .subtype(pdf_writer::types::AnnotationType::Link)
                        .rect(link.rect)
                        .border(0.0, 0.0, 0.0, None);
                    // PDF/UA 7.18.5-2: a description of the link (Word omits it).
                    let text = link.text.trim();
                    if !text.is_empty() {
                        annot.contents(TextStr(text));
                    }
                    if let Some(node) = link.node {
                        annot.struct_parent((n + tagged_annots.len()) as i32);
                        tagged_annots.push((node, page, annot_ref));
                    }
                    if let Some((dest_page_idx, dest_y)) = dest {
                        debug_assert!(
                            dest_page_idx < page_ids.len(),
                            "bookmark '{}' points to page {dest_page_idx} but only {} pages exist",
                            link.url,
                            page_ids.len(),
                        );
                        let safe_idx = dest_page_idx.min(page_ids.len().saturating_sub(1));
                        annot
                            .action()
                            .action_type(pdf_writer::types::ActionType::GoTo)
                            .destination()
                            .page(page_ids[safe_idx])
                            .xyz(0.0, dest_y, None);
                    } else {
                        annot
                            .action()
                            .action_type(pdf_writer::types::ActionType::Uri)
                            .uri(Str(link.url.as_bytes()));
                    }
                    Some(annot_ref)
                })
                .collect()
        })
        .collect();

    // BTreeSet: object numbers must not follow the per-process HashSet order.
    let all_alpha_values: BTreeSet<u8> = all_page_alpha_states
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
                        let encode: Vec<f32> = sub_refs.iter().flat_map(|_| [0.0, 1.0]).collect();

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
                    let half_len =
                        ((spec.w / 2.0 * cos_a).powi(2) + (spec.h / 2.0 * sin_a).powi(2)).sqrt();
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

    let mut all_deferred_shapes = all_deferred_shapes.into_iter();
    for (i, c) in all_contents.into_iter().enumerate() {
        let mut body_raw = super::tagging::finish_artifact_content(c);
        // Anchored shapes paint above the page's text layer, pre-sorted by
        // relativeHeight at flush time
        // Each buffer opens inside an artifact (tagging::artifact_content) so
        // a textbox's text can be tagged.
        for (_, shape) in all_deferred_shapes.next().into_iter().flatten() {
            body_raw.extend_from_slice(&super::tagging::finish_artifact_content(shape));
        }

        // When the document has comments, Word's PDF export renders body
        // content scaled (so glyphs become 9.16pt from 12pt) and shifted down
        // so the first line sits at roughly the pane's top. We replicate that
        // by wrapping the body content stream in a `q s 0 0 s tx ty cm ... Q`
        // transform. The comment pane and connectors stay in unscaled PDF
        // coordinates and are rendered into a separate stream appended after.
        let (.., zoom_si) = page_section_indices[i];
        let zoom_sp = &doc.sections[zoom_si].properties;
        let (zoom, zoom_tx, zoom_ty) = page_zoom(
            zoom_sp.page_width,
            zoom_sp.page_height,
            zoom_sp.margin_right,
        );
        let (scale_prefix, scale_suffix) = if has_any_comments {
            (
                format!("q {zoom} 0 0 {zoom} {zoom_tx} {zoom_ty} cm\n").into_bytes(),
                b"\nQ\n".to_vec(),
            )
        } else {
            (Vec::new(), Vec::new())
        };

        // §17.6.23 vAlign: translate the body block down by the precomputed
        // offset. Wraps the body (and the comment-scale wrapper) but not the
        // header/footer or page-border streams, which stay page-fixed.
        let valign_off = valign_offsets.get(i).copied().unwrap_or(0.0);
        let (valign_prefix, valign_suffix) = if valign_off > 0.01 {
            (
                format!("q 1 0 0 1 0 {:.3} cm\n", -valign_off).into_bytes(),
                b"\nQ\n".to_vec(),
            )
        } else {
            (Vec::new(), Vec::new())
        };

        let mut pane_raw = Vec::new();
        if let Some((body_key, label_key)) = &comment_keys {
            let transformed: Vec<(u32, f32, f32, f32)> = all_page_comment_anchors[i]
                .iter()
                .map(|(id, x, y, fs)| (*id, zoom_tx + zoom * x, zoom_ty + zoom * y, zoom * fs))
                .collect();
            let mut pane_content = Content::new();
            render_comment_pane(
                &mut pane_content,
                &doc.comments,
                &transformed,
                zoom_sp,
                (zoom, zoom_tx, zoom_ty),
                seen_fonts,
                (body_key, label_key, doc.comment_font_size),
            );
            pane_raw = pane_content.finish().to_vec();
        }

        let mut combined: Vec<u8> = Vec::new();
        if let Some(hf) = all_hf_contents[i].take() {
            // Word zooms headers and footers with the body (door_air_cooling_unit_spec).
            combined.extend_from_slice(&scale_prefix);
            super::tagging::wrap_artifact(&mut combined, hf.finish().as_slice(), true);
            combined.extend_from_slice(&scale_suffix);
        }
        // Page border box (page coords, unscaled — outside the comment-scale
        // wrapper). @display gates which pages get it; is_first = first page of
        // its section, matching ST_PageBorderDisplay firstPage/notFirstPage.
        {
            let (_, is_first, content_si) = page_section_indices[i];
            let sp = &doc.sections[content_si].properties;
            if let Some(pb) = &sp.page_borders {
                let show = match pb.display {
                    PageBorderDisplay::AllPages => true,
                    PageBorderDisplay::FirstPage => is_first,
                    PageBorderDisplay::NotFirstPage => !is_first,
                };
                if show {
                    let borders = render_page_borders(pb, sp).finish();
                    super::tagging::wrap_artifact(&mut combined, borders.as_slice(), false);
                }
            }
        }
        combined.extend_from_slice(&valign_prefix);
        combined.extend_from_slice(&scale_prefix);
        combined.extend_from_slice(body_raw.as_slice());
        combined.extend_from_slice(&scale_suffix);
        combined.extend_from_slice(&valign_suffix);
        if !pane_raw.is_empty() {
            combined.push(b'\n');
            super::tagging::wrap_artifact(&mut combined, &pane_raw, false);
        }
        let compressed = miniz_oxide::deflate::compress_to_vec_zlib(&combined, 6);
        pdf.stream(content_ids[i], &compressed)
            .filter(Filter::FlateDecode);
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
            item.title(TextStr(&entry.title)).parent(parent_ref);

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

            item.dest()
                .page(page_ids[entry.page_idx])
                .xyz(0.0, entry.y_position, None);
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

    // The language most of the text is in (`document_lang`); passages in
    // another one carry their own `/Lang`.
    let lang = tags.lang.as_str();
    // Claim PDF/UA-1 only when every machine check that depends on the
    // document is known to pass: a title, alt on every Figure, heading order,
    // all fonts embedded (no standard-14 fallback) and no .notdef drawn. The
    // accessibility test fails any claiming PDF that veraPDF faults.
    let claims_ua = !tags.is_empty()
        && doc.title.as_deref().is_some_and(|t| !t.trim().is_empty())
        && tags.ua_structure_ok()
        && seen_fonts
            .values()
            .all(|f| f.char_to_gid.is_some() && !f.drew_notdef.get());
    let metadata_id = alloc();
    pdf.metadata(metadata_id, xmp_packet(doc, lang, claims_ua).as_bytes());
    let struct_root = (!tags.is_empty()).then(|| tags.write(pdf, alloc, &page_ids, &tagged_annots));
    {
        let mut catalog = pdf.catalog(catalog_id);
        catalog.pages(pages_id);
        if let Some(oid) = outline_id {
            catalog
                .outlines(oid)
                .page_mode(pdf_writer::types::PageMode::UseOutlines);
        }
        catalog.lang(TextStr(lang)).metadata(metadata_id);
        catalog.viewer_preferences().display_doc_title(true);
        if let Some(root) = struct_root {
            catalog.pair(Name(b"StructTreeRoot"), root);
            catalog.mark_info().marked(true);
        }
    }

    if doc.title.is_some()
        || doc.author.is_some()
        || doc.subject.is_some()
        || doc.keywords.is_some()
    {
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
        info.producer(TextStr("docxide-pdf"));
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
            .contents(content_ids[i])
            .tab_order(pdf_writer::types::TabOrder::StructureOrder);
        if let Some(key) = tags.struct_parents(i) {
            page.struct_parents(key);
        }
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
                for pct in alpha_set.iter().copied().collect::<BTreeSet<u8>>() {
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
