use std::collections::HashMap;
use std::io::{Read, Seek};

use crate::model::Comment;

use super::{WML_NS, read_zip_text};

pub(super) fn parse_comments<R: Read + Seek>(
    zip: &mut zip::ZipArchive<R>,
) -> HashMap<u32, Comment> {
    let mut out = HashMap::new();
    let Some(xml_text) = read_zip_text(zip, "word/comments.xml") else {
        return out;
    };
    let Ok(xml) = roxmltree::Document::parse(&xml_text) else {
        return out;
    };
    let root = xml.root_element();
    let mut display_index: u32 = 0;
    for node in root.children() {
        if node.tag_name().namespace() != Some(WML_NS)
            || node.tag_name().name() != "comment"
        {
            continue;
        }
        let Some(id) = node
            .attribute((WML_NS, "id"))
            .and_then(|v| v.parse::<u32>().ok())
        else {
            continue;
        };
        let author = node
            .attribute((WML_NS, "author"))
            .unwrap_or("")
            .to_string();
        let initials = node
            .attribute((WML_NS, "initials"))
            .unwrap_or("")
            .to_string();
        let text = collect_text(node);
        display_index += 1;
        out.insert(
            id,
            Comment {
                id,
                author,
                initials,
                text,
                display_index,
            },
        );
    }
    out
}

fn collect_text(node: roxmltree::Node) -> String {
    let mut buf = String::new();
    for d in node.descendants() {
        if d.tag_name().namespace() != Some(WML_NS) {
            continue;
        }
        match d.tag_name().name() {
            "t" => {
                if let Some(t) = d.text() {
                    buf.push_str(t);
                }
            }
            "tab" => buf.push('\t'),
            "br" => buf.push('\n'),
            "p" => {
                if !buf.is_empty() && !buf.ends_with('\n') {
                    buf.push('\n');
                }
            }
            _ => {}
        }
    }
    buf.trim_end_matches('\n').to_string()
}
