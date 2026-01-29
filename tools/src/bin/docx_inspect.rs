//! Inspect the raw contents of a DOCX file (which is a ZIP).
//!
//! Usage:
//!   docx-inspect <file.docx>                    list all ZIP entries
//!   docx-inspect <file.docx> <internal/path>    dump file (XML auto-formatted via xmllint)
//!   docx-inspect <file.docx> --grep <pattern>   search all XML/rels files

use std::fs;
use std::io::{self, Read, Write};
use std::process::{Command, Stdio};
use zip::ZipArchive;

fn main() {
    let args: Vec<String> = std::env::args().collect();

    if args.len() < 2 {
        eprintln!("Usage:");
        eprintln!("  docx-inspect <file.docx>                    list ZIP entries");
        eprintln!("  docx-inspect <file.docx> <internal/path>    dump file");
        eprintln!("  docx-inspect <file.docx> --grep <pattern>   search XML/rels files");
        std::process::exit(1);
    }

    let file = fs::File::open(&args[1]).unwrap_or_else(|e| {
        eprintln!("Cannot open '{}': {e}", args[1]);
        std::process::exit(1);
    });
    let mut archive = ZipArchive::new(file).unwrap_or_else(|e| {
        eprintln!("Not a valid ZIP/DOCX: {e}");
        std::process::exit(1);
    });

    match args.get(2).map(String::as_str) {
        None => list_entries(&mut archive),
        Some("--grep") => {
            let pattern = args.get(3).unwrap_or_else(|| {
                eprintln!("--grep requires a pattern");
                std::process::exit(1);
            });
            grep_entries(&mut archive, pattern);
        }
        Some(path) => dump_entry(&mut archive, path),
    }
}

fn list_entries(archive: &mut ZipArchive<fs::File>) {
    println!("{:>9}  {}", "bytes", "path");
    println!("{}", "─".repeat(55));
    for i in 0..archive.len() {
        let entry = archive.by_index(i).unwrap();
        if !entry.is_dir() {
            println!("{:>9}  {}", entry.size(), entry.name());
        }
    }
}

fn dump_entry(archive: &mut ZipArchive<fs::File>, path: &str) {
    let mut entry = archive.by_name(path).unwrap_or_else(|_| {
        eprintln!("'{}' not found in archive", path);
        eprintln!("Run without a path argument to list available entries.");
        std::process::exit(1);
    });

    let mut content = Vec::new();
    entry.read_to_end(&mut content).unwrap();

    let is_xml = path.ends_with(".xml") || path.ends_with(".rels");
    if is_xml {
        // Try xmllint for pretty-printing; fall through to raw on failure
        if let Ok(mut child) = Command::new("xmllint")
            .args(["--format", "-"])
            .stdin(Stdio::piped())
            .stdout(Stdio::inherit())
            .stderr(Stdio::null())
            .spawn()
        {
            if let Some(mut stdin) = child.stdin.take() {
                stdin.write_all(&content).unwrap();
            }
            child.wait().unwrap();
            return;
        }
    }

    io::stdout().write_all(&content).unwrap();
}

fn grep_entries(archive: &mut ZipArchive<fs::File>, pattern: &str) {
    let mut found = 0;
    for i in 0..archive.len() {
        let mut entry = archive.by_index(i).unwrap();
        let name = entry.name().to_owned();
        if name.ends_with(".xml") || name.ends_with(".rels") {
            let mut content = String::new();
            entry.read_to_string(&mut content).unwrap_or(0);
            for (n, line) in content.lines().enumerate() {
                if line.contains(pattern) {
                    println!("{}:{}: {}", name, n + 1, line.trim());
                    found += 1;
                }
            }
        }
    }
    if found == 0 {
        eprintln!("No matches for '{pattern}'");
    }
}
