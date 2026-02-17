use clap::Parser;
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "docxside-pdf", about = "Convert DOCX files to PDF")]
struct Args {
    /// Input DOCX file
    input: PathBuf,
    /// Output PDF file (defaults to input with .pdf extension)
    output: Option<PathBuf>,
}

fn available_path(path: PathBuf) -> PathBuf {
    if !path.exists() {
        return path;
    }
    let stem = path.file_stem().unwrap_or_default().to_string_lossy().to_string();
    let ext = path.extension().map(|e| format!(".{}", e.to_string_lossy())).unwrap_or_default();
    let parent = path.parent().unwrap_or(std::path::Path::new("."));
    let mut n = 2;
    loop {
        let candidate = parent.join(format!("{stem}({n}){ext}"));
        if !candidate.exists() {
            return candidate;
        }
        n += 1;
    }
}

fn main() {
    env_logger::init();
    let args = Args::parse();

    if !args.input.exists() {
        eprintln!("Error: file not found: {}", args.input.display());
        std::process::exit(1);
    }
    if !args.input.is_file() {
        eprintln!("Error: not a file: {}", args.input.display());
        std::process::exit(1);
    }

    let output = args
        .output
        .unwrap_or_else(|| args.input.with_extension("pdf"));
    let output = available_path(output);

    let t0 = std::time::Instant::now();
    if let Err(e) = docxside_pdf::convert_docx_to_pdf(&args.input, &output) {
        eprintln!("Error: {e}");
        std::process::exit(1);
    }
    let elapsed = t0.elapsed();
    println!("Converted to {} in {:.0?}", output.display(), elapsed);
}
