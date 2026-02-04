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

    if let Err(e) = docxside_pdf::convert_docx_to_pdf(&args.input, &output) {
        eprintln!("Error: {e}");
        std::process::exit(1);
    }
}
