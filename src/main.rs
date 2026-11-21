use clap::Parser;
use std::path::PathBuf;

#[derive(Parser)]
#[command(name = "docxide-pdf", about = "Convert DOCX files to PDF")]
struct Args {
    /// Input DOCX file
    input: PathBuf,
    /// Output PDF file (defaults to input with .pdf extension)
    output: Option<PathBuf>,
    /// Print tracked changes as Word's All Markup view (insertions
    /// underlined, deletions struck through, change bars in the margin)
    /// instead of the final text (same as DOCXIDE_REVISIONS=markup)
    #[arg(long)]
    markup: bool,
    /// Keep every picture's source pixels instead of resampling pictures
    /// above 192 ppi down to 192 ppi like Word's online export (same as
    /// DOCXIDE_NO_RESAMPLE=1)
    #[arg(long)]
    no_resample: bool,
    /// Font directory searched before the system ones; repeatable, the first
    /// wins (replaces DOCXIDE_FONTS)
    #[arg(long, value_name = "DIR")]
    fonts: Vec<PathBuf>,
    /// Scan the font directories afresh instead of reading the font index
    /// disk cache (same as DOCXIDE_NO_FONT_CACHE=1)
    #[arg(long)]
    no_font_cache: bool,
    /// Draw the float wrap zones on the page (same as DOCXIDE_DEBUG_WRAP=1)
    #[arg(long)]
    debug_wrap: bool,
}

fn available_path(path: PathBuf) -> PathBuf {
    if !path.exists() {
        return path;
    }
    let stem = path
        .file_stem()
        .unwrap_or_default()
        .to_string_lossy()
        .to_string();
    let ext = path
        .extension()
        .map(|e| format!(".{}", e.to_string_lossy()))
        .unwrap_or_default();
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

    let output = available_path(
        args.output
            .unwrap_or_else(|| args.input.with_extension("pdf")),
    )
    .with_extension("pdf");

    let t0 = std::time::Instant::now();
    let mut options = docxide_pdf::ConvertOptions::from_env();
    if args.markup {
        options = options.revisions(docxide_pdf::RevisionView::Markup);
    }
    if args.no_resample {
        options = options.resample_pictures(false);
    }
    if !args.fonts.is_empty() {
        options = options.font_dirs(args.fonts);
    }
    if args.no_font_cache {
        options = options.font_cache(false);
    }
    if args.debug_wrap {
        options = options.debug_wrap(true);
    }
    let fail = |e: &dyn std::fmt::Display| -> ! {
        eprintln!("Error: {e}");
        std::process::exit(1);
    };
    let docx = std::fs::read(&args.input)
        .unwrap_or_else(|e| fail(&format_args!("{e}: {}", args.input.display())));
    let pdf = docxide_pdf::convert(&docx, &options)
        .unwrap_or_else(|e| fail(&e))
        .pdf;
    let t_render = t0.elapsed();
    std::fs::write(&output, &pdf)
        .unwrap_or_else(|e| fail(&format_args!("{e}: {}", output.display())));
    log::info!(
        "Timing: render={:.1}ms, write={:.1}ms, total={:.1}ms (output {} bytes)",
        t_render.as_secs_f64() * 1000.0,
        (t0.elapsed() - t_render).as_secs_f64() * 1000.0,
        t0.elapsed().as_secs_f64() * 1000.0,
        pdf.len(),
    );
    println!("Converted to {} in {:.0?}", output.display(), t0.elapsed());
}
