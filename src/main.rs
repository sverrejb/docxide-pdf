use clap::Parser;
use clap::builder::FalseyValueParser;
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
    /// instead of the final text (or DOCXIDE_REVISIONS=markup)
    #[arg(long)]
    markup: bool,
    /// Keep every picture's source pixels instead of resampling pictures
    /// above 192 ppi down to 192 ppi like Word's online export
    #[arg(long, env = "DOCXIDE_NO_RESAMPLE", value_parser = FalseyValueParser::new())]
    no_resample: bool,
    /// Font directory searched before the system ones; repeatable, the first
    /// wins (replaces DOCXIDE_FONTS, a path list like PATH)
    #[arg(long, value_name = "DIR")]
    fonts: Vec<PathBuf>,
    /// Scan the font directories afresh instead of reading the font index
    /// disk cache
    #[arg(long, env = "DOCXIDE_NO_FONT_CACHE", value_parser = FalseyValueParser::new())]
    no_font_cache: bool,
    /// Draw the float wrap zones on the page
    #[arg(long, env = "DOCXIDE_DEBUG_WRAP", value_parser = FalseyValueParser::new())]
    debug_wrap: bool,
    /// Write larger PDF streams faster, for a local preview
    #[arg(long)]
    fast_compression: bool,
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
            .unwrap_or_else(|| args.input.clone())
            .with_extension("pdf"),
    );

    let t0 = std::time::Instant::now();
    // Not a flag's on/off value, and a path list clap can't split with
    // std::env::split_paths, so these two variables are read here.
    let markup = args.markup
        || std::env::var("DOCXIDE_REVISIONS")
            .is_ok_and(|v| v.trim().eq_ignore_ascii_case("markup"));
    let mut fonts = args.fonts;
    if fonts.is_empty()
        && let Some(dirs) = std::env::var_os("DOCXIDE_FONTS")
    {
        fonts = std::env::split_paths(&dirs)
            .filter(|dir| !dir.as_os_str().is_empty())
            .collect();
    }
    let view = if markup {
        docxide_pdf::RevisionView::Markup
    } else {
        docxide_pdf::RevisionView::Final
    };
    let compression = if args.fast_compression
        || std::env::var("DOCXIDE_COMPRESSION").is_ok_and(|v| v.trim().eq_ignore_ascii_case("fast"))
    {
        docxide_pdf::Compression::Fast
    } else {
        docxide_pdf::Compression::Smallest
    };
    let options = docxide_pdf::ConvertOptions::new()
        .revisions(view)
        .resample_pictures(!args.no_resample)
        .font_dirs(fonts)
        .font_cache(!args.no_font_cache)
        .debug_wrap(args.debug_wrap)
        .compression(compression);
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
