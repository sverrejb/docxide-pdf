//! Background child processes driven from the GUI: the visual_comparison test
//! rig for one case (in the main tree) and headless `claude -p` runs that each
//! work one annotation in their own git worktree, so several can run at once.

use crate::annotations::{Annotation, ImageSource};
use crate::cases::CaseInfo;
use std::fs::File;
use std::os::unix::process::CommandExt;
use std::path::{Path, PathBuf};
use std::process::{Child, Command, Stdio};
use std::time::Instant;

#[derive(Clone, Copy)]
pub enum JobKind {
    Regen,
    Fix(u64),
}

#[derive(Clone, Copy)]
pub enum FixOutcome {
    Done,
    Failed,
    Stopped,
}

pub struct Job {
    pub kind: JobKind,
    pub child: Child,
    pub started: Instant,
}

impl Job {
    pub fn elapsed(&self) -> String {
        let s = self.started.elapsed().as_secs();
        format!("{}m {:02}s", s / 60, s % 60)
    }

    pub fn fixing(&self) -> Option<u64> {
        match self.kind {
            JobKind::Fix(id) => Some(id),
            JobKind::Regen => None,
        }
    }

    /// Kill the job and everything it started. Fixes run as their own process
    /// group, so a negative pid reaches the cargo test Claude may be inside.
    pub fn stop(mut self) {
        let _ = Command::new("kill").arg("--").arg(format!("-{}", self.child.id())).status();
        let _ = self.child.kill();
        let _ = self.child.wait();
    }
}

/// Re-run the visual_comparison test rig for one case.
pub fn regen(root: &Path, case: &CaseInfo) -> std::io::Result<Job> {
    let (group, name) = case.group_and_name();
    // Force a fresh conversion — ensure_generated_pdf otherwise skips on mtime.
    let _ = std::fs::remove_file(case.dir.join("generated.pdf"));
    let child = Command::new("cargo")
        .args(["test", "--test", "suite", "--", "--nocapture", "visual_comparison"])
        .env("DOCXIDE_CASE", name)
        .env("DOCXIDE_GROUP", group)
        .current_dir(root)
        .spawn()?;
    Ok(Job { kind: JobKind::Regen, child, started: Instant::now() })
}

pub fn fix_log_path(output_dir: &Path, ann_id: u64) -> PathBuf {
    output_dir.join(format!("fix-{ann_id}.log"))
}

pub fn worktree_path(root: &Path, ann_id: u64) -> PathBuf {
    root.join(".worktrees").join(format!("fix-{ann_id}"))
}

pub fn fix_branch(ann_id: u64) -> String {
    format!("fix/ann-{ann_id}")
}

/// Sets up the worktree and then becomes the claude process, so the GUI never
/// blocks on the setup. A re-run of the same note starts over from HEAD.
/// The worktree lacks two gitignored things the run needs: the vendored fonts
/// (without them every score collapses to the Arial fallback) and the annotation
/// store, which Claude marks fixed in its copy for the GUI to mirror afterwards.
/// Builds share the main tree's target dir instead of cloning it (half a million
/// files): cargo keys artifacts by source path, so worktree and main builds
/// coexist there and reuse the compiled dependencies.
const FIX_SCRIPT: &str = r#"set -e
root=$1 wt=$2 branch=$3 prompt=$4
git -C "$root" worktree remove --force "$wt" 2>/dev/null || true
git -C "$root" worktree add -B "$branch" "$wt"
ln -s ../../fonts "$wt/fonts"
mkdir -p "$wt/tests/output"
cp "$root/tests/output/annotations.json" "$wt/tests/output/"
cd "$wt"
exec env CARGO_TARGET_DIR="$root/target" claude --permission-mode acceptEdits -p "$prompt"
"#;

/// Hand one annotation to a headless Claude session (same flags as ralph/lisa.sh)
/// in its own worktree; its stdout and stderr go to `log`.
pub fn fix(root: &Path, log: &Path, case: &CaseInfo, ann: &Annotation) -> std::io::Result<Job> {
    let (group, name) = case.group_and_name();
    let source = match ann.source {
        ImageSource::Reference => "reference (Word)",
        ImageSource::Generated => "generated (ours)",
    };
    let prompt = format!(
        "Fix annotation id {id} in tests/output/annotations.json for fixture {group}/{name}.\n\
         \n\
         Files: tests/fixtures/{group}/{name}/input.docx (input), \
         tests/fixtures/{group}/{name}/reference.pdf (Word's output). Run \
         ./tools/run-tests.sh --group {group} --case {name} first to produce \
         tests/output/{group}/{name}/generated.pdf (ours) and the page PNGs under \
         tests/output/{group}/{name}/reference/, generated/ and diff/.\n\
         Location: page {page}, on the {source} image, at ({x:.0}, {y:.0}) pt \
         with x from the left and y from the bottom of the page.\n\
         Note: {note}\n\
         \n\
         Analyze the annotation against the input, our output and the reference; use the \
         coordinates to pinpoint the problem; consult the specs via the local-rag MCP. Plan a \
         fix and implement it now, do not defer or reason yourself out of it. Verify with \
         ./tools/run-tests.sh --group {group} --case {name}, then the full ./tools/run-tests.sh \
         for regressions. If the problem is fixed without substantial regressions, commit on \
         the current branch and set \"fixed\": true on annotation id {id} in \
         tests/output/annotations.json. If it is not fixable, append what you learned to \
         ralph/annotations_progress.md and commit that.",
        id = ann.id,
        page = ann.page + 1,
        x = ann.x_pt,
        y = ann.y_pt,
        note = ann.note,
    );
    let log = File::create(log)?;
    let child = Command::new("sh")
        .arg("-c")
        .arg(FIX_SCRIPT)
        .arg("fix")
        .arg(root)
        .arg(worktree_path(root, ann.id))
        .arg(fix_branch(ann.id))
        .arg(&prompt)
        .process_group(0)
        .stdin(Stdio::null())
        .stdout(Stdio::from(log.try_clone()?))
        .stderr(Stdio::from(log))
        .spawn()?;
    Ok(Job { kind: JobKind::Fix(ann.id), child, started: Instant::now() })
}
