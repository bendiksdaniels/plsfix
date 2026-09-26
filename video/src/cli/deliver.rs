//! deliver.rs: `plsfix-video deliver` copies the full render to the Desktop as
//! "pls,fix video.mp4", only after the audit comes back green on that file, and never over a
//! copy Daniel changed: an existing copy must match video/delivered.sha256 (the last delivery's
//! hash), else it moves to video/build/backup/<date>/ and the delivery stops to ask.
//! `all` = capture, render, deliver: the whole build in one command.

use crate::cli::render::output;
use crate::cli::{audit, build_dir, capture, render, video_dir};
use crate::io::sha::sha256;
use anyhow::{bail, Context, Result};
use std::path::{Path, PathBuf};

pub const FILE_NAME: &str = "pls,fix video.mp4";

/// Where copies go: the Desktop (the playbook's default; pls,fix has no presentation folder).
pub fn targets() -> Result<Vec<PathBuf>> {
    let home = PathBuf::from(std::env::var("HOME").context("deliver: HOME unset")?);
    Ok(vec![home.join("Desktop").join(FILE_NAME)])
}

fn record() -> PathBuf {
    video_dir().join("delivered.sha256")
}

pub fn run() -> Result<()> {
    let found = audit::findings()?;
    if !found.is_empty() {
        for f in &found {
            println!("audit: FINDING {f}");
        }
        bail!("deliver: the audit has {} finding(s); nothing copied", found.len());
    }
    let last = std::fs::read_to_string(record()).ok().map(|s| s.trim().to_string());
    for dst in targets()? {
        guard(&dst, last.as_deref())?;
    }
    for dst in targets()? {
        std::fs::copy(output(false), &dst).with_context(|| format!("deliver: copy to {}", dst.display()))?;
        let mb = std::fs::metadata(&dst)?.len() as f64 / 1_048_576.0;
        println!("deliver: audit green; {} ({mb:.1} MB)", dst.display());
    }
    let hash = sha256(&output(false))?;
    std::fs::write(record(), format!("{hash}\n"))?;
    println!("deliver: recorded {} in {}", &hash[..12], record().display());
    Ok(())
}

/// An existing copy that is not the last delivery is his: it moves aside and the delivery stops.
fn guard(dst: &Path, last: Option<&str>) -> Result<()> {
    if !dst.exists() {
        return Ok(());
    }
    let now = sha256(dst)?;
    if last == Some(now.as_str()) {
        return Ok(());
    }
    let day = std::process::Command::new("date").arg("+%Y-%m-%d").output()?;
    let dir = build_dir().join("backup").join(String::from_utf8_lossy(&day.stdout).trim());
    std::fs::create_dir_all(&dir)?;
    let moved = dir.join(dst.file_name().unwrap_or_default());
    std::fs::rename(dst, &moved).with_context(|| format!("deliver: move {} aside", dst.display()))?;
    bail!("deliver: {} was not the last delivery (changed by hand?); moved to {}; ask before delivering over it", dst.display(), moved.display())
}

pub fn all(workers: u32, track: Option<&Path>) -> Result<()> {
    capture::run()?;
    render::run(false, workers, track)?;
    run()
}
