//! loudness.rs: EBU R128 loudness of a file's audio, measured by ffmpeg's own meter (integrated
//! LUFS and true peak in dBTP), for setting the soundtrack's level and for the audit.

use anyhow::{bail, Context, Result};
use std::path::Path;
use std::process::Command;

#[derive(Debug, Clone, Copy)]
pub struct Loudness {
    pub integrated: f64,
    pub true_peak: f64,
}

pub fn measure(path: &Path) -> Result<Loudness> {
    let out = Command::new("ffmpeg")
        .args(["-nostats", "-hide_banner", "-i"])
        .arg(path)
        .args(["-map", "0:a:0", "-filter:a", "ebur128=peak=true", "-f", "null", "-"])
        .output()
        .context("loudness: spawn ffmpeg")?;
    if !out.status.success() {
        bail!("loudness {}: {}", path.display(), String::from_utf8_lossy(&out.stderr).trim());
    }
    let log = String::from_utf8_lossy(&out.stderr);
    // the summary at the end: "I:  -16.0 LUFS" and, under "True peak:", "Peak:  -1.6 dBFS"
    let summary = log.rsplit("Summary:").next().unwrap_or("");
    let value = |key: &str| summary.lines().find_map(|l| l.trim().strip_prefix(key).and_then(|r| r.split_whitespace().next()).and_then(|v| v.parse::<f64>().ok()));
    match (value("I:"), value("Peak:")) {
        (Some(integrated), Some(true_peak)) => Ok(Loudness { integrated, true_peak }),
        _ => bail!("loudness {}: no R128 summary in ffmpeg's output", path.display()),
    }
}
