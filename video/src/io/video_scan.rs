//! video_scan.rs: reads a rendered file back for the audit: frozen and black stretches (ffmpeg's
//! own detectors) and single frames or PNGs decoded to raw RGB with the BT.709 matrix, the same
//! way a player decodes them.

use anyhow::{bail, Context, Result};
use std::path::Path;
use std::process::Command;

/// A stretch ffmpeg reported: start and length in seconds.
#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Stretch {
    pub start: f64,
    pub seconds: f64,
}

fn ffmpeg_stderr(args: &[&str]) -> Result<String> {
    let out = Command::new("ffmpeg").args(args).output().context("ffmpeg scan: spawn")?;
    if !out.status.success() {
        bail!("ffmpeg scan {}: {}", args.join(" "), String::from_utf8_lossy(&out.stderr).trim());
    }
    Ok(String::from_utf8_lossy(&out.stderr).into_owned())
}

/// Reads `key_start: <x>` and `key_duration: <y>` pairs from a detector's log.
fn stretches(log: &str, start_key: &str, dur_key: &str) -> Vec<Stretch> {
    let value = |line: &str, key: &str| line.split(key).nth(1).and_then(|r| r.split_whitespace().next()).and_then(|v| v.parse::<f64>().ok());
    let mut out = Vec::new();
    let mut start = None;
    for line in log.lines() {
        if let Some(s) = value(line, start_key) {
            start = Some(s);
        }
        if let (Some(s), Some(d)) = (start, value(line, dur_key)) {
            out.push(Stretch { start: s, seconds: d });
            start = None;
        }
    }
    out
}

/// Stretches where the picture does not change (noise under -66 dB) for at least `min` s.
pub fn freezes(path: &Path, min: f64) -> Result<Vec<Stretch>> {
    let p = path.to_string_lossy();
    let log = ffmpeg_stderr(&["-v", "info", "-i", &p, "-vf", &format!("freezedetect=n=0.0005:d={min}"), "-map", "0:v", "-f", "null", "-"])?;
    Ok(stretches(&log, "lavfi.freezedetect.freeze_start:", "lavfi.freezedetect.freeze_duration:"))
}

/// Stretches of near-black picture at least `min` s long.
pub fn blacks(path: &Path, min: f64) -> Result<Vec<Stretch>> {
    let p = path.to_string_lossy();
    let log = ffmpeg_stderr(&["-v", "info", "-i", &p, "-vf", &format!("blackdetect=d={min}:pix_th=0.06"), "-f", "null", "-"])?;
    Ok(stretches(&log, "black_start:", "black_duration:"))
}

fn rgb(args: &[&str]) -> Result<Vec<u8>> {
    let out = Command::new("ffmpeg").args(args).output().context("ffmpeg decode: spawn")?;
    if !out.status.success() {
        bail!("ffmpeg decode {}: {}", args.join(" "), String::from_utf8_lossy(&out.stderr).trim());
    }
    Ok(out.stdout)
}

/// The frame shown at `t` seconds, as RGB bytes (BT.709, limited range, as tagged).
pub fn frame_rgb(path: &Path, t: f64) -> Result<Vec<u8>> {
    let p = path.to_string_lossy();
    rgb(&["-v", "error", "-ss", &format!("{t:.4}"), "-i", &p, "-frames:v", "1", "-vf", "scale=in_color_matrix=bt709:in_range=tv,format=rgb24", "-f", "rawvideo", "-"])
}

/// `dur` seconds of a file's first audio stream from `t0`, mixed to mono, 48 kHz floats.
pub fn audio_mono(path: &Path, t0: f64, dur: f64) -> Result<Vec<f32>> {
    let p = path.to_string_lossy();
    let raw = rgb(&["-v", "error", "-ss", &format!("{t0:.4}"), "-t", &format!("{dur:.4}"), "-i", &p, "-map", "0:a:0", "-ac", "1", "-ar", "48000", "-f", "f32le", "-"])?;
    Ok(raw.chunks_exact(4).map(|b| f32::from_le_bytes([b[0], b[1], b[2], b[3]])).collect())
}

/// A PNG's pixels as RGB bytes.
pub fn png_rgb(path: &Path) -> Result<Vec<u8>> {
    let p = path.to_string_lossy();
    rgb(&["-v", "error", "-i", &p, "-f", "rawvideo", "-pix_fmt", "rgb24", "-"])
}
