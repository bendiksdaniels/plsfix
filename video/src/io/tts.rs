//! tts.rs: one voice-over line through edge-tts (Microsoft's neural voices, Latvian included;
//! the maintained client is Python's, called as a CLI the way ffmpeg is), cached by voice, rate
//! and text in build/voice-cache/, so a re-render asks the service only for lines that changed.
use crate::core::voice::cache_key;
use anyhow::{bail, Context, Result};
use std::path::{Path, PathBuf};
use std::process::Command;

/// The spoken line as an mp3 in `cache`, made on first use.
pub fn speak(voice: &str, rate: i32, text: &str, cache: &Path) -> Result<PathBuf> {
    let out = cache.join(format!("{}.mp3", cache_key(voice, rate, text)));
    if out.is_file() && std::fs::metadata(&out)?.len() > 0 {
        return Ok(out);
    }
    std::fs::create_dir_all(cache).with_context(|| format!("tts: mkdir {}", cache.display()))?;
    let part = out.with_extension("part.mp3");
    let res = Command::new("edge-tts")
        .args(["--voice", voice, &format!("--rate={rate:+}%"), "--text", text, "--write-media"])
        .arg(&part)
        .output()
        .context("tts: edge-tts not found (brew install edge-tts, or pipx install edge-tts)")?;
    if !res.status.success() || !part.is_file() {
        bail!("tts: edge-tts refused \"{text}\" ({voice}, {rate:+}%): {}", String::from_utf8_lossy(&res.stderr).trim());
    }
    std::fs::rename(&part, &out)?;
    Ok(out)
}
