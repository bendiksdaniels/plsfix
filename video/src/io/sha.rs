//! sha.rs: a file's SHA-256 as lowercase hex, through macOS's own `shasum` (no hashing crate for
//! two call sites). Used by the provenance gate and by the delivery's overwrite guard.

use anyhow::{bail, Context, Result};
use std::path::Path;
use std::process::Command;

pub fn sha256(path: &Path) -> Result<String> {
    let out = Command::new("shasum").args(["-a", "256"]).arg(path).output().with_context(|| format!("sha256 {}", path.display()))?;
    let text = String::from_utf8_lossy(&out.stdout);
    match text.split_whitespace().next() {
        Some(h) if out.status.success() && h.len() == 64 => Ok(h.to_string()),
        _ => bail!("sha256 {}: shasum said {text:?}", path.display()),
    }
}
