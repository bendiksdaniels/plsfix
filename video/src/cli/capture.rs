//! capture.rs: `plsfix-video capture`: builds the demo workbook, brings up the web rig (the
//! signed-in scratch Chrome and the manifest server; what it starts it stops), runs the shot
//! lists in video/capture/ against Excel and PowerPoint for the web, checks every shot and
//! writes video/build/shots/shots.json. One log line per stage with counts.
//! Node drives the browser here, not the engine's Rust CDP: the pane runs in cross-origin frames
//! that Playwright (already the rig's tool, scripts/rig/) reaches and the Rust client cannot.

use crate::cli::{build_dir, repo_dir, video_dir};
use crate::core::shots::{png_size, validate, CSS_H, CSS_W, DPR};
use crate::io::comp_page::write_file;
use crate::io::rig::{self, Rig};
use anyhow::{bail, Context, Result};
use serde_json::{json, Value};
use std::process::Command;

pub fn run() -> Result<()> {
    let repo = repo_dir();
    let status = Command::new("cargo")
        .args(["run", "--quiet", "--manifest-path"])
        .arg(repo.join("demo/Cargo.toml"))
        .status()
        .context("capture: build the demo workbook")?;
    if !status.success() {
        bail!("capture: the demo workbook build failed");
    }
    let out = build_dir().join("shots");
    if out.exists() {
        std::fs::remove_dir_all(&out)?;
    }
    std::fs::create_dir_all(&out)?;
    let _rig = Rig::up(&repo, &build_dir())?;
    let status = Command::new("node")
        .arg(video_dir().join("capture/run.mjs"))
        .arg(&out)
        .env("PLSFIX_RIG_SITE", rig::site()?)
        .current_dir(&repo)
        .status()
        .context("capture: run the shot lists")?;
    if !status.success() {
        bail!("capture: the shot lists failed (see the lines above)");
    }
    let text = std::fs::read_to_string(out.join("capture.json")).context("capture: read capture.json")?;
    let mut record: Value = serde_json::from_str(&text).context("capture: parse capture.json")?;
    let mut findings = validate(&record);
    let (w, h) = ((CSS_W * DPR) as u32, (CSS_H * DPR) as u32);
    for (name, shot) in record["shots"].as_object_mut().into_iter().flatten() {
        let bytes = std::fs::read(out.join(format!("{name}.png"))).with_context(|| format!("capture {name}: read the PNG"))?;
        match png_size(&bytes) {
            Some(size) if size == (w, h) => {}
            other => findings.push(format!("capture {name}: {other:?} pixels, want {w} x {h}")),
        }
        shot["w"] = json!(w);
        shot["h"] = json!(h);
        shot["css_w"] = json!(CSS_W);
        shot["css_h"] = json!(CSS_H);
        shot["dpr"] = json!(DPR);
    }
    if !findings.is_empty() {
        for f in &findings {
            println!("capture: FINDING {f}");
        }
        bail!("capture: {} finding(s); shots.json not written", findings.len());
    }
    let n = record["shots"].as_object().map_or(0, |s| s.len());
    write_file(&out.join("shots.json"), serde_json::to_string_pretty(&record)?.as_bytes())?;
    println!("capture: {n} shots checked -> {} (pane {})", out.display(), record["pane_version"]);
    Ok(())
}
