//! audit.rs: `plsfix-video audit [--lang en]`: the gate on a cut's full render. The file's
//! format (probe_check), motion (no freeze over 1,5 s, no black), colour (the browser's frame
//! and the decoded video agree at fixed points), determinism (a frame never depends on the
//! frames drawn before it), sound (loudness, and the track in sync with its master), voice (every
//! line inside its slot and clear of the music), layout, copy, and provenance. Exit 1 on any finding.

use crate::cli::render::output;
use crate::cli::{build_dir, music, repo_dir, video_dir};
use crate::io::sha::sha256;
use crate::core::colour::{max_diff, patch_mean, pixels_differing, MAX_CONTENT_PIXELS, NOISE_LEVEL, PROBES, TOLERANCE};
use crate::core::copy_lint::lint;
use crate::core::lang::Lang;
use crate::core::voice::{MIN_OVER_MUSIC_DB, TAIL_S};
use crate::core::shots::validate;
use crate::core::frames::Timeline;
use crate::core::probe_check;
use crate::core::sync::lag;
use crate::core::viewport::Viewport;
use crate::io::comp_page::{write_file, CompPage};
use crate::io::{ffmpeg, loudness, static_server, video_scan};
use anyhow::{bail, Context, Result};
use serde_json::Value;
use std::path::Path;

pub fn run(lang: Lang) -> Result<()> {
    let found = findings(lang)?;
    for f in &found {
        println!("audit: FINDING {f}");
    }
    if !found.is_empty() {
        bail!("audit: {} finding(s)", found.len());
    }
    println!("audit: green (format, motion, colour, determinism, sound, voice, layout, copy, provenance)");
    Ok(())
}

/// Every finding on the current full render of cut `lang`; empty means green.
pub fn findings(lang: Lang) -> Result<Vec<String>> {
    let out = output(false, lang);
    if !out.exists() {
        bail!("audit: {} is missing; run render first", out.display());
    }
    let port = static_server::serve(video_dir())?;
    let mut page = CompPage::open(port, Viewport::STAGE, lang)?;
    let tl = Timeline::new(page.meta.duration, page.meta.fps);
    let mut f = probe_check::check(&ffmpeg::probe(&out)?, tl.frames as u64);
    f.extend(motion(&out, tl.seconds())?);
    f.extend(colours(&out, &mut page)?);
    f.extend(determinism(port, tl.seconds(), lang)?);
    f.extend(sound(&out, lang)?);
    f.extend(narration(lang)?);
    f.extend(layout(lang)?);
    f.extend(copy(lang)?);
    f.extend(provenance()?);
    println!("audit: checked {} ({} frames, {:.2}s)", out.display(), tl.frames, tl.seconds());
    Ok(f)
}

/// No still stretch over 1,5 s before the end card's hold; no black outside the fades.
fn motion(path: &Path, seconds: f64) -> Result<Vec<String>> {
    let mut f = Vec::new();
    for s in video_scan::freezes(path, 1.5)? {
        println!("audit: still stretch {:.2}s from {:.2}s", s.seconds, s.start);
        if s.start < seconds - 2.0 {
            f.push(format!("motion: the picture stands still {:.2}s from {:.2}s", s.seconds, s.start));
        }
    }
    for s in video_scan::blacks(path, 0.2)? {
        if s.start > 0.5 && s.start + s.seconds < seconds - 0.8 {
            f.push(format!("motion: {:.2}s of black from {:.2}s", s.seconds, s.start));
        }
    }
    Ok(f)
}

/// The decoded video matches the browser's own frame at every probe point.
fn colours(path: &Path, page: &mut CompPage) -> Result<Vec<String>> {
    let mut f = Vec::new();
    for &(t, x, y, what) in PROBES {
        let png = page.frame_png(t)?;
        let still = build_dir().join(format!("audit-probe-{t:.2}.png"));
        write_file(&still, &png)?;
        let a = patch_mean(&video_scan::png_rgb(&still)?, 1920, x, y, 4).context("colour: probe outside the still")?;
        let b = patch_mean(&video_scan::frame_rgb(path, t)?, 1920, x, y, 4).context("colour: probe outside the frame")?;
        let d = max_diff(a, b);
        println!("audit: colour at {t}s ({x},{y}) {what}: browser {a:.0?}, video {b:.0?}, off by {d:.1}");
        if d > TOLERANCE {
            f.push(format!("colour at {t}s ({x},{y}) {what}: browser {a:.0?}, video {b:.0?}, off by {d:.1}"));
        }
    }
    Ok(f)
}

/// Frames the determinism check draws twice, from both directions: two inside each act.
const TWICE: &[f64] = &[2.5, 5.0, 11.0, 21.0, 33.0, 52.0, 66.0, 75.0, 82.0, 85.0];

/// Two fresh pages draw the same frames, one forward and one backward after visiting the end;
/// every frame must match its twin up to Chrome's own dithering (a render worker only ever sees
/// its own slice, so a frame that depends on earlier frames differs between workers).
fn determinism(port: u16, seconds: f64, lang: Lang) -> Result<Vec<String>> {
    let mut fwd = CompPage::open(port, Viewport::STAGE, lang)?;
    let mut back = CompPage::open(port, Viewport::STAGE, lang)?;
    back.frame_png(seconds - 0.05)?;
    let first: Vec<Vec<u8>> = TWICE.iter().map(|&t| fwd.frame_png(t)).collect::<Result<_>>()?;
    let mut f = Vec::new();
    for (i, &t) in TWICE.iter().enumerate().rev() {
        let twin = back.frame_png(t)?;
        if twin == first[i] {
            continue;
        }
        // the pair is kept, named by its time, only when it is a finding (the evidence)
        let (pa, pb) = (build_dir().join(format!("audit-twin-{t:.2}-a.png")), build_dir().join(format!("audit-twin-{t:.2}-b.png")));
        write_file(&pa, &first[i])?;
        write_file(&pb, &twin)?;
        let (a, b) = (video_scan::png_rgb(&pa)?, video_scan::png_rgb(&pb)?);
        let n = pixels_differing(&a, &b, NOISE_LEVEL);
        if n > MAX_CONTENT_PIXELS {
            f.push(format!("determinism: the frame at {t}s changes with the frames drawn before it ({n} pixels differ by more than {NOISE_LEVEL}; {})", pa.display()));
        } else {
            let _ = (std::fs::remove_file(&pa), std::fs::remove_file(&pb));
        }
    }
    println!("audit: determinism: {} frames drawn in both orders", TWICE.len());
    Ok(f)
}

/// The soundtrack in the file: online-video loudness, peaks under -1 dBTP, and sample-aligned
/// with the cut's soundtrack it was muxed from (checked in the hero and at the PowerPoint launch).
fn sound(path: &Path, lang: Lang) -> Result<Vec<String>> {
    let mut f = Vec::new();
    let l = loudness::measure(path)?;
    println!("audit: sound: {:.1} LUFS, true peak {:.1} dBTP", l.integrated, l.true_peak);
    if (l.integrated - music::TARGET_LUFS).abs() > 1.5 {
        f.push(format!("sound: {:.1} LUFS, want {} +-1,5", l.integrated, music::TARGET_LUFS));
    }
    if l.true_peak > -1.0 {
        f.push(format!("sound: true peak {:.1} dBTP, want at most -1,0", l.true_peak));
    }
    for t0 in [5.0, 64.5] {
        let master = video_scan::audio_mono(&build_dir().join(lang.audio()), t0, 1.0)?;
        let file = video_scan::audio_mono(path, t0, 1.0)?;
        let d = lag(&master, &file, 480);
        println!("audit: sound: in sync at {t0}s within {d} samples");
        if d.unsigned_abs() > 96 {
            f.push(format!("sound: the track is {:.1} ms off its master at {t0}s", d as f64 / 48.0));
        }
    }
    Ok(f)
}

/// The narrator's report from the cut's last render: every line inside its slot and at least
/// MIN_OVER_MUSIC_DB over the music under it. No report = a render with --no-voice.
fn narration(lang: Lang) -> Result<Vec<String>> {
    let path = build_dir().join(lang.voice_report());
    if !path.is_file() {
        println!("audit: voice: none (rendered with --no-voice)");
        return Ok(Vec::new());
    }
    let report = read_json(&path)?;
    let mut f = Vec::new();
    let lines = report["lines"].as_array().cloned().unwrap_or_default();
    if lines.is_empty() {
        f.push("voice: the report lists no lines".to_string());
    }
    let mut closest = f64::INFINITY;
    for l in &lines {
        let key = l["key"].as_str().unwrap_or("?");
        let (from, to, start, end) = (l["from"].as_f64().unwrap_or(0.0), l["to"].as_f64().unwrap_or(0.0), l["start"].as_f64().unwrap_or(0.0), l["end"].as_f64().unwrap_or(f64::MAX));
        if start < from || end > to - TAIL_S + 1e-3 {
            f.push(format!("voice: {key} speaks {start:.2}-{end:.2}s, outside its slot {from:.2}-{to:.2}s"));
        }
        let over = l["over_music_db"].as_f64().unwrap_or(f64::NEG_INFINITY);
        closest = closest.min(over);
        if over < MIN_OVER_MUSIC_DB {
            f.push(format!("voice: {key} stands only {over:.1} dB over the music, want {MIN_OVER_MUSIC_DB}"));
        }
    }
    println!("audit: voice: {} lines by {}, the closest {closest:.1} dB over the music", lines.len(), report["voice"].as_str().unwrap_or("?"));
    Ok(f)
}

fn read_json(path: &Path) -> Result<Value> {
    let text = std::fs::read_to_string(path).with_context(|| format!("read {}", path.display()))?;
    serde_json::from_str(&text).with_context(|| format!("parse {}", path.display()))
}

/// The page's own layout report from the cut's last render must be empty.
fn layout(lang: Lang) -> Result<Vec<String>> {
    let report = read_json(&build_dir().join(lang.layout()))?;
    Ok(report.as_array().into_iter().flatten().map(|p| format!("layout at {}s {}: {}", p["time"], p["id"].as_str().unwrap_or("?"), p["problem"].as_str().unwrap_or("?"))).collect())
}

/// Every caption on screen and every line the narrator speaks: the cut's copy and voice files.
fn copy(lang: Lang) -> Result<Vec<String>> {
    let mut f = Vec::new();
    for file in [lang.copy_file(), lang.voice_file()] {
        let copy = read_json(&video_dir().join("comp").join(file))?;
        for (key, v) in copy.as_object().into_iter().flatten() {
            f.extend(lint(key, v.as_str().unwrap_or("")));
        }
    }
    Ok(f)
}

/// The screens came from the demo workbook and deck in Office for the web, never live data:
/// the record's own rules, and the workbook the shots were taken on is the demo on disk now.
fn provenance() -> Result<Vec<String>> {
    let m = read_json(&build_dir().join("shots/shots.json"))?;
    let mut f = validate(&m);
    let demo = sha256(&repo_dir().join("demo/out/pls,fix Demo Model.xlsx"))?;
    if m["workbook"]["sha256"].as_str() != Some(demo.as_str()) {
        f.push("provenance: the shots were taken on another build of the demo workbook; capture again".into());
    }
    Ok(f)
}
