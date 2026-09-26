//! render.rs: `plsfix-video render [--draft] [--workers N] [--track <file>]`: the composition,
//! frame by frame, in N parallel Chromes into ffmpeg parts, joined, then the soundtrack (written
//! first, from the same page's cues, or the picked track) muxed under it:
//! video/build/plsfix-video.mp4 (draft: 960x540 at 30 fps, plsfix-video-draft.mp4; the English
//! cut plsfix-video-en.mp4). Also writes the page's layout report.

use crate::cli::{build_dir, music, video_dir};
use crate::core::frames::{chunks, Timeline};
use crate::core::lang::Lang;
use crate::core::viewport::Viewport;
use crate::io::comp_page::{write_file, CompPage};
use crate::io::ffmpeg;
use crate::io::render_worker::{self, Job};
use crate::io::static_server;
use anyhow::{bail, Context, Result};
use std::path::{Path, PathBuf};
use std::thread;
use std::time::Instant;

/// The finished file for a full or a draft render of cut `lang`.
pub fn output(draft: bool, lang: Lang) -> PathBuf {
    build_dir().join(lang.mp4(draft))
}

pub fn run(draft: bool, workers: u32, track: Option<&Path>, lang: Lang) -> Result<()> {
    let started = Instant::now();
    music::run(track)?;
    let port = static_server::serve(video_dir())?;
    let (meta, layout) = {
        let mut probe = CompPage::open(port, Viewport::STAGE, lang)?;
        let layout = probe.cdp.eval("window.__layoutCheck()").context("render: __layoutCheck")?;
        (probe.meta, layout)
    };
    write_file(&build_dir().join(lang.layout()), serde_json::to_string_pretty(&layout)?.as_bytes())?;
    let problems = layout.as_array().map_or(0, Vec::len);
    let fps = if draft { 30 } else { meta.fps };
    let vp = if draft { Viewport { dpr: 0.5, ..Viewport::STAGE } } else { Viewport::STAGE };
    let tl = Timeline::new(meta.duration, fps);
    let parts_dir = build_dir().join("parts");
    if parts_dir.exists() {
        std::fs::remove_dir_all(&parts_dir)?;
    }
    std::fs::create_dir_all(&parts_dir)?;
    let jobs: Vec<Job> = chunks(tl.frames, workers).into_iter().enumerate().map(|(index, frames)| Job {
        index, frames, fps, vp, port, lang, out: parts_dir.join(format!("part-{index:02}.mp4")),
    }).collect();
    println!("render: {} frames ({:.2}s at {fps} fps, {}x{}) on {} workers; layout problems: {problems}", tl.frames, tl.seconds(), (vp.w as f64 * vp.dpr) as u32, (vp.h as f64 * vp.dpr) as u32, jobs.len());
    let results: Vec<Result<u32>> = thread::scope(|s| {
        let handles: Vec<_> = jobs.iter().map(|job| s.spawn(move || render_worker::run(job))).collect();
        handles.into_iter().map(|h| h.join().unwrap_or_else(|_| bail!("render worker panicked"))).collect()
    });
    let mut written = 0;
    for r in results {
        written += r?;
    }
    if written != tl.frames {
        bail!("render: wrote {written} frames, the timeline has {}", tl.frames);
    }
    let parts: Vec<PathBuf> = jobs.iter().map(|j| j.out.clone()).collect();
    let out = output(draft, lang);
    let silent = out.with_extension("silent.mp4");
    ffmpeg::concat(&parts, &silent)?;
    ffmpeg::mux(&silent, &music::path(), &out)?;
    std::fs::remove_file(&silent).ok();
    println!("render: {} ({written} frames) in {:.1}s", out.display(), started.elapsed().as_secs_f64());
    Ok(())
}
