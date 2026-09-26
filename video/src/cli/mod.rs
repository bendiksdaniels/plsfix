//! cli: the `plsfix-video` commands and the folders they share.
//! capture = real screens from Excel and PowerPoint for the web; music = the soundtrack (or a picked track); render = the MP4
//! (music first); audit = the gate; deliver = to the Desktop after a green audit; all = capture,
//! render, deliver in a row; still = single frames for QA. `--lang en` picks the English cut
//! (the same picture and sound, the English copy file) for render, audit, deliver, all, still.
//! Paths come from this crate's own folder (video/), so the tool works from any directory.

mod audit;
mod capture;
mod deliver;
mod music;
mod render;
mod still;

use anyhow::Result;
use crate::core::lang::Lang;
use clap::{Parser, Subcommand};
use std::path::PathBuf;
use std::process::ExitCode;

#[derive(Parser)]
#[command(name = "plsfix-video", about = "Build the pls,fix motion-graphic video")]
struct Cli {
    #[command(subcommand)]
    cmd: Cmd,
    /// The cut: lv (Latvian captions, the default) or en (English captions).
    #[arg(long, global = true, default_value = "lv")]
    lang: Lang,
}

#[derive(Subcommand)]
enum Cmd {
    /// Capture the real pls,fix screens in Excel and PowerPoint for the web (the rig) to video/build/shots/.
    Capture,
    /// Write the soundtrack (score + sound design from the picture's cues) to video/build/music.wav.
    Music {
        /// Use this audio file (mp3, m4a, wav...) instead of the score: cut to length, faded, levelled.
        #[arg(long)]
        track: Option<PathBuf>,
    },
    /// Render the composition to video/build/plsfix-video.mp4 (draft: 960x540 at 30 fps).
    Render {
        #[arg(long)]
        draft: bool,
        #[arg(long, default_value_t = 5)]
        workers: u32,
        /// Use this audio file instead of the score (see `music --track`).
        #[arg(long)]
        track: Option<PathBuf>,
    },
    /// Check the full render: format, motion, colour, determinism, sound, layout, copy, provenance (exit 1 on any).
    Audit,
    /// Audit, then copy the full render to the Desktop as "pls,fix video.mp4" (English: "pls,fix video EN.mp4").
    Deliver,
    /// Capture, render and deliver in one go.
    All {
        #[arg(long, default_value_t = 5)]
        workers: u32,
        /// Use this audio file instead of the score (see `music --track`).
        #[arg(long)]
        track: Option<PathBuf>,
    },
    /// Render single frames at the given times (seconds) to video/build/stills/.
    Still {
        times: Vec<f64>,
    },
}

/// video/ (this crate's folder).
pub fn video_dir() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
}

/// The pls,fix repository root (video/..).
pub fn repo_dir() -> PathBuf {
    video_dir().parent().map(PathBuf::from).unwrap_or_else(video_dir)
}

/// video/build: every generated file.
pub fn build_dir() -> PathBuf {
    video_dir().join("build")
}

pub fn run() -> ExitCode {
    let cli = Cli::parse();
    let result: Result<()> = match cli.cmd {
        Cmd::Capture => capture::run(),
        Cmd::Music { track } => music::run(track.as_deref()),
        Cmd::Render { draft, workers, track } => render::run(draft, workers, track.as_deref(), cli.lang),
        Cmd::Audit => audit::run(cli.lang),
        Cmd::Deliver => deliver::run(cli.lang),
        Cmd::All { workers, track } => deliver::all(workers, track.as_deref(), cli.lang),
        Cmd::Still { times } => still::run(&times, cli.lang),
    };
    match result {
        Ok(()) => ExitCode::SUCCESS,
        Err(e) => {
            eprintln!("error: {e:#}");
            ExitCode::FAILURE
        }
    }
}
