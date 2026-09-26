//! score.rs: the music. 120 BPM in C major, one chord per bar (F, Am, C, G, then again),
//! sections taken from the picture's section cues and snapped to the nearest bar line, and the
//! notes each section plays. Invariant: every note starts on the 16th-note grid of 120 BPM, so a
//! cut the picture makes on a beat lands on a beat of the music.

use crate::core::sound::cues::{Cue, Kind};

pub const BEAT: f64 = 0.5;
pub const BAR: f64 = 2.0;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Section {
    Intro,
    Groove,
    Lift,
    Break,
    Launch,
    Outro,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Inst {
    Pad,
    Bass,
    Kick,
    Hat,
    Clap,
    Pluck,
}

#[derive(Debug, Clone, Copy, PartialEq)]
pub struct Note {
    pub inst: Inst,
    pub t: f64,
    pub len: f64,
    pub midi: f64,
    pub vel: f64,
    pub pan: f64,
    /// the pad's filter cutoff in Hz (how open the sound is)
    pub bright: f64,
}

// per chord (F, Am, C, G): the pad's voicing, the bass root, the melody's four tones
const PAD: [[f64; 4]; 4] = [[53.0, 57.0, 60.0, 64.0], [52.0, 57.0, 60.0, 64.0], [52.0, 55.0, 59.0, 62.0], [50.0, 55.0, 59.0, 64.0]];
const ROOT: [f64; 4] = [41.0, 45.0, 36.0, 43.0];
const ARP: [[f64; 4]; 4] = [[65.0, 69.0, 72.0, 76.0], [69.0, 72.0, 76.0, 79.0], [72.0, 76.0, 79.0, 83.0], [67.0, 71.0, 74.0, 81.0]];
const ARP_STEPS: [usize; 8] = [0, 1, 2, 3, 2, 1, 2, 3];
const CHORD_C: usize = 2;

fn section_named(name: &str) -> Option<Section> {
    Some(match name {
        "intro" => Section::Intro,
        "groove" => Section::Groove,
        "lift" => Section::Lift,
        "break" => Section::Break,
        "launch" => Section::Launch,
        "outro" => Section::Outro,
        _ => return None,
    })
}

/// Section starts in bars, from the section cues snapped to the nearest bar line.
pub fn sections(cues: &[Cue]) -> Result<Vec<(u32, Section)>, String> {
    let mut out: Vec<(u32, Section)> = Vec::new();
    for c in cues.iter().filter(|c| c.kind == Kind::Section) {
        let s = section_named(&c.name).ok_or_else(|| format!("section cue at {}s: unknown section {}", c.t, c.name))?;
        let bar = (c.t / BAR).round() as u32;
        out.retain(|(b, _)| *b != bar);
        out.push((bar, s));
    }
    out.sort_by_key(|(b, _)| *b);
    if out.first().map(|(b, _)| *b) != Some(0) {
        return Err("sections: the first section must start at 0 s".into());
    }
    Ok(out)
}

pub fn section_at(sections: &[(u32, Section)], bar: u32) -> Section {
    sections.iter().rev().find(|(b, _)| *b <= bar).map(|(_, s)| *s).unwrap_or(Section::Intro)
}

fn bright(s: Section) -> f64 {
    match s {
        Section::Intro => 1000.0,
        Section::Groove => 2000.0,
        Section::Lift => 2800.0,
        Section::Break => 2200.0,
        Section::Launch => 2500.0,
        Section::Outro => 2600.0,
    }
}

fn note(inst: Inst, t: f64, len: f64, midi: f64, vel: f64, pan: f64) -> Note {
    Note { inst, t, len, midi, vel, pan, bright: 0.0 }
}

/// Every note of the music for a video of `seconds`.
pub fn notes(sections: &[(u32, Section)], seconds: f64) -> Vec<Note> {
    let bars = (seconds / BAR).ceil() as u32;
    let mut out = Vec::new();
    for bar in 0..bars {
        let s = section_at(sections, bar);
        let t0 = bar as f64 * BAR;
        if s == Section::Outro {
            if bar == 0 || section_at(sections, bar - 1) != Section::Outro {
                outro(&mut out, t0, seconds - t0);
            }
            continue;
        }
        let c = (bar % 4) as usize;
        for (i, &m) in PAD[c].iter().enumerate() {
            let vel = if s == Section::Intro { 0.75 } else { 0.8 };
            out.push(Note { bright: bright(s), ..note(Inst::Pad, t0, BAR, m, vel, [-0.5, -0.15, 0.15, 0.5][i]) });
        }
        bass(&mut out, s, c, t0);
        drums(&mut out, s, t0);
        if !(s == Section::Intro && bar == 0) {
            melody(&mut out, s, c, t0);
        }
    }
    out
}

fn bass(out: &mut Vec<Note>, s: Section, c: usize, t0: f64) {
    match s {
        Section::Groove | Section::Lift | Section::Launch => {
            for k in 0..8 {
                out.push(note(Inst::Bass, t0 + k as f64 * 0.25, 0.2, ROOT[c], if k % 2 == 0 { 1.0 } else { 0.7 }, 0.0));
            }
        }
        Section::Break => out.push(note(Inst::Bass, t0, 1.9, ROOT[c], 0.8, 0.0)),
        Section::Intro | Section::Outro => {}
    }
}

fn drums(out: &mut Vec<Note>, s: Section, t0: f64) {
    let (kicks, kick_vel, hats, ghost, clap) = match s {
        Section::Groove => (4, 1.0, 0.75, false, false),
        Section::Lift => (4, 1.0, 0.8, true, true),
        Section::Launch => (4, 0.9, 0.7, false, false),
        Section::Break => (2, 0.8, 0.35, false, false),
        Section::Intro | Section::Outro => return,
    };
    for k in 0..kicks {
        out.push(note(Inst::Kick, t0 + k as f64 * (BAR / kicks as f64), 0.5, 0.0, kick_vel, 0.0));
    }
    for k in 0..4 {
        out.push(note(Inst::Hat, t0 + k as f64 * BEAT + 0.25, 0.1, 0.0, hats, if k % 2 == 0 { -0.2 } else { 0.2 }));
        if ghost {
            out.push(note(Inst::Hat, t0 + k as f64 * BEAT + 0.375, 0.1, 0.0, 0.3, 0.35));
        }
    }
    if clap {
        out.push(note(Inst::Clap, t0 + BEAT, 0.4, 0.0, 0.8, 0.0));
        out.push(note(Inst::Clap, t0 + 3.0 * BEAT, 0.4, 0.0, 0.8, 0.0));
    }
}

fn melody(out: &mut Vec<Note>, s: Section, c: usize, t0: f64) {
    let base = match s {
        Section::Intro => 0.35,
        Section::Groove | Section::Launch => 0.55,
        Section::Lift => 0.65,
        Section::Break => 0.6,
        Section::Outro => return,
    };
    for (k, &step) in ARP_STEPS.iter().enumerate() {
        let accent = if k % 2 == 0 { 0.1 } else { 0.0 };
        out.push(note(Inst::Pluck, t0 + k as f64 * 0.25, 0.25, ARP[c][step], base + accent, if k % 2 == 0 { -0.25 } else { 0.25 }));
    }
}

/// The end: one long C chord under the logo and a falling run of five notes.
fn outro(out: &mut Vec<Note>, t0: f64, left: f64) {
    let len = (left - 0.6).max(1.0);
    for (i, &m) in PAD[CHORD_C].iter().enumerate() {
        out.push(Note { bright: bright(Section::Outro), ..note(Inst::Pad, t0, len, m, 0.8, [-0.5, -0.15, 0.15, 0.5][i]) });
    }
    out.push(note(Inst::Bass, t0, len.min(3.0), ROOT[CHORD_C], 0.75, 0.0));
    for (k, &m) in [86.0, 83.0, 79.0, 76.0, 72.0].iter().enumerate() {
        out.push(note(Inst::Pluck, t0 + k as f64 * 0.125, 0.25, m, 0.5, if k % 2 == 0 { -0.3 } else { 0.3 }));
    }
}
