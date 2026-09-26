//! bed.rs: the calm bed, the music under a narrator. The same four chords as the groove score
//! (F, Am, C, G) but one every two bars, a dark pad and a sustained bass under a few soft keys
//! notes per chord; no hats, no claps, no arpeggio, and a quiet kick on beat one only where the
//! picture lifts. Invariant: every note starts on the 16th-note grid, like the groove score.
use crate::core::sound::score::{section_at, Inst, Note, Section, BAR, BEAT};

/// Bars per chord: the harmony moves every four seconds, not every two.
const BARS_PER_CHORD: u32 = 2;
const CHORD_C: usize = 2;
// per chord (F, Am, C, G): the pad's voicing, the bass root, the keys' three tones (mid register)
const PAD: [[f64; 4]; 4] = [[53.0, 57.0, 60.0, 64.0], [52.0, 57.0, 60.0, 64.0], [52.0, 55.0, 60.0, 64.0], [50.0, 55.0, 59.0, 62.0]];
const ROOT: [f64; 4] = [41.0, 45.0, 36.0, 43.0];
const KEYS: [[f64; 3]; 4] = [[69.0, 72.0, 76.0], [72.0, 76.0, 79.0], [67.0, 72.0, 76.0], [71.0, 74.0, 79.0]];
const SPREAD: [f64; 4] = [-0.5, -0.15, 0.15, 0.5];

/// How open the pad sounds per section (its filter in Hz): darker than the groove score's.
fn bright(s: Section) -> f64 {
    match s {
        Section::Intro => 800.0,
        Section::Groove => 1100.0,
        Section::Lift => 1300.0,
        Section::Break => 1000.0,
        Section::Launch => 1200.0,
        Section::Outro => 1000.0,
    }
}

fn note(inst: Inst, t: f64, len: f64, midi: f64, vel: f64, pan: f64) -> Note {
    Note { inst, t, len, midi, vel, pan, bright: 0.0 }
}

/// Every note of the bed for a video of `seconds`.
pub fn notes(sections: &[(u32, Section)], seconds: f64) -> Vec<Note> {
    let bars = (seconds / BAR).ceil() as u32;
    let mut out = Vec::new();
    let mut bar = 0;
    while bar < bars {
        let s = section_at(sections, bar);
        let t0 = f64::from(bar) * BAR;
        if s == Section::Outro {
            outro(&mut out, t0, seconds - t0);
            break;
        }
        // a chord lasts two bars, or up to the next section, whichever comes first
        let span = (1..BARS_PER_CHORD).take_while(|k| section_at(sections, bar + k) == s).count() as u32 + 1;
        let len = f64::from(span) * BAR;
        let c = ((bar / BARS_PER_CHORD) % 4) as usize;
        for (i, &m) in PAD[c].iter().enumerate() {
            let vel = if s == Section::Intro { 0.6 } else { 0.7 };
            out.push(Note { bright: bright(s), ..note(Inst::Pad, t0, len, m, vel, SPREAD[i]) });
        }
        if s != Section::Intro {
            out.push(note(Inst::Bass, t0, len - 0.2, ROOT[c], 0.5, 0.0));
        }
        if matches!(s, Section::Lift | Section::Launch) {
            for k in 0..span {
                out.push(note(Inst::Kick, t0 + f64::from(k) * BAR, 0.5, 0.0, 0.35, 0.0));
            }
        }
        keys(&mut out, s, c, t0, len, bar == 0);
        bar += span;
    }
    out
}

/// A few soft keys notes per chord: the third on the downbeat, then the fifth and the top.
fn keys(out: &mut Vec<Note>, s: Section, c: usize, t0: f64, len: f64, first: bool) {
    let steps: &[(f64, usize, f64)] = match s {
        Section::Intro if first => &[],
        Section::Intro | Section::Break => &[(0.0, 0, 0.3)],
        Section::Lift => &[(0.0, 0, 0.42), (3.0 * BEAT, 1, 0.34), (6.0 * BEAT, 2, 0.32), (7.0 * BEAT, 1, 0.26)],
        _ => &[(0.0, 0, 0.4), (3.0 * BEAT, 1, 0.32), (6.0 * BEAT, 2, 0.3)],
    };
    for &(dt, tone, vel) in steps {
        if dt < len {
            out.push(note(Inst::Keys, t0 + dt, 2.0, KEYS[c][tone], vel, if tone == 1 { -0.25 } else { 0.25 }));
        }
    }
}

/// The end: one long C chord under the logo and three keys notes falling to rest.
fn outro(out: &mut Vec<Note>, t0: f64, left: f64) {
    let len = (left - 0.6).max(1.0);
    for (i, &m) in PAD[CHORD_C].iter().enumerate() {
        out.push(Note { bright: bright(Section::Outro), ..note(Inst::Pad, t0, len, m, 0.7, SPREAD[i]) });
    }
    out.push(note(Inst::Bass, t0, len.min(4.0), ROOT[CHORD_C], 0.45, 0.0));
    for (k, &m) in [79.0, 76.0, 72.0].iter().enumerate() {
        out.push(note(Inst::Keys, t0 + 0.5 + f64::from(k as u32) * 0.75, 2.5, m, 0.34, if k % 2 == 0 { -0.2 } else { 0.2 }));
    }
}
