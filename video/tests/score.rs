//! score.rs tests: sections snap to the nearest bar, every note sits on the 16th grid, the
//! intro has no drums and the outro plays its chord once.
use plsfix_video::core::sound::cues::{Cue, Kind};
use plsfix_video::core::sound::score::{notes, section_at, sections, Inst, Section};

fn section(t: f64, name: &str) -> Cue {
    Cue { t, kind: Kind::Section, dur: 0.5, pan: 0.0, pitch: 76.0, gain: 0.0, name: name.into() }
}

fn real() -> Vec<(u32, Section)> {
    sections(&[section(0.0, "intro"), section(5.3, "groove"), section(19.25, "lift"), section(40.95, "break"), section(48.05, "launch"), section(53.8, "outro")]).unwrap()
}

#[test]
fn sections_snap_to_the_nearest_bar() {
    let s = real();
    assert_eq!(s.iter().map(|(b, _)| *b).collect::<Vec<_>>(), vec![0, 3, 10, 20, 24, 27]);
    assert_eq!(section_at(&s, 2), Section::Intro);
    assert_eq!(section_at(&s, 3), Section::Groove);
    assert!(sections(&[section(5.3, "groove")]).unwrap_err().contains("start at 0"));
}

#[test]
fn every_note_is_on_the_grid_and_the_intro_has_no_drums() {
    let n = notes(&real(), 58.5);
    for x in &n {
        let steps = x.t / 0.125;
        assert!((steps - steps.round()).abs() < 1e-9, "{x:?} is off the 16th grid");
    }
    assert!(!n.iter().any(|x| matches!(x.inst, Inst::Kick | Inst::Hat | Inst::Clap) && x.t < 6.0));
    let outro_pads = n.iter().filter(|x| x.inst == Inst::Pad && x.t >= 54.0).count();
    assert_eq!(outro_pads, 4, "one four-note chord under the end card");
}
