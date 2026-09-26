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

// The calm bed (the default under a narrator): the video's own sections.
fn plsfix() -> Vec<(u32, Section)> {
    sections(&[section(0.0, "intro"), section(6.4, "groove"), section(26.0, "lift"), section(60.0, "break"), section(64.0, "launch"), section(80.0, "outro")]).unwrap()
}

#[test]
fn the_bed_is_on_the_grid_and_plays_none_of_the_busy_voices() {
    use plsfix_video::core::sound::score::{arrange, Style};
    let bed = arrange(Style::Bed, &plsfix(), 87.5);
    for n in &bed {
        let steps = n.t / 0.125;
        assert!((steps - steps.round()).abs() < 1e-9, "{:?} at {}s is off the grid", n.inst, n.t);
        assert!(!matches!(n.inst, Inst::Hat | Inst::Clap | Inst::Pluck), "the bed plays no {:?}", n.inst);
    }
    let onsets = bed.iter().filter(|n| matches!(n.inst, Inst::Keys | Inst::Kick | Inst::Bass)).count() as f64;
    assert!(onsets / 87.5 < 2.0, "the bed stays sparse: {onsets} onsets in 87,5 s");
    let groove = arrange(Style::Groove, &plsfix(), 87.5);
    assert!(groove.len() > 4 * bed.len(), "the groove score is the busy one");
}

#[test]
fn the_bed_changes_chord_every_two_bars() {
    use plsfix_video::core::sound::score::{arrange, Style};
    let pad_starts: Vec<f64> = arrange(Style::Bed, &plsfix(), 87.5).iter().filter(|n| n.inst == Inst::Pad).map(|n| n.t).collect();
    let mut starts = pad_starts.clone();
    starts.dedup();
    let long = starts.windows(2).filter(|w| (w[1] - w[0] - 4.0).abs() < 1e-9).count();
    assert!(long * 10 >= starts.len() * 8, "most chords last 4 s: {starts:?}");
}
