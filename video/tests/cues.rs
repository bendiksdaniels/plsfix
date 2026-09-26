//! cues.rs tests: the page's cue list parses, sorts, and refuses what the soundtrack cannot play.
use plsfix_video::core::sound::cues::{parse, Kind};
use serde_json::json;

#[test]
fn cues_parse_and_sort() {
    let v = json!([{ "t": 9.7, "kind": "tap" }, { "t": 0, "kind": "section", "name": "intro" }, { "t": 5.3, "kind": "impact", "gain": -2 }]);
    let c = parse(&v, 58.5).unwrap();
    assert_eq!(c.iter().map(|x| x.kind).collect::<Vec<_>>(), vec![Kind::Section, Kind::Impact, Kind::Tap]);
    assert_eq!(c[1].gain, -2.0);
}

#[test]
fn bad_cues_are_refused_by_name() {
    assert!(parse(&json!([{ "t": 1, "kind": "gong" }]), 58.5).unwrap_err().contains("unknown kind gong"));
    assert!(parse(&json!([{ "t": 60, "kind": "tap" }]), 58.5).unwrap_err().contains("outside the video"));
    assert!(parse(&json!([{ "t": 1, "kind": "section" }]), 58.5).unwrap_err().contains("needs a name"));
}
