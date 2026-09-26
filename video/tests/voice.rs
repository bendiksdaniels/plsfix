//! voice.rs tests: slots, the pace that fits, trimming, the words a voice is handed, and the
//! music's ducking under speech; plus both narration files kept to the same keys.
use plsfix_video::core::voice::*;
use serde_json::json;

#[test]
fn slots_must_be_inside_sorted_and_roomy() {
    let ok = json!([{ "key": "a", "from": 1.0, "to": 3.0 }, { "key": "b", "from": 3.0, "to": 5.0 }]);
    assert_eq!(parse_slots(&ok, 10.0).unwrap().len(), 2);
    let overlap = json!([{ "key": "a", "from": 1.0, "to": 3.5 }, { "key": "b", "from": 3.0, "to": 5.0 }]);
    assert!(parse_slots(&overlap, 10.0).unwrap_err().contains("before a closes"));
    let outside = json!([{ "key": "a", "from": 9.0, "to": 11.0 }]);
    assert!(parse_slots(&outside, 10.0).is_err());
    let tiny = json!([{ "key": "a", "from": 1.0, "to": 1.4 }]);
    assert!(parse_slots(&tiny, 10.0).unwrap_err().contains("to speak in"));
}

#[test]
fn the_slowest_rate_that_fits() {
    assert_eq!(fit_rate(2.0, 2.5), Some(0));
    assert_eq!(fit_rate(2.2, 2.0), Some(10));
    assert_eq!(fit_rate(4.0, 2.0), None);
}

#[test]
fn trimming_keeps_the_spoken_part() {
    let x = [0.0, 0.001, 0.2, -0.3, 0.1, 0.002, 0.0];
    assert_eq!(trim(&x), (2, 5));
    assert_eq!(trim(&[0.0; 4]), (0, 0));
}

#[test]
fn a_voice_gets_words_only() {
    assert_eq!(speakable("  Atkal «plīz  fiks»? *x* "), "Atkal plīz fiks? x");
    assert_ne!(cache_key("v", 0, "a"), cache_key("v", 5, "a"));
}

#[test]
fn music_ducks_under_speech_and_holds_through_short_gaps() {
    let sr = 1000.0;
    let g = duck_gains(&[(1.0, 2.0), (2.5, 3.0), (6.0, 7.0)], 9000, sr, &DUCK);
    let db = |i: usize| 20.0 * f64::from(g[i]).log10();
    assert!((db(100) - DUCK.base_db).abs() < 0.01, "before any speech: the base level");
    assert!((db(1500) - DUCK.duck_db).abs() < 0.01, "under a line: ducked");
    assert!((db(2250) - DUCK.duck_db).abs() < 0.01, "a 0.5 s gap stays down");
    assert!((db(4500) - DUCK.base_db).abs() < 0.01, "a long gap comes back up");
    assert!(db(5850) < DUCK.base_db && db(5850) > DUCK.duck_db, "the attack ramps before a line");
}

#[test]
fn both_narrations_carry_the_same_lines_and_pass_the_lint() {
    let read = |lang: &str| -> serde_json::Map<String, serde_json::Value> {
        let path = format!("{}/comp/voice.{lang}.json", env!("CARGO_MANIFEST_DIR"));
        serde_json::from_str(&std::fs::read_to_string(path).unwrap()).unwrap()
    };
    let (lv, en) = (read("lv"), read("en"));
    assert_eq!(lv.keys().collect::<Vec<_>>(), en.keys().collect::<Vec<_>>());
    for (key, v) in lv.iter().chain(en.iter()) {
        assert!(plsfix_video::core::copy_lint::lint(key, v.as_str().unwrap()).is_empty(), "{key}");
    }
}
