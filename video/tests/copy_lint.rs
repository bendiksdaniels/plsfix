//! copy_lint.rs tests: each rule fires on its own case, and the real copy file passes.
use plsfix_video::core::copy_lint::lint;

#[test]
fn every_rule_fires() {
    assert!(lint("k", "Viss \u{2014} vienuviet").iter().any(|f| f.contains("em dash")));
    assert!(lint("k", "Atver cilni").iter().any(|f| f.contains("ciln")));
    assert!(lint("k", "Mūsu lietotne").iter().any(|f| f.contains("aplikācija")));
    assert!(lint("k", "Pēc aplēses").iter().any(|f| f.contains("novērtējums")));
    assert!(lint("k", "Kohorta").iter().any(|f| f.contains("banned")));
    assert!(lint("k", "Divas  atstarpes").iter().any(|f| f.contains("double space")));
    assert!(lint("k", " ").iter().any(|f| f.contains("empty")));
}

#[test]
fn plain_latvian_and_ranges_pass() {
    assert!(lint("k", "Pārvelc kartīti, statuss mainās uzreiz").is_empty());
    assert!(lint("k", "Safari → Share → Add to Home Screen").is_empty());
    assert!(lint("k", "2026–2027").is_empty(), "an en dash in a range is fine");
}

#[test]
fn the_real_copy_file_passes() {
    let text = std::fs::read_to_string(concat!(env!("CARGO_MANIFEST_DIR"), "/comp/copy.lv.json")).unwrap();
    let copy: serde_json::Value = serde_json::from_str(&text).unwrap();
    for (key, v) in copy.as_object().unwrap() {
        assert!(lint(key, v.as_str().unwrap()).is_empty(), "{key}");
    }
}
