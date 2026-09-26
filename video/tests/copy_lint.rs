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

fn copy_of(lang: &str) -> serde_json::Map<String, serde_json::Value> {
    let path = format!("{}/comp/copy.{lang}.json", env!("CARGO_MANIFEST_DIR"));
    let text = std::fs::read_to_string(&path).unwrap();
    serde_json::from_str::<serde_json::Value>(&text).unwrap().as_object().unwrap().clone()
}

#[test]
fn both_real_copy_files_pass() {
    for lang in ["lv", "en"] {
        for (key, v) in copy_of(lang) {
            assert!(lint(&key, v.as_str().unwrap()).is_empty(), "{lang} {key}");
        }
    }
}

#[test]
fn the_cuts_carry_the_same_keys() {
    let (lv, en) = (copy_of("lv"), copy_of("en"));
    let missing: Vec<_> = lv.keys().filter(|k| !en.contains_key(*k)).chain(en.keys().filter(|k| !lv.contains_key(*k))).collect();
    assert!(missing.is_empty(), "keys in one cut only: {missing:?}");
}
