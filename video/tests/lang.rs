//! lang.rs tests: each cut's file names, and parsing the flag.
use plsfix_video::core::lang::Lang;

#[test]
fn the_latvian_cut_keeps_the_names_it_shipped_with() {
    assert_eq!(Lang::Lv.mp4(false), "plsfix-video.mp4");
    assert_eq!(Lang::Lv.mp4(true), "plsfix-video-draft.mp4");
    assert_eq!(Lang::Lv.layout(), "layout.json");
    assert_eq!(Lang::Lv.desktop_name(), "pls,fix video.mp4");
    assert_eq!(Lang::Lv.record(), "delivered.sha256");
    assert_eq!(Lang::Lv.copy_file(), "copy.lv.json");
    assert_eq!(Lang::Lv.still(9.9), "still-009.90.png");
}

#[test]
fn the_english_cut_never_overwrites_the_latvian_files() {
    assert_eq!(Lang::En.mp4(false), "plsfix-video-en.mp4");
    assert_eq!(Lang::En.mp4(true), "plsfix-video-en-draft.mp4");
    assert_eq!(Lang::En.layout(), "layout-en.json");
    assert_eq!(Lang::En.desktop_name(), "pls,fix video EN.mp4");
    assert_eq!(Lang::En.record(), "delivered.en.sha256");
    assert_eq!(Lang::En.copy_file(), "copy.en.json");
    assert_eq!(Lang::En.still(9.9), "still-en-009.90.png");
}

#[test]
fn the_flag_parses() {
    assert_eq!("en".parse::<Lang>(), Ok(Lang::En));
    assert_eq!("LV".parse::<Lang>(), Ok(Lang::Lv));
    assert!("de".parse::<Lang>().is_err());
}
