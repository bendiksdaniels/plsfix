//! shots.rs tests: a capture record keeps its rules (the demo source, the host, the pane
//! version, a file per shot, rectangles that are real and inside the frame), and a PNG's size
//! comes out of its IHDR.
use plsfix_video::core::shots::{check_rect, png_size, validate};
use serde_json::{json, Value};

fn good() -> Value {
    json!({
        "source": "npm run demo:build", "host": "Office for the web", "pane_version": "v2.9.001",
        "workbook": { "sha256": "a".repeat(64) }, "deck": { "sha256": "b".repeat(64) },
        "shots": {
            "x-pnl": { "file": "x-pnl.png", "rects": { "block": [43, 250, 873, 360], "chips": [[1, 2, 3, 4], [5, 6, 7, 8]] } }
        }
    })
}

#[test]
fn a_good_record_passes() {
    assert!(validate(&good()).is_empty(), "{:?}", validate(&good()));
}

#[test]
fn provenance_is_held() {
    let mut r = good();
    r["source"] = json!("seed --showcase");
    r["host"] = json!("127.0.0.1");
    r["pane_version"] = json!("");
    r["deck"]["sha256"] = json!("short");
    let f = validate(&r).join(" | ");
    for want in ["want npm run demo:build", "want Office for the web", "no pane version", "sha256 for the deck"] {
        assert!(f.contains(want), "missing {want} in {f}");
    }
}

#[test]
fn a_bad_rect_names_its_shot_and_key() {
    let mut r = good();
    r["shots"]["x-pnl"]["rects"]["block"] = json!([43, 250, 0, 360]);
    r["shots"]["x-pnl"]["rects"]["chips"] = json!([[1400, 800, 90, 20]]);
    let f = validate(&r).join(" | ");
    assert!(f.contains("x-pnl: rect block is empty"), "{f}");
    assert!(f.contains("x-pnl: rect chips leaves the frame"), "{f}");
    assert!(check_rect("s", "k", [0.0, 0.0, 1440.0, 900.0]).is_ok());
}

#[test]
fn a_png_size_comes_from_ihdr() {
    let mut png = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR".to_vec();
    png.extend_from_slice(&2880u32.to_be_bytes());
    png.extend_from_slice(&1800u32.to_be_bytes());
    assert_eq!(png_size(&png), Some((2880, 1800)));
    assert_eq!(png_size(b"GIF89a not a png at all...."), None);
}
