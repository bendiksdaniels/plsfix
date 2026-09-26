//! shots.rs: the rules the capture record (`capture.json`, then `shots.json`) must keep, as
//! pure checks on its JSON: it names the demo build, the host and the pane version; every shot
//! has a file and only non-empty rectangles inside the frame. `png_size` reads a PNG's IHDR.
//! Invariant: a rectangle is capture CSS px of the cropped 1440 x 900 frame.

use serde_json::Value;

/// A capture: 1440 x 900 CSS px (the Office header cropped) at 2x.
pub const CSS_W: f64 = 1440.0;
pub const CSS_H: f64 = 900.0;
pub const DPR: f64 = 2.0;
/// The one demo source and host the video may show.
pub const SOURCE: &str = "npm run demo:build";
pub const HOST: &str = "Office for the web";

/// A rectangle that is empty or leaves the frame would put an animation on nothing.
pub fn check_rect(shot: &str, key: &str, r: [f64; 4]) -> Result<(), String> {
    if r[2] < 1.0 || r[3] < 1.0 {
        return Err(format!("capture {shot}: rect {key} is empty ({} x {})", r[2], r[3]));
    }
    if r[0] < -1.0 || r[1] < -1.0 || r[0] + r[2] > CSS_W + 1.0 || r[1] + r[3] > CSS_H + 1.0 {
        return Err(format!("capture {shot}: rect {key} leaves the frame ({r:?})"));
    }
    Ok(())
}

/// Every finding on a capture record; empty means it may be used.
pub fn validate(record: &Value) -> Vec<String> {
    let mut f = Vec::new();
    if record["source"] != SOURCE {
        f.push(format!("provenance: shots came from {}, want {SOURCE}", record["source"]));
    }
    if record["host"] != HOST {
        f.push(format!("provenance: shots came from host {}, want {HOST}", record["host"]));
    }
    if record["pane_version"].as_str().is_none_or(|v| !v.starts_with('v')) {
        f.push(format!("provenance: no pane version recorded ({})", record["pane_version"]));
    }
    for doc in ["workbook", "deck"] {
        if record[doc]["sha256"].as_str().is_none_or(|s| s.len() != 64) {
            f.push(format!("provenance: no sha256 for the {doc}"));
        }
    }
    let Some(shots) = record["shots"].as_object().filter(|s| !s.is_empty()) else {
        f.push("capture: no shots".into());
        return f;
    };
    for (name, shot) in shots {
        if shot["file"].as_str() != Some(&format!("{name}.png")) {
            f.push(format!("capture {name}: file is {}, want {name}.png", shot["file"]));
        }
        for (key, v) in shot["rects"].as_object().into_iter().flatten() {
            let list: Vec<&Value> = if v.get(0).is_some_and(Value::is_array) { v.as_array().into_iter().flatten().collect() } else { vec![v] };
            if list.is_empty() {
                f.push(format!("capture {name}: rect list {key} is empty"));
            }
            for r in list {
                match rect4(r) {
                    Some(r) => f.extend(check_rect(name, key, r).err()),
                    None => f.push(format!("capture {name}: rect {key} is not [x, y, w, h]: {r}")),
                }
            }
        }
    }
    f
}

fn rect4(v: &Value) -> Option<[f64; 4]> {
    let a = v.as_array().filter(|a| a.len() == 4)?;
    Some([a[0].as_f64()?, a[1].as_f64()?, a[2].as_f64()?, a[3].as_f64()?])
}

/// A PNG's width and height from its IHDR chunk, or None when the bytes are not a PNG.
pub fn png_size(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 24 || &bytes[..8] != b"\x89PNG\r\n\x1a\n" || &bytes[12..16] != b"IHDR" {
        return None;
    }
    let be = |i: usize| u32::from_be_bytes([bytes[i], bytes[i + 1], bytes[i + 2], bytes[i + 3]]);
    Some((be(16), be(20)))
}
