//! probe_check.rs tests: a good probe passes; each broken property is named; the soundtrack
//! must be there, AAC stereo 48 kHz, as long as the picture.
use plsfix_video::core::probe_check::check;
use serde_json::{json, Value};

fn good() -> Value {
    json!({
        "streams": [
            { "codec_type": "video", "codec_name": "h264", "pix_fmt": "yuv420p", "r_frame_rate": "60/1",
              "color_space": "bt709", "color_primaries": "bt709", "color_transfer": "bt709", "color_range": "tv",
              "width": 1920, "height": 1080, "nb_read_frames": "5160" },
            { "codec_type": "audio", "codec_name": "aac", "sample_rate": "48000", "channels": 2, "duration": "86.000000" }
        ],
        "format": { "duration": "86.000000" }
    })
}

#[test]
fn a_good_file_passes() {
    assert!(check(&good(), 5160).is_empty(), "{:?}", check(&good(), 5160));
}

#[test]
fn each_broken_property_is_named() {
    let mut p = good();
    p["streams"][0]["width"] = json!(1080);
    p["streams"][0]["height"] = json!(1920);
    p["streams"][0]["r_frame_rate"] = json!("30/1");
    p["streams"][0]["pix_fmt"] = json!("yuv444p");
    p["streams"][0]["color_primaries"] = json!("unknown");
    p["format"]["duration"] = json!("91.5");
    let f = check(&p, 5160).join(" | ");
    for want in ["size", "r_frame_rate", "pix_fmt", "color_primaries", "duration"] {
        assert!(f.contains(want), "missing {want} in {f}");
    }
}

#[test]
fn a_frame_count_off_by_one_is_a_finding() {
    assert!(check(&good(), 5161).join(" | ").contains("frames: 5160"));
}

#[test]
fn the_soundtrack_must_be_there_and_right() {
    let mut silent = good();
    silent["streams"].as_array_mut().unwrap().pop();
    assert!(check(&silent, 5160).join(" | ").contains("0 audio streams"));
    let mut bad = good();
    bad["streams"][1]["codec_name"] = json!("mp3");
    bad["streams"][1]["channels"] = json!(1);
    bad["streams"][1]["duration"] = json!("85.0");
    let f = check(&bad, 5160).join(" | ");
    for want in ["audio codec", "48000 x 2", "85.00s long"] {
        assert!(f.contains(want), "missing {want} in {f}");
    }
}
