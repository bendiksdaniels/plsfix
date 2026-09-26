//! probe_check.rs: what the finished file must be, checked against ffprobe's JSON: H.264,
//! 1920 x 1080, 60 fps, yuv420p, BT.709 tagged, limited range, 80 to 90 seconds, exactly the
//! timeline's frame count, and one AAC stereo 48 kHz soundtrack as long as the picture.

use serde_json::Value;

pub const WIDTH: u64 = 1920;
pub const HEIGHT: u64 = 1080;
pub const FPS: &str = "60/1";
/// Daniel's ask for this video (26.09.2026): all the features, under 90 s.
pub const MIN_SECONDS: f64 = 80.0;
pub const MAX_SECONDS: f64 = 90.0;

/// Findings for a probe of the final file whose timeline has `frames` frames.
pub fn check(probe: &Value, frames: u64) -> Vec<String> {
    let mut out = Vec::new();
    let streams = probe["streams"].as_array().cloned().unwrap_or_default();
    let video: Vec<&Value> = streams.iter().filter(|s| s["codec_type"] == "video").collect();
    if video.len() != 1 {
        out.push(format!("streams: {} video streams, want 1", video.len()));
        return out;
    }
    let v = video[0];
    let want = |field: &str, value: &str, out: &mut Vec<String>| {
        if v[field].as_str() != Some(value) {
            out.push(format!("{field}: {}, want {value}", v[field]));
        }
    };
    want("codec_name", "h264", &mut out);
    want("pix_fmt", "yuv420p", &mut out);
    want("r_frame_rate", FPS, &mut out);
    want("color_space", "bt709", &mut out);
    want("color_primaries", "bt709", &mut out);
    want("color_transfer", "bt709", &mut out);
    want("color_range", "tv", &mut out);
    if v["width"].as_u64() != Some(WIDTH) || v["height"].as_u64() != Some(HEIGHT) {
        out.push(format!("size: {}x{}, want {WIDTH}x{HEIGHT}", v["width"], v["height"]));
    }
    let seconds: f64 = probe["format"]["duration"].as_str().and_then(|d| d.parse().ok()).unwrap_or(0.0);
    if !(MIN_SECONDS..=MAX_SECONDS).contains(&seconds) {
        out.push(format!("duration: {seconds:.2}s, want {MIN_SECONDS}-{MAX_SECONDS}s"));
    }
    let counted: u64 = v["nb_read_frames"].as_str().and_then(|n| n.parse().ok()).unwrap_or(0);
    if counted != frames {
        out.push(format!("frames: {counted} in the file, the timeline has {frames}"));
    }
    out.extend(audio(&streams, seconds));
    out
}

/// The soundtrack: exactly one AAC stereo 48 kHz stream, within 0,1 s of the picture's length.
fn audio(streams: &[Value], seconds: f64) -> Vec<String> {
    let a: Vec<&Value> = streams.iter().filter(|s| s["codec_type"] == "audio").collect();
    if a.len() != 1 {
        return vec![format!("audio: {} audio streams, want 1 (the soundtrack)", a.len())];
    }
    let mut out = Vec::new();
    if a[0]["codec_name"] != "aac" {
        out.push(format!("audio codec: {}, want aac", a[0]["codec_name"]));
    }
    if a[0]["sample_rate"] != "48000" || a[0]["channels"] != 2 {
        out.push(format!("audio: {} Hz x {} channels, want 48000 x 2", a[0]["sample_rate"], a[0]["channels"]));
    }
    let len: f64 = a[0]["duration"].as_str().and_then(|d| d.parse().ok()).unwrap_or(0.0);
    if (len - seconds).abs() > 0.1 {
        out.push(format!("audio: {len:.2}s long, the picture is {seconds:.2}s"));
    }
    out
}
