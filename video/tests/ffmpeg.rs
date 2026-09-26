//! ffmpeg.rs tests: the encoder always converts to BT.709 limited range 4:2:0 H.264 and tags
//! the file, so a player cannot guess the colours differently from the browser.
use plsfix_video::io::ffmpeg::encoder_args;
use std::path::Path;

#[test]
fn encoder_converts_and_tags_bt709() {
    let a = encoder_args(60, Path::new("out.mp4")).join(" ");
    for want in ["-c:v libx264", "-pix_fmt yuv420p", "-colorspace bt709", "-color_primaries bt709", "-color_trc bt709", "-color_range tv", "scale=out_color_matrix=bt709:out_range=tv", "setparams=color_primaries=bt709:color_trc=bt709", "-framerate 60", "-g 120"] {
        assert!(a.contains(want), "missing {want} in {a}");
    }
    assert!(a.ends_with("out.mp4"));
}
