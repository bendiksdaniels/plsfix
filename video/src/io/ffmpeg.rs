//! ffmpeg.rs: the encoder a render worker pipes PNG frames into, the lossless join of the
//! parts, the soundtrack mux, decoding a picked track, and ffprobe. Owns the colour contract: RGB frames become BT.709 limited-range 4:2:0
//! and the file says so (untagged video shows a different green in QuickTime).

use anyhow::{bail, Context, Result};
use serde_json::Value;
use std::io::Write;
use std::path::{Path, PathBuf};
use std::process::{Child, ChildStdin, Command, Stdio};

/// The x264 arguments for one part: PNG frames on stdin at `fps`, BT.709 out, to `out`.
pub fn encoder_args(fps: u32, out: &Path) -> Vec<String> {
    let args = [
        "-v", "error", "-y", "-f", "image2pipe", "-c:v", "png", "-framerate", &fps.to_string(), "-i", "-",
        // setparams tags every frame too: x264 takes primaries and transfer from the frames
        "-vf", "scale=out_color_matrix=bt709:out_range=tv,format=yuv420p,setparams=color_primaries=bt709:color_trc=bt709:colorspace=bt709:range=tv",
        "-c:v", "libx264", "-preset", "slow", "-crf", "16", "-tune", "animation",
        "-profile:v", "high", "-level:v", "4.2", "-g", &(fps * 2).to_string(), "-pix_fmt", "yuv420p",
        "-colorspace", "bt709", "-color_primaries", "bt709", "-color_trc", "bt709", "-color_range", "tv",
    ];
    let mut v: Vec<String> = args.iter().map(|s| s.to_string()).collect();
    v.push(out.display().to_string());
    v
}

/// One running encoder; frames go in with `write`, `finish` waits and checks the exit.
pub struct Encoder {
    child: Child,
    stdin: Option<ChildStdin>,
    out: PathBuf,
}

impl Encoder {
    pub fn start(fps: u32, out: &Path) -> Result<Encoder> {
        let child = Command::new("ffmpeg")
            .args(encoder_args(fps, out))
            .stdin(Stdio::piped())
            .stdout(Stdio::null())
            .stderr(Stdio::piped())
            .spawn()
            .with_context(|| format!("ffmpeg encode {}: spawn", out.display()))?;
        let mut child = child;
        let stdin = child.stdin.take();
        Ok(Encoder { child, stdin, out: out.to_path_buf() })
    }

    pub fn write(&mut self, png: &[u8]) -> Result<()> {
        let stdin = self.stdin.as_mut().context("ffmpeg encode: stdin already closed")?;
        stdin.write_all(png).with_context(|| format!("ffmpeg encode {}: write frame", self.out.display()))
    }

    pub fn finish(mut self) -> Result<()> {
        drop(self.stdin.take());
        let out = self.child.wait_with_output().with_context(|| format!("ffmpeg encode {}: wait", self.out.display()))?;
        if !out.status.success() {
            bail!("ffmpeg encode {}: {}", self.out.display(), String::from_utf8_lossy(&out.stderr).trim());
        }
        Ok(())
    }
}

/// Joins same-settings parts into `out` without re-encoding, index moved to the front.
pub fn concat(parts: &[PathBuf], out: &Path) -> Result<()> {
    let list = out.with_extension("parts.txt");
    let body: String = parts.iter().map(|p| format!("file '{}'\n", p.display())).collect();
    std::fs::write(&list, body)?;
    let st = Command::new("ffmpeg")
        .args(["-v", "error", "-y", "-f", "concat", "-safe", "0", "-i"])
        .arg(&list)
        .args(["-c", "copy", "-movflags", "+faststart"])
        .arg(out)
        .status()
        .context("ffmpeg concat: spawn")?;
    if !st.success() {
        bail!("ffmpeg concat into {}: failed ({st})", out.display());
    }
    std::fs::remove_file(&list).ok();
    Ok(())
}

/// Puts the soundtrack under the picture: video copied, audio to AAC 192 kb/s, 48 kHz stereo.
pub fn mux(video: &Path, audio: &Path, out: &Path) -> Result<()> {
    let st = Command::new("ffmpeg")
        .args(["-v", "error", "-y", "-i"])
        .arg(video)
        .arg("-i")
        .arg(audio)
        .args(["-map", "0:v:0", "-map", "1:a:0", "-c:v", "copy", "-c:a", "aac", "-b:a", "192k", "-ar", "48000", "-ac", "2", "-shortest", "-movflags", "+faststart"])
        .arg(out)
        .status()
        .context("ffmpeg mux: spawn")?;
    if !st.success() {
        bail!("ffmpeg mux into {}: failed ({st})", out.display());
    }
    Ok(())
}

/// Any audio file ffmpeg can open, as 48 kHz stereo floats (left, right), at most `seconds` long.
pub fn decode(path: &Path, seconds: f64) -> Result<(Vec<f32>, Vec<f32>)> {
    let out = Command::new("ffmpeg")
        .args(["-v", "error", "-i"])
        .arg(path)
        .args(["-map", "0:a:0", "-t", &format!("{seconds:.3}"), "-ac", "2", "-ar", "48000", "-f", "f32le", "-"])
        .output()
        .context("ffmpeg decode: spawn")?;
    if !out.status.success() {
        bail!("ffmpeg decode {}: {}", path.display(), String::from_utf8_lossy(&out.stderr).trim());
    }
    let (mut l, mut r) = (Vec::new(), Vec::new());
    for frame in out.stdout.chunks_exact(8) {
        l.push(f32::from_le_bytes([frame[0], frame[1], frame[2], frame[3]]));
        r.push(f32::from_le_bytes([frame[4], frame[5], frame[6], frame[7]]));
    }
    Ok((l, r))
}

/// ffprobe's JSON for the file's streams and format, with an exact frame count.
pub fn probe(path: &Path) -> Result<Value> {
    let out = Command::new("ffprobe")
        .args(["-v", "error", "-count_frames", "-print_format", "json", "-show_streams", "-show_format"])
        .arg(path)
        .output()
        .context("ffprobe: spawn")?;
    if !out.status.success() {
        bail!("ffprobe {}: {}", path.display(), String::from_utf8_lossy(&out.stderr).trim());
    }
    serde_json::from_slice(&out.stdout).with_context(|| format!("ffprobe {}: not JSON", path.display()))
}
