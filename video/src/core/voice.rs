//! voice.rs: the voice-over's rules, pure. A line's slot comes from the page's timing tables;
//! the narrator speaks at the slowest rate that fits it; the music ducks under every line (and
//! stays down through short gaps, so it never pumps between phrases). Invariant: a line never
//! starts before its slot opens nor ends after it closes.
use crate::core::lang::Lang;
use serde_json::Value;

/// The voice starts this long after its slot opens, and ends at least TAIL_S before it closes.
pub const LEAD_S: f64 = 0.08;
pub const TAIL_S: f64 = 0.05;
/// edge-tts `--rate` steps tried, in percent faster than the voice's own pace.
pub const RATES: &[i32] = &[0, 5, 10, 15, 20];
/// A clip's ends quieter than this are trimmed (about -48 dBFS).
pub const SILENCE: f32 = 0.004;
/// The narrator's loudness before the final mix is levelled.
pub const VOICE_LUFS: f64 = -16.0;
/// A line must stand this far over the music under it (RMS over the line, dB).
pub const MIN_OVER_MUSIC_DB: f64 = 12.0;

/// The narrator per cut: Microsoft's neural voices (male, warm; `--voice` names another).
pub fn default_voice(lang: Lang) -> &'static str {
    match lang {
        Lang::Lv => "lv-LV-NilsNeural",
        Lang::En => "en-US-AndrewMultilingualNeural",
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Slot {
    pub key: String,
    pub from: f64,
    pub to: f64,
}

impl Slot {
    /// Seconds the spoken line may last.
    pub fn room(&self) -> f64 {
        self.to - self.from - LEAD_S - TAIL_S
    }
}

/// The page's `__voice()` list: sorted, inside the video, never overlapping, each with room.
pub fn parse_slots(v: &Value, seconds: f64) -> Result<Vec<Slot>, String> {
    let list = v.as_array().ok_or("voice slots: not a list")?;
    let mut out: Vec<Slot> = Vec::with_capacity(list.len());
    for s in list {
        let key = s["key"].as_str().ok_or("voice slot: no key")?.to_string();
        let (from, to) = (s["from"].as_f64().unwrap_or(-1.0), s["to"].as_f64().unwrap_or(-1.0));
        if from < 0.0 || to > seconds || to <= from {
            return Err(format!("voice slot {key}: {from}-{to}s is not inside the video's {seconds}s"));
        }
        if let Some(prev) = out.last() {
            if from < prev.to {
                return Err(format!("voice slot {key} opens at {from}s, before {} closes at {}s", prev.key, prev.to));
            }
        }
        let slot = Slot { key, from, to };
        if slot.room() < 0.5 {
            return Err(format!("voice slot {}: only {:.2}s to speak in", slot.key, slot.room()));
        }
        out.push(slot);
    }
    Ok(out)
}

/// The slowest rate in RATES at which a line `d0` seconds long at +0 % fits `room`.
pub fn fit_rate(d0: f64, room: f64) -> Option<i32> {
    RATES.iter().copied().find(|r| d0 / (1.0 + f64::from(*r) / 100.0) <= room)
}

/// The spoken part of a clip: indices of its first and one past its last non-silent sample.
pub fn trim(x: &[f32]) -> (usize, usize) {
    let first = x.iter().position(|v| v.abs() > SILENCE).unwrap_or(0);
    let last = x.iter().rposition(|v| v.abs() > SILENCE).map_or(first, |i| i + 1);
    (first, last.max(first))
}

/// What a line hands to the speech service: the words only (no quotes, stars or brackets,
/// which a voice would read out or stumble on), single-spaced.
pub fn speakable(text: &str) -> String {
    let kept: String = text.chars().filter(|c| !matches!(c, '*' | '«' | '»' | '"' | '“' | '”' | '[' | ']' | '#')).collect();
    kept.split_whitespace().collect::<Vec<_>>().join(" ")
}

/// A stable name for a spoken line in the cache (FNV-1a over voice, rate and text).
pub fn cache_key(voice: &str, rate: i32, text: &str) -> String {
    let mut h: u64 = 0xcbf2_9ce4_8422_2325;
    for b in format!("{voice}|{rate}|{text}").bytes() {
        h ^= u64::from(b);
        h = h.wrapping_mul(0x0100_0000_01b3);
    }
    format!("{h:016x}")
}

/// How the music sits under the voice: `base_db` between lines, `duck_db` under them, ramps of
/// `attack_s` before a line and `release_s` after it, gaps under `bridge_s` held down.
pub struct Duck {
    pub base_db: f64,
    pub duck_db: f64,
    pub attack_s: f64,
    pub release_s: f64,
    pub bridge_s: f64,
}

pub const DUCK: Duck = Duck { base_db: -6.0, duck_db: -17.0, attack_s: 0.3, release_s: 0.7, bridge_s: 1.6 };

/// The music's gain per sample for spoken spans (seconds, sorted).
pub fn duck_gains(spans: &[(f64, f64)], n: usize, sr: f64, d: &Duck) -> Vec<f32> {
    let mut merged: Vec<(f64, f64)> = Vec::new();
    for &(a, b) in spans {
        match merged.last_mut() {
            Some(last) if a - last.1 < d.bridge_s => last.1 = last.1.max(b),
            _ => merged.push((a, b)),
        }
    }
    let mut depth = vec![0.0f64; n];
    for &(a, b) in &merged {
        let lo = (((a - d.attack_s) * sr).floor().max(0.0) as usize).min(n);
        let hi = (((b + d.release_s) * sr).ceil() as usize).min(n);
        for (i, v) in depth.iter_mut().enumerate().take(hi).skip(lo) {
            let t = i as f64 / sr;
            let amount = if t < a { 1.0 - (a - t) / d.attack_s } else if t > b { 1.0 - (t - b) / d.release_s } else { 1.0 };
            *v = v.max(amount.clamp(0.0, 1.0));
        }
    }
    depth.iter().map(|x| (10f64.powf((d.base_db + (d.duck_db - d.base_db) * x) / 20.0)) as f32).collect()
}

/// Root-mean-square level of a stretch of samples, in dB (very quiet reads as -120).
pub fn rms_db(x: &[f32]) -> f64 {
    if x.is_empty() {
        return -120.0;
    }
    let ms = x.iter().map(|v| f64::from(*v) * f64::from(*v)).sum::<f64>() / x.len() as f64;
    if ms <= 1e-12 { -120.0 } else { 10.0 * ms.log10() }
}
