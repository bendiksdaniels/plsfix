//! voice.rs: `plsfix-video voice [--lang en] [--voice <name>]`: the narrator over the music.
//! Reads the voice-over slots from the page (the scenes' own timing tables) and the words from
//! comp/voice.<lang>.json, speaks each line at the slowest rate that fits its slot, places it,
//! then mixes it over build/music.wav ducked under speech: build/voice[-en].wav (the narrator),
//! build/audio[-en].wav (the cut's soundtrack, what render muxes), build/voice[-en].json (what
//! each line did, read by the audit). `render` runs it after the music unless --no-voice.

use crate::cli::{build_dir, music, video_dir};
use crate::core::dsp::dynamics::{db_to_gain, gain};
use crate::core::dsp::SR;
use crate::core::lang::Lang;
use crate::core::sound::bus::Bus;
use crate::core::viewport::Viewport;
use crate::core::voice::{self as rules, Slot, DUCK, LEAD_S, RATES, VOICE_LUFS};
use crate::io::comp_page::CompPage;
use crate::io::{ffmpeg, loudness, static_server, tts, wav};
use anyhow::{bail, Context, Result};
use serde_json::{json, Map, Value};
use std::path::Path;

/// One placed line: its slot, where it sits (seconds) and the rate it was spoken at.
struct Placed {
    slot: Slot,
    start: f64,
    end: f64,
    rate: i32,
}

pub fn run(lang: Lang, speaker: Option<&str>) -> Result<()> {
    let speaker = speaker.unwrap_or(rules::default_voice(lang));
    let music_path = music::path();
    if !music_path.is_file() {
        bail!("voice: {} is missing; run music first", music_path.display());
    }
    let port = static_server::serve(video_dir())?;
    let (seconds, slots) = {
        let mut page = CompPage::open(port, Viewport::STAGE, lang)?;
        let seconds = page.meta.duration;
        (seconds, rules::parse_slots(&page.cdp.eval("window.__voice()")?, seconds).map_err(anyhow::Error::msg)?)
    };
    let words = words(lang, &slots)?;
    let n = (seconds * SR).round() as usize;
    let mut stem = Bus::new(n);
    let mut placed = Vec::with_capacity(slots.len());
    for s in slots {
        let (clip, rate) = spoken(speaker, &words[&s.key], &s)?;
        let start = s.from + LEAD_S;
        let i0 = (start * SR).round() as usize;
        for (k, x) in clip.iter().enumerate() {
            stem.add(i0 + k, *x, *x);
        }
        let end = start + clip.len() as f64 / SR;
        println!("voice: {} {start:.2}-{end:.2}s in {:.2}-{:.2}s ({rate:+}%)", s.key, s.from, s.to);
        placed.push(Placed { slot: s, start, end, rate });
    }
    // the narrator at its own level first, then the music ducked under every line
    let stem_path = build_dir().join(lang.voice_stem());
    wav::write(&stem_path, &stem, SR as u32)?;
    let own = loudness::measure(&stem_path)?;
    gain(&mut stem.l, &mut stem.r, db_to_gain(VOICE_LUFS - own.integrated));
    wav::write(&stem_path, &stem, SR as u32)?;
    let (ml, mr) = ffmpeg::decode(&music_path, seconds)?;
    let spans: Vec<(f64, f64)> = placed.iter().map(|p| (p.start, p.end)).collect();
    let g = rules::duck_gains(&spans, n, SR, &DUCK);
    let bed = |i: usize, ch: &[f32]| ch.get(i).copied().unwrap_or(0.0) * g[i];
    let mut mix = Bus::new(n);
    for i in 0..n {
        mix.l[i] = bed(i, &ml) + stem.l[i];
        mix.r[i] = bed(i, &mr) + stem.r[i];
    }
    // how far each line stands over the music under it (the final levelling scales both alike)
    let lines: Vec<Value> = placed.iter().map(|p| {
        let (i, j) = (((p.start * SR) as usize).min(n), ((p.end * SR) as usize).min(n));
        let music_db = rules::rms_db(&(i..j).map(|k| bed(k, &ml)).collect::<Vec<f32>>());
        let over = rules::rms_db(&stem.l[i..j]) - music_db;
        json!({ "key": p.slot.key, "text": words[&p.slot.key], "from": p.slot.from, "to": p.slot.to,
            "start": p.start, "end": p.end, "rate": p.rate, "over_music_db": (over * 10.0).round() / 10.0 })
    }).collect();
    let out = build_dir().join(lang.audio());
    let l = music::level(mix, &out, "voice")?;
    let report = json!({ "voice": speaker, "lufs": l.integrated, "true_peak": l.true_peak, "lines": lines });
    let report_path = build_dir().join(lang.voice_report());
    std::fs::write(&report_path, serde_json::to_string_pretty(&report)?).with_context(|| format!("voice: write {}", report_path.display()))?;
    println!("voice: {} lines by {speaker}; {:.1} LUFS, true peak {:.1} dBTP -> {}", placed.len(), l.integrated, l.true_peak, out.display());
    Ok(())
}

/// The cut's words, one per slot, and no line without a slot.
fn words(lang: Lang, slots: &[Slot]) -> Result<Map<String, Value>> {
    let path = video_dir().join("comp").join(lang.voice_file());
    let text = std::fs::read_to_string(&path).with_context(|| format!("voice: read {}", path.display()))?;
    let words: Map<String, Value> = serde_json::from_str(&text).with_context(|| format!("voice: parse {}", path.display()))?;
    for s in slots {
        if words.get(&s.key).and_then(Value::as_str).is_none_or(|t| t.trim().is_empty()) {
            bail!("voice: {} has no words for the slot {}", path.display(), s.key);
        }
    }
    if let Some(stray) = words.keys().find(|k| !slots.iter().any(|s| &s.key == *k)) {
        bail!("voice: {} has {stray}, which no scene gives a slot", path.display());
    }
    Ok(words)
}

/// The line spoken and trimmed, at the slowest rate that fits its slot.
fn spoken(speaker: &str, text: &Value, s: &Slot) -> Result<(Vec<f32>, i32)> {
    let say = rules::speakable(text.as_str().unwrap_or_default());
    let cache = build_dir().join("voice-cache");
    let first = clip(&tts::speak(speaker, 0, &say, &cache)?)?;
    let d0 = first.len() as f64 / SR;
    let Some(rate) = rules::fit_rate(d0, s.room()) else {
        bail!("voice: {} lasts {d0:.2}s at its own pace; its slot holds {:.2}s even at {:+}%: shorten the line", s.key, s.room(), RATES[RATES.len() - 1]);
    };
    if rate == 0 {
        return Ok((first, 0));
    }
    let faster = clip(&tts::speak(speaker, rate, &say, &cache)?)?;
    let d = faster.len() as f64 / SR;
    if d > s.room() {
        bail!("voice: {} lasts {d:.2}s at {rate:+}%; its slot holds {:.2}s: shorten the line", s.key, s.room());
    }
    Ok((faster, rate))
}

/// A spoken mp3 as mono samples, silence trimmed, with 8 ms fades so no edge clicks.
fn clip(mp3: &Path) -> Result<Vec<f32>> {
    let (l, _) = ffmpeg::decode(mp3, 60.0)?;
    let (a, b) = rules::trim(&l);
    let mut x = l[a..b].to_vec();
    let fade = ((0.008 * SR) as usize).min(x.len() / 2);
    let len = x.len();
    for k in 0..fade {
        let g = k as f32 / fade as f32;
        x[k] *= g;
        x[len - 1 - k] *= g;
    }
    Ok(x)
}
