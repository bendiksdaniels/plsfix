//! lang.rs: the video's two cuts, Latvian (the default) and English. A cut is only its copy file:
//! picture, timing and sound are the same, so each cut's files carry the language in their names
//! (the Latvian cut keeps the plain names it shipped with).
use std::str::FromStr;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Lang {
    Lv,
    En,
}

impl Lang {
    pub fn code(self) -> &'static str {
        match self {
            Lang::Lv => "lv",
            Lang::En => "en",
        }
    }

    /// comp/<file>: the captions the page loads (its URL carries `?lang=<code>`).
    pub fn copy_file(self) -> String {
        format!("copy.{}.json", self.code())
    }

    /// build/<file>: the full render, or the draft.
    pub fn mp4(self, draft: bool) -> String {
        format!("plsfix-video{}{}.mp4", self.tag(), if draft { "-draft" } else { "" })
    }

    /// build/<file>: the page's layout report, written by render and read by the audit.
    pub fn layout(self) -> String {
        format!("layout{}.json", self.tag())
    }

    /// build/stills/<file>: a QA frame at `t` seconds.
    pub fn still(self, t: f64) -> String {
        format!("still{}-{t:06.2}.png", self.tag())
    }

    /// The delivered copy on the Desktop.
    pub fn desktop_name(self) -> &'static str {
        match self {
            Lang::Lv => "pls,fix video.mp4",
            Lang::En => "pls,fix video EN.mp4",
        }
    }

    /// video/<file>: the last delivery's sha256, committed with the work.
    pub fn record(self) -> &'static str {
        match self {
            Lang::Lv => "delivered.sha256",
            Lang::En => "delivered.en.sha256",
        }
    }

    fn tag(self) -> &'static str {
        match self {
            Lang::Lv => "",
            Lang::En => "-en",
        }
    }
}

impl FromStr for Lang {
    type Err = String;

    fn from_str(s: &str) -> Result<Lang, String> {
        match s.to_ascii_lowercase().as_str() {
            "lv" => Ok(Lang::Lv),
            "en" => Ok(Lang::En),
            other => Err(format!("no cut in {other:?}: lv or en")),
        }
    }
}
