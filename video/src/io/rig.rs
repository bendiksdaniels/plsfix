//! rig.rs: the web rig the capture drives (scripts/rig/README.md): the scratch Chrome profile
//! signed in to Office, headless on CDP 9222, and the manifest server on 3001. `Rig::up` reuses
//! whatever already answers and starts the rest; Drop stops only what it started. The OneDrive
//! site comes from PLSFIX_RIG_SITE or the profile's own history, never from the repository.

use crate::io::chrome::CHROME;
use crate::io::http;
use anyhow::{bail, Context, Result};
use std::fs::File;
use std::net::TcpStream;
use std::path::{Path, PathBuf};
use std::process::{Child, Command};
use std::thread::sleep;
use std::time::{Duration, Instant};

pub const CDP_PORT: u16 = 9222;
pub const MANIFEST_PORT: u16 = 3001;
const DEV_CERT: &str = ".office-addin-dev-certs/localhost.crt";

fn home() -> Result<PathBuf> {
    Ok(PathBuf::from(std::env::var("HOME").context("rig: HOME unset")?))
}

/// The scratch profile the rig README sets up (signed in to Office by hand, once).
pub fn profile() -> Result<PathBuf> {
    Ok(home()?.join(".cache/plsfix-rig-chrome"))
}

pub struct Rig {
    started: Vec<(&'static str, Child)>,
}

impl Rig {
    /// Makes sure the manifest server and the scratch Chrome answer; logs go to `log_dir`.
    pub fn up(repo: &Path, log_dir: &Path) -> Result<Rig> {
        let mut rig = Rig { started: Vec::new() };
        std::fs::create_dir_all(log_dir)?;
        if !port_open(MANIFEST_PORT) {
            let log = File::create(log_dir.join("rig-manifest.log"))?;
            let child = Command::new("node")
                .arg(repo.join("scripts/rig/serve-manifest.mjs"))
                .current_dir(repo)
                .stdout(log.try_clone()?)
                .stderr(log)
                .spawn()
                .context("rig: start the manifest server")?;
            rig.started.push(("manifest server", child));
            wait_for("manifest server", || port_open(MANIFEST_PORT))?;
        }
        if !cdp_answers() {
            let log = File::create(log_dir.join("rig-chrome.log"))?;
            let agent = format!(
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/{}.0.0.0 Safari/537.36",
                chrome_major()?
            );
            let child = Command::new(CHROME)
                .args(["--headless=new", &format!("--remote-debugging-port={CDP_PORT}")])
                .arg(format!("--user-data-dir={}", profile()?.display()))
                .arg(format!("--ignore-certificate-errors-spki-list={}", spki()?))
                .args(["--no-first-run", "--no-default-browser-check", "--window-size=1440,948", "--force-color-profile=srgb"])
                .arg(format!("--user-agent={agent}"))
                .arg("about:blank")
                .stdout(log.try_clone()?)
                .stderr(log)
                .spawn()
                .context("rig: start the scratch Chrome")?;
            rig.started.push(("scratch Chrome", child));
            wait_for("scratch Chrome", cdp_answers)?;
        }
        let names: Vec<&str> = rig.started.iter().map(|(n, _)| *n).collect();
        println!("capture: rig up (started: {})", if names.is_empty() { "nothing".into() } else { names.join(", ") });
        Ok(rig)
    }
}

impl Drop for Rig {
    fn drop(&mut self) {
        for (name, child) in &mut self.started {
            child.kill().ok();
            child.wait().ok();
            println!("capture: stopped the {name} it started");
        }
    }
}

fn port_open(port: u16) -> bool {
    TcpStream::connect(("127.0.0.1", port)).is_ok()
}

fn cdp_answers() -> bool {
    matches!(http::get(CDP_PORT, "/json/version"), Ok((200, _)))
}

fn wait_for(what: &str, ok: impl Fn() -> bool) -> Result<()> {
    let t0 = Instant::now();
    while t0.elapsed() < Duration::from_secs(30) {
        if ok() {
            return Ok(());
        }
        sleep(Duration::from_millis(300));
    }
    bail!("rig: the {what} did not answer within 30 s")
}

/// Chrome's major version, for a user agent that does not say HeadlessChrome.
fn chrome_major() -> Result<String> {
    let out = Command::new(CHROME).arg("--version").output().context("rig: chrome --version")?;
    let text = String::from_utf8_lossy(&out.stdout);
    let major = text.split_whitespace().last().and_then(|v| v.split('.').next()).filter(|m| m.parse::<u32>().is_ok());
    major.map(str::to_string).with_context(|| format!("rig: no version in {text:?}"))
}

/// The dev certificate's SPKI hash: Chrome trusts the manifest server's certificate by it.
fn spki() -> Result<String> {
    let cert = home()?.join(DEV_CERT);
    let pipeline = format!(
        "openssl x509 -in '{}' -pubkey -noout | openssl pkey -pubin -outform der | openssl dgst -sha256 -binary | base64",
        cert.display()
    );
    let out = Command::new("sh").args(["-c", &pipeline]).output().context("rig: spki")?;
    let hash = String::from_utf8_lossy(&out.stdout).trim().to_string();
    if !out.status.success() || hash.len() != 44 {
        bail!("rig: no SPKI for {} (run `npm start` once for the dev certificates)", cert.display());
    }
    Ok(hash)
}

/// The OneDrive site the demo files go to: PLSFIX_RIG_SITE, or the newest one the scratch
/// profile visited (`https://<tenant>-my.sharepoint.com/personal/<user>`).
pub fn site() -> Result<String> {
    if let Ok(s) = std::env::var("PLSFIX_RIG_SITE") {
        return Ok(s);
    }
    let db = profile()?.join("Default/History");
    let query = "select url from urls where url like 'https://%-my.sharepoint.com/%personal/%' order by last_visit_time desc limit 20";
    let out = Command::new("sqlite3")
        .arg(format!("file:{}?immutable=1", db.display()))
        .arg(query)
        .output()
        .context("rig: read the profile history")?;
    let text = String::from_utf8_lossy(&out.stdout);
    text.lines()
        .find_map(site_of)
        .context("rig: no OneDrive site in the scratch profile's history; set PLSFIX_RIG_SITE")
}

/// `https://<host>/[:x:/r/]personal/<user>/...` -> `https://<host>/personal/<user>`.
pub fn site_of(url: &str) -> Option<String> {
    let rest = url.strip_prefix("https://")?;
    let (host, path) = rest.split_once('/')?;
    let at = path.find("personal/")?;
    let user = path[at + "personal/".len()..].split('/').next().filter(|u| !u.is_empty())?;
    Some(format!("https://{host}/personal/{user}"))
}
