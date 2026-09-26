//! cdp.rs: a raw Chrome DevTools Protocol client over one page's WebSocket.
//! Owns request ids and reply matching. Events are read as plain JSON and only console errors
//! and exceptions are kept, so a newer Chrome's protocol changes cannot break the client.

use anyhow::{bail, Context, Result};
use base64::Engine;
use serde_json::{json, Value};
use std::net::TcpStream;
use std::thread::sleep;
use std::time::{Duration, Instant};
use tungstenite::stream::MaybeTlsStream;
use tungstenite::{Message, WebSocket};

/// How long one CDP reply may take before the call fails (a hung page, a dead Chrome).
const REPLY_TIMEOUT: Duration = Duration::from_secs(60);

pub struct Cdp {
    ws: WebSocket<MaybeTlsStream<TcpStream>>,
    next_id: u64,
    errors: Vec<String>,
}

impl Cdp {
    /// Connects to a page target's `webSocketDebuggerUrl`.
    pub fn connect(ws_url: &str) -> Result<Cdp> {
        let (ws, _) = tungstenite::connect(ws_url).with_context(|| format!("cdp connect {ws_url}"))?;
        if let MaybeTlsStream::Plain(s) = ws.get_ref() {
            s.set_read_timeout(Some(REPLY_TIMEOUT)).context("cdp connect: read timeout")?;
        }
        Ok(Cdp { ws, next_id: 0, errors: Vec::new() })
    }

    /// Sends one command and returns its `result`; events read meanwhile are noted, not returned.
    pub fn call(&mut self, method: &str, params: Value) -> Result<Value> {
        self.next_id += 1;
        let id = self.next_id;
        let msg = json!({ "id": id, "method": method, "params": params }).to_string();
        self.ws.send(Message::text(msg)).with_context(|| format!("cdp send {method}"))?;
        loop {
            let text = match self.ws.read().with_context(|| format!("cdp read {method}"))? {
                Message::Text(t) => t.to_string(),
                Message::Binary(b) => String::from_utf8_lossy(&b).into_owned(),
                Message::Close(_) => bail!("cdp {method}: Chrome closed the socket"),
                _ => continue,
            };
            let v: Value = serde_json::from_str(&text).with_context(|| format!("cdp {method}: reply is not JSON"))?;
            if v.get("id").and_then(Value::as_u64) == Some(id) {
                if let Some(e) = v.get("error") {
                    bail!("cdp {method}: {e}");
                }
                return Ok(v.get("result").cloned().unwrap_or(Value::Null));
            }
            self.note_event(&v);
        }
    }

    /// Keeps uncaught exceptions and console errors for `take_errors`.
    fn note_event(&mut self, v: &Value) {
        let p = &v["params"];
        match v.get("method").and_then(Value::as_str) {
            Some("Runtime.exceptionThrown") => {
                let d = &p["exceptionDetails"];
                let text = d["exception"]["description"].as_str().or(d["text"].as_str()).unwrap_or("?");
                self.errors.push(format!("exception: {text}"));
            }
            Some("Runtime.consoleAPICalled") if p["type"] == "error" => {
                let args: Vec<String> = p["args"].as_array().into_iter().flatten()
                    .map(|a| a["value"].as_str().map(str::to_string).unwrap_or_else(|| a["description"].to_string()))
                    .collect();
                self.errors.push(format!("console.error: {}", args.join(" ")));
            }
            Some("Log.entryAdded") if p["entry"]["level"] == "error" => {
                let e = &p["entry"];
                self.errors.push(format!("log: {} {}", e["text"].as_str().unwrap_or("?"), e["url"].as_str().unwrap_or("")));
            }
            _ => {}
        }
    }

    /// The exceptions and console errors seen since the last call of this function.
    pub fn take_errors(&mut self) -> Vec<String> {
        std::mem::take(&mut self.errors)
    }

    /// Evaluates `js` and returns its value (JSON); a thrown error fails with its message.
    /// Only this tool's own scripts run here, in a throwaway headless page (CDP Runtime.evaluate).
    pub fn eval(&mut self, js: &str) -> Result<Value> {
        self.evaluate(js, false)
    }

    /// Like `eval`, awaiting the promise `js` returns.
    pub fn eval_async(&mut self, js: &str) -> Result<Value> {
        self.evaluate(js, true)
    }

    fn evaluate(&mut self, js: &str, await_promise: bool) -> Result<Value> {
        let r = self.call("Runtime.evaluate", json!({ "expression": js, "returnByValue": true, "awaitPromise": await_promise }))?;
        if let Some(d) = r.get("exceptionDetails") {
            let text = d["exception"]["description"].as_str().or(d["text"].as_str()).unwrap_or("?");
            let head: String = js.chars().take(80).collect();
            bail!("page script threw: {text} (in: {head})");
        }
        Ok(r["result"].get("value").cloned().unwrap_or(Value::Null))
    }

    /// Polls `js` until it is truthy; fails after `timeout` naming `what`.
    pub fn wait_for(&mut self, js: &str, timeout: Duration, what: &str) -> Result<()> {
        let start = Instant::now();
        loop {
            if self.eval(js)?.as_bool() == Some(true) {
                return Ok(());
            }
            if start.elapsed() > timeout {
                bail!("waited {}s for {what}", timeout.as_secs());
            }
            sleep(Duration::from_millis(100));
        }
    }

    /// Loads `url` and waits for `document.readyState` to be complete.
    pub fn navigate(&mut self, url: &str) -> Result<()> {
        self.call("Page.navigate", json!({ "url": url }))?;
        sleep(Duration::from_millis(50));
        self.wait_for("document.readyState === 'complete'", Duration::from_secs(20), &format!("load of {url}"))
    }

    /// A PNG of the current viewport.
    pub fn screenshot_png(&mut self) -> Result<Vec<u8>> {
        let r = self.call("Page.captureScreenshot", json!({ "format": "png", "optimizeForSpeed": true }))?;
        let data = r["data"].as_str().context("cdp Page.captureScreenshot: no data")?;
        base64::engine::general_purpose::STANDARD.decode(data).context("cdp Page.captureScreenshot: bad base64")
    }
}
