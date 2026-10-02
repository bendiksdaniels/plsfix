//! The add-in manifest for whoever runs this server: the shipped
//! `manifest.prod.xml` verbatim, or the same file re-pointed at
//! `MODELIS_PUBLIC_URL` with an add-in id of its own, so a self-hosted copy
//! never collides with the hosted add-in on one machine.
//! Invariant: a rewrite touches the base URL, the AppDomain and the Id, nothing else,
//! and every interpolated URL is XML-escaped.

use std::{env, fs, io, path::PathBuf};

use uuid::Uuid;

/// Where the committed manifest points; every URL under it is rewritten.
pub const HOSTED_BASE: &str = "https://dbautomatizacijas.com/modelis/";
/// The AppDomain the committed manifest trusts.
pub const HOSTED_ORIGIN: &str = "https://dbautomatizacijas.com";

/// The file to serve and, for a self-hosted copy, the URL it is served from.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct ManifestSource {
    pub file: PathBuf,
    pub public_url: Option<String>,
}

impl Default for ManifestSource {
    fn default() -> Self {
        ManifestSource {
            file: PathBuf::from("manifest.prod.xml"),
            public_url: None,
        }
    }
}

impl ManifestSource {
    /// `MODELIS_MANIFEST` (default `manifest.prod.xml`, relative to the working
    /// directory) and `MODELIS_PUBLIC_URL` (unset or empty = the hosted add-in).
    pub fn from_env() -> Self {
        let file = env::var("MODELIS_MANIFEST")
            .map(PathBuf::from)
            .unwrap_or_else(|_| Self::default().file);
        let public_url = env::var("MODELIS_PUBLIC_URL")
            .ok()
            .map(|url| url.trim().to_string())
            .filter(|url| !url.is_empty());
        ManifestSource { file, public_url }
    }
}

/// The manifest as this server hands it out.
pub fn render(source: &ManifestSource) -> io::Result<String> {
    let xml = fs::read_to_string(&source.file)?;
    Ok(match &source.public_url {
        Some(url) => rewrite(&xml, url),
        None => xml,
    })
}

/// The committed manifest re-pointed at `public_url`: every hosted URL, the
/// AppDomain and the Id. The hosted URL itself changes nothing.
pub fn rewrite(xml: &str, public_url: &str) -> String {
    let base = with_slash(public_url);
    if base == HOSTED_BASE {
        return xml.to_string();
    }
    let origin = xml_text(&origin_of(&base));
    let escaped_base = xml_text(&base);
    let repointed = xml.replace(HOSTED_BASE, &escaped_base).replace(
        &format!("<AppDomain>{HOSTED_ORIGIN}</AppDomain>"),
        &format!("<AppDomain>{origin}</AppDomain>"),
    );
    replace_id(&repointed, &addin_id(&base))
}

fn xml_text(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// A stable id per host: the same URL always gets the same add-in, another
/// URL another one (UUID v5 over the URL namespace, upper case like Office's own).
pub fn addin_id(base: &str) -> String {
    Uuid::new_v5(&Uuid::NAMESPACE_URL, with_slash(base).as_bytes())
        .hyphenated()
        .to_string()
        .to_uppercase()
}

fn with_slash(url: &str) -> String {
    let url = url.trim();
    if url.ends_with('/') {
        url.to_string()
    } else {
        format!("{url}/")
    }
}

/// `https://host:port/any/path/` -> `https://host:port`.
fn origin_of(base: &str) -> String {
    match base.find("://") {
        Some(at) => {
            let rest = &base[at + 3..];
            let host_end = rest.find('/').unwrap_or(rest.len());
            format!("{}{}", &base[..at + 3], &rest[..host_end])
        }
        None => base.trim_end_matches('/').to_string(),
    }
}

/// The text between the first `<Id>` and `</Id>`; a manifest without one is
/// handed back untouched rather than guessed at.
fn replace_id(xml: &str, id: &str) -> String {
    let (Some(open), Some(close)) = (xml.find("<Id>"), xml.find("</Id>")) else {
        return xml.to_string();
    };
    if close < open {
        return xml.to_string();
    }
    format!("{}<Id>{id}{}", &xml[..open], &xml[close..])
}
