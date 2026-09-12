//! Manifest route tests, mirroring `src/manifest.rs`: the committed file
//! verbatim for the hosted add-in, re-pointed with an id of its own for any
//! other host, and the shipped manifest.prod.xml holding the shape the
//! rewrite relies on.

use std::{path::PathBuf, sync::Arc};

use axum::{
    body::Body,
    http::{header, Request, StatusCode},
};
use http_body_util::BodyExt;
use plsfix_server::{
    app,
    manifest::{addin_id, rewrite, ManifestSource, HOSTED_BASE},
    relay::AppState,
    store::Store,
};
use tower::ServiceExt;

const HOSTED_ID: &str = "FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E";

const SAMPLE: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<!-- Served LIVE from https://dbautomatizacijas.com/modelis/ -->
<OfficeApp>
  <Id>FF1B34D8-DD7D-4B39-8FA9-6248CA09DB6E</Id>
  <IconUrl DefaultValue="https://dbautomatizacijas.com/modelis/assets/icon-32.png"/>
  <SupportUrl DefaultValue="https://github.com/bendiksdaniels/plsfix"/>
  <AppDomains>
    <AppDomain>https://dbautomatizacijas.com</AppDomain>
  </AppDomains>
  <SourceLocation DefaultValue="https://dbautomatizacijas.com/modelis/taskpane.html"/>
</OfficeApp>
"#;

fn temp_root() -> PathBuf {
    let root = std::env::temp_dir().join(format!("modelis-manifest-{}", std::process::id()));
    std::fs::create_dir_all(root.join("dist")).unwrap();
    std::fs::write(
        root.join("dist").join("taskpane.html"),
        "<title>pls,fix</title>",
    )
    .unwrap();
    root
}

fn sample_file(name: &str) -> PathBuf {
    let file = temp_root().join(format!("{name}.xml"));
    std::fs::write(&file, SAMPLE).unwrap();
    file
}

async fn get(source: ManifestSource) -> (StatusCode, Option<String>, Option<String>, String) {
    let state = Arc::new(AppState {
        manifest: source,
        ..AppState::new(Store::in_memory().unwrap())
    });
    let response = app(temp_root().join("dist"), state)
        .oneshot(Request::get("/manifest.xml").body(Body::empty()).unwrap())
        .await
        .unwrap();
    let status = response.status();
    let text = |name: header::HeaderName| {
        response
            .headers()
            .get(name)
            .map(|value| value.to_str().unwrap().to_string())
    };
    let content_type = text(header::CONTENT_TYPE);
    let cache = text(header::CACHE_CONTROL);
    let body = response.into_body().collect().await.unwrap().to_bytes();
    (
        status,
        content_type,
        cache,
        String::from_utf8_lossy(&body).to_string(),
    )
}

#[tokio::test]
async fn the_hosted_manifest_is_served_verbatim() {
    let source = ManifestSource {
        file: sample_file("verbatim"),
        public_url: None,
    };
    let (status, content_type, cache, body) = get(source).await;
    assert_eq!(status, StatusCode::OK);
    assert_eq!(body, SAMPLE);
    assert!(content_type.unwrap().starts_with("application/xml"));
    assert_eq!(cache.as_deref(), Some("no-cache"));
}

#[tokio::test]
async fn a_self_hosted_manifest_is_repointed() {
    let source = ManifestSource {
        file: sample_file("repointed"),
        public_url: Some("https://models.example.com/plsfix".to_string()),
    };
    let (status, _, _, body) = get(source).await;
    assert_eq!(status, StatusCode::OK);
    assert!(body.contains(r#"DefaultValue="https://models.example.com/plsfix/taskpane.html""#));
    assert!(body.contains(r#"DefaultValue="https://models.example.com/plsfix/assets/icon-32.png""#));
    assert!(body.contains("<AppDomain>https://models.example.com</AppDomain>"));
    assert!(
        body.contains(r#"<SupportUrl DefaultValue="https://github.com/bendiksdaniels/plsfix"/>"#)
    );
    assert!(!body.contains("dbautomatizacijas"), "{body}");
    let expected = addin_id("https://models.example.com/plsfix/");
    assert!(body.contains(&format!("<Id>{expected}</Id>")), "{body}");
    assert_ne!(expected, HOSTED_ID);
}

#[test]
fn the_id_is_stable_per_host_and_differs_between_hosts() {
    let a = addin_id("https://models.example.com/plsfix/");
    assert_eq!(a, addin_id("https://models.example.com/plsfix"));
    assert_eq!(a.len(), 36);
    assert_eq!(a, a.to_uppercase());
    assert_ne!(a, addin_id("https://other.example.com/"));
    assert_ne!(a, HOSTED_ID);
}

#[test]
fn the_hosted_url_changes_nothing() {
    assert_eq!(rewrite(SAMPLE, HOSTED_BASE), SAMPLE);
    assert_eq!(rewrite(SAMPLE, HOSTED_BASE.trim_end_matches('/')), SAMPLE);
}

#[test]
fn a_port_stays_in_the_origin() {
    let out = rewrite(SAMPLE, "http://localhost:8804");
    assert!(
        out.contains("<AppDomain>http://localhost:8804</AppDomain>"),
        "{out}"
    );
    assert!(out.contains(r#"DefaultValue="http://localhost:8804/taskpane.html""#));
}

#[tokio::test]
async fn a_missing_file_is_a_404() {
    let source = ManifestSource {
        file: temp_root().join("nowhere.xml"),
        public_url: Some("https://models.example.com/".to_string()),
    };
    let (status, _, _, _) = get(source).await;
    assert_eq!(status, StatusCode::NOT_FOUND);
}

/// The committed manifest.prod.xml, the file production serves: the rewrite
/// must find the hosted base, the AppDomain and the Id in it, or a self-hosted
/// copy would keep pointing at the hosted add-in.
#[test]
fn the_shipped_manifest_has_the_shape_the_rewrite_needs() {
    let shipped = std::fs::read_to_string("../manifest.prod.xml").unwrap();
    assert!(shipped.contains(HOSTED_BASE));
    assert!(shipped.contains(&format!("<Id>{HOSTED_ID}</Id>")));
    assert_eq!(rewrite(&shipped, HOSTED_BASE), shipped);
    let out = rewrite(&shipped, "https://example.test/");
    assert!(
        !out.contains("dbautomatizacijas"),
        "a hosted URL survived the rewrite"
    );
    assert_eq!(out.matches("<Id>").count(), 1);
    assert!(out.contains(&format!("<Id>{}</Id>", addin_id("https://example.test/"))));
    for page in [
        "taskpane.html",
        "pptpane.html",
        "functions.js",
        "functions.json",
    ] {
        assert!(
            out.contains(&format!("https://example.test/{page}")),
            "{page} missing"
        );
    }
    assert!(out.contains("<AppDomain>https://example.test</AppDomain>"));
}
