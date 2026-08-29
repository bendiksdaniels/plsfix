// The gate on the generated manual: it builds the whole document in memory and
// checks what a reader would notice if it broke. Owns the language rules for
// the deliverable, the chapter list and the picture count.

use std::io::Read;

/// Builds the document once and hands back the text of word/document.xml.
fn document_xml() -> String {
    let bytes = manual::render::build();
    let mut archive =
        zip::ZipArchive::new(std::io::Cursor::new(bytes)).expect("the build is a zip");
    let mut entry = archive
        .by_name("word/document.xml")
        .expect("word/document.xml");
    let mut xml = String::new();
    entry.read_to_string(&mut xml).expect("read document.xml");
    xml
}

/// The text a reader sees, with every tag taken out.
fn plain_text(xml: &str) -> String {
    let mut text = String::with_capacity(xml.len());
    let mut inside_tag = false;
    for character in xml.chars() {
        match character {
            '<' => inside_tag = true,
            '>' => inside_tag = false,
            other if !inside_tag => text.push(other),
            _ => {}
        }
    }
    text
}

#[test]
fn every_chapter_heading_is_in_the_document() {
    let text = plain_text(&document_xml());
    for chapter in manual::content::chapters() {
        assert!(
            text.contains(chapter.title),
            "chapter heading missing: {}",
            chapter.title
        );
        for section in chapter.sections {
            assert!(
                text.contains(section.title),
                "section heading missing: {}",
                section.title
            );
        }
    }
    assert!(text.contains("Saturs"), "the contents page is missing");
}

#[test]
fn one_picture_per_referenced_screenshot() {
    let xml = document_xml();
    let pictures = xml.matches("<pic:pic").count();
    let referenced = manual::content::images();
    assert_eq!(
        pictures,
        referenced.len(),
        "the document embeds {pictures} pictures for {} referenced screenshots",
        referenced.len()
    );
    let alts = xml.matches("descr=\"").count();
    assert_eq!(alts, pictures, "every picture needs alt text");
}

#[test]
fn every_referenced_screenshot_exists() {
    for file in manual::content::images() {
        let path = manual::shots::folder().join(file);
        assert!(path.is_file(), "missing screenshot: {}", path.display());
    }
}

#[test]
fn the_latvian_rules_hold() {
    let text = plain_text(&document_xml());
    assert!(!text.contains('\u{2014}'), "an em dash slipped in");
    assert!(!text.contains('\u{2013}'), "an en dash slipped in");
    for banned in ["aplēse", "Aplēse", "aplēses"] {
        assert!(!text.contains(banned), "banned word in the text: {banned}");
    }
    for chunk in manual::content::prose() {
        assert!(!chunk.contains('\u{2014}'), "an em dash in: {chunk}");
        assert!(
            !chunk.to_lowercase().contains("aplēs"),
            "a banned word in: {chunk}"
        );
    }
}

#[test]
fn the_version_is_printed_and_substituted() {
    let text = plain_text(&document_xml());
    let version = manual::meta::version();
    assert!(text.contains(&version), "the cover must name {version}");
    assert!(
        !text.contains("{version}"),
        "a version placeholder was left unfilled"
    );
}

#[test]
fn every_shortcut_carries_a_latvian_gloss() {
    for shortcut in manual::meta::shortcuts() {
        assert!(
            !manual::content::reference::gloss(&shortcut.id).is_empty(),
            "no Latvian gloss for {}",
            shortcut.id
        );
    }
}
