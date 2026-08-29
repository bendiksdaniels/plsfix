// Two things docx-rs cannot write itself, patched into the packed archive:
// alt text on every picture and an Arial fallback for Aptos. Owns the only
// place that edits raw OOXML; the invariant is that every other entry in the
// zip is copied through untouched.

use std::io::{Cursor, Read, Write};

/// What docx-rs writes for every inline picture, one per image, in order.
const DOC_PR: &str = r#"<wp:docPr id="1" name="Figure" />"#;
/// The font table entry Word reads before it substitutes a missing font.
const ARIAL_ENTRY: &str = r#"<w:font w:name="Arial">"#;
const APTOS_ENTRY: &str = concat!(
    r#"<w:font w:name="Aptos"><w:altName w:val="Arial" />"#,
    r#"<w:charset w:val="00" /><w:family w:val="swiss" />"#,
    r#"<w:pitch w:val="variable" /></w:font>"#,
);

/// Rewrites the two parts that need patching and returns the finished archive.
pub fn finish(packed: &[u8], alts: &[String]) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(Cursor::new(packed)).expect("read the packed docx");
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    let names: Vec<String> = archive.file_names().map(str::to_string).collect();
    for name in names {
        let mut entry = archive.by_name(&name).expect("zip entry");
        match name.as_str() {
            "word/document.xml" => write(&mut out, &name, &alt_text(&read(&mut entry), alts)),
            "word/fontTable.xml" => write(&mut out, &name, &font_fallback(&read(&mut entry))),
            _ => out.raw_copy_file(entry).expect("copy zip entry"),
        }
    }
    out.finish().expect("finish the zip").into_inner()
}

fn read(entry: &mut impl Read) -> String {
    let mut text = String::new();
    entry.read_to_string(&mut text).expect("read zip entry");
    text
}

fn write(out: &mut zip::ZipWriter<Cursor<Vec<u8>>>, name: &str, body: &str) {
    let options: zip::write::FileOptions<'_, ()> =
        zip::write::FileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    out.start_file(name, options).expect("start zip entry");
    out.write_all(body.as_bytes()).expect("write zip entry");
}

/// Gives each picture its own id, name and alt text, in document order.
fn alt_text(document: &str, alts: &[String]) -> String {
    let mut out = String::with_capacity(document.len() + alts.len() * 120);
    let mut rest = document;
    for (index, alt) in alts.iter().enumerate() {
        let Some(at) = rest.find(DOC_PR) else { break };
        out.push_str(&rest[..at]);
        out.push_str(&format!(
            r#"<wp:docPr id="{id}" name="Attēls {id}" descr="{alt}" />"#,
            id = index + 1,
            alt = escape(alt),
        ));
        rest = &rest[at + DOC_PR.len()..];
    }
    out.push_str(rest);
    out
}

/// Declares Aptos with Arial as its alternate, so a machine without Aptos
/// still renders the document in a matching sans face.
fn font_fallback(font_table: &str) -> String {
    if font_table.contains(r#"w:name="Aptos""#) {
        return font_table.to_string();
    }
    font_table.replacen(ARIAL_ENTRY, &format!("{APTOS_ENTRY}{ARIAL_ENTRY}"), 1)
}

fn escape(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn every_picture_gets_its_own_alt_text() {
        let document = format!("<a>{DOC_PR}</a><b>{DOC_PR}</b>");
        let alts = vec!["Pirmais".to_string(), r#"Otrais "citāts""#.to_string()];
        let out = alt_text(&document, &alts);
        assert!(out.contains(r#"id="1" name="Attēls 1" descr="Pirmais""#));
        assert!(out.contains(r#"descr="Otrais &quot;citāts&quot;""#));
        assert!(!out.contains(DOC_PR));
    }

    #[test]
    fn the_font_table_learns_about_aptos() {
        let table = format!("<w:fonts>{ARIAL_ENTRY}</w:font></w:fonts>");
        let out = font_fallback(&table);
        assert!(out.contains(r#"<w:altName w:val="Arial" />"#));
        assert_eq!(font_fallback(&out), out);
    }
}
