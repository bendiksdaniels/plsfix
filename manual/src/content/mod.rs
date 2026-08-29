// The manual as data: chapters, sections and blocks, with no Word knowledge at
// all. Owns every sentence the reader sees and the order they come in; the
// invariant is that nothing here renders itself, so `render.rs` is the only
// place that knows what a heading or a picture looks like.

mod charts;
mod install;
mod intro;
mod links;
pub mod reference;
mod tools;
mod workbook;

/// One piece of a section. Text is Latvian; UI labels stay English in quotes.
pub enum Block {
    /// A plain paragraph.
    Para(&'static str),
    /// Numbered steps, one string per step.
    Steps(&'static [&'static str]),
    /// Unnumbered points, one string per point.
    Bullets(&'static [&'static str]),
    /// A table with a header row.
    Table {
        head: &'static [&'static str],
        rows: &'static [&'static [&'static str]],
    },
    /// The shortcut table, built from `public/shortcuts.json` at build time.
    Shortcuts,
    /// A screenshot from `manual/shots`, with the alt text it carries in Word.
    Image {
        file: &'static str,
        alt: &'static str,
    },
}

pub struct Section {
    pub title: &'static str,
    pub blocks: Vec<Block>,
}

pub struct Chapter {
    pub title: &'static str,
    pub sections: Vec<Section>,
}

/// The picture on the title page.
pub const COVER_IMAGE: &str = "icon.png";
pub const COVER_ALT: &str = "pls,fix logotips: komats uz tumši zila kvadrāta.";

pub fn section(title: &'static str, blocks: Vec<Block>) -> Section {
    Section { title, blocks }
}

/// Every chapter in reading order.
pub fn chapters() -> Vec<Chapter> {
    vec![
        intro::what_it_is(),
        install::setup(),
        intro::pane(),
        tools::tools(),
        workbook::workbook(),
        links::links(),
        workbook::brand(),
        reference::shortcuts(),
        reference::problems(),
        reference::version(),
    ]
}

/// Every picture the document embeds, cover first, in document order. The test
/// counts the pictures in the built file against this list.
pub fn images() -> Vec<&'static str> {
    let mut files = vec![COVER_IMAGE];
    for chapter in chapters() {
        for section in chapter.sections {
            for block in section.blocks {
                if let Block::Image { file, .. } = block {
                    files.push(file);
                }
            }
        }
    }
    files
}

/// Every sentence in the document, for the language gates in the test.
pub fn prose() -> Vec<String> {
    let mut out: Vec<String> = Vec::new();
    for chapter in chapters() {
        out.push(chapter.title.to_string());
        for section in chapter.sections {
            out.push(section.title.to_string());
            for block in section.blocks {
                collect(&block, &mut out);
            }
        }
    }
    out
}

fn collect(block: &Block, out: &mut Vec<String>) {
    match block {
        Block::Para(text) => out.push((*text).to_string()),
        Block::Steps(items) | Block::Bullets(items) => {
            out.extend(items.iter().map(|i| (*i).to_string()));
        }
        Block::Table { head, rows } => {
            out.extend(head.iter().map(|h| (*h).to_string()));
            for row in *rows {
                out.extend(row.iter().map(|c| (*c).to_string()));
            }
        }
        Block::Image { alt, .. } => out.push((*alt).to_string()),
        Block::Shortcuts => {}
    }
}
