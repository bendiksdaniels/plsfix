//! copy_lint.rs: the rules every line of on-screen text keeps (Daniel's deliverable rules): no
//! em dash, none of the office-Latvian words he rejected, no double spaces, no empty line.
//! A finding names the copy key and the rule.

/// Word stems he rejected, matched case-insensitively anywhere in a word.
pub const BANNED: &[(&str, &str)] = &[
    ("cilne", "say what a person at a desk says, not Microsoft's word for a tab"),
    ("ciln", "say what a person at a desk says, not Microsoft's word for a tab"),
    ("lietotn", "the app is \"aplikācija\", never \"lietotne\""),
    ("aplēs", "\"aplēse\" was rejected: say \"novērtējums\""),
    ("kohort", "\"kohorta\" is banned"),
];

/// Every rule `text` (the copy for `key`) breaks, one line each.
pub fn lint(key: &str, text: &str) -> Vec<String> {
    let mut out = Vec::new();
    if text.trim().is_empty() {
        out.push(format!("{key}: empty"));
    }
    if text.contains('\u{2014}') {
        out.push(format!("{key}: an em dash; use a comma, colon or parentheses"));
    }
    if text.contains("  ") {
        out.push(format!("{key}: a double space"));
    }
    let lower = text.to_lowercase();
    let mut seen = std::collections::BTreeSet::new();
    for (stem, why) in BANNED {
        if lower.contains(stem) && seen.insert(*why) {
            out.push(format!("{key}: \"{stem}\" ({why})"));
        }
    }
    out
}
