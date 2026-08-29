// The version the manual describes, taken from package.json at build time so
// the cover and the footer can never drift from the add-in. Owns the one
// vMAJOR.MINOR.PATCH format, the same three-digit patch the pane footer shows.

const PACKAGE_JSON: &str = include_str!("../../package.json");
const SHORTCUTS_JSON: &str = include_str!("../../public/shortcuts.json");

/// The add-in version as the pane writes it, for example v2.2.006.
pub fn version() -> String {
    let package: serde_json::Value =
        serde_json::from_str(PACKAGE_JSON).expect("parse package.json");
    let semver = package["version"]
        .as_str()
        .expect("package.json has no version");
    format(semver)
}

fn format(semver: &str) -> String {
    let parts: Vec<&str> = semver.split('.').collect();
    assert_eq!(parts.len(), 3, "not a semver version: {semver}");
    format!("v{}.{}.{:0>3}", parts[0], parts[1], parts[2])
}

/// One keyboard shortcut: the English action name, the keys, and the action id
/// the Latvian gloss is looked up by.
pub struct Shortcut {
    pub id: String,
    pub name: String,
    pub keys: String,
}

/// Every shortcut in `public/shortcuts.json`, in the order the file lists them.
pub fn shortcuts() -> Vec<Shortcut> {
    let file: serde_json::Value =
        serde_json::from_str(SHORTCUTS_JSON).expect("parse shortcuts.json");
    let actions = file["actions"].as_array().expect("actions array");
    let bindings = file["shortcuts"].as_array().expect("shortcuts array");
    bindings
        .iter()
        .filter_map(|binding| shortcut(binding, actions))
        .collect()
}

fn shortcut(binding: &serde_json::Value, actions: &[serde_json::Value]) -> Option<Shortcut> {
    let id = binding["action"].as_str()?;
    let keys = binding["key"]["default"].as_str()?;
    let name = actions
        .iter()
        .find(|action| action["id"].as_str() == Some(id))?["name"]
        .as_str()?;
    Some(Shortcut {
        id: id.to_string(),
        name: name.to_string(),
        keys: keys.to_string(),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn the_patch_is_padded_to_three_digits() {
        assert_eq!(format("2.2.6"), "v2.2.006");
        assert_eq!(format("10.1.234"), "v10.1.234");
    }

    #[test]
    fn every_shortcut_names_its_action() {
        let list = shortcuts();
        assert!(list.len() > 30, "expected the full shortcut set");
        assert!(
            list.iter()
                .all(|s| !s.name.is_empty() && !s.keys.is_empty())
        );
    }
}
