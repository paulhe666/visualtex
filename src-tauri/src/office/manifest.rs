use crate::office::state::OFFICE_PORT;

pub const WORD_ADDIN_ID: &str = "d6fcb260-4c37-4f73-a173-cf24674f81f2";
pub const POWERPOINT_ADDIN_ID: &str = "a6d13cf2-54e8-4dfa-a20c-15de864ab3c5";
pub const WORD_MANIFEST_FILE: &str = "d6fcb260-4c37-4f73-a173-cf24674f81f2.VisualTeX.Word.xml";
pub const POWERPOINT_MANIFEST_FILE: &str =
    "a6d13cf2-54e8-4dfa-a20c-15de864ab3c5.VisualTeX.PowerPoint.xml";
pub const LEGACY_WORD_MANIFEST_FILE: &str = "VisualTeX.Word.xml";
pub const LEGACY_POWERPOINT_MANIFEST_FILE: &str = "VisualTeX.PowerPoint.xml";
// Revision 3 forces Office for Mac to rebuild its add-in command/runtime cache
// after the Word native-return and PowerPoint native-finalization fixes.
const MAC_MANIFEST_REVISION: &str = "3";

const WORD_TEMPLATE: &str =
    include_str!("../../../office/macos/manifests/visualtex-word.template.xml");
const POWERPOINT_TEMPLATE: &str =
    include_str!("../../../office/macos/manifests/visualtex-powerpoint.template.xml");

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ManifestHost {
    Word,
    PowerPoint,
}

impl ManifestHost {
    pub fn file_name(self) -> &'static str {
        match self {
            Self::Word => WORD_MANIFEST_FILE,
            Self::PowerPoint => POWERPOINT_MANIFEST_FILE,
        }
    }

    pub fn addin_id(self) -> &'static str {
        match self {
            Self::Word => WORD_ADDIN_ID,
            Self::PowerPoint => POWERPOINT_ADDIN_ID,
        }
    }

    fn template(self) -> &'static str {
        match self {
            Self::Word => WORD_TEMPLATE,
            Self::PowerPoint => POWERPOINT_TEMPLATE,
        }
    }
}

pub fn manifest_version() -> String {
    let mut parts = env!("CARGO_PKG_VERSION")
        .split('.')
        .map(|part| {
            part.chars()
                .take_while(|character| character.is_ascii_digit())
                .collect::<String>()
        })
        .filter(|part| !part.is_empty())
        .take(4)
        .collect::<Vec<_>>();
    while parts.len() < 4 {
        parts.push("0".to_string());
    }
    parts[3] = MAC_MANIFEST_REVISION.to_string();
    parts.join(".")
}

pub fn companion_origin() -> String {
    format!("https://127.0.0.1:{OFFICE_PORT}")
}

pub fn render_manifest(host: ManifestHost) -> Result<String, String> {
    let origin = companion_origin();
    let rendered = host
        .template()
        .replace("{{WORD_ADDIN_ID}}", WORD_ADDIN_ID)
        .replace("{{POWERPOINT_ADDIN_ID}}", POWERPOINT_ADDIN_ID)
        .replace("{{MANIFEST_VERSION}}", &manifest_version())
        .replace("{{COMPANION_ORIGIN}}", &origin);
    if rendered.contains("{{") || rendered.contains("}}") {
        return Err(format!(
            "{} still contains unresolved placeholders",
            host.file_name()
        ));
    }
    for forbidden in [
        "appsforoffice.microsoft.com",
        "github.com",
        "githubusercontent.com",
        "googleapis.com",
        "gstatic.com",
        "unpkg.com",
        "jsdelivr.net",
        "cloudflare.com",
    ] {
        if rendered.contains(forbidden) {
            return Err(format!(
                "{} contains forbidden remote domain: {forbidden}",
                host.file_name()
            ));
        }
    }
    if !rendered.contains(&format!("<Id>{}</Id>", host.addin_id())) {
        return Err(format!(
            "{} contains the wrong add-in GUID",
            host.file_name()
        ));
    }
    if !rendered.contains("<Permissions>ReadWriteDocument</Permissions>") {
        return Err(format!(
            "{} is missing ReadWriteDocument permission",
            host.file_name()
        ));
    }
    if !rendered.contains("<CustomTab id=\"VisualTeX.Tab\">")
        || rendered.contains("<OfficeTab id=\"TabHome\">")
    {
        return Err(format!(
            "{} must expose the independent VisualTeX ribbon tab",
            host.file_name()
        ));
    }
    if host == ManifestHost::Word
        && !rendered.contains("<FunctionName>VisualTeX.UpdateEquationNumbers</FunctionName>")
    {
        return Err(format!(
            "{} is missing the equation numbering refresh command",
            host.file_name()
        ));
    }
    Ok(rendered)
}

#[cfg(test)]
mod tests {
    use super::*;
    use uuid::Uuid;

    #[test]
    fn manifest_guids_are_distinct_permanent_uuid_v4_values() {
        let word = Uuid::parse_str(WORD_ADDIN_ID).expect("Word GUID");
        let powerpoint = Uuid::parse_str(POWERPOINT_ADDIN_ID).expect("PowerPoint GUID");
        assert_eq!(word.get_version_num(), 4);
        assert_eq!(powerpoint.get_version_num(), 4);
        assert_ne!(word, powerpoint);
    }

    #[test]
    fn rendered_manifests_use_fixed_local_origin_and_four_part_version() {
        let expected_version = format!("<Version>{}</Version>", manifest_version());
        for host in [ManifestHost::Word, ManifestHost::PowerPoint] {
            let manifest = render_manifest(host).expect("render manifest");
            assert!(manifest.contains("https://127.0.0.1:43127/bridge/index.html"));
            assert!(manifest.contains("https://127.0.0.1:43127/icons/icon-32.png"));
            assert!(manifest.contains(&expected_version));
            assert!(!manifest.contains("localhost"));
            assert!(!manifest.contains("{{"));
        }
    }

    #[test]
    fn hosts_and_permissions_are_correct() {
        let word = render_manifest(ManifestHost::Word).unwrap();
        assert!(word.contains("<Host Name=\"Document\" />"));
        assert!(word.contains("<Host xsi:type=\"Document\">"));
        assert!(!word.contains("<Host Name=\"Presentation\" />"));

        let powerpoint = render_manifest(ManifestHost::PowerPoint).unwrap();
        assert!(powerpoint.contains("<Host Name=\"Presentation\" />"));
        assert!(powerpoint.contains("<Host xsi:type=\"Presentation\">"));
        assert!(!powerpoint.contains("<Host Name=\"Document\" />"));
    }

    #[test]
    fn manifests_use_custom_ribbon_tab_and_word_exposes_numbering_command() {
        let word = render_manifest(ManifestHost::Word).unwrap();
        assert!(word.contains("<CustomTab id=\"VisualTeX.Tab\">"));
        assert!(word.contains("VisualTeX.UpdateEquationNumbers"));
        assert!(!word.contains("<OfficeTab id=\"TabHome\">"));

        let powerpoint = render_manifest(ManifestHost::PowerPoint).unwrap();
        assert!(powerpoint.contains("<CustomTab id=\"VisualTeX.Tab\">"));
        assert!(!powerpoint.contains("VisualTeX.UpdateEquationNumbers"));
        assert!(!powerpoint.contains("<OfficeTab id=\"TabHome\">"));
    }
}
