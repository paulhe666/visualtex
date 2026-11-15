//! Checking the compiled add-ins before they are installed: zip entries, VBA markers, the main content type and the
//! manifest.

use super::*;

pub(super) fn bytes_contain(haystack: &[u8], needle: &[u8]) -> bool {
    !needle.is_empty()
        && haystack
            .windows(needle.len())
            .any(|window| window == needle)
}

pub(super) fn sha256_hex(bytes: &[u8]) -> String {
    format!("{:x}", Sha256::digest(bytes))
}

#[cfg(target_os = "macos")]
pub(super) fn validate_zip_entries(path: &Path, required_entries: &[&str]) -> Result<(), String> {
    let integrity = Command::new("/usr/bin/unzip")
        .arg("-tqq")
        .arg(path)
        .output()
        .map_err(|error| format!("Unable to validate {} with unzip: {error}", path.display()))?;
    if !integrity.status.success() {
        let detail = String::from_utf8_lossy(&integrity.stderr)
            .trim()
            .to_string();
        return Err(if detail.is_empty() {
            format!(
                "Compiled add-in {} is a corrupt OOXML package",
                path.display()
            )
        } else {
            format!("Compiled add-in {} is corrupt: {detail}", path.display())
        });
    }
    let listing = Command::new("/usr/bin/unzip")
        .arg("-Z1")
        .arg(path)
        .output()
        .map_err(|error| format!("Unable to inspect {} with unzip: {error}", path.display()))?;
    if !listing.status.success() {
        return Err(format!(
            "Unable to list OOXML entries in {}",
            path.display()
        ));
    }
    let listing_text = String::from_utf8_lossy(&listing.stdout);
    let entries = listing_text
        .lines()
        .map(str::trim)
        .collect::<std::collections::HashSet<_>>();
    for entry in required_entries {
        if !entries.contains(entry) {
            return Err(format!(
                "Compiled add-in {} does not contain required OOXML entry {entry}",
                path.display()
            ));
        }
    }
    Ok(())
}

#[cfg(not(target_os = "macos"))]
pub(super) fn validate_zip_entries(_path: &Path, _required_entries: &[&str]) -> Result<(), String> {
    Ok(())
}

#[cfg(target_os = "macos")]
pub(super) fn validate_vba_markers(path: &Path, vba_entry: &str, markers: &[&str]) -> Result<(), String> {
    let output = Command::new("/usr/bin/unzip")
        .args(["-p"])
        .arg(path)
        .arg(vba_entry)
        .output()
        .map_err(|error| {
            format!(
                "Unable to inspect VBA modules in {}: {error}",
                path.display()
            )
        })?;
    if !output.status.success() {
        return Err(format!(
            "Unable to read {vba_entry} from {}",
            path.display()
        ));
    }
    for marker in markers {
        if !bytes_contain(&output.stdout, marker.as_bytes()) {
            return Err(format!(
                "Compiled add-in {} is missing required VBA marker {marker}",
                path.display()
            ));
        }
    }
    Ok(())
}

#[cfg(not(target_os = "macos"))]
pub(super) fn validate_vba_markers(_path: &Path, _vba_entry: &str, _markers: &[&str]) -> Result<(), String> {
    Ok(())
}

pub(super) fn validate_main_content_type(path: &Path, expected_name: &str) -> Result<(), String> {
    let (part_name, content_type) = match expected_name {
        WORD_ADDIN_NAME => (WORD_MAIN_PART, WORD_MAIN_CONTENT_TYPE),
        POWERPOINT_ADDIN_NAME => (POWERPOINT_MAIN_PART, POWERPOINT_MAIN_CONTENT_TYPE),
        _ => return Err(format!("Unsupported compiled add-in name {expected_name}")),
    };
    let output = Command::new("/usr/bin/unzip")
        .args(["-p"])
        .arg(path)
        .arg(CONTENT_TYPES_ZIP_PATTERN)
        .output()
        .map_err(|error| {
            format!(
                "Unable to inspect the OOXML main content type in {}: {error}",
                path.display()
            )
        })?;
    if !output.status.success() {
        return Err(format!(
            "Unable to read {CONTENT_TYPES_ENTRY} from {}",
            path.display()
        ));
    }
    let xml = String::from_utf8_lossy(&output.stdout);
    let valid = xml
        .split("<Override")
        .any(|segment| segment.contains(part_name) && segment.contains(content_type));
    if valid {
        Ok(())
    } else {
        Err(format!(
            "Compiled add-in {} is not a real Office add-in package: expected {content_type} for {part_name}",
            path.display()
        ))
    }
}

pub(super) fn validate_compiled_addin(
    path: &Path,
    expected_name: &str,
    expected_vba_entry: &str,
) -> Result<Vec<u8>, String> {
    if path.file_name().and_then(|value| value.to_str()) != Some(expected_name) {
        return Err(format!("Compiled add-in must be named {expected_name}"));
    }
    let bytes = fs::read(path)
        .map_err(|error| format!("Compiled add-in {} is missing: {error}", path.display()))?;
    if bytes.len() < 1024 {
        return Err(format!(
            "Compiled add-in {} is too small to contain a VBA project",
            path.display()
        ));
    }
    if !bytes.starts_with(b"PK\x03\x04") {
        return Err(format!(
            "Compiled add-in {} is not a valid OOXML package",
            path.display()
        ));
    }
    let required_entries = [CONTENT_TYPES_ENTRY, expected_vba_entry, CUSTOM_UI_ENTRY];
    for entry in required_entries {
        if !bytes_contain(&bytes, entry.as_bytes()) {
            return Err(format!(
                "Compiled add-in {} does not contain required OOXML entry {entry}",
                path.display()
            ));
        }
    }
    validate_zip_entries(path, &required_entries)?;
    validate_main_content_type(path, expected_name)?;
    let required_vba_markers: &[&str] = match expected_name {
        WORD_ADDIN_NAME => &[
            "VTProtocol",
            "VTOfficePaths",
            "VTWordAdapter",
            "VTWordEvents",
            "VisualTeX_InitializeWordHost",
            "App_WindowBeforeDoubleClick",
            "VTWordRibbonApplyImageFontSizePreset",
            "VTWordRibbonApplyConversionFontSize",
            env!("CARGO_PKG_VERSION"),
            WORD_VBA_SOURCE_REVISION,
        ],
        POWERPOINT_ADDIN_NAME => &[
            "VTProtocol",
            "VTOfficePaths",
            "VTPowerPointAdapter",
            "VTPowerPointEvents",
            "VTPowerPointNativeMath",
            "VTPowerPointNativeTarget",
            "VisualTeX_NewInlineNativeEquation",
            "VisualTeX_NewDisplayNativeEquation",
            "VisualTeX_CopyNativeParagraphContext",
            "VTDispatchPowerPointNativeEquation",
            "VTTryResolveNativeMathZone",
            "VisualTeX_CopyPowerPointNativeEditSource",
            "VisualTeX_EditNativeAtScreenPoint",
            "VTPowerPointRibbonInlineOmml",
            "VTPowerPointRibbonDisplayOmml",
            "Auto_Open",
            "App_WindowBeforeDoubleClick",
            "App_WindowSelectionChange",
            "VTPowerPointRibbonApplyFormulaFontSizePreset",
            "VTPowerPointRedraw",
            "VisualTeX_ApplyPresentationItem",
            "VisualTeX_CopyPresentationEquation",
            "VTPowerPointRibbonRedrawOmmlSelection",
            env!("CARGO_PKG_VERSION"),
            POWERPOINT_VBA_SOURCE_REVISION,
        ],
        _ => return Err(format!("Unsupported compiled add-in name {expected_name}")),
    };
    validate_vba_markers(path, expected_vba_entry, required_vba_markers)?;
    Ok(bytes)
}

pub(super) fn read_addin_manifest(root: &Path) -> Result<AddinManifest, String> {
    let manifest_path = root.join("resources").join(ADDIN_MANIFEST_NAME);
    let manifest_bytes = fs::read(&manifest_path)
        .map_err(|error| format!("Compiled add-in manifest is missing: {error}"))?;
    if manifest_bytes.is_empty() || manifest_bytes.len() > 64 * 1024 {
        return Err("Compiled add-in manifest has an invalid size".to_string());
    }
    let manifest: AddinManifest = serde_json::from_slice(&manifest_bytes)
        .map_err(|error| format!("Compiled add-in manifest is invalid: {error}"))?;
    if manifest.schema_version != 1 || manifest.plugin_version != env!("CARGO_PKG_VERSION") {
        return Err("Compiled add-in manifest version does not match VisualTeX".to_string());
    }
    Ok(manifest)
}

pub(super) fn validate_compiled_artifacts(root: &Path) -> Result<(PathBuf, PathBuf), String> {
    let manifest = read_addin_manifest(root)?;
    let (word, powerpoint) = source_artifacts(root);
    for (path, name, vba_entry) in [
        (&word, WORD_ADDIN_NAME, WORD_VBA_ENTRY),
        (&powerpoint, POWERPOINT_ADDIN_NAME, POWERPOINT_VBA_ENTRY),
    ] {
        let bytes = validate_compiled_addin(path, name, vba_entry)?;
        let expected = manifest
            .files
            .get(name)
            .ok_or_else(|| format!("Compiled add-in manifest is missing {name}"))?;
        if expected.sha256.len() != 64
            || !expected
                .sha256
                .bytes()
                .all(|value| value.is_ascii_hexdigit())
            || !expected.sha256.eq_ignore_ascii_case(&sha256_hex(&bytes))
        {
            return Err(format!("Compiled add-in checksum failed for {name}"));
        }
    }
    Ok((word, powerpoint))
}

pub(super) fn source_artifacts(root: &Path) -> (PathBuf, PathBuf) {
    (
        root.join("resources/VisualTeX.dotm"),
        root.join("resources/VisualTeX.ppam"),
    )
}

pub(super) fn compiled_artifacts_available(root: &Path) -> bool {
    validate_compiled_artifacts(root).is_ok()
}
