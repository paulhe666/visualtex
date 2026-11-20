//! The math fonts of Word's OMML font menu, installed into the user's font folder with the add-ins.

use super::*;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum MathFontInstallOutcome {
    /// The font folder already holds the same bytes; nothing was written.
    Current,
    /// The font was missing and is now installed.
    Installed,
    /// A file of the same name held other bytes (an earlier build); replaced.
    Replaced,
}

impl MathFontInstallOutcome {
    pub(super) fn label(self) -> &'static str {
        match self {
            Self::Current => "already installed",
            Self::Installed => "installed",
            Self::Replaced => "replaced an older copy",
        }
    }
}

/// The bundled math fonts with the SHA-256 that addins.json pins for each.
pub(super) fn bundled_math_fonts(root: &Path) -> Result<Vec<(PathBuf, String)>, String> {
    let manifest = read_addin_manifest(root)?;
    BUNDLED_MATH_FONTS
        .iter()
        .map(|name| {
            let key = format!("{BUNDLED_MATH_FONT_DIRECTORY}/{name}");
            let sha256 = manifest
                .files
                .get(&key)
                .map(|file| file.sha256.to_ascii_lowercase())
                .filter(|value| {
                    value.len() == 64 && value.bytes().all(|byte| byte.is_ascii_hexdigit())
                })
                .ok_or_else(|| format!("The add-in manifest does not pin {key}"))?;
            Ok((root.join(BUNDLED_MATH_FONT_DIRECTORY).join(name), sha256))
        })
        .collect()
}

/// Copies one bundled font into `fonts_directory` unless the file of its name
/// there already has its bytes. Only that one file name is ever written; the
/// user's other fonts are never touched.
pub(super) fn install_math_font(
    source: &Path,
    expected_sha256: &str,
    fonts_directory: &Path,
) -> Result<MathFontInstallOutcome, String> {
    let name = source
        .file_name()
        .ok_or_else(|| format!("Bundled font {} has no file name", source.display()))?;
    let bytes = fs::read(source)
        .map_err(|error| format!("Bundled font {} is missing: {error}", source.display()))?;
    if sha256_hex(&bytes) != expected_sha256 {
        return Err(format!(
            "Bundled font {} failed its SHA-256 check",
            source.display()
        ));
    }
    let destination = fonts_directory.join(name);
    let existing = match fs::read(&destination) {
        Ok(existing) => Some(existing),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => None,
        Err(error) => return Err(format!("Unable to read {}: {error}", destination.display())),
    };
    if existing
        .as_deref()
        .is_some_and(|existing| sha256_hex(existing) == expected_sha256)
    {
        return Ok(MathFontInstallOutcome::Current);
    }
    atomic_write(&destination, &bytes, 0o644)?;
    let installed = fs::read(&destination)
        .map_err(|error| format!("Unable to verify {}: {error}", destination.display()))?;
    if sha256_hex(&installed) != expected_sha256 {
        return Err(format!(
            "Installed font {} does not match its VisualTeX source",
            destination.display()
        ));
    }
    Ok(if existing.is_some() {
        MathFontInstallOutcome::Replaced
    } else {
        MathFontInstallOutcome::Installed
    })
}

/// Every bundled math font, installed into `fonts_directory`, with its result.
pub(super) fn install_math_fonts_into(
    root: &Path,
    fonts_directory: &Path,
) -> Result<Vec<(String, Result<MathFontInstallOutcome, String>)>, String> {
    Ok(bundled_math_fonts(root)?
        .into_iter()
        .map(|(source, sha256)| {
            let name = source
                .file_name()
                .map(|value| value.to_string_lossy().into_owned())
                .unwrap_or_default();
            (name, install_math_font(&source, &sha256, fonts_directory))
        })
        .collect())
}

/// Installs the fonts of Word's OMML font menu into ~/Library/Fonts, where
/// macOS activates them for every app (Word lists a new font once restarted).
/// Never fails the add-in operation: every result goes to the maintenance log.
pub(super) fn install_bundled_math_fonts(root: &Path, operation: &str) {
    append_office_maintenance_log(operation, "install-math-fonts", "begin", "");
    let (failed, detail) = match user_home()
        .map(|home| home.join(USER_FONT_DIRECTORY))
        .and_then(|fonts_directory| {
            install_math_fonts_into(root, &fonts_directory)
                .map(|results| (fonts_directory, results))
        }) {
        Ok((fonts_directory, results)) => {
            let mut failed = false;
            let mut detail = format!("fonts_directory={}", fonts_directory.display());
            for (name, result) in results {
                let outcome = match result {
                    Ok(outcome) => outcome.label().to_string(),
                    Err(error) => {
                        failed = true;
                        format!("failed: {error}")
                    }
                };
                detail.push_str(&format!("\n{name}: {outcome}"));
            }
            (failed, detail)
        }
        Err(error) => (true, error),
    };
    append_office_maintenance_log(
        operation,
        "install-math-fonts",
        if failed { "error" } else { "ok" },
        &detail,
    );
}

/// The same fonts with the WPS add-in: WPS's equation font menu offers them too (office::wps::bridge). Only writes
/// a font that is missing or older; the result goes to the maintenance log, never to the add-in installation.
pub(crate) fn install_math_fonts_with_wps_addin(app: &AppHandle) {
    if !cfg!(target_os = "macos") {
        return;
    }
    match resource_root(app) {
        Ok(root) => install_bundled_math_fonts(&root, "wps-install"),
        Err(error) => append_office_maintenance_log("wps-install", "install-math-fonts", "error", &error),
    }
}
