use serde::Serialize;

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SystemMathFontProbe {
    requested_family: String,
    resolved_family: String,
    available: bool,
}

#[derive(Debug, Clone, Copy, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SystemMathGlyphMetrics {
    width_em: f64,
    ascent_em: f64,
    descent_em: f64,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct SystemMathGlyphOutline {
    character: String,
    requested_family: String,
    resolved_family: String,
    fallback_used: bool,
    glyph_id: u16,
    path: String,
    metrics: SystemMathGlyphMetrics,
}

const DEFAULT_MATH_FONT_FALLBACKS: &[&str] = &[
    "Cambria Math",
    "STIX Two Math",
    "Latin Modern Math",
    "XITS Math",
    "Segoe UI Symbol",
    "Times New Roman",
];

fn normalize_font_family(value: &str) -> String {
    value
        .chars()
        .filter(|character| !character.is_whitespace() && *character != '-' && *character != '_')
        .flat_map(char::to_lowercase)
        .collect()
}

fn validate_font_family(value: &str) -> Result<String, String> {
    let trimmed = value.trim();
    if trimmed.is_empty()
        || trimmed.chars().count() > 128
        || trimmed
            .chars()
            .any(|character| character.is_control() || matches!(character, '"' | '\'' | '<' | '>'))
    {
        return Err("The requested system math font family is invalid.".to_string());
    }
    Ok(trimmed.to_string())
}

fn validate_character(value: &str) -> Result<String, String> {
    let normalized = value.trim();
    if normalized.chars().count() != 1
        || normalized.chars().any(char::is_control)
        || normalized.encode_utf16().count() > 2
    {
        return Err("The requested system math glyph is invalid.".to_string());
    }
    Ok(normalized.to_string())
}

#[cfg(target_os = "windows")]
mod windows {
    use super::*;
    use std::collections::HashSet;
    use std::env;
    use std::fs;
    use std::path::{Path, PathBuf};
    use std::sync::OnceLock;
    use ttf_parser::{Face, OutlineBuilder};

    const MAX_FONT_FILE_BYTES: u64 = 128 * 1024 * 1024;
    const MAX_OUTLINE_BYTES: usize = 240_000;

    #[derive(Debug, Clone)]
    struct FontFaceEntry {
        path: PathBuf,
        face_index: u32,
        families: Vec<String>,
        display_family: String,
    }

    static FONT_CATALOG: OnceLock<Vec<FontFaceEntry>> = OnceLock::new();

    fn font_directories() -> Vec<PathBuf> {
        let mut directories = Vec::new();
        if let Some(windows_dir) = env::var_os("WINDIR") {
            directories.push(PathBuf::from(windows_dir).join("Fonts"));
        }
        if let Some(local_app_data) = env::var_os("LOCALAPPDATA") {
            directories.push(
                PathBuf::from(local_app_data)
                    .join("Microsoft")
                    .join("Windows")
                    .join("Fonts"),
            );
        }
        directories
    }

    fn supported_font_file(path: &Path) -> bool {
        path.extension()
            .and_then(|extension| extension.to_str())
            .is_some_and(|extension| {
                matches!(
                    extension.to_ascii_lowercase().as_str(),
                    "ttf" | "otf" | "ttc" | "otc"
                )
            })
    }

    fn face_family_names(face: &Face<'_>) -> Vec<String> {
        let mut names = Vec::new();
        for name in face.names() {
            // OpenType name IDs 1 and 16 are Family and Typographic Family.
            if name.name_id != 1 && name.name_id != 16 {
                continue;
            }
            let Some(value) = name.to_string() else {
                continue;
            };
            let value = value.trim();
            if value.is_empty() || value.chars().count() > 128 {
                continue;
            }
            if !names.iter().any(|existing: &String| {
                normalize_font_family(existing) == normalize_font_family(value)
            }) {
                names.push(value.to_string());
            }
        }
        names
    }

    fn scan_font_file(path: &Path, entries: &mut Vec<FontFaceEntry>) {
        let Ok(metadata) = fs::metadata(path) else {
            return;
        };
        if !metadata.is_file() || metadata.len() == 0 || metadata.len() > MAX_FONT_FILE_BYTES {
            return;
        }
        let Ok(data) = fs::read(path) else {
            return;
        };
        let face_count = ttf_parser::fonts_in_collection(&data).unwrap_or(1).max(1);
        for face_index in 0..face_count {
            let Ok(face) = Face::parse(&data, face_index) else {
                continue;
            };
            let families = face_family_names(&face);
            let Some(display_family) = families.first().cloned() else {
                continue;
            };
            entries.push(FontFaceEntry {
                path: path.to_path_buf(),
                face_index,
                families,
                display_family,
            });
        }
    }

    fn build_font_catalog() -> Vec<FontFaceEntry> {
        let mut entries = Vec::new();
        let mut seen_paths = HashSet::new();
        for directory in font_directories() {
            let Ok(children) = fs::read_dir(directory) else {
                continue;
            };
            for child in children.flatten() {
                let path = child.path();
                if !supported_font_file(&path) {
                    continue;
                }
                let canonical = fs::canonicalize(&path).unwrap_or_else(|_| path.clone());
                let identity = canonical.to_string_lossy().to_ascii_lowercase();
                if !seen_paths.insert(identity) {
                    continue;
                }
                scan_font_file(&canonical, &mut entries);
            }
        }
        entries
    }

    fn catalog() -> &'static [FontFaceEntry] {
        FONT_CATALOG.get_or_init(build_font_catalog).as_slice()
    }

    fn resolve_font(requested: &str) -> Option<FontFaceEntry> {
        let normalized = normalize_font_family(requested);
        catalog()
            .iter()
            .find(|entry| {
                entry
                    .families
                    .iter()
                    .any(|family| normalize_font_family(family) == normalized)
            })
            .cloned()
    }

    fn format_number(value: f64) -> String {
        let normalized = if value.abs() < 0.000_000_5 { 0.0 } else { value };
        let mut output = format!("{normalized:.5}");
        while output.contains('.') && output.ends_with('0') {
            output.pop();
        }
        if output.ends_with('.') {
            output.pop();
        }
        output
    }

    #[derive(Debug, Clone, Copy)]
    enum RawPathCommand {
        Move(f32, f32),
        Line(f32, f32),
        Quad(f32, f32, f32, f32),
        Cubic(f32, f32, f32, f32, f32, f32),
        Close,
    }

    #[derive(Default)]
    struct RawPathWriter {
        commands: Vec<RawPathCommand>,
    }

    impl OutlineBuilder for RawPathWriter {
        fn move_to(&mut self, x: f32, y: f32) {
            self.commands.push(RawPathCommand::Move(x, y));
        }

        fn line_to(&mut self, x: f32, y: f32) {
            self.commands.push(RawPathCommand::Line(x, y));
        }

        fn quad_to(&mut self, x1: f32, y1: f32, x: f32, y: f32) {
            self.commands.push(RawPathCommand::Quad(x1, y1, x, y));
        }

        fn curve_to(&mut self, x1: f32, y1: f32, x2: f32, y2: f32, x: f32, y: f32) {
            self.commands
                .push(RawPathCommand::Cubic(x1, y1, x2, y2, x, y));
        }

        fn close(&mut self) {
            self.commands.push(RawPathCommand::Close);
        }
    }

    fn write_path(
        commands: &[RawPathCommand],
        scale: f64,
        shift_x: f64,
        ascent: f64,
    ) -> String {
        let point = |x: f32, y: f32| {
            (
                format_number(f64::from(x) * scale + shift_x),
                format_number(ascent - f64::from(y) * scale),
            )
        };
        let mut output = String::with_capacity(commands.len().saturating_mul(24));
        for command in commands {
            match *command {
                RawPathCommand::Move(x, y) => {
                    let (x, y) = point(x, y);
                    output.push_str(&format!("M{x} {y}"));
                }
                RawPathCommand::Line(x, y) => {
                    let (x, y) = point(x, y);
                    output.push_str(&format!("L{x} {y}"));
                }
                RawPathCommand::Quad(x1, y1, x, y) => {
                    let (x1, y1) = point(x1, y1);
                    let (x, y) = point(x, y);
                    output.push_str(&format!("Q{x1} {y1} {x} {y}"));
                }
                RawPathCommand::Cubic(x1, y1, x2, y2, x, y) => {
                    let (x1, y1) = point(x1, y1);
                    let (x2, y2) = point(x2, y2);
                    let (x, y) = point(x, y);
                    output.push_str(&format!("C{x1} {y1} {x2} {y2} {x} {y}"));
                }
                RawPathCommand::Close => output.push('Z'),
            }
        }
        output
    }

    fn outline_from_entry(
        entry: FontFaceEntry,
        character: &str,
        primary_requested_family: &str,
        candidate_family: &str,
    ) -> Result<SystemMathGlyphOutline, String> {
        let data = fs::read(&entry.path).map_err(|error| {
            format!(
                "VisualTeX could not read the resolved font {}: {error}",
                entry.path.display()
            )
        })?;
        let face = Face::parse(&data, entry.face_index).map_err(|error| {
            format!(
                "VisualTeX could not parse the resolved font {}: {error:?}",
                entry.path.display()
            )
        })?;
        let scalar = character
            .chars()
            .next()
            .ok_or_else(|| "The requested system math glyph is empty.".to_string())?;
        let glyph = face.glyph_index(scalar).ok_or_else(|| {
            format!(
                "{} does not contain the requested mathematical glyph {}.",
                entry.display_family, character
            )
        })?;
        let mut writer = RawPathWriter::default();
        let bounds = face.outline_glyph(glyph, &mut writer).ok_or_else(|| {
            format!(
                "{} could not expose a vector outline for {}.",
                entry.display_family, character
            )
        })?;
        if writer.commands.is_empty() {
            return Err(format!(
                "{} returned an empty vector outline for {}.",
                entry.display_family, character
            ));
        }

        let units_per_em = f64::from(face.units_per_em());
        if !units_per_em.is_finite() || units_per_em <= 0.0 {
            return Err("The resolved system math font has invalid units per em.".to_string());
        }
        let scale = 1000.0 / units_per_em;
        let min_x = f64::from(bounds.x_min) * scale;
        let min_y = f64::from(bounds.y_min) * scale;
        let max_x = f64::from(bounds.x_max) * scale;
        let max_y = f64::from(bounds.y_max) * scale;
        if ![min_x, min_y, max_x, max_y].into_iter().all(f64::is_finite)
            || max_x <= min_x
            || max_y <= min_y
        {
            return Err(format!(
                "{} returned invalid vector bounds for {}.",
                entry.display_family, character
            ));
        }

        let advance = face
            .glyph_hor_advance(glyph)
            .map(f64::from)
            .unwrap_or_else(|| f64::from(bounds.x_max - bounds.x_min))
            * scale;
        let shift_x = -min_x.min(0.0);
        let ascent = max_y.max(20.0);
        let descent = (-min_y).max(0.0);
        let width = (advance + shift_x).max(max_x + shift_x).max(20.0);
        let path = write_path(&writer.commands, scale, shift_x, ascent);
        if path.is_empty() || path.len() > MAX_OUTLINE_BYTES {
            return Err("The system mathematical glyph outline is empty or too complex.".to_string());
        }

        Ok(SystemMathGlyphOutline {
            character: character.to_string(),
            requested_family: primary_requested_family.to_string(),
            resolved_family: entry.display_family,
            fallback_used: normalize_font_family(primary_requested_family)
                != normalize_font_family(candidate_family),
            glyph_id: glyph.0,
            path,
            metrics: SystemMathGlyphMetrics {
                width_em: width / 1000.0,
                ascent_em: ascent / 1000.0,
                descent_em: descent / 1000.0,
            },
        })
    }

    pub(super) fn probe_fonts(font_families: &[String]) -> Vec<SystemMathFontProbe> {
        font_families
            .iter()
            .filter_map(|requested| {
                let requested_family = validate_font_family(requested).ok()?;
                match resolve_font(&requested_family) {
                    Some(font) => Some(SystemMathFontProbe {
                        requested_family,
                        resolved_family: font.display_family,
                        available: true,
                    }),
                    None => Some(SystemMathFontProbe {
                        requested_family,
                        resolved_family: String::new(),
                        available: false,
                    }),
                }
            })
            .collect()
    }

    pub(super) fn extract_outline(
        requested_families: &[String],
        character: &str,
    ) -> Result<SystemMathGlyphOutline, String> {
        let character = validate_character(character)?;
        let mut candidates = Vec::new();
        for family in requested_families {
            let family = validate_font_family(family)?;
            if !candidates.iter().any(|candidate: &String| {
                normalize_font_family(candidate) == normalize_font_family(&family)
            }) {
                candidates.push(family);
            }
        }
        for fallback in DEFAULT_MATH_FONT_FALLBACKS {
            if !candidates.iter().any(|candidate| {
                normalize_font_family(candidate) == normalize_font_family(fallback)
            }) {
                candidates.push((*fallback).to_string());
            }
        }
        let primary = candidates
            .first()
            .cloned()
            .ok_or_else(|| "No system math font family was requested.".to_string())?;
        let mut errors = Vec::new();
        for candidate in candidates {
            let Some(font) = resolve_font(&candidate) else {
                errors.push(format!("Windows could not find the requested family {candidate}."));
                continue;
            };
            match outline_from_entry(font, &character, &primary, &candidate) {
                Ok(outline) => return Ok(outline),
                Err(error) => errors.push(error),
            }
        }
        Err(format!(
            "VisualTeX could not find a Windows math font containing {}. {}",
            character,
            errors.join(" ")
        ))
    }
}

#[tauri::command]
pub(crate) fn probe_system_math_fonts(
    font_families: Vec<String>,
) -> Result<Vec<SystemMathFontProbe>, String> {
    if font_families.len() > 32 {
        return Err("Too many system math font families were requested.".to_string());
    }
    #[cfg(target_os = "windows")]
    {
        return Ok(windows::probe_fonts(&font_families));
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = font_families;
        Err("System mathematical glyph extraction is unavailable on this platform.".to_string())
    }
}

#[tauri::command]
pub(crate) fn extract_system_math_glyph(
    font_families: Vec<String>,
    character: String,
) -> Result<SystemMathGlyphOutline, String> {
    if font_families.is_empty() || font_families.len() > 16 {
        return Err("A bounded system math font fallback list is required.".to_string());
    }
    #[cfg(target_os = "windows")]
    {
        return windows::extract_outline(&font_families, &character);
    }
    #[cfg(not(target_os = "windows"))]
    {
        let _ = (font_families, character);
        Err("System mathematical glyph extraction is unavailable on this platform.".to_string())
    }
}

#[cfg(all(test, target_os = "windows"))]
mod tests {
    use super::*;

    #[test]
    fn rejects_multiple_unicode_scalars_for_one_glyph_request() {
        assert!(validate_character("AB").is_err());
        assert!(validate_character("αβ").is_err());
        assert_eq!(validate_character("𝛼").as_deref(), Ok("𝛼"));
    }

    #[test]
    fn rejects_a_missing_font_without_accepting_substitution() {
        let probes = windows::probe_fonts(&[
            "VisualTeX Definitely Missing Math Font".to_string(),
        ]);
        assert_eq!(probes.len(), 1);
        assert!(!probes[0].available);
    }

    #[test]
    fn extracts_a_crisp_cambria_math_outline_when_available() {
        let probes = windows::probe_fonts(&["Cambria Math".to_string()]);
        if !probes.first().is_some_and(|probe| probe.available) {
            return;
        }
        let outline = windows::extract_outline(&["Cambria Math".to_string()], "∫")
            .expect("Cambria Math should expose an integral vector outline");
        assert_eq!(normalize_font_family(&outline.resolved_family), "cambriamath");
        assert!(outline.path.starts_with('M'));
        assert!(!outline.path.contains("NaN"));
        assert!(outline.metrics.width_em > 0.05);
        assert!(outline.metrics.ascent_em > 0.1);
    }
}
