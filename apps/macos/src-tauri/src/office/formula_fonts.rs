//! The fonts VisualTeX draws formulas in now - the math font (KaTeX or newtxmath, stored as "times"), the Chinese
//! font and the fonts for letters, Greek letters and digits over the math font - and whether a formula was drawn in
//! others: Update Equation Format draws those again from their LaTeX.
//!
//! The fonts are the Settings' in the app's web storage, which only its windows read: the main window and the
//! resident Office editors publish them at start and after every change (src/office/shared/formulaFontPublisher.ts);
//! they are kept in `office/formula-fonts.json` of the app's data folder, so a resident started later knows them
//! before its windows load. The Office add-ins ask through the resident's inbox (macos_offline::fast_open,
//! `<session>.fonts`) and the WPS bridge (`/formulas/font-check`). The rule is the shared formulaFontsDiffer's
//! (shared/visualtex-app/src/office/shared/formulaFontRefresh.ts) and the Windows add-ins' FormulaFontSettings.

use std::path::PathBuf;
use std::sync::RwLock;

use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};
use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::office::sessions::VisualTeXFormulaMetadata;

const FONTS_FILE: &str = "Library/Application Support/com.visualtex.studio/office/formula-fonts.json";
const MAX_FACE_CHARS: usize = 128;
const FACE_ROLES: [&str; 3] = ["letter", "greek", "digit"];
/// The Office add-ins check one document's picture formulas at a time.
pub(crate) const MAX_CHECKED_FORMULAS: usize = 4000;

/// The faces: a font name for each role, "" for the math font's.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormulaFaces {
    #[serde(default)]
    pub letter: String,
    #[serde(default)]
    pub greek: String,
    #[serde(default)]
    pub digit: String,
}

impl FormulaFaces {
    fn role(&self, role: &str) -> &str {
        match role {
            "letter" => &self.letter,
            "greek" => &self.greek,
            _ => &self.digit,
        }
    }
}

/// The fonts formulas are drawn in now.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct FormulaFonts {
    pub letter_font: String,
    pub chinese_font: String,
    #[serde(default)]
    pub faces: FormulaFaces,
}

impl FormulaFonts {
    fn validate(&self) -> Result<(), String> {
        if !matches!(self.letter_font.as_str(), "katex" | "times") {
            return Err("The formula math font must be katex or times".to_string());
        }
        if self.chinese_font.is_empty()
            || self.chinese_font.len() > 64
            || !self.chinese_font.bytes().all(|byte| byte.is_ascii_alphanumeric())
        {
            return Err("The formula Chinese font is invalid".to_string());
        }
        for role in FACE_ROLES {
            let name = self.faces.role(role);
            if name.chars().count() > MAX_FACE_CHARS || name.chars().any(char::is_control) {
                return Err(format!("The formula {role} font name is invalid"));
            }
        }
        Ok(())
    }
}

static CURRENT: RwLock<Option<FormulaFonts>> = RwLock::new(None);

fn fonts_path() -> Option<PathBuf> {
    let home = std::env::var_os("HOME").map(PathBuf::from).filter(|path| path.is_absolute())?;
    Some(home.join(FONTS_FILE))
}

fn read_fonts_file() -> Option<FormulaFonts> {
    let bytes = std::fs::read(fonts_path()?).ok().filter(|bytes| bytes.len() <= 16 * 1024)?;
    let fonts: FormulaFonts = serde_json::from_slice(&bytes).ok()?;
    fonts.validate().ok().map(|_| fonts)
}

/// The fonts formulas are drawn in now: the ones published last, or the ones kept from an earlier run; None when no
/// window has published them yet.
pub fn current() -> Option<FormulaFonts> {
    if let Some(fonts) = CURRENT.read().ok().and_then(|current| current.clone()) {
        return Some(fonts);
    }
    let fonts = read_fonts_file()?;
    if let Ok(mut current) = CURRENT.write() {
        current.get_or_insert_with(|| fonts.clone());
    }
    Some(fonts)
}

/// Takes the fonts a window published and keeps them for the next run.
pub fn publish(fonts: FormulaFonts) -> Result<(), String> {
    let fonts = FormulaFonts {
        letter_font: fonts.letter_font.trim().to_string(),
        chinese_font: fonts.chinese_font.trim().to_string(),
        faces: FormulaFaces {
            letter: fonts.faces.letter.trim().to_string(),
            greek: fonts.faces.greek.trim().to_string(),
            digit: fonts.faces.digit.trim().to_string(),
        },
    };
    fonts.validate()?;
    if read_fonts_file().as_ref() != Some(&fonts) {
        let path = fonts_path().ok_or("Unable to resolve the VisualTeX data folder")?;
        if let Some(parent) = path.parent() {
            std::fs::create_dir_all(parent).map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
        }
        let json = serde_json::to_vec(&fonts).map_err(|error| error.to_string())?;
        crate::office::macos_offline::atomic_write_runtime(&path, &json, 0o600)?;
    }
    if let Ok(mut current) = CURRENT.write() {
        *current = Some(fonts);
    }
    Ok(())
}

#[tauri::command]
pub fn macos_office_publish_formula_fonts(fonts: FormulaFonts) -> Result<(), String> {
    publish(fonts)
}

/// A formula's record of its faces as the metadata of either app writes it: an object whose roles are names of at
/// most 128 characters (or null).
pub(crate) fn valid_faces_record(value: &Value) -> bool {
    let Some(object) = value.as_object() else {
        return false;
    };
    FACE_ROLES.iter().all(|role| match object.get(*role) {
        None | Some(Value::Null) => true,
        Some(Value::String(name)) => name.chars().count() <= MAX_FACE_CHARS,
        Some(_) => false,
    })
}

fn same_name(recorded: &str, current: &str) -> bool {
    recorded.trim().to_lowercase() == current.trim().to_lowercase()
}

/// CJK ideographs, compatibility ideographs, CJK punctuation and full-width forms: what the Chinese font draws.
fn contains_chinese(latex: &str) -> bool {
    latex.chars().any(|character| {
        matches!(character, '\u{3400}'..='\u{9FFF}' | '\u{F900}'..='\u{FAFF}' | '\u{3000}'..='\u{303F}' | '\u{FF00}'..='\u{FFEF}')
    })
}

/// Whether the formula was drawn in other fonts than `current`. A record the formula lacks counts as what VisualTeX
/// drew without it: KaTeX, the system Chinese font, no faces. The Chinese font counts only for a formula with
/// Chinese text, which is all it changes.
pub fn differ(metadata: &VisualTeXFormulaMetadata, current: &FormulaFonts) -> bool {
    let letter = metadata
        .formula_letter_font
        .as_deref()
        .map(str::trim)
        .filter(|font| !font.is_empty())
        .unwrap_or("katex");
    if !same_name(letter, &current.letter_font) {
        return true;
    }
    let recorded_face = |role: &str| {
        metadata
            .formula_faces
            .as_ref()
            .and_then(|faces| faces.get(role))
            .and_then(Value::as_str)
            .unwrap_or("")
            .to_string()
    };
    if FACE_ROLES.iter().any(|role| !same_name(&recorded_face(*role), current.faces.role(*role))) {
        return true;
    }
    let chinese = metadata
        .formula_chinese_font
        .as_deref()
        .map(str::trim)
        .filter(|font| !font.is_empty())
        .unwrap_or("system");
    !same_name(chinese, &current.chinese_font) && contains_chinese(&metadata.latex)
}

/// The fonts in a few words, for the add-ins' confirmations: "KaTeX; letters, Greek letters, digits: …".
pub fn describe(current: &FormulaFonts, english: bool) -> String {
    let math = if current.letter_font == "times" { "newtxmath" } else { "KaTeX" };
    let names = [("letter", "字母", "letters"), ("greek", "希腊字母", "Greek letters"), ("digit", "数字", "digits")];
    let parts: Vec<String> = names
        .iter()
        .filter(|(role, _, _)| !current.faces.role(role).trim().is_empty())
        .map(|(role, zh, en)| {
            let name = current.faces.role(role).trim();
            if english { format!("{en}: {name}") } else { format!("{zh}：{name}") }
        })
        .collect();
    if parts.is_empty() {
        math.to_string()
    } else if english {
        format!("{math}; {}", parts.join(", "))
    } else {
        format!("{math}；{}", parts.join("，"))
    }
}

/// Which of the encoded metadata strings name formulas drawn in other fonts than `current` (one that cannot be read
/// is not a VisualTeX formula this can redraw).
pub fn stale_formulas(metadata: &[String], current: &FormulaFonts) -> Vec<bool> {
    metadata
        .iter()
        .map(|encoded| {
            crate::office::macos_offline::decode_metadata(encoded)
                .map(|decoded| !decoded.latex.trim().is_empty() && differ(&decoded, current))
                .unwrap_or(false)
        })
        .collect()
}

/// The answer to an Office add-in's font check (`<session>.fonts`): the manifest rows `protocolVersion`, `sessionId`,
/// `imageCount` and `image<n>Metadata` (n from 1); the answer rows `operation=fontCheck`, `staleCount`, `stale` (the
/// 1-based numbers of the formulas drawn in other fonts, comma separated) and `fontsEn` / `fontsZh` (Base64URL of
/// the description), or `error=<message>`.
pub fn answer_font_check(manifest: &str) -> String {
    answer_font_check_with(manifest, current())
}

fn answer_font_check_with(manifest: &str, current: Option<FormulaFonts>) -> String {
    let result = (|| -> Result<String, String> {
        let mut values = std::collections::HashMap::new();
        for row in manifest.lines().filter(|row| !row.is_empty()) {
            let (key, value) = row.split_once('=').ok_or("The font check contains an invalid row")?;
            if key.is_empty() || !key.bytes().all(|byte| byte.is_ascii_alphanumeric()) || values.insert(key, value).is_some() {
                return Err("The font check contains an invalid key".to_string());
            }
        }
        if values.get("protocolVersion").copied() != Some("1") {
            return Err("The font check protocol version is invalid".to_string());
        }
        let session_id = values.get("sessionId").copied().ok_or("The font check names no Session")?;
        if session_id.len() != 36 || uuid::Uuid::parse_str(session_id).is_err() {
            return Err("The font check Session id is invalid".to_string());
        }
        let count = values
            .get("imageCount")
            .and_then(|value| value.parse::<usize>().ok())
            .filter(|count| (1..=MAX_CHECKED_FORMULAS).contains(count))
            .ok_or("The font check must name 1 to 4000 formulas")?;
        let metadata = (1..=count)
            .map(|index| {
                values
                    .get(format!("image{index}Metadata").as_str())
                    .map(|value| value.to_string())
                    .ok_or_else(|| format!("The font check is missing formula {index}"))
            })
            .collect::<Result<Vec<_>, _>>()?;
        let current = current.ok_or("VisualTeX has not read its formula fonts yet")?;
        let stale = stale_formulas(&metadata, &current);
        let numbers: Vec<String> = stale
            .iter()
            .enumerate()
            .filter(|(_, stale)| **stale)
            .map(|(index, _)| (index + 1).to_string())
            .collect();
        Ok(format!(
            "protocolVersion=1\nsessionId={session_id}\noperation=fontCheck\nimageCount={count}\nstaleCount={}\nstale={}\nfontsEn={}\nfontsZh={}\n",
            numbers.len(),
            numbers.join(","),
            URL_SAFE_NO_PAD.encode(describe(&current, true).as_bytes()),
            URL_SAFE_NO_PAD.encode(describe(&current, false).as_bytes()),
        ))
    })();
    result.unwrap_or_else(|error| format!("error={}\n", error.replace(['\r', '\n'], " ")))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn metadata(latex: &str, letter: Option<&str>, chinese: Option<&str>, faces: Option<Value>) -> VisualTeXFormulaMetadata {
        serde_json::from_value(serde_json::json!({
            "schema": "visualtex-formula",
            "schemaVersion": 1,
            "formulaId": "12345678-1234-4234-9234-123456789abc",
            "title": "Formula",
            "latex": latex,
            "lines": [{ "id": "12345678-1234-4234-9234-123456789abd", "latex": latex }],
            "codeFormat": "raw",
            "displayMode": "block",
            "formulaLetterFont": letter,
            "formulaChineseFont": chinese,
            "formulaFaces": faces,
            "createdWithVersion": "1",
            "updatedWithVersion": "1",
            "createdAt": "1",
            "updatedAt": "1",
        }))
        .unwrap()
    }

    fn fonts(letter: &str, chinese: &str, faces: [&str; 3]) -> FormulaFonts {
        FormulaFonts {
            letter_font: letter.to_string(),
            chinese_font: chinese.to_string(),
            faces: FormulaFaces { letter: faces[0].to_string(), greek: faces[1].to_string(), digit: faces[2].to_string() },
        }
    }

    #[test]
    fn a_missing_record_is_what_visualtex_drew_without_it() {
        let plain = fonts("katex", "system", ["", "", ""]);
        assert!(!differ(&metadata("x+y", None, None, None), &plain));
        assert!(differ(&metadata("x+y", None, None, None), &fonts("times", "system", ["", "", ""])));
        let times_faces = fonts("katex", "system", ["Times New Roman", "Times New Roman", "Times New Roman"]);
        assert!(differ(&metadata("x+y", Some("katex"), Some("system"), None), &times_faces));
        let recorded = serde_json::json!({ "letter": "times new roman ", "greek": "Times New Roman", "digit": "Times New Roman" });
        assert!(!differ(&metadata("x+y", Some("KaTeX"), None, Some(recorded)), &times_faces));
        let partial = serde_json::json!({ "letter": "Times New Roman" });
        assert!(differ(&metadata("x+y", None, None, Some(partial)), &times_faces));
        // an earlier version's letter font is drawn in KaTeX now: drawn again
        assert!(differ(&metadata("x", Some("cambria"), None, None), &plain));
    }

    #[test]
    fn the_chinese_font_counts_only_for_chinese_text() {
        let kaiti = fonts("katex", "kaiti", ["", "", ""]);
        assert!(!differ(&metadata("x^2", Some("katex"), Some("system"), None), &kaiti));
        assert!(differ(&metadata("\\text{速度}=v", Some("katex"), Some("system"), None), &kaiti));
        assert!(differ(&metadata("a\u{FF0C}b", Some("katex"), None, None), &kaiti));
        assert!(!differ(&metadata("\\text{速度}", Some("katex"), Some("kaiti"), None), &kaiti));
    }

    #[test]
    fn descriptions_name_the_math_font_and_the_faces() {
        assert_eq!(describe(&fonts("katex", "system", ["", "", ""]), true), "KaTeX");
        assert_eq!(
            describe(&fonts("times", "system", ["Times New Roman", "", "Arial"]), true),
            "newtxmath; letters: Times New Roman, digits: Arial"
        );
        assert_eq!(
            describe(&fonts("katex", "system", ["Times New Roman", "Times New Roman", ""]), false),
            "KaTeX；字母：Times New Roman，希腊字母：Times New Roman"
        );
    }

    #[test]
    fn faces_records_are_objects_of_short_names() {
        assert!(valid_faces_record(&serde_json::json!({ "letter": "Times New Roman", "greek": "", "digit": null })));
        assert!(valid_faces_record(&serde_json::json!({})));
        assert!(!valid_faces_record(&serde_json::json!("Times")));
        assert!(!valid_faces_record(&serde_json::json!({ "letter": 3 })));
        assert!(!valid_faces_record(&serde_json::json!({ "digit": "x".repeat(129) })));
    }

    #[test]
    fn published_fonts_are_validated() {
        assert!(fonts("katex", "system", ["", "", ""]).validate().is_ok());
        assert!(fonts("cambria", "system", ["", "", ""]).validate().is_err());
        assert!(fonts("katex", "", ["", "", ""]).validate().is_err());
        assert!(fonts("katex", "system", ["a\nb", "", ""]).validate().is_err());
    }

    #[test]
    fn a_font_check_answers_with_the_stale_numbers() {
        let plain = Some(fonts("katex", "system", ["", "", ""]));
        let session = "sessionId=12345678-1234-4234-9234-123456789abc";
        assert!(answer_font_check_with("protocolVersion=1\nsessionId=not-a-uuid\nimageCount=1\nimage1Metadata=x\n", plain.clone())
            .starts_with("error="));
        assert!(answer_font_check_with(&format!("protocolVersion=1\n{session}\nimageCount=0\n"), plain.clone()).starts_with("error="));
        assert!(answer_font_check_with(&format!("protocolVersion=1\n{session}\nimageCount=2\nimage1Metadata=x\n"), plain.clone())
            .starts_with("error="));
        let katex = crate::office::macos_offline::encode_metadata(&metadata("x+y", Some("katex"), Some("system"), None)).unwrap();
        let times = crate::office::macos_offline::encode_metadata(&metadata("x+y", Some("times"), Some("system"), None)).unwrap();
        let manifest = format!(
            "protocolVersion=1\n{session}\nimageCount=3\nimage1Metadata={katex}\nimage2Metadata=broken\nimage3Metadata={times}\n"
        );
        assert!(answer_font_check_with(&manifest, None).starts_with("error="));
        let answer = answer_font_check_with(&manifest, plain);
        assert!(answer.contains("operation=fontCheck\n"), "{answer}");
        assert!(answer.contains("\nstaleCount=1\nstale=3\n"), "{answer}");
        assert!(answer.contains(&format!("\nfontsEn={}\n", URL_SAFE_NO_PAD.encode("KaTeX"))), "{answer}");
        let answer = answer_font_check_with(&manifest, Some(fonts("times", "system", ["", "", ""])));
        assert!(answer.contains("\nstaleCount=1\nstale=1\n"), "{answer}");
    }
}
