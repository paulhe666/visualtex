use serde::Serialize;
use std::process::Command;

const WORD_METADATA_PREFIX: &str = "visualtex:v1:deflate:";
const WORD_FIELD_SEPARATOR: &str = "<VISUALTEX_WORD_FIELD>";

const APPLY_INLINE_BASELINE_SCRIPT: &str = r#"
on run argv
    if (count of argv) is not 2 then error "Expected a Word font position and formula marker"
    -- Never pass a bare negative number to osascript: values such as "-5"
    -- are parsed by osascript itself as command-line options before AppleScript
    -- receives argv. Named arguments keep both values intact.
    set rawPositionArgument to (item 1 of argv) as text
    if rawPositionArgument does not start with "position=" then error "Invalid Word font position argument"
    set requestedPosition to (text 10 thru -1 of rawPositionArgument) as integer
    set rawMarkerArgument to (item 2 of argv) as text
    if rawMarkerArgument does not start with "marker=" then error "Invalid Word formula marker argument"
    set expectedMarker to text 8 thru -1 of rawMarkerArgument

    tell application "Microsoft Word"
        if not (exists active document) then error "Microsoft Word has no active document"
        set matchingShapeIndex to 0
        set matchingShapeCount to 0
        -- Do not keep Word object references inside an AppleScript list. Word
        -- 16.89 can coerce that list back into the live `every inline shape`
        -- collection, after which `count` is sent to the collection object and
        -- fails with -1708. Track only integer indexes and resolve the durable
        -- picture after the search has completed.
        --
        -- Office.js has already synchronized the picture, but Word for Mac can
        -- publish its native inline-shape collection a fraction later. Retry a
        -- bounded number of times and identify the exact formula by metadata;
        -- never depend on the insertion point remaining beside the picture.
        repeat with attemptIndex from 1 to 12
            set matchingShapeIndex to 0
            set matchingShapeCount to 0
            set documentShapeCount to count of inline shapes of active document
            repeat with shapeIndex from 1 to documentShapeCount
                set candidateShape to inline shape shapeIndex of active document
                try
                    if ((alternative text of candidateShape) as text) is expectedMarker then
                        set matchingShapeIndex to shapeIndex
                        set matchingShapeCount to matchingShapeCount + 1
                    end if
                end try
            end repeat
            if matchingShapeCount is 1 then exit repeat
            if matchingShapeCount is greater than 1 then error "Microsoft Word contains duplicate VisualTeX formula markers"
            delay 0.08
        end repeat
        if matchingShapeCount is not 1 then error "Microsoft Word could not locate the VisualTeX formula written by Office.js"

        set formulaPicture to inline shape matchingShapeIndex of active document
        set formulaFont to font object of text object of formulaPicture
        set font position of formulaFont to requestedPosition
        set appliedPosition to font position of formulaFont
        if appliedPosition is not requestedPosition then error "Microsoft Word did not persist the requested baseline position"
        -- Word for Mac can coerce a control-character separator to `missing value`
        -- when the expression is evaluated inside its application dictionary.
        -- A fixed ASCII token is unambiguous here because every returned field
        -- is numeric and the metadata marker validation excludes `<` and `>`.
        set fieldSeparator to "<VISUALTEX_WORD_FIELD>"
        return (appliedPosition as text) & fieldSeparator & (width of formulaPicture as text) & fieldSeparator & (height of formulaPicture as text) & fieldSeparator & (matchingShapeIndex as text)
    end tell
end run
"#;

#[derive(Debug, Clone, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct WordInlineBaselineResult {
    pub applied_position: i32,
    pub width: f64,
    pub height: f64,
    pub matched_shape_index: u32,
}

fn validate_position(position: i32) -> Result<i32, String> {
    // Word's Font.Position is an integer point offset. This range is far wider
    // than any formula can legitimately require while preventing malformed
    // requests from applying extreme document formatting.
    if !(-256..=0).contains(&position) {
        return Err("Word inline baseline position must be between -256 and 0 points".to_string());
    }
    Ok(position)
}

fn validate_formula_marker(marker: &str) -> Result<&str, String> {
    if !marker.starts_with(WORD_METADATA_PREFIX)
        || marker.len() > 32 * 1024
        || !marker
            .bytes()
            .all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b':' | b'-' | b'_'))
    {
        return Err("Word formula marker is invalid".to_string());
    }
    Ok(marker)
}

fn position_argument(position: i32) -> String {
    format!("position={position}")
}

fn marker_argument(marker: &str) -> String {
    format!("marker={marker}")
}

pub fn apply_inline_baseline(
    position: i32,
    formula_marker: &str,
) -> Result<WordInlineBaselineResult, String> {
    let position = validate_position(position)?;
    let formula_marker = validate_formula_marker(formula_marker)?;
    let output = Command::new("/usr/bin/osascript")
        .arg("-e")
        .arg(APPLY_INLINE_BASELINE_SCRIPT)
        .arg(position_argument(position))
        .arg(marker_argument(formula_marker))
        .output()
        .map_err(|error| format!("Unable to launch Microsoft Word AppleScript: {error}"))?;
    if !output.status.success() {
        let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
        return Err(if detail.is_empty() {
            "Microsoft Word rejected the native inline baseline update".to_string()
        } else {
            format!("Microsoft Word rejected the native inline baseline update: {detail}")
        });
    }
    let output = String::from_utf8_lossy(&output.stdout);
    let fields = output
        .trim()
        .split(WORD_FIELD_SEPARATOR)
        .collect::<Vec<_>>();
    if fields.len() != 4 {
        return Err(format!(
            "Microsoft Word returned an invalid inline baseline verification payload: {}",
            output.trim()
        ));
    }
    let applied_position = fields[0]
        .parse::<i32>()
        .map_err(|error| format!("Unable to verify Microsoft Word baseline position: {error}"))?;
    let parse_real = |value: &str, label: &str| {
        value
            .trim()
            .replace(',', ".")
            .parse::<f64>()
            .map_err(|error| format!("Unable to verify Microsoft Word {label}: {error}"))
    };
    let width = parse_real(fields[1], "formula width")?;
    let height = parse_real(fields[2], "formula height")?;
    let matched_shape_index = fields[3]
        .parse::<u32>()
        .map_err(|error| format!("Unable to verify Microsoft Word shape index: {error}"))?;
    if applied_position != position {
        return Err(format!(
            "Microsoft Word persisted baseline position {applied_position} instead of {position}"
        ));
    }
    if !width.is_finite() || !height.is_finite() || width <= 0.0 || height <= 0.0 {
        return Err("Microsoft Word returned invalid formula dimensions after baseline finalization".to_string());
    }
    Ok(WordInlineBaselineResult {
        applied_position,
        width,
        height,
        matched_shape_index,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const MARKER: &str = "visualtex:v1:deflate:AbC_123-def";

    #[test]
    fn native_word_script_targets_exact_marker_and_verifies_write() {
        assert!(APPLY_INLINE_BASELINE_SCRIPT.contains("start with \"position=\""));
        assert!(APPLY_INLINE_BASELINE_SCRIPT.contains("start with \"marker=\""));
        assert!(APPLY_INLINE_BASELINE_SCRIPT.contains("alternative text of candidateShape"));
        assert!(APPLY_INLINE_BASELINE_SCRIPT.contains("repeat with attemptIndex from 1 to 12"));
        assert!(APPLY_INLINE_BASELINE_SCRIPT.contains("set documentShapeCount to count of inline shapes of active document"));
        assert!(!APPLY_INLINE_BASELINE_SCRIPT.contains("set matchingShapes to {}"));
        assert!(!APPLY_INLINE_BASELINE_SCRIPT.contains("caretStart - 1"));
        assert!(
            APPLY_INLINE_BASELINE_SCRIPT.contains("font object of text object of formulaPicture")
        );
        assert!(APPLY_INLINE_BASELINE_SCRIPT
            .contains("font position of formulaFont to requestedPosition"));
        assert!(APPLY_INLINE_BASELINE_SCRIPT.contains("appliedPosition is not requestedPosition"));
        assert!(APPLY_INLINE_BASELINE_SCRIPT.contains(WORD_FIELD_SEPARATOR));
        assert!(!APPLY_INLINE_BASELINE_SCRIPT.contains("ASCII character 31"));
    }

    #[test]
    fn baseline_position_validation_accepts_safe_downward_offsets() {
        assert_eq!(validate_position(0), Ok(0));
        assert_eq!(validate_position(-4), Ok(-4));
        assert_eq!(validate_position(-256), Ok(-256));
        assert!(validate_position(1).is_err());
        assert!(validate_position(-257).is_err());
    }

    #[test]
    fn marker_validation_accepts_only_visualtex_metadata_payloads() {
        assert_eq!(validate_formula_marker(MARKER), Ok(MARKER));
        assert!(validate_formula_marker("other:v1:deflate:abc").is_err());
        assert!(validate_formula_marker("visualtex:v1:deflate:abc=").is_err());
        assert!(validate_formula_marker("visualtex:v1:deflate:abc\n").is_err());
    }

    #[test]
    fn negative_position_is_not_passed_as_an_osascript_option() {
        let position = position_argument(-5);
        let marker = marker_argument(MARKER);
        assert_eq!(position, "position=-5");
        assert_eq!(marker, format!("marker={MARKER}"));
        assert!(!position.starts_with('-'));
        assert!(!marker.starts_with('-'));
    }
}
