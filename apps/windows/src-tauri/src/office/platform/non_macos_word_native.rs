use serde::Serialize;

#[derive(Debug, Clone, Serialize, PartialEq, Eq)]
#[serde(rename_all = "camelCase")]
pub struct WordInlineBaselineResult {
    pub applied_position: i32,
}

pub fn apply_inline_baseline(
    _position: i32,
    _formula_marker: &str,
) -> Result<WordInlineBaselineResult, String> {
    Err(
        "macOS Microsoft Word native baseline integration is unavailable on this platform"
            .to_string(),
    )
}
