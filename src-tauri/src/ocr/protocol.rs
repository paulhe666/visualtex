use serde::{Deserialize, Serialize};

pub const DEFAULT_MODEL: &str = "PP-FormulaNet_plus-M";
pub const MAX_IMAGE_BYTES: usize = 12 * 1024 * 1024;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct RecognizeFormulaRequest {
    pub file_name: String,
    pub mime_type: String,
    pub base64_data: String,
    pub model: Option<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FormulaOcrResult {
    pub latex: String,
    pub raw_latex: String,
    pub model: String,
    pub elapsed_ms: u64,
    pub warnings: Vec<String>,
}

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OcrStatus {
    pub available: bool,
    pub running: bool,
    pub python_path: Option<String>,
    pub script_path: Option<String>,
    pub message: String,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
pub struct OcrError {
    pub code: String,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub detail: Option<String>,
}

impl OcrError {
    pub fn new(code: impl Into<String>, message: impl Into<String>) -> Self {
        Self {
            code: code.into(),
            message: message.into(),
            detail: None,
        }
    }

    pub fn with_detail(
        code: impl Into<String>,
        message: impl Into<String>,
        detail: impl Into<String>,
    ) -> Self {
        Self {
            code: code.into(),
            message: message.into(),
            detail: Some(detail.into()),
        }
    }
}

#[derive(Debug, Deserialize)]
pub struct SidecarErrorResponse {
    pub code: String,
    pub message: String,
    pub detail: Option<String>,
}

#[derive(Debug, Deserialize)]
pub struct SidecarResponse {
    pub id: String,
    pub ok: bool,
    pub status: Option<String>,
    pub latex: Option<String>,
    pub raw_latex: Option<String>,
    pub model: Option<String>,
    pub elapsed_ms: Option<u64>,
    #[serde(default)]
    pub warnings: Vec<String>,
    pub error: Option<SidecarErrorResponse>,
}

pub fn validate_model(model: &str) -> Result<(), OcrError> {
    match model {
        "PP-FormulaNet_plus-M" | "PP-FormulaNet_plus-L" | "PP-FormulaNet-S" => Ok(()),
        _ => Err(OcrError::new(
            "INVALID_REQUEST",
            format!("Unsupported formula OCR model: {model}"),
        )),
    }
}
