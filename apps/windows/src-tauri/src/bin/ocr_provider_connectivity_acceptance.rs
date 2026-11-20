use serde::{Deserialize, Serialize};

pub(crate) const MAX_IMAGE_BYTES: usize = 20 * 1024 * 1024;

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrImageRequest {
    pub(crate) bytes: Vec<u8>,
    pub(crate) extension: String,
    pub(crate) model: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub(crate) struct OcrFormulaResult {
    pub(crate) latex: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrRecognitionResult {
    pub(crate) provider: String,
    pub(crate) model: String,
    pub(crate) elapsed_ms: u64,
    pub(crate) processed_width: u32,
    pub(crate) processed_height: u32,
    pub(crate) background_inverted: bool,
    pub(crate) background_luminance: f64,
    pub(crate) formulas: Vec<OcrFormulaResult>,
}

fn emit_recognition_progress(
    _app: &tauri::AppHandle,
    _request_id: &str,
    _stage: &str,
    _message: impl Into<String>,
    _model: &str,
) {
}

#[path = "../ocr_provider.rs"]
mod ocr_provider;

fn main() -> Result<(), String> {
    let arguments = std::env::args().skip(1).collect::<Vec<_>>();
    let public_probe = arguments.iter().any(|argument| argument == "--public");
    let saved_paddle_image = arguments
        .windows(2)
        .find(|pair| pair[0] == "--saved-paddle")
        .map(|pair| std::path::PathBuf::from(&pair[1]));
    let repeats = arguments
        .windows(2)
        .find(|pair| pair[0] == "--repeat")
        .and_then(|pair| pair[1].parse::<usize>().ok())
        .unwrap_or(3);
    let paddle_model_override = arguments
        .windows(2)
        .find(|pair| pair[0] == "--model")
        .map(|pair| pair[1].as_str());
    let existing_no_proxy = std::env::var("NO_PROXY")
        .or_else(|_| std::env::var("no_proxy"))
        .unwrap_or_default();
    let mut no_proxy = existing_no_proxy.trim().to_string();
    for host in ["127.0.0.1", "localhost"] {
        if !no_proxy.split(',').any(|item| item.trim().eq_ignore_ascii_case(host)) {
            if !no_proxy.is_empty() {
                no_proxy.push(',');
            }
            no_proxy.push_str(host);
        }
    }
    std::env::set_var("NO_PROXY", no_proxy);

    let runtime = tokio::runtime::Builder::new_multi_thread()
        .enable_all()
        .build()
        .map_err(|error| format!("Unable to create OCR provider acceptance runtime: {error}"))?;
    if let Some(image_path) = saved_paddle_image {
        runtime.block_on(ocr_provider::run_saved_paddle_latency_diagnostic(
            &image_path,
            repeats,
            paddle_model_override,
        ))?;
    } else {
        runtime.block_on(ocr_provider::run_provider_connectivity_acceptance(public_probe))?;
        println!("VisualTeX Windows OCR provider connectivity acceptance passed.");
    }
    Ok(())
}
