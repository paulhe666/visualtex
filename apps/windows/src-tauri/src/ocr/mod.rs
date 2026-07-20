mod manager;
mod protocol;

use std::time::Duration;

use base64::{engine::general_purpose::STANDARD, Engine as _};
use manager::{resolve_runtime, OcrManager};
use protocol::{
    validate_model, FormulaOcrResult, OcrError, OcrStatus, RecognizeFormulaRequest,
    SidecarResponse, DEFAULT_MODEL, MAX_IMAGE_BYTES,
};
use serde_json::json;
use tauri::{AppHandle, Manager, State};
use tokio::{fs, sync::Mutex};
use uuid::Uuid;

pub struct OcrState {
    manager: Mutex<OcrManager>,
}

impl Default for OcrState {
    fn default() -> Self {
        Self {
            manager: Mutex::new(OcrManager::default()),
        }
    }
}

#[tauri::command]
pub async fn ocr_status(app: AppHandle, state: State<'_, OcrState>) -> OcrStatus {
    let runtime = resolve_runtime(&app);
    let running = state.manager.lock().await.is_running();

    match runtime {
        Ok(paths) => OcrStatus {
            available: true,
            running,
            python_path: Some(paths.python.display().to_string()),
            script_path: Some(paths.script.display().to_string()),
            message: if running {
                "Formula OCR runtime is running".to_owned()
            } else {
                "Formula OCR runtime is ready".to_owned()
            },
        },
        Err(error) => OcrStatus {
            available: false,
            running: false,
            python_path: None,
            script_path: None,
            message: error.message,
        },
    }
}

#[tauri::command]
pub async fn warmup_formula_ocr(
    app: AppHandle,
    state: State<'_, OcrState>,
    model: Option<String>,
) -> Result<OcrStatus, OcrError> {
    let model = model.unwrap_or_else(|| DEFAULT_MODEL.to_owned());
    validate_model(&model)?;
    let request_id = Uuid::new_v4().to_string();
    let payload = json!({
        "id": request_id,
        "action": "warmup",
        "model": model,
        "device": "cpu"
    });

    state
        .manager
        .lock()
        .await
        .request(&app, payload, Duration::from_secs(600))
        .await?;

    let runtime = resolve_runtime(&app)?;
    Ok(OcrStatus {
        available: true,
        running: true,
        python_path: Some(runtime.python.display().to_string()),
        script_path: Some(runtime.script.display().to_string()),
        message: "Formula OCR model is ready".to_owned(),
    })
}

#[tauri::command]
pub async fn recognize_formula_image(
    app: AppHandle,
    state: State<'_, OcrState>,
    request: RecognizeFormulaRequest,
) -> Result<FormulaOcrResult, OcrError> {
    if !request.mime_type.starts_with("image/") {
        return Err(OcrError::new(
            "UNSUPPORTED_IMAGE",
            "Only PNG, JPEG, and WebP images are supported",
        ));
    }

    let model = request.model.unwrap_or_else(|| DEFAULT_MODEL.to_owned());
    validate_model(&model)?;
    let bytes = decode_base64(&request.base64_data)?;
    if bytes.is_empty() {
        return Err(OcrError::new("UNSUPPORTED_IMAGE", "The image is empty"));
    }
    if bytes.len() > MAX_IMAGE_BYTES {
        return Err(OcrError::new(
            "IMAGE_TOO_LARGE",
            "The formula image exceeds 12 MiB",
        ));
    }

    let extension = detect_image_extension(&bytes).ok_or_else(|| {
        OcrError::new(
            "UNSUPPORTED_IMAGE",
            "The selected file is not a valid PNG, JPEG, or WebP image",
        )
    })?;

    let cache_dir = app
        .path()
        .app_cache_dir()
        .map_err(|error| {
            OcrError::with_detail(
                "INFERENCE_FAILED",
                "Unable to resolve the application cache directory",
                error.to_string(),
            )
        })?
        .join("ocr-input");
    fs::create_dir_all(&cache_dir).await.map_err(|error| {
        OcrError::with_detail(
            "INFERENCE_FAILED",
            "Unable to create OCR temporary directory",
            error.to_string(),
        )
    })?;

    let safe_stem = sanitize_file_stem(&request.file_name);
    let temp_path = cache_dir.join(format!(
        "{}-{}.{}",
        safe_stem,
        Uuid::new_v4(),
        extension
    ));
    fs::write(&temp_path, &bytes).await.map_err(|error| {
        OcrError::with_detail(
            "INFERENCE_FAILED",
            "Unable to write OCR temporary image",
            error.to_string(),
        )
    })?;

    let request_id = Uuid::new_v4().to_string();
    let payload = json!({
        "id": request_id,
        "action": "recognize",
        "image_path": temp_path,
        "model": model,
        "device": "cpu"
    });

    let response = state
        .manager
        .lock()
        .await
        .request(&app, payload, Duration::from_secs(180))
        .await;
    let _ = fs::remove_file(&temp_path).await;

    sidecar_to_result(response?)
}

fn sidecar_to_result(response: SidecarResponse) -> Result<FormulaOcrResult, OcrError> {
    let latex = response.latex.ok_or_else(|| {
        OcrError::new("EMPTY_RESULT", "The OCR process returned no LaTeX")
    })?;
    Ok(FormulaOcrResult {
        raw_latex: response.raw_latex.unwrap_or_else(|| latex.clone()),
        latex,
        model: response.model.unwrap_or_else(|| DEFAULT_MODEL.to_owned()),
        elapsed_ms: response.elapsed_ms.unwrap_or_default(),
        warnings: response.warnings,
    })
}

fn decode_base64(value: &str) -> Result<Vec<u8>, OcrError> {
    let encoded = value.split_once(',').map(|(_, data)| data).unwrap_or(value);
    STANDARD.decode(encoded).map_err(|error| {
        OcrError::with_detail(
            "INVALID_REQUEST",
            "Unable to decode the selected image",
            error.to_string(),
        )
    })
}

fn detect_image_extension(bytes: &[u8]) -> Option<&'static str> {
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        Some("png")
    } else if bytes.starts_with(&[0xff, 0xd8, 0xff]) {
        Some("jpg")
    } else if bytes.len() >= 12 && &bytes[..4] == b"RIFF" && &bytes[8..12] == b"WEBP" {
        Some("webp")
    } else {
        None
    }
}

fn sanitize_file_stem(file_name: &str) -> String {
    let stem = file_name
        .rsplit_once('.')
        .map(|(stem, _)| stem)
        .unwrap_or(file_name);
    let sanitized: String = stem
        .chars()
        .filter(|character| character.is_ascii_alphanumeric() || *character == '-' || *character == '_')
        .take(48)
        .collect();
    if sanitized.is_empty() {
        "formula".to_owned()
    } else {
        sanitized
    }
}

#[cfg(test)]
mod tests {
    use super::{detect_image_extension, sanitize_file_stem};

    #[test]
    fn detects_png() {
        assert_eq!(detect_image_extension(b"\x89PNG\r\n\x1a\nrest"), Some("png"));
    }

    #[test]
    fn sanitizes_file_name() {
        assert_eq!(sanitize_file_stem("my formula (1).png"), "myformula1");
    }
}
