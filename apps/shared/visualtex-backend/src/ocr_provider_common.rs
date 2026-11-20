//! The remote OCR providers as both backends use them: the providers and their limits, reading the configuration file,
//! normalizing settings, the HTTP requests and responses, and the recognition clients for OpenAI-compatible
//! services, Ollama, Mathpix, SimpleTex and PaddleOCR AI Studio with their output parsing. Each backend includes this
//! file as ocr::provider::common and re-exports it; how secrets are kept (Windows DPAPI, the macOS keychain), the
//! proxy, the configuration views and the provider state stay in each backend.

use super::*;

pub(crate) const LOCAL_PROVIDER: &str = "local";

pub(crate) const OPENAI_COMPATIBLE_PROVIDER: &str = "openai-compatible";

pub(crate) const OLLAMA_PROVIDER: &str = "ollama";

pub(crate) const MATHPIX_PROVIDER: &str = "mathpix";

pub(crate) const PADDLEOCR_PROVIDER: &str = "paddleocr";

pub(crate) const SIMPLETEX_PROVIDER: &str = "simpletex";

pub(super) const CONFIGURATION_SCHEMA_VERSION: u32 = 2;

pub(super) const CONFIGURATION_FILE: &str = "providers.json";

pub(super) const MAX_RESPONSE_BYTES: usize = 4 * 1024 * 1024;

pub(super) const MATHPIX_MAX_BASE64_IMAGE_BYTES: usize = 2 * 1024 * 1024;

pub(super) const PADDLEOCR_JOBS_URL: &str = "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs";

pub(super) const PADDLEOCR_RESULT_MAX_BYTES: usize = 16 * 1024 * 1024;

pub(super) const PADDLEOCR_POLL_INTERVAL: Duration = Duration::from_secs(1);

pub(super) const PADDLEOCR_MAX_POLL_ATTEMPTS: usize = 120;

pub(super) const PADDLEOCR_SUBMIT_TIMEOUT: Duration = Duration::from_secs(20);

pub(super) const PADDLEOCR_STATUS_TIMEOUT: Duration = Duration::from_secs(8);

pub(super) const PADDLEOCR_RESULT_TIMEOUT: Duration = Duration::from_secs(20);

pub(super) const PADDLEOCR_TOTAL_TIMEOUT: Duration = Duration::from_secs(120);

pub(super) const SIMPLETEX_STANDARD_URL: &str = "https://server.simpletex.cn/api/latex_ocr";

pub(super) const SIMPLETEX_TURBO_URL: &str = "https://server.simpletex.cn/api/latex_ocr_turbo";

pub(super) const SIMPLETEX_REQUEST_TIMEOUT: Duration = Duration::from_secs(45);

pub(super) const MAX_FORMULAS: usize = 64;

pub(super) const MAX_FORMULA_CHARS: usize = 200_000;

pub(super) const DEFAULT_PROMPT: &str = "Read every mathematical formula in this image in visual order. Return JSON only in the exact form {\"formulas\":[{\"latex\":\"...\"}]}. Return each independent visual formula row as a separate formulas-array item, while keeping a matrix or cases construction as one item. Use valid LaTeX without markdown fences or surrounding dollar delimiters. Preserve symbols, superscripts, subscripts, matrices, cases, alignment, and line structure. Do not explain the result.";

#[derive(Clone, Default)]
pub(crate) struct OcrProviderState {
    pub(super) configuration: Arc<Mutex<Option<StoredOcrProviderConfiguration>>>,
    pub(super) remote_recognition_running: Arc<AtomicBool>,
}

pub(super) struct RemoteRecognitionLease {
    pub(super) running: Arc<AtomicBool>,
}

pub(super) struct RemoteRecognitionProgress {
    pub(super) app: AppHandle,
    pub(super) request_id: String,
    pub(super) ui_model: String,
}

impl RemoteRecognitionProgress {
    pub(super) fn new(app: &AppHandle, ui_model: &str) -> Self {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .map(|duration| duration.as_millis())
            .unwrap_or_default();
        Self {
            app: app.clone(),
            request_id: format!("remote-{}-{nonce}", std::process::id()),
            ui_model: ui_model.to_string(),
        }
    }

    pub(super) fn emit(&self, stage: &str, message: impl Into<String>) {
        crate::ocr::emit_recognition_progress(
            &self.app,
            &self.request_id,
            stage,
            message,
            &self.ui_model,
        );
    }
}

impl Drop for RemoteRecognitionLease {
    fn drop(&mut self) {
        self.running.store(false, Ordering::SeqCst);
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrProviderConfigurationView {
    pub(super) active_provider: String,
    pub(super) open_ai_compatible: OpenAiCompatibleConfigurationView,
    pub(super) ollama: OllamaConfigurationView,
    pub(super) mathpix: MathpixConfigurationView,
    pub(super) paddle_ocr: PaddleOcrConfigurationView,
    pub(super) simple_tex: SimpleTexConfigurationView,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct OllamaConfigurationView {
    pub(super) base_url: String,
    pub(super) model: String,
    pub(super) prompt: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrProviderConfigurationUpdate {
    pub(super) active_provider: String,
    pub(super) open_ai_compatible: OpenAiCompatibleConfigurationUpdate,
    pub(super) ollama: OllamaConfigurationUpdate,
    pub(super) mathpix: MathpixConfigurationUpdate,
    #[serde(default)]
    pub(super) paddle_ocr: PaddleOcrConfigurationUpdate,
    #[serde(default)]
    pub(super) simple_tex: SimpleTexConfigurationUpdate,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct OpenAiCompatibleConfigurationUpdate {
    pub(super) protocol: String,
    pub(super) base_url: String,
    pub(super) model: String,
    pub(super) prompt: String,
    #[serde(default)]
    pub(super) api_key: Option<String>,
    #[serde(default)]
    pub(super) clear_api_key: bool,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct OllamaConfigurationUpdate {
    pub(super) base_url: String,
    pub(super) model: String,
    pub(super) prompt: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct MathpixConfigurationUpdate {
    pub(super) base_url: String,
    pub(super) app_id: String,
    #[serde(default)]
    pub(super) app_key: Option<String>,
    #[serde(default)]
    pub(super) clear_app_key: bool,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct PaddleOcrConfigurationUpdate {
    pub(super) model: String,
    #[serde(default)]
    pub(super) access_token: Option<String>,
    #[serde(default)]
    pub(super) clear_access_token: bool,
}

impl Default for PaddleOcrConfigurationUpdate {
    fn default() -> Self {
        Self {
            model: "PaddleOCR-VL-1.6".to_string(),
            access_token: None,
            clear_access_token: false,
        }
    }
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct SimpleTexConfigurationUpdate {
    pub(super) model: String,
    #[serde(default)]
    pub(super) access_token: Option<String>,
    #[serde(default)]
    pub(super) clear_access_token: bool,
}

impl Default for SimpleTexConfigurationUpdate {
    fn default() -> Self {
        Self {
            model: "standard".to_string(),
            access_token: None,
            clear_access_token: false,
        }
    }
}

pub(super) fn remote_provider_progress_message(provider: &str, stage: &str) -> Option<&'static str> {
    match (provider, stage) {
        (OPENAI_COMPATIBLE_PROVIDER, "api-submit") => {
            Some("正在准备向 OpenAI 兼容 API 提交图片…")
        }
        (OPENAI_COMPATIBLE_PROVIDER, "api-inference") => {
            Some("OpenAI 兼容 API 正在上传图片并识别公式…")
        }
        (OPENAI_COMPATIBLE_PROVIDER, "api-result") => {
            Some("OpenAI 兼容 API 已返回结果，正在写入编辑器…")
        }
        (OLLAMA_PROVIDER, "api-submit") => Some("正在准备向 Ollama 提交图片…"),
        (OLLAMA_PROVIDER, "api-inference") => Some("Ollama 正在上传图片并识别公式…"),
        (OLLAMA_PROVIDER, "api-result") => Some("Ollama 已返回结果，正在写入编辑器…"),
        (MATHPIX_PROVIDER, "api-submit") => Some("正在准备向 Mathpix 提交图片…"),
        (MATHPIX_PROVIDER, "api-inference") => Some("Mathpix 正在上传图片并识别公式…"),
        (MATHPIX_PROVIDER, "api-result") => Some("Mathpix 已返回结果，正在写入编辑器…"),
        (PADDLEOCR_PROVIDER, "api-submit") => Some("正在提交图片到 PaddleOCR…"),
        (SIMPLETEX_PROVIDER, "api-submit") => Some("正在准备向 SimpleTex 提交图片…"),
        (SIMPLETEX_PROVIDER, "api-inference") => {
            Some("SimpleTex 正在上传图片并识别公式…")
        }
        (SIMPLETEX_PROVIDER, "api-result") => {
            Some("SimpleTex 已返回结果，正在写入编辑器…")
        }
        _ => None,
    }
}

pub(super) fn emit_remote_provider_progress(
    progress: Option<&RemoteRecognitionProgress>,
    provider: &str,
    stage: &str,
) {
    if let (Some(progress), Some(message)) = (
        progress,
        remote_provider_progress_message(provider, stage),
    ) {
        progress.emit(stage, message);
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct StoredOcrProviderConfiguration {
    pub(super) schema_version: u32,
    pub(super) active_provider: String,
    pub(super) open_ai_compatible: StoredOpenAiCompatibleConfiguration,
    pub(super) ollama: StoredOllamaConfiguration,
    pub(super) mathpix: StoredMathpixConfiguration,
    #[serde(default)]
    pub(super) paddle_ocr: StoredPaddleOcrConfiguration,
    #[serde(default)]
    pub(super) simple_tex: StoredSimpleTexConfiguration,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct StoredOllamaConfiguration {
    pub(super) base_url: String,
    pub(super) model: String,
    pub(super) prompt: String,
}

#[derive(Debug)]
pub(super) enum ConfigurationReadError {
    Invalid(String),
    Io(String),
}

impl From<ConfigurationReadError> for String {
    fn from(error: ConfigurationReadError) -> Self {
        match error {
            ConfigurationReadError::Invalid(message) | ConfigurationReadError::Io(message) => {
                message
            }
        }
    }
}

pub(super) fn read_configuration_file(
    app: &AppHandle,
) -> Result<StoredOcrProviderConfiguration, ConfigurationReadError> {
    let path = configuration_path(app).map_err(ConfigurationReadError::Io)?;
    let configuration = migrate_legacy_paddleocr_configuration(match fs::read(&path) {
        Ok(bytes) => {
            serde_json::from_slice::<StoredOcrProviderConfiguration>(&bytes).map_err(|error| {
                ConfigurationReadError::Invalid(format!(
                    "The OCR provider configuration is damaged ({}): {error}",
                    path.display()
                ))
            })?
        }
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => {
            StoredOcrProviderConfiguration::default()
        }
        Err(error) => {
            return Err(ConfigurationReadError::Io(format!(
                "Unable to read the OCR provider configuration ({}): {error}",
                path.display()
            )))
        }
    });
    validate_stored_configuration(&configuration).map_err(|error| {
        ConfigurationReadError::Invalid(format!(
            "The OCR provider configuration is invalid ({}): {error}",
            path.display()
        ))
    })?;
    Ok(configuration)
}

pub(super) fn configuration_path(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX app-data directory: {error}"))?
        .join("ocr")
        .join(CONFIGURATION_FILE))
}

pub(super) fn normalize_provider(value: &str) -> Result<&'static str, String> {
    match value.trim().to_ascii_lowercase().as_str() {
        LOCAL_PROVIDER => Ok(LOCAL_PROVIDER),
        OPENAI_COMPATIBLE_PROVIDER | "openai" => Ok(OPENAI_COMPATIBLE_PROVIDER),
        OLLAMA_PROVIDER => Ok(OLLAMA_PROVIDER),
        MATHPIX_PROVIDER => Ok(MATHPIX_PROVIDER),
        PADDLEOCR_PROVIDER | "paddle" | "paddle-ocr" => Ok(PADDLEOCR_PROVIDER),
        SIMPLETEX_PROVIDER | "simple-tex" => Ok(SIMPLETEX_PROVIDER),
        _ => Err("Unsupported OCR provider".to_string()),
    }
}

pub(super) fn normalize_openai_protocol(value: &str) -> Result<&'static str, String> {
    match value.trim().to_ascii_lowercase().as_str() {
        "responses" => Ok("responses"),
        "chat-completions" | "chat_completions" | "chat" => Ok("chat-completions"),
        _ => Err("Unsupported OpenAI-compatible protocol".to_string()),
    }
}

pub(super) fn normalize_short_text(
    value: &str,
    label: &str,
    max_chars: usize,
    required: bool,
) -> Result<String, String> {
    let normalized = value.trim().to_string();
    if required && normalized.is_empty() {
        return Err(format!("{label} is required"));
    }
    if normalized.chars().count() > max_chars {
        return Err(format!("{label} is too long"));
    }
    if normalized.chars().any(|character| character.is_control()) {
        return Err(format!("{label} contains an unsupported control character"));
    }
    Ok(normalized)
}

pub(super) fn normalize_prompt(value: &str) -> Result<String, String> {
    let normalized = value.trim();
    if normalized.is_empty() {
        return Ok(DEFAULT_PROMPT.to_string());
    }
    if normalized.chars().count() > 8_000 {
        return Err("OCR API prompt is too long".to_string());
    }
    Ok(normalized.to_string())
}

pub(super) fn normalize_base_url(value: &str) -> Result<String, String> {
    let normalized = value.trim().trim_end_matches('/').to_string();
    if normalized.is_empty() {
        return Err("OCR provider base URL is required".to_string());
    }
    validate_url(&normalized)?;
    Ok(normalized)
}

pub(super) fn normalize_paddleocr_model(value: &str) -> Result<&'static str, String> {
    match value.trim() {
        "PaddleOCR-VL-1.6" => Ok("PaddleOCR-VL-1.6"),
        _ => Err("Unsupported PaddleOCR AI Studio model".to_string()),
    }
}

pub(super) fn normalize_simpletex_model(value: &str) -> Result<&'static str, String> {
    match value.trim().to_ascii_lowercase().as_str() {
        "standard" | "latex_ocr" => Ok("standard"),
        "turbo" | "latex_ocr_turbo" => Ok("turbo"),
        _ => Err("Unsupported SimpleTex formula model".to_string()),
    }
}

pub(super) fn validate_secret_transport(value: &str, has_secret: bool, label: &str) -> Result<(), String> {
    if has_secret && value.starts_with("http://") && !is_loopback_url(value)? {
        return Err(format!(
            "Refusing to send {label} over plain HTTP to a non-local OCR endpoint"
        ));
    }
    Ok(())
}

pub(super) struct SensitiveString(pub(super) String);

impl SensitiveString {
    pub(super) fn as_str(&self) -> &str {
        &self.0
    }
}

pub(super) async fn recognize_openai_compatible(
    client: &reqwest::Client,
    configuration: &StoredOpenAiCompatibleConfiguration,
    api_key: Option<&str>,
    data_url: &str,
) -> Result<Vec<String>, String> {
    let schema = formula_json_schema();
    let endpoint = append_endpoint(
        &configuration.base_url,
        if configuration.protocol == "responses" {
            "responses"
        } else {
            "chat/completions"
        },
    );
    let mut body = if configuration.protocol == "responses" {
        json!({
            "model": configuration.model,
            "input": [{
                "role": "user",
                "content": [
                    {"type": "input_text", "text": configuration.prompt},
                    {"type": "input_image", "image_url": data_url}
                ]
            }],
            "text": {
                "format": {
                    "type": "json_schema",
                    "name": "visualtex_formula_ocr",
                    "strict": true,
                    "schema": schema
                }
            },
            "max_output_tokens": 4096
        })
    } else {
        json!({
            "model": configuration.model,
            "messages": [{
                "role": "user",
                "content": [
                    {"type": "text", "text": configuration.prompt},
                    {"type": "image_url", "image_url": {"url": data_url}}
                ]
            }],
            "response_format": {
                "type": "json_schema",
                "json_schema": {
                    "name": "visualtex_formula_ocr",
                    "strict": true,
                    "schema": schema
                }
            },
            "temperature": 0
        })
    };

    let first = post_json(client, &endpoint, api_key, &[], &body).await?;
    let response = if first.status == 400 && structured_output_rejected(&first.text) {
        if configuration.protocol == "responses" {
            if let Some(object) = body.as_object_mut() {
                object.remove("text");
            }
        } else if let Some(object) = body.as_object_mut() {
            object.remove("response_format");
        }
        post_json(client, &endpoint, api_key, &[], &body).await?
    } else {
        first
    };
    ensure_success(&response, "OpenAI-compatible OCR")?;
    let value = response
        .json
        .as_ref()
        .ok_or_else(|| "OpenAI-compatible OCR returned non-JSON output".to_string())?;
    let content = if configuration.protocol == "responses" {
        extract_responses_text(value)
    } else {
        extract_chat_completions_text(value)
    }
    .ok_or_else(|| "OpenAI-compatible OCR response contains no output text".to_string())?;
    parse_formula_output(&content)
}

pub(super) async fn recognize_ollama(
    client: &reqwest::Client,
    configuration: &StoredOllamaConfiguration,
    data_base64: &str,
) -> Result<Vec<String>, String> {
    let endpoint = append_endpoint(&configuration.base_url, "api/chat");
    let mut body = json!({
        "model": configuration.model,
        "stream": false,
        "format": formula_json_schema(),
        "messages": [{
            "role": "user",
            "content": configuration.prompt,
            "images": [data_base64]
        }],
        "options": {"temperature": 0}
    });
    let first = post_json(client, &endpoint, None, &[], &body).await?;
    let response = if first.status == 400 && structured_output_rejected(&first.text) {
        if let Some(object) = body.as_object_mut() {
            object.insert("format".to_string(), Value::String("json".to_string()));
        }
        post_json(client, &endpoint, None, &[], &body).await?
    } else {
        first
    };
    ensure_success(&response, "Ollama OCR")?;
    let value = response
        .json
        .as_ref()
        .ok_or_else(|| "Ollama returned non-JSON output".to_string())?;
    let content = value
        .pointer("/message/content")
        .and_then(Value::as_str)
        .or_else(|| value.get("response").and_then(Value::as_str))
        .ok_or_else(|| "Ollama response contains no message content".to_string())?;
    parse_formula_output(content)
}

pub(super) async fn recognize_simpletex(
    client: &reqwest::Client,
    configuration: &StoredSimpleTexConfiguration,
    access_token: &str,
    image_bytes: &[u8],
    mime: &str,
    extension: &str,
) -> Result<Vec<String>, String> {
    recognize_simpletex_at(
        client,
        configuration,
        access_token,
        image_bytes,
        mime,
        extension,
        simpletex_endpoint(&configuration.model)?,
        SIMPLETEX_REQUEST_TIMEOUT,
    )
    .await
}

#[allow(clippy::too_many_arguments)]
pub(super) async fn recognize_simpletex_at(
    client: &reqwest::Client,
    configuration: &StoredSimpleTexConfiguration,
    access_token: &str,
    image_bytes: &[u8],
    mime: &str,
    extension: &str,
    endpoint: &str,
    timeout: Duration,
) -> Result<Vec<String>, String> {
    normalize_simpletex_model(&configuration.model)?;
    let file_extension = extension.trim().trim_start_matches('.');
    let file_name = if file_extension.is_empty() {
        "visualtex-formula.png".to_string()
    } else {
        format!("visualtex-formula.{file_extension}")
    };
    let part = reqwest::multipart::Part::bytes(image_bytes.to_vec())
        .file_name(file_name)
        .mime_str(mime)
        .map_err(|error| format!("Unable to prepare the SimpleTex image upload: {error}"))?;
    let response = client
        .post(endpoint)
        .header("token", access_token)
        .multipart(reqwest::multipart::Form::new().part("file", part))
        .timeout(timeout)
        .send()
        .await
        .map_err(|error| {
            if error.is_timeout() {
                format!(
                    "SimpleTex formula recognition timed out after {} seconds",
                    timeout.as_secs_f64()
                )
            } else {
                format!("SimpleTex formula recognition request failed: {error}")
            }
        })?;
    let response = read_http_response(
        response,
        MAX_RESPONSE_BYTES,
        "SimpleTex formula recognition",
    )
    .await?;
    ensure_success(&response, "SimpleTex formula recognition")?;
    let value = response
        .json
        .as_ref()
        .ok_or_else(|| "SimpleTex returned non-JSON output".to_string())?;
    parse_simpletex_formula_output(value)
}

pub(super) fn parse_simpletex_formula_output(value: &Value) -> Result<Vec<String>, String> {
    if value.get("status").and_then(Value::as_bool) != Some(true) {
        return Err(format!(
            "SimpleTex formula recognition failed: {}",
            extract_api_error(value).unwrap_or_else(|| "unknown API error".to_string())
        ));
    }
    let latex = value
        .pointer("/res/latex")
        .and_then(Value::as_str)
        .filter(|text| !text.trim().is_empty() && text.trim() != "[EMPTY]")
        .ok_or_else(|| "SimpleTex returned no usable formula".to_string())?;
    normalize_formula_candidates(vec![latex.to_string()])
}

pub(super) fn paddleocr_progress_for_state(state: &str) -> Option<(&'static str, &'static str)> {
    match state {
        "pending" => Some(("api-queued", "PaddleOCR 任务正在排队…")),
        "running" => Some(("api-inference", "PaddleOCR 正在识别公式…")),
        "done" => Some(("api-result", "PaddleOCR 识别完成，正在读取结果…")),
        _ => None,
    }
}

#[cfg_attr(not(test), allow(dead_code))]
pub(super) async fn recognize_paddleocr_aistudio_at(
    client: &reqwest::Client,
    configuration: &StoredPaddleOcrConfiguration,
    access_token: &str,
    image_bytes: &[u8],
    mime: &str,
    extension: &str,
    jobs_url: &str,
) -> Result<Vec<String>, String> {
    recognize_paddleocr_aistudio_at_with_progress(
        client,
        configuration,
        access_token,
        image_bytes,
        mime,
        extension,
        jobs_url,
        None,
    )
    .await
}

#[allow(clippy::too_many_arguments)]
pub(super) async fn recognize_paddleocr_aistudio_at_with_progress(
    client: &reqwest::Client,
    configuration: &StoredPaddleOcrConfiguration,
    access_token: &str,
    image_bytes: &[u8],
    mime: &str,
    extension: &str,
    jobs_url: &str,
    progress: Option<&RemoteRecognitionProgress>,
) -> Result<Vec<String>, String> {
    let model = normalize_paddleocr_model(&configuration.model)?;
    let api_model = model;
    let optional_payload = if model == "PP-StructureV3" {
        json!({
            "useDocOrientationClassify": false,
            "useDocUnwarping": false,
            "useTextlineOrientation": false,
            "useChartRecognition": false,
            "useFormulaRecognition": true,
            "useTableRecognition": false,
            "useSealRecognition": false
        })
    } else {
        json!({
            "useDocOrientationClassify": false,
            "useDocUnwarping": false,
            "useLayoutDetection": true,
            "useChartRecognition": false,
            "showFormulaNumber": false,
            "prettifyMarkdown": false
        })
    };
    let options = serde_json::to_string(&optional_payload)
        .map_err(|error| format!("Unable to encode PaddleOCR options: {error}"))?;
    let file_extension = extension.trim().trim_start_matches('.');
    let file_name = if file_extension.is_empty() {
        "visualtex-formula.png".to_string()
    } else {
        format!("visualtex-formula.{file_extension}")
    };
    let part = reqwest::multipart::Part::bytes(image_bytes.to_vec())
        .file_name(file_name)
        .mime_str(mime)
        .map_err(|error| format!("Unable to prepare the PaddleOCR image upload: {error}"))?;
    let form = reqwest::multipart::Form::new()
        .text("model", api_model.to_string())
        .text("optionalPayload", options)
        .part("file", part);
    let request_started = Instant::now();
    let submit = client
        .post(jobs_url)
        .bearer_auth(access_token)
        .multipart(form)
        .timeout(PADDLEOCR_SUBMIT_TIMEOUT)
        .send()
        .await
        .map_err(|error| format!("PaddleOCR job submission failed: {error}"))?;
    let submit = read_http_response(submit, MAX_RESPONSE_BYTES, "PaddleOCR job submission").await?;
    if submit.status == 429 {
        return Err("PaddleOCR AI Studio daily quota has been exhausted. The official quota is currently 3000 pages per user per model per day; try again after the quota resets or select the other PaddleOCR document-parsing model.".to_string());
    }
    ensure_success(&submit, "PaddleOCR AI Studio")?;
    let submit_json = submit
        .json
        .as_ref()
        .ok_or_else(|| "PaddleOCR job submission returned non-JSON output".to_string())?;
    ensure_paddleocr_api_code(submit_json, "PaddleOCR job submission")?;
    let job_id = submit_json
        .pointer("/data/jobId")
        .and_then(Value::as_str)
        .filter(|value| !value.trim().is_empty())
        .ok_or_else(|| "PaddleOCR job submission returned no jobId".to_string())?;
    let status_url = format!("{}/{job_id}", jobs_url.trim_end_matches('/'));
    if let Some(progress) = progress {
        progress.emit("api-queued", "图片已提交，正在等待 PaddleOCR 处理…");
    }
    let mut last_reported_state = String::new();

    for _ in 0..PADDLEOCR_MAX_POLL_ATTEMPTS {
        if request_started.elapsed() >= PADDLEOCR_TOTAL_TIMEOUT {
            return Err("PaddleOCR AI Studio did not finish within 120 seconds".to_string());
        }
        let status = client
            .get(&status_url)
            .bearer_auth(access_token)
            .timeout(PADDLEOCR_STATUS_TIMEOUT)
            .send()
            .await
            .map_err(|error| format!("PaddleOCR job status request failed: {error}"))?;
        let status = read_http_response(status, MAX_RESPONSE_BYTES, "PaddleOCR job status").await?;
        if status.status == 429 {
            return Err("PaddleOCR AI Studio request quota has been exhausted. Try again after the quota resets.".to_string());
        }
        ensure_success(&status, "PaddleOCR AI Studio")?;
        let status_json = status
            .json
            .as_ref()
            .ok_or_else(|| "PaddleOCR job status returned non-JSON output".to_string())?;
        ensure_paddleocr_api_code(status_json, "PaddleOCR job status")?;
        let state = status_json
            .pointer("/data/state")
            .and_then(Value::as_str)
            .ok_or_else(|| "PaddleOCR job status contains no state".to_string())?;
        if state != last_reported_state {
            if let Some(progress) = progress {
                if let Some((stage, message)) = paddleocr_progress_for_state(state) {
                    progress.emit(stage, message);
                }
            }
            last_reported_state = state.to_string();
        }
        match state {
            "done" => {
                let result_url = status_json
                    .pointer("/data/resultUrl/jsonUrl")
                    .and_then(Value::as_str)
                    .filter(|value| !value.trim().is_empty())
                    .ok_or_else(|| "PaddleOCR completed without a JSON result URL".to_string())?;
                return download_and_parse_paddleocr_result_with_client(client, result_url).await;
            }
            "failed" => {
                let reason = status_json
                    .pointer("/data/errorMsg")
                    .and_then(Value::as_str)
                    .filter(|value| !value.trim().is_empty())
                    .unwrap_or("unknown remote processing error");
                return Err(format!("PaddleOCR AI Studio job failed: {reason}"));
            }
            "pending" | "running" => {
                tokio::time::sleep(PADDLEOCR_POLL_INTERVAL).await;
            }
            other => {
                return Err(format!("PaddleOCR AI Studio returned an unknown job state: {other}"));
            }
        }
    }
    Err("PaddleOCR AI Studio did not finish within 120 seconds".to_string())
}

pub(super) fn ensure_paddleocr_api_code(value: &Value, operation: &str) -> Result<(), String> {
    if value.get("code").and_then(Value::as_i64).unwrap_or(0) == 0 {
        return Ok(());
    }
    Err(format!(
        "{operation} failed: {}",
        extract_api_error(value).unwrap_or_else(|| "unknown API error".to_string())
    ))
}

pub(super) async fn download_and_parse_paddleocr_result_with_client(
    client: &reqwest::Client,
    result_url: &str,
) -> Result<Vec<String>, String> {
    validate_paddleocr_result_url(result_url)?;
    let response = client
        .get(result_url)
        .timeout(PADDLEOCR_RESULT_TIMEOUT)
        .send()
        .await
        .map_err(|error| format!("Unable to download the PaddleOCR result: {error}"))?;
    let response = read_http_response(response, PADDLEOCR_RESULT_MAX_BYTES, "PaddleOCR result").await?;
    ensure_success(&response, "PaddleOCR result download")?;
    parse_paddleocr_result_text(&response.text)
}

pub(super) fn parse_paddleocr_result_text(text: &str) -> Result<Vec<String>, String> {
    if let Ok(value) = serde_json::from_str::<Value>(text.trim()) {
        if let Ok(formulas) = parse_paddleocr_formula_output(&value) {
            return Ok(formulas);
        }
    }
    let mut formulas = Vec::new();
    for line in text.lines().map(str::trim).filter(|line| !line.is_empty()) {
        let Ok(value) = serde_json::from_str::<Value>(line) else {
            continue;
        };
        if let Ok(page_formulas) = parse_paddleocr_formula_output(&value) {
            for formula in page_formulas {
                push_unique_formula_candidate(&mut formulas, formula);
            }
        }
    }
    normalize_formula_candidates(formulas)
}

pub(super) async fn recognize_mathpix(
    client: &reqwest::Client,
    configuration: &StoredMathpixConfiguration,
    app_key: &str,
    data_url: &str,
) -> Result<Vec<String>, String> {
    let endpoint = append_endpoint(&configuration.base_url, "v3/text");
    let body = json!({
        "src": data_url,
        "formats": ["latex_styled", "text"],
        "math_inline_delimiters": ["$", "$"],
        "rm_spaces": true,
        "metadata": {"improve_mathpix": false}
    });
    let headers = [
        ("app_id", configuration.app_id.as_str()),
        ("app_key", app_key),
    ];
    let response = post_json(client, &endpoint, None, &headers, &body).await?;
    ensure_success(&response, "Mathpix OCR")?;
    let value = response
        .json
        .as_ref()
        .ok_or_else(|| "Mathpix returned non-JSON output".to_string())?;
    if let Some(error) = value.get("error").and_then(Value::as_str) {
        if !error.trim().is_empty() {
            return Err(format!("Mathpix OCR failed: {error}"));
        }
    }
    let content = value
        .get("latex_styled")
        .and_then(Value::as_str)
        .filter(|text| !text.trim().is_empty())
        .or_else(|| value.get("text").and_then(Value::as_str))
        .ok_or_else(|| "Mathpix response contains no LaTeX or text output".to_string())?;
    parse_formula_output(content)
}

pub(super) fn parse_paddleocr_formula_output(value: &Value) -> Result<Vec<String>, String> {
    let pages = value
        .pointer("/result/layoutParsingResults")
        .and_then(Value::as_array)
        .ok_or_else(|| "PaddleOCR AI Studio response contains no layoutParsingResults".to_string())?;

    let mut candidates = Vec::new();
    for page in pages {
        if let Some(pruned) = page.get("prunedResult") {
            collect_paddleocr_structured_formulas(pruned, &mut candidates);
        }
    }
    if candidates.is_empty() {
        for page in pages {
            if let Some(markdown) = page.pointer("/markdown/text").and_then(Value::as_str) {
                for formula in extract_markdown_math(markdown) {
                    push_unique_formula_candidate(&mut candidates, formula);
                }
            }
        }
    }
    normalize_formula_candidates(candidates)
}

pub(super) fn extract_markdown_math(markdown: &str) -> Vec<String> {
    let mut formulas = Vec::new();
    let mut cursor = 0usize;
    while cursor < markdown.len() {
        let rest = &markdown[cursor..];
        let (opening, closing) = if rest.starts_with("$$") {
            ("$$", "$$")
        } else if rest.starts_with("\\[") {
            ("\\[", "\\]")
        } else if rest.starts_with("\\(") {
            ("\\(", "\\)")
        } else {
            cursor += rest.chars().next().map(char::len_utf8).unwrap_or(1);
            continue;
        };
        let content_start = cursor + opening.len();
        let Some(relative_end) = markdown[content_start..].find(closing) else {
            cursor = content_start;
            continue;
        };
        let content_end = content_start + relative_end;
        push_unique_formula_candidate(
            &mut formulas,
            markdown[content_start..content_end].to_string(),
        );
        cursor = content_end + closing.len();
    }

    if formulas.is_empty() {
        let mut cursor = 0usize;
        while let Some(opening_offset) = markdown[cursor..].find('$') {
            let opening = cursor + opening_offset;
            if markdown[opening..].starts_with("$$") {
                cursor = opening + 2;
                continue;
            }
            let content_start = opening + 1;
            let Some(closing_offset) = markdown[content_start..].find('$') else {
                break;
            };
            let closing = content_start + closing_offset;
            if closing > content_start {
                push_unique_formula_candidate(
                    &mut formulas,
                    markdown[content_start..closing].to_string(),
                );
            }
            cursor = closing + 1;
        }
    }
    formulas
}

pub(super) fn collect_paddleocr_structured_formulas(value: &Value, output: &mut Vec<String>) {
    match value {
        Value::Array(items) => {
            for item in items {
                collect_paddleocr_structured_formulas(item, output);
            }
        }
        Value::Object(object) => {
            let label = object
                .get("block_label")
                .or_else(|| object.get("blockLabel"))
                .and_then(Value::as_str)
                .unwrap_or_default()
                .to_ascii_lowercase();
            if label.contains("formula") || label.contains("equation") {
                if let Some(content) = object
                    .get("block_content")
                    .or_else(|| object.get("blockContent"))
                    .and_then(Value::as_str)
                {
                    push_unique_formula_candidate(output, content.to_string());
                    return;
                }
            }
            if let Some(formula) = object
                .get("rec_formula")
                .or_else(|| object.get("recFormula"))
                .and_then(Value::as_str)
            {
                push_unique_formula_candidate(output, formula.to_string());
            }
            for child in object.values() {
                collect_paddleocr_structured_formulas(child, output);
            }
        }
        _ => {}
    }
}

pub(super) fn push_unique_formula_candidate(output: &mut Vec<String>, candidate: String) {
    let normalized = candidate.trim();
    if normalized.is_empty() || output.iter().any(|item| item.trim() == normalized) {
        return;
    }
    output.push(normalized.to_string());
}

pub(super) fn append_endpoint(base_url: &str, route: &str) -> String {
    let base = base_url.trim_end_matches('/');
    let normalized_route = route.trim_start_matches('/');
    if base
        .to_ascii_lowercase()
        .ends_with(&format!("/{normalized_route}").to_ascii_lowercase())
    {
        base.to_string()
    } else {
        format!("{base}/{normalized_route}")
    }
}

pub(super) struct HttpJsonResponse {
    pub(super) status: u16,
    pub(super) text: String,
    pub(super) json: Option<Value>,
}

pub(super) async fn read_http_response(
    mut response: reqwest::Response,
    max_bytes: usize,
    label: &str,
) -> Result<HttpJsonResponse, String> {
    let status = response.status().as_u16();
    if response
        .content_length()
        .is_some_and(|length| length > max_bytes as u64)
    {
        return Err(format!("{label} is larger than the allowed response limit"));
    }
    let mut bytes = Vec::with_capacity(
        response
            .content_length()
            .unwrap_or(0)
            .min(max_bytes as u64) as usize,
    );
    while let Some(chunk) = response
        .chunk()
        .await
        .map_err(|error| format!("Unable to read {label}: {error}"))?
    {
        if bytes.len().saturating_add(chunk.len()) > max_bytes {
            return Err(format!("{label} is larger than the allowed response limit"));
        }
        bytes.extend_from_slice(&chunk);
    }
    let text = String::from_utf8_lossy(&bytes).trim().to_string();
    let json = serde_json::from_slice::<Value>(&bytes).ok();
    Ok(HttpJsonResponse { status, text, json })
}

pub(super) fn ensure_success(response: &HttpJsonResponse, provider: &str) -> Result<(), String> {
    if (200..300).contains(&response.status) {
        return Ok(());
    }
    let detail = response
        .json
        .as_ref()
        .and_then(extract_api_error)
        .unwrap_or_else(|| truncate_for_error(&response.text, 800));
    Err(if detail.is_empty() {
        format!("{provider} returned HTTP {}", response.status)
    } else {
        format!("{provider} returned HTTP {}: {detail}", response.status)
    })
}

pub(super) fn structured_output_rejected(text: &str) -> bool {
    let normalized = text.to_ascii_lowercase();
    normalized.contains("response_format")
        || normalized.contains("json_schema")
        || normalized.contains("structured output")
        || normalized.contains("unknown field") && normalized.contains("format")
}

pub(super) fn extract_responses_text(value: &Value) -> Option<String> {
    if let Some(text) = value.get("output_text").and_then(Value::as_str) {
        if !text.trim().is_empty() {
            return Some(text.to_string());
        }
    }
    let mut parts = Vec::new();
    for item in value.get("output")?.as_array()? {
        if let Some(content) = item.get("content").and_then(Value::as_array) {
            for entry in content {
                if let Some(text) = entry.get("text").and_then(Value::as_str) {
                    if !text.trim().is_empty() {
                        parts.push(text.to_string());
                    }
                }
            }
        }
    }
    (!parts.is_empty()).then(|| parts.join("\n"))
}

pub(super) fn extract_chat_completions_text(value: &Value) -> Option<String> {
    let content = value.pointer("/choices/0/message/content")?;
    if let Some(text) = content.as_str() {
        return Some(text.to_string());
    }
    let mut parts = Vec::new();
    for entry in content.as_array()? {
        if let Some(text) = entry.get("text").and_then(Value::as_str) {
            if !text.trim().is_empty() {
                parts.push(text.to_string());
            }
        }
    }
    (!parts.is_empty()).then(|| parts.join("\n"))
}

pub(super) fn formula_json_schema() -> Value {
    json!({
        "type": "object",
        "additionalProperties": false,
        "properties": {
            "formulas": {
                "type": "array",
                "maxItems": MAX_FORMULAS,
                "items": {
                    "type": "object",
                    "additionalProperties": false,
                    "properties": {"latex": {"type": "string"}},
                    "required": ["latex"]
                }
            }
        },
        "required": ["formulas"]
    })
}

pub(super) fn parse_formula_output(raw: &str) -> Result<Vec<String>, String> {
    let cleaned = strip_markdown_fence(raw.trim());
    if cleaned.is_empty() {
        return Err("OCR API returned an empty result".to_string());
    }
    let parsed = serde_json::from_str::<Value>(cleaned)
        .ok()
        .or_else(|| extract_embedded_json(cleaned));
    let candidates = if let Some(value) = parsed {
        formula_candidates_from_json(&value)
    } else {
        vec![cleaned.to_string()]
    };
    normalize_formula_candidates(candidates)
}

pub(super) fn formula_candidates_from_json(value: &Value) -> Vec<String> {
    if let Some(formulas) = value.get("formulas").and_then(Value::as_array) {
        return formulas
            .iter()
            .filter_map(|item| {
                item.as_str().map(str::to_string).or_else(|| {
                    item.get("latex")
                        .and_then(Value::as_str)
                        .map(str::to_string)
                })
            })
            .collect();
    }
    if let Some(latex) = value.get("latex").and_then(Value::as_str) {
        return vec![latex.to_string()];
    }
    if let Some(text) = value.get("text").and_then(Value::as_str) {
        return vec![text.to_string()];
    }
    if let Some(items) = value.as_array() {
        return items
            .iter()
            .filter_map(|item| {
                item.as_str().map(str::to_string).or_else(|| {
                    item.get("latex")
                        .and_then(Value::as_str)
                        .map(str::to_string)
                })
            })
            .collect();
    }
    Vec::new()
}

pub(super) fn normalize_formula_candidates(candidates: Vec<String>) -> Result<Vec<String>, String> {
    let mut formulas = Vec::new();
    let mut total_chars = 0usize;
    for candidate in candidates.into_iter().take(MAX_FORMULAS) {
        let normalized = strip_formula_wrapper(strip_markdown_fence(candidate.trim())).trim();
        if normalized.is_empty() {
            continue;
        }
        let char_count = normalized.chars().count();
        total_chars = total_chars.saturating_add(char_count);
        if total_chars > MAX_FORMULA_CHARS {
            return Err("OCR API formula output is too large".to_string());
        }
        formulas.push(normalized.to_string());
    }
    if formulas.is_empty() {
        return Err("OCR API returned no usable formula".to_string());
    }
    Ok(formulas)
}

pub(super) fn strip_markdown_fence(value: &str) -> &str {
    let trimmed = value.trim();
    if !trimmed.starts_with("```") || !trimmed.ends_with("```") {
        return trimmed;
    }
    let Some(first_newline) = trimmed.find('\n') else {
        return trimmed;
    };
    let content = &trimmed[first_newline + 1..trimmed.len() - 3];
    content.trim()
}

pub(super) fn extract_embedded_json(value: &str) -> Option<Value> {
    let start = value.find('{')?;
    let end = value.rfind('}')?;
    (end > start)
        .then(|| serde_json::from_str::<Value>(&value[start..=end]).ok())
        .flatten()
}

pub(super) fn strip_formula_wrapper(value: &str) -> &str {
    let trimmed = value.trim();
    for (prefix, suffix) in [("$$", "$$"), ("\\[", "\\]"), ("\\(", "\\)"), ("$", "$")] {
        if trimmed.starts_with(prefix)
            && trimmed.ends_with(suffix)
            && trimmed.len() >= prefix.len() + suffix.len()
        {
            return trimmed[prefix.len()..trimmed.len() - suffix.len()].trim();
        }
    }
    trimmed
}

pub(super) fn truncate_for_error(value: &str, max_chars: usize) -> String {
    let mut result = value.chars().take(max_chars).collect::<String>();
    if value.chars().count() > max_chars {
        result.push('…');
    }
    result
}

pub(super) fn mime_for_extension(extension: &str) -> Result<&'static str, String> {
    match extension
        .trim()
        .trim_start_matches('.')
        .to_ascii_lowercase()
        .as_str()
    {
        "png" => Ok("image/png"),
        "jpg" | "jpeg" => Ok("image/jpeg"),
        "webp" => Ok("image/webp"),
        "bmp" => Ok("image/bmp"),
        "tif" | "tiff" => Ok("image/tiff"),
        _ => Err("Unsupported OCR image extension".to_string()),
    }
}

pub(super) fn image_dimensions(bytes: &[u8]) -> (u32, u32) {
    if bytes.len() >= 24 && bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        return (
            u32::from_be_bytes([bytes[16], bytes[17], bytes[18], bytes[19]]),
            u32::from_be_bytes([bytes[20], bytes[21], bytes[22], bytes[23]]),
        );
    }
    if bytes.len() >= 26 && bytes.starts_with(b"BM") {
        let width = i32::from_le_bytes([bytes[18], bytes[19], bytes[20], bytes[21]]);
        let height = i32::from_le_bytes([bytes[22], bytes[23], bytes[24], bytes[25]]);
        return (width.unsigned_abs(), height.unsigned_abs());
    }
    if bytes.len() >= 30 && bytes.starts_with(b"RIFF") && &bytes[8..12] == b"WEBP" {
        if &bytes[12..16] == b"VP8X" {
            let width = 1
                + u32::from(bytes[24])
                + (u32::from(bytes[25]) << 8)
                + (u32::from(bytes[26]) << 16);
            let height = 1
                + u32::from(bytes[27])
                + (u32::from(bytes[28]) << 8)
                + (u32::from(bytes[29]) << 16);
            return (width, height);
        }
    }
    jpeg_dimensions(bytes).unwrap_or((0, 0))
}

pub(super) fn jpeg_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
    if bytes.len() < 4 || bytes[0] != 0xff || bytes[1] != 0xd8 {
        return None;
    }
    let mut index = 2usize;
    while index + 4 <= bytes.len() {
        while index < bytes.len() && bytes[index] == 0xff {
            index += 1;
        }
        if index >= bytes.len() {
            return None;
        }
        let marker = bytes[index];
        index += 1;
        if marker == 0xd8 || marker == 0xd9 || marker == 0x01 {
            continue;
        }
        if index + 2 > bytes.len() {
            return None;
        }
        let segment_length = u16::from_be_bytes([bytes[index], bytes[index + 1]]) as usize;
        if segment_length < 2 || index + segment_length > bytes.len() {
            return None;
        }
        let is_start_of_frame = matches!(
            marker,
            0xc0 | 0xc1
                | 0xc2
                | 0xc3
                | 0xc5
                | 0xc6
                | 0xc7
                | 0xc9
                | 0xca
                | 0xcb
                | 0xcd
                | 0xce
                | 0xcf
        );
        if is_start_of_frame && segment_length >= 7 {
            let height = u16::from_be_bytes([bytes[index + 3], bytes[index + 4]]) as u32;
            let width = u16::from_be_bytes([bytes[index + 5], bytes[index + 6]]) as u32;
            return Some((width, height));
        }
        index += segment_length;
    }
    None
}
