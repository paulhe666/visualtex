use super::{OcrFormulaResult, OcrImageRequest, OcrRecognitionResult, MAX_IMAGE_BYTES};
use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::fs;
use std::path::PathBuf;
#[cfg(target_os = "macos")]
use std::process::Command;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};
use tauri::{AppHandle, Manager};

pub(crate) const LOCAL_PROVIDER: &str = "local";
pub(crate) const OPENAI_COMPATIBLE_PROVIDER: &str = "openai-compatible";
pub(crate) const OLLAMA_PROVIDER: &str = "ollama";
pub(crate) const MATHPIX_PROVIDER: &str = "mathpix";
pub(crate) const PADDLEOCR_PROVIDER: &str = "paddleocr";
pub(crate) const SIMPLETEX_PROVIDER: &str = "simpletex";

const CONFIGURATION_SCHEMA_VERSION: u32 = 2;
const CONFIGURATION_FILE: &str = "providers.json";
const KEYCHAIN_SERVICE: &str = "com.visualtex.studio.ocr-providers";
const OPENAI_API_KEY_ACCOUNT: &str = "openai-compatible-api-key";
const MATHPIX_APP_KEY_ACCOUNT: &str = "mathpix-app-key";
const PADDLEOCR_ACCESS_TOKEN_ACCOUNT: &str = "paddleocr-aistudio-access-token";
const SIMPLETEX_ACCESS_TOKEN_ACCOUNT: &str = "simpletex-user-access-token";
const PADDLEOCR_JOBS_URL: &str = "https://paddleocr.aistudio-app.com/api/v2/ocr/jobs";
const PADDLEOCR_RESULT_MAX_BYTES: usize = 16 * 1024 * 1024;
const PADDLEOCR_POLL_INTERVAL: Duration = Duration::from_secs(1);
const PADDLEOCR_MAX_POLL_ATTEMPTS: usize = 120;
const PADDLEOCR_SUBMIT_TIMEOUT: Duration = Duration::from_secs(20);
const PADDLEOCR_STATUS_TIMEOUT: Duration = Duration::from_secs(8);
const PADDLEOCR_RESULT_TIMEOUT: Duration = Duration::from_secs(20);
const PADDLEOCR_TOTAL_TIMEOUT: Duration = Duration::from_secs(120);
const SIMPLETEX_STANDARD_URL: &str = "https://server.simpletex.cn/api/latex_ocr";
const SIMPLETEX_TURBO_URL: &str = "https://server.simpletex.cn/api/latex_ocr_turbo";
const SIMPLETEX_REQUEST_TIMEOUT: Duration = Duration::from_secs(45);
const MAX_RESPONSE_BYTES: usize = 4 * 1024 * 1024;
const MATHPIX_MAX_BASE64_IMAGE_BYTES: usize = 2 * 1024 * 1024;
const MAX_FORMULAS: usize = 64;
const MAX_FORMULA_CHARS: usize = 200_000;
const MAX_SECRET_CHARS: usize = 8_192;
const DEFAULT_PROMPT: &str = "Read every mathematical formula in this image in visual order. Return JSON only in the exact form {\"formulas\":[{\"latex\":\"...\"}]}. Return each independent visual formula row as a separate formulas-array item, while keeping a matrix or cases construction as one item. Use valid LaTeX without markdown fences or surrounding dollar delimiters. Preserve symbols, superscripts, subscripts, matrices, cases, alignment, and line structure. Do not explain the result.";

#[derive(Clone, Default)]
pub(crate) struct OcrProviderState {
    configuration: Arc<Mutex<Option<StoredOcrProviderConfiguration>>>,
    remote_recognition_running: Arc<AtomicBool>,
}

struct RemoteRecognitionLease {
    running: Arc<AtomicBool>,
}

struct RemoteRecognitionProgress {
    app: AppHandle,
    request_id: String,
    ui_model: String,
}

impl RemoteRecognitionProgress {
    fn new(app: &AppHandle, ui_model: &str) -> Self {
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

    fn emit(&self, stage: &str, message: impl Into<String>) {
        super::emit_recognition_progress(
            &self.app,
            &self.request_id,
            stage,
            message,
            &self.ui_model,
        );
    }
}

fn remote_provider_progress_message(provider: &str, stage: &str) -> Option<&'static str> {
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

fn emit_remote_provider_progress(
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

impl Drop for RemoteRecognitionLease {
    fn drop(&mut self) {
        self.running.store(false, Ordering::SeqCst);
    }
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrProviderConfigurationView {
    active_provider: String,
    open_ai_compatible: OpenAiCompatibleConfigurationView,
    ollama: OllamaConfigurationView,
    mathpix: MathpixConfigurationView,
    paddle_ocr: PaddleOcrConfigurationView,
    simple_tex: SimpleTexConfigurationView,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct OpenAiCompatibleConfigurationView {
    protocol: String,
    base_url: String,
    model: String,
    prompt: String,
    has_api_key: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct OllamaConfigurationView {
    base_url: String,
    model: String,
    prompt: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct MathpixConfigurationView {
    base_url: String,
    app_id: String,
    has_app_key: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct PaddleOcrConfigurationView {
    model: String,
    has_access_token: bool,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct SimpleTexConfigurationView {
    model: String,
    has_access_token: bool,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrProviderConfigurationUpdate {
    active_provider: String,
    open_ai_compatible: OpenAiCompatibleConfigurationUpdate,
    ollama: OllamaConfigurationUpdate,
    mathpix: MathpixConfigurationUpdate,
    #[serde(default)]
    paddle_ocr: PaddleOcrConfigurationUpdate,
    #[serde(default)]
    simple_tex: SimpleTexConfigurationUpdate,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OpenAiCompatibleConfigurationUpdate {
    protocol: String,
    base_url: String,
    model: String,
    prompt: String,
    #[serde(default)]
    api_key: Option<String>,
    #[serde(default)]
    clear_api_key: bool,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct OllamaConfigurationUpdate {
    base_url: String,
    model: String,
    prompt: String,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct MathpixConfigurationUpdate {
    base_url: String,
    app_id: String,
    #[serde(default)]
    app_key: Option<String>,
    #[serde(default)]
    clear_app_key: bool,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
struct PaddleOcrConfigurationUpdate {
    model: String,
    #[serde(default)]
    access_token: Option<String>,
    #[serde(default)]
    clear_access_token: bool,
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
struct SimpleTexConfigurationUpdate {
    model: String,
    #[serde(default)]
    access_token: Option<String>,
    #[serde(default)]
    clear_access_token: bool,
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

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredOcrProviderConfiguration {
    schema_version: u32,
    active_provider: String,
    open_ai_compatible: StoredOpenAiCompatibleConfiguration,
    ollama: StoredOllamaConfiguration,
    mathpix: StoredMathpixConfiguration,
    #[serde(default)]
    paddle_ocr: StoredPaddleOcrConfiguration,
    #[serde(default)]
    simple_tex: StoredSimpleTexConfiguration,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredOpenAiCompatibleConfiguration {
    protocol: String,
    base_url: String,
    model: String,
    prompt: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredOllamaConfiguration {
    base_url: String,
    model: String,
    prompt: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredMathpixConfiguration {
    base_url: String,
    app_id: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredPaddleOcrConfiguration {
    model: String,
}

impl Default for StoredPaddleOcrConfiguration {
    fn default() -> Self {
        Self {
            model: "PaddleOCR-VL-1.6".to_string(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredSimpleTexConfiguration {
    model: String,
}

impl Default for StoredSimpleTexConfiguration {
    fn default() -> Self {
        Self {
            model: "standard".to_string(),
        }
    }
}

impl Default for StoredOcrProviderConfiguration {
    fn default() -> Self {
        Self {
            schema_version: CONFIGURATION_SCHEMA_VERSION,
            active_provider: LOCAL_PROVIDER.to_string(),
            open_ai_compatible: StoredOpenAiCompatibleConfiguration {
                protocol: "responses".to_string(),
                base_url: "https://api.openai.com/v1".to_string(),
                model: String::new(),
                prompt: DEFAULT_PROMPT.to_string(),
            },
            ollama: StoredOllamaConfiguration {
                base_url: "http://127.0.0.1:11434".to_string(),
                model: String::new(),
                prompt: DEFAULT_PROMPT.to_string(),
            },
            mathpix: StoredMathpixConfiguration {
                base_url: "https://api.mathpix.com".to_string(),
                app_id: String::new(),
            },
            paddle_ocr: StoredPaddleOcrConfiguration::default(),
            simple_tex: StoredSimpleTexConfiguration::default(),
        }
    }
}

impl OcrProviderState {
    pub(crate) fn configuration(
        &self,
        app: &AppHandle,
    ) -> Result<OcrProviderConfigurationView, String> {
        let configuration = self.load(app)?;
        configuration_view(&configuration)
    }

    pub(crate) fn active_provider(&self, app: &AppHandle) -> Result<String, String> {
        Ok(self.load(app)?.active_provider)
    }

    pub(crate) fn save_configuration(
        &self,
        app: &AppHandle,
        update: OcrProviderConfigurationUpdate,
    ) -> Result<OcrProviderConfigurationView, String> {
        let _previous = self.load_for_update(app)?;
        let active_provider = normalize_provider(&update.active_provider)?.to_string();
        let protocol = normalize_openai_protocol(&update.open_ai_compatible.protocol)?.to_string();
        let openai_base_url = normalize_base_url(&update.open_ai_compatible.base_url)?;
        let ollama_base_url = normalize_base_url(&update.ollama.base_url)?;
        let mathpix_base_url = normalize_base_url(&update.mathpix.base_url)?;
        let openai_model = normalize_short_text(
            &update.open_ai_compatible.model,
            "OpenAI-compatible model",
            160,
            false,
        )?;
        let ollama_model = normalize_short_text(&update.ollama.model, "Ollama model", 160, false)?;
        let openai_prompt = normalize_prompt(&update.open_ai_compatible.prompt)?;
        let ollama_prompt = normalize_prompt(&update.ollama.prompt)?;
        let mathpix_app_id =
            normalize_short_text(&update.mathpix.app_id, "Mathpix app_id", 240, false)?;
        let paddle_ocr_model = normalize_paddleocr_model(&update.paddle_ocr.model)?.to_string();
        let simpletex_model = normalize_simpletex_model(&update.simple_tex.model)?.to_string();

        let current_openai_secret = secret_exists(OPENAI_API_KEY_ACCOUNT)?;
        let current_mathpix_secret = secret_exists(MATHPIX_APP_KEY_ACCOUNT)?;
        let current_paddle_ocr_secret = secret_exists(PADDLEOCR_ACCESS_TOKEN_ACCOUNT)?;
        let current_simpletex_secret = secret_exists(SIMPLETEX_ACCESS_TOKEN_ACCOUNT)?;
        let openai_replacement = normalize_secret_update(update.open_ai_compatible.api_key)?;
        let mathpix_replacement = normalize_secret_update(update.mathpix.app_key)?;
        let paddle_ocr_replacement = normalize_secret_update(update.paddle_ocr.access_token)?;
        let simpletex_replacement = normalize_secret_update(update.simple_tex.access_token)?;
        let has_openai_secret = if update.open_ai_compatible.clear_api_key {
            false
        } else {
            openai_replacement.is_some() || current_openai_secret
        };
        let has_mathpix_secret = if update.mathpix.clear_app_key {
            false
        } else {
            mathpix_replacement.is_some() || current_mathpix_secret
        };
        let has_paddle_ocr_secret = if update.paddle_ocr.clear_access_token {
            false
        } else {
            paddle_ocr_replacement.is_some() || current_paddle_ocr_secret
        };
        let has_simpletex_secret = if update.simple_tex.clear_access_token {
            false
        } else {
            simpletex_replacement.is_some() || current_simpletex_secret
        };

        validate_secret_transport(&openai_base_url, has_openai_secret, "API key")?;
        validate_secret_transport(&mathpix_base_url, has_mathpix_secret, "app_key")?;

        let next = StoredOcrProviderConfiguration {
            schema_version: CONFIGURATION_SCHEMA_VERSION,
            active_provider,
            open_ai_compatible: StoredOpenAiCompatibleConfiguration {
                protocol,
                base_url: openai_base_url,
                model: openai_model,
                prompt: openai_prompt,
            },
            ollama: StoredOllamaConfiguration {
                base_url: ollama_base_url,
                model: ollama_model,
                prompt: ollama_prompt,
            },
            mathpix: StoredMathpixConfiguration {
                base_url: mathpix_base_url,
                app_id: mathpix_app_id,
            },
            paddle_ocr: StoredPaddleOcrConfiguration {
                model: paddle_ocr_model,
            },
            simple_tex: StoredSimpleTexConfiguration {
                model: simpletex_model,
            },
        };
        validate_active_provider_configuration(
            &next,
            has_mathpix_secret,
            has_paddle_ocr_secret,
            has_simpletex_secret,
        )?;

        apply_secret_update(
            OPENAI_API_KEY_ACCOUNT,
            openai_replacement.as_deref(),
            update.open_ai_compatible.clear_api_key,
        )?;
        apply_secret_update(
            MATHPIX_APP_KEY_ACCOUNT,
            mathpix_replacement.as_deref(),
            update.mathpix.clear_app_key,
        )?;
        apply_secret_update(
            PADDLEOCR_ACCESS_TOKEN_ACCOUNT,
            paddle_ocr_replacement.as_deref(),
            update.paddle_ocr.clear_access_token,
        )?;
        apply_secret_update(
            SIMPLETEX_ACCESS_TOKEN_ACCOUNT,
            simpletex_replacement.as_deref(),
            update.simple_tex.clear_access_token,
        )?;
        persist_configuration(app, &next)?;
        let mut cache = self
            .configuration
            .lock()
            .map_err(|_| "OCR provider configuration state is unavailable".to_string())?;
        *cache = Some(next.clone());
        configuration_view(&next)
    }

    pub(crate) async fn recognize_remote(
        &self,
        app: &AppHandle,
        request: &OcrImageRequest,
    ) -> Result<OcrRecognitionResult, String> {
        if request.bytes.is_empty() {
            return Err("The OCR image is empty".to_string());
        }
        if request.bytes.len() > MAX_IMAGE_BYTES {
            return Err("The OCR image is larger than the 20 MB limit".to_string());
        }
        self.remote_recognition_running
            .compare_exchange(false, true, Ordering::SeqCst, Ordering::SeqCst)
            .map_err(|_| "Another remote OCR request is already running".to_string())?;
        let _lease = RemoteRecognitionLease {
            running: self.remote_recognition_running.clone(),
        };
        let configuration = self.load(app)?;
        let has_mathpix_secret = secret_exists(MATHPIX_APP_KEY_ACCOUNT)?;
        let has_paddle_ocr_secret = secret_exists(PADDLEOCR_ACCESS_TOKEN_ACCOUNT)?;
        let has_simpletex_secret = secret_exists(SIMPLETEX_ACCESS_TOKEN_ACCOUNT)?;
        validate_active_provider_configuration(
            &configuration,
            has_mathpix_secret,
            has_paddle_ocr_secret,
            has_simpletex_secret,
        )?;
        let progress = RemoteRecognitionProgress::new(app, &request.model);
        emit_remote_provider_progress(
            Some(&progress),
            &configuration.active_provider,
            "api-submit",
        );
        recognize_remote_inner(&configuration, request, Some(&progress)).await
    }

    fn load(&self, app: &AppHandle) -> Result<StoredOcrProviderConfiguration, String> {
        if let Some(configuration) = self.cached_configuration()? {
            return Ok(configuration);
        }
        let configuration = read_configuration_file(app)?;
        self.cache_configuration(configuration.clone())?;
        Ok(configuration)
    }

    fn load_for_update(&self, app: &AppHandle) -> Result<StoredOcrProviderConfiguration, String> {
        if let Some(configuration) = self.cached_configuration()? {
            return Ok(configuration);
        }
        match read_configuration_file(app) {
            Ok(configuration) => Ok(configuration),
            Err(ConfigurationReadError::Invalid(_)) => {
                Ok(StoredOcrProviderConfiguration::default())
            }
            Err(ConfigurationReadError::Io(error)) => Err(error),
        }
    }

    fn cached_configuration(&self) -> Result<Option<StoredOcrProviderConfiguration>, String> {
        self.configuration
            .lock()
            .map_err(|_| "OCR provider configuration state is unavailable".to_string())
            .map(|configuration| configuration.clone())
    }

    fn cache_configuration(
        &self,
        configuration: StoredOcrProviderConfiguration,
    ) -> Result<(), String> {
        let mut cache = self
            .configuration
            .lock()
            .map_err(|_| "OCR provider configuration state is unavailable".to_string())?;
        *cache = Some(configuration);
        Ok(())
    }
}

fn configuration_view(
    configuration: &StoredOcrProviderConfiguration,
) -> Result<OcrProviderConfigurationView, String> {
    Ok(OcrProviderConfigurationView {
        active_provider: configuration.active_provider.clone(),
        open_ai_compatible: OpenAiCompatibleConfigurationView {
            protocol: configuration.open_ai_compatible.protocol.clone(),
            base_url: configuration.open_ai_compatible.base_url.clone(),
            model: configuration.open_ai_compatible.model.clone(),
            prompt: configuration.open_ai_compatible.prompt.clone(),
            has_api_key: secret_exists(OPENAI_API_KEY_ACCOUNT)?,
        },
        ollama: OllamaConfigurationView {
            base_url: configuration.ollama.base_url.clone(),
            model: configuration.ollama.model.clone(),
            prompt: configuration.ollama.prompt.clone(),
        },
        mathpix: MathpixConfigurationView {
            base_url: configuration.mathpix.base_url.clone(),
            app_id: configuration.mathpix.app_id.clone(),
            has_app_key: secret_exists(MATHPIX_APP_KEY_ACCOUNT)?,
        },
        paddle_ocr: PaddleOcrConfigurationView {
            model: normalize_paddleocr_model(&configuration.paddle_ocr.model)?.to_string(),
            has_access_token: secret_exists(PADDLEOCR_ACCESS_TOKEN_ACCOUNT)?,
        },
        simple_tex: SimpleTexConfigurationView {
            model: normalize_simpletex_model(&configuration.simple_tex.model)?.to_string(),
            has_access_token: secret_exists(SIMPLETEX_ACCESS_TOKEN_ACCOUNT)?,
        },
    })
}

#[derive(Debug)]
enum ConfigurationReadError {
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

fn read_configuration_file(
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

fn configuration_path(app: &AppHandle) -> Result<PathBuf, String> {
    Ok(app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX app-data directory: {error}"))?
        .join("ocr")
        .join(CONFIGURATION_FILE))
}

fn persist_configuration(
    app: &AppHandle,
    configuration: &StoredOcrProviderConfiguration,
) -> Result<(), String> {
    let path = configuration_path(app)?;
    let parent = path
        .parent()
        .ok_or_else(|| "OCR provider configuration path has no parent directory".to_string())?;
    fs::create_dir_all(parent).map_err(|error| {
        format!(
            "Unable to create the OCR provider configuration directory ({}): {error}",
            parent.display()
        )
    })?;
    let bytes = serde_json::to_vec_pretty(configuration)
        .map_err(|error| format!("Unable to serialize OCR provider configuration: {error}"))?;
    fs::write(&path, bytes).map_err(|error| {
        format!(
            "Unable to save the OCR provider configuration ({}): {error}",
            path.display()
        )
    })?;
    #[cfg(unix)]
    {
        use std::os::unix::fs::PermissionsExt;
        let _ = fs::set_permissions(parent, fs::Permissions::from_mode(0o700));
        let _ = fs::set_permissions(&path, fs::Permissions::from_mode(0o600));
    }
    Ok(())
}

fn keychain_entry(account: &str) -> Result<keyring::Entry, String> {
    keyring::Entry::new(KEYCHAIN_SERVICE, account)
        .map_err(|error| format!("Unable to access the macOS Keychain for OCR: {error}"))
}

fn secret_exists(account: &str) -> Result<bool, String> {
    match keychain_entry(account)?.get_password() {
        Ok(mut secret) => {
            let exists = !secret.is_empty();
            unsafe { secret.as_mut_vec().fill(0) };
            Ok(exists)
        }
        Err(keyring::Error::NoEntry) => Ok(false),
        Err(error) => Err(format!(
            "Unable to read the OCR secret from macOS Keychain: {error}"
        )),
    }
}

fn read_secret(account: &str) -> Result<Option<SensitiveString>, String> {
    match keychain_entry(account)?.get_password() {
        Ok(secret) if secret.is_empty() => Ok(None),
        Ok(secret) => Ok(Some(SensitiveString(secret))),
        Err(keyring::Error::NoEntry) => Ok(None),
        Err(error) => Err(format!(
            "Unable to read the OCR secret from macOS Keychain: {error}"
        )),
    }
}

fn normalize_secret_update(value: Option<String>) -> Result<Option<String>, String> {
    let value = value.unwrap_or_default();
    let normalized = value.trim();
    if normalized.is_empty() {
        return Ok(None);
    }
    if normalized.chars().count() > MAX_SECRET_CHARS {
        return Err("OCR provider secret is too long".to_string());
    }
    Ok(Some(normalized.to_string()))
}

fn apply_secret_update(
    account: &str,
    replacement: Option<&str>,
    clear: bool,
) -> Result<(), String> {
    let entry = keychain_entry(account)?;
    if clear {
        match entry.delete_credential() {
            Ok(()) | Err(keyring::Error::NoEntry) => {}
            Err(error) => {
                return Err(format!(
                    "Unable to remove the OCR secret from macOS Keychain: {error}"
                ))
            }
        }
        return Ok(());
    }
    if let Some(secret) = replacement {
        entry
            .set_password(secret)
            .map_err(|error| format!("Unable to save the OCR secret in macOS Keychain: {error}"))?;
    }
    Ok(())
}

struct SensitiveString(String);

impl SensitiveString {
    fn as_str(&self) -> &str {
        &self.0
    }
}

impl Drop for SensitiveString {
    fn drop(&mut self) {
        unsafe { self.0.as_mut_vec().fill(0) };
    }
}

fn normalize_provider(value: &str) -> Result<&'static str, String> {
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

fn normalize_openai_protocol(value: &str) -> Result<&'static str, String> {
    match value.trim().to_ascii_lowercase().as_str() {
        "responses" => Ok("responses"),
        "chat-completions" | "chat_completions" | "chat" => Ok("chat-completions"),
        _ => Err("Unsupported OpenAI-compatible protocol".to_string()),
    }
}

fn normalize_short_text(
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

fn normalize_prompt(value: &str) -> Result<String, String> {
    let normalized = value.trim();
    if normalized.is_empty() {
        return Ok(DEFAULT_PROMPT.to_string());
    }
    if normalized.chars().count() > 8_000 {
        return Err("OCR API prompt is too long".to_string());
    }
    Ok(normalized.to_string())
}

fn normalize_base_url(value: &str) -> Result<String, String> {
    let normalized = value.trim().trim_end_matches('/').to_string();
    if normalized.is_empty() {
        return Err("OCR provider base URL is required".to_string());
    }
    validate_url(&normalized)?;
    Ok(normalized)
}

fn normalize_paddleocr_model(value: &str) -> Result<&'static str, String> {
    match value.trim() {
        "PaddleOCR-VL-1.6" => Ok("PaddleOCR-VL-1.6"),
        _ => Err("Unsupported PaddleOCR AI Studio model".to_string()),
    }
}

fn normalize_simpletex_model(value: &str) -> Result<&'static str, String> {
    match value.trim().to_ascii_lowercase().as_str() {
        "standard" | "latex_ocr" => Ok("standard"),
        "turbo" | "latex_ocr_turbo" => Ok("turbo"),
        _ => Err("Unsupported SimpleTex formula model".to_string()),
    }
}

fn simpletex_endpoint(model: &str) -> Result<&'static str, String> {
    match normalize_simpletex_model(model)? {
        "standard" => Ok(SIMPLETEX_STANDARD_URL),
        "turbo" => Ok(SIMPLETEX_TURBO_URL),
        _ => Err("Unsupported SimpleTex formula model".to_string()),
    }
}

fn migrate_legacy_paddleocr_configuration(
    mut configuration: StoredOcrProviderConfiguration,
) -> StoredOcrProviderConfiguration {
    if configuration.schema_version == 1 {
        configuration.paddle_ocr.model = match configuration.paddle_ocr.model.as_str() {
            "PaddleOCR-VL-1.5" | "PaddleOCR-VL" | "PP-StructureV3" => {
                "PaddleOCR-VL-1.6".to_string()
            }
            "PaddleOCR-VL-1.6" => "PaddleOCR-VL-1.6".to_string(),
            other => other.to_string(),
        };
        configuration.schema_version = CONFIGURATION_SCHEMA_VERSION;
    }
    configuration
}

fn validate_stored_configuration(
    configuration: &StoredOcrProviderConfiguration,
) -> Result<(), String> {
    if configuration.schema_version != CONFIGURATION_SCHEMA_VERSION {
        return Err(format!(
            "Unsupported OCR provider configuration schema: {}",
            configuration.schema_version
        ));
    }
    normalize_provider(&configuration.active_provider)?;
    normalize_openai_protocol(&configuration.open_ai_compatible.protocol)?;
    validate_url(&configuration.open_ai_compatible.base_url)?;
    validate_url(&configuration.ollama.base_url)?;
    validate_url(&configuration.mathpix.base_url)?;
    normalize_paddleocr_model(&configuration.paddle_ocr.model)?;
    normalize_simpletex_model(&configuration.simple_tex.model)?;
    Ok(())
}

fn validate_active_provider_configuration(
    configuration: &StoredOcrProviderConfiguration,
    has_mathpix_secret: bool,
    has_paddle_ocr_secret: bool,
    has_simpletex_secret: bool,
) -> Result<(), String> {
    validate_stored_configuration(configuration)?;
    match configuration.active_provider.as_str() {
        LOCAL_PROVIDER => Ok(()),
        OPENAI_COMPATIBLE_PROVIDER => {
            normalize_short_text(
                &configuration.open_ai_compatible.model,
                "OpenAI-compatible model",
                160,
                true,
            )?;
            Ok(())
        }
        OLLAMA_PROVIDER => {
            normalize_short_text(&configuration.ollama.model, "Ollama model", 160, true)?;
            Ok(())
        }
        MATHPIX_PROVIDER => {
            normalize_short_text(&configuration.mathpix.app_id, "Mathpix app_id", 240, true)?;
            if !has_mathpix_secret {
                return Err("Mathpix app_key is required".to_string());
            }
            Ok(())
        }
        PADDLEOCR_PROVIDER => {
            normalize_paddleocr_model(&configuration.paddle_ocr.model)?;
            if !has_paddle_ocr_secret {
                return Err("PaddleOCR AI Studio access token is required".to_string());
            }
            Ok(())
        }
        SIMPLETEX_PROVIDER => {
            normalize_simpletex_model(&configuration.simple_tex.model)?;
            if !has_simpletex_secret {
                return Err("SimpleTex user access token is required".to_string());
            }
            Ok(())
        }
        _ => Err("Unsupported OCR provider".to_string()),
    }
}

fn validate_url(value: &str) -> Result<(), String> {
    let url =
        reqwest::Url::parse(value).map_err(|error| format!("Invalid OCR provider URL: {error}"))?;
    if url.scheme() != "http" && url.scheme() != "https" {
        return Err("OCR provider URL must use http or https".to_string());
    }
    if url.host_str().is_none() {
        return Err("OCR provider URL must include a host".to_string());
    }
    if !url.username().is_empty() || url.password().is_some() {
        return Err("OCR provider URL must not contain embedded credentials".to_string());
    }
    if url.query().is_some() || url.fragment().is_some() {
        return Err("OCR provider base URL must not contain a query or fragment".to_string());
    }
    Ok(())
}

fn is_loopback_url(value: &str) -> Result<bool, String> {
    let url =
        reqwest::Url::parse(value).map_err(|error| format!("Invalid OCR provider URL: {error}"))?;
    let Some(host) = url.host_str() else {
        return Ok(false);
    };
    Ok(host.eq_ignore_ascii_case("localhost")
        || host == "::1"
        || host.starts_with("127.")
        || host == "0:0:0:0:0:0:0:1")
}

fn validate_secret_transport(value: &str, has_secret: bool, label: &str) -> Result<(), String> {
    if has_secret && value.starts_with("http://") && !is_loopback_url(value)? {
        return Err(format!(
            "Refusing to send {label} over plain HTTP to a non-local OCR endpoint"
        ));
    }
    Ok(())
}

#[derive(Debug, Clone, Default, PartialEq, Eq)]
struct MacosSystemProxySettings {
    https_proxy: Option<String>,
}

fn parse_scutil_proxy_value(output: &str, key: &str) -> Option<String> {
    output.lines().find_map(|line| {
        let (candidate, value) = line.trim().split_once(':')?;
        (candidate.trim() == key).then(|| value.trim().to_string())
    })
}

fn format_proxy_url(host: &str, port: &str) -> Option<String> {
    let host = host.trim();
    let port = port.trim();
    if host.is_empty()
        || port
            .parse::<u16>()
            .ok()
            .filter(|value| *value > 0)
            .is_none()
    {
        return None;
    }
    let host = if host.contains(':') && !host.starts_with('[') {
        format!("[{host}]")
    } else {
        host.to_string()
    };
    Some(format!("http://{host}:{port}"))
}

fn parse_macos_system_proxy_settings(output: &str) -> MacosSystemProxySettings {
    let https_enabled =
        parse_scutil_proxy_value(output, "HTTPSEnable").is_some_and(|value| value == "1");
    let https_proxy = https_enabled
        .then(|| {
            let host = parse_scutil_proxy_value(output, "HTTPSProxy")?;
            let port = parse_scutil_proxy_value(output, "HTTPSPort")?;
            format_proxy_url(&host, &port)
        })
        .flatten();
    MacosSystemProxySettings { https_proxy }
}

#[cfg(target_os = "macos")]
fn macos_system_proxy_settings() -> MacosSystemProxySettings {
    let output = Command::new("/usr/sbin/scutil").arg("--proxy").output();
    let Ok(output) = output else {
        return MacosSystemProxySettings::default();
    };
    if !output.status.success() {
        return MacosSystemProxySettings::default();
    }
    parse_macos_system_proxy_settings(&String::from_utf8_lossy(&output.stdout))
}

#[cfg(not(target_os = "macos"))]
fn macos_system_proxy_settings() -> MacosSystemProxySettings {
    MacosSystemProxySettings::default()
}

fn apply_macos_system_https_proxy(
    mut builder: reqwest::ClientBuilder,
) -> Result<reqwest::ClientBuilder, String> {
    if let Some(proxy_url) = macos_system_proxy_settings().https_proxy {
        let proxy = reqwest::Proxy::https(&proxy_url)
            .map_err(|error| format!("Unable to configure the macOS HTTPS proxy: {error}"))?;
        builder = builder.proxy(proxy);
    }
    Ok(builder)
}

fn build_ocr_http_client(
    timeout: Duration,
    redirect: reqwest::redirect::Policy,
    user_agent: String,
) -> Result<reqwest::Client, String> {
    let builder = reqwest::Client::builder()
        .connect_timeout(Duration::from_secs(15))
        .timeout(timeout)
        .redirect(redirect)
        .user_agent(user_agent);
    apply_macos_system_https_proxy(builder)?
        .build()
        .map_err(|error| format!("Unable to initialize the OCR API client: {error}"))
}

async fn recognize_remote_inner(
    configuration: &StoredOcrProviderConfiguration,
    request: &OcrImageRequest,
    progress: Option<&RemoteRecognitionProgress>,
) -> Result<OcrRecognitionResult, String> {
    let started = Instant::now();
    let mime = mime_for_extension(&request.extension)?;
    let data_base64 = BASE64_STANDARD.encode(&request.bytes);
    let data_url = format!("data:{mime};base64,{data_base64}");
    let (processed_width, processed_height) = image_dimensions(&request.bytes);
    let client = build_ocr_http_client(
        Duration::from_secs(120),
        reqwest::redirect::Policy::none(),
        format!("VisualTeX/{} OCR", env!("CARGO_PKG_VERSION")),
    )?;

    let (provider, model, formulas) = match configuration.active_provider.as_str() {
        OPENAI_COMPATIBLE_PROVIDER => {
            let api_key = read_secret(OPENAI_API_KEY_ACCOUNT)?;
            validate_secret_transport(
                &configuration.open_ai_compatible.base_url,
                api_key.is_some(),
                "API key",
            )?;
            emit_remote_provider_progress(
                progress,
                OPENAI_COMPATIBLE_PROVIDER,
                "api-inference",
            );
            let formulas = recognize_openai_compatible(
                &client,
                &configuration.open_ai_compatible,
                api_key.as_ref().map(SensitiveString::as_str),
                &data_url,
            )
            .await?;
            (
                OPENAI_COMPATIBLE_PROVIDER,
                configuration.open_ai_compatible.model.clone(),
                formulas,
            )
        }
        OLLAMA_PROVIDER => {
            emit_remote_provider_progress(progress, OLLAMA_PROVIDER, "api-inference");
            let formulas = recognize_ollama(&client, &configuration.ollama, &data_base64).await?;
            (
                OLLAMA_PROVIDER,
                configuration.ollama.model.clone(),
                formulas,
            )
        }
        MATHPIX_PROVIDER => {
            if data_url.len() > MATHPIX_MAX_BASE64_IMAGE_BYTES {
                return Err("Mathpix accepts at most 2 MB for a base64-encoded image. Crop the formula more tightly or use a smaller PNG/JPEG image.".to_string());
            }
            let app_key = read_secret(MATHPIX_APP_KEY_ACCOUNT)?
                .ok_or_else(|| "Mathpix app_key is required".to_string())?;
            validate_secret_transport(&configuration.mathpix.base_url, true, "app_key")?;
            emit_remote_provider_progress(progress, MATHPIX_PROVIDER, "api-inference");
            let formulas =
                recognize_mathpix(&client, &configuration.mathpix, app_key.as_str(), &data_url)
                    .await?;
            (MATHPIX_PROVIDER, "Mathpix Text API".to_string(), formulas)
        }
        PADDLEOCR_PROVIDER => {
            let access_token = read_secret(PADDLEOCR_ACCESS_TOKEN_ACCOUNT)?
                .ok_or_else(|| "PaddleOCR AI Studio access token is required".to_string())?;
            let paddle_model = normalize_paddleocr_model(&configuration.paddle_ocr.model)?;
            let formulas = recognize_paddleocr_aistudio(
                &client,
                &configuration.paddle_ocr,
                access_token.as_str(),
                &request.bytes,
                mime,
                &request.extension,
                progress,
            )
            .await?;
            (PADDLEOCR_PROVIDER, paddle_model.to_string(), formulas)
        }
        SIMPLETEX_PROVIDER => {
            let access_token = read_secret(SIMPLETEX_ACCESS_TOKEN_ACCOUNT)?
                .ok_or_else(|| "SimpleTex user access token is required".to_string())?;
            let simpletex_model = normalize_simpletex_model(&configuration.simple_tex.model)?;
            emit_remote_provider_progress(progress, SIMPLETEX_PROVIDER, "api-inference");
            let formulas = recognize_simpletex(
                &client,
                &configuration.simple_tex,
                access_token.as_str(),
                &request.bytes,
                mime,
                &request.extension,
            )
            .await?;
            (
                SIMPLETEX_PROVIDER,
                format!("SimpleTex V2.5 ({simpletex_model})"),
                formulas,
            )
        }
        _ => return Err("The active OCR provider is local, not remote".to_string()),
    };
    if formulas.is_empty() {
        return Err("OCR API returned no usable formula".to_string());
    }
    if provider != PADDLEOCR_PROVIDER {
        emit_remote_provider_progress(progress, provider, "api-result");
    }
    Ok(OcrRecognitionResult {
        provider: provider.to_string(),
        model,
        elapsed_ms: started.elapsed().as_millis().min(u64::MAX as u128) as u64,
        processed_width,
        processed_height,
        background_inverted: false,
        background_luminance: 0.0,
        formulas: formulas
            .into_iter()
            .map(|latex| OcrFormulaResult { latex })
            .collect(),
    })
}

async fn recognize_openai_compatible(
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

async fn recognize_ollama(
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

async fn recognize_simpletex(
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
async fn recognize_simpletex_at(
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

fn parse_simpletex_formula_output(value: &Value) -> Result<Vec<String>, String> {
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

fn paddleocr_progress_for_state(state: &str) -> Option<(&'static str, &'static str)> {
    match state {
        "pending" => Some(("api-queued", "PaddleOCR 任务正在排队…")),
        "running" => Some(("api-inference", "PaddleOCR 正在识别公式…")),
        "done" => Some(("api-result", "PaddleOCR 识别完成，正在读取结果…")),
        _ => None,
    }
}

async fn recognize_paddleocr_aistudio(
    client: &reqwest::Client,
    configuration: &StoredPaddleOcrConfiguration,
    access_token: &str,
    image_bytes: &[u8],
    mime: &str,
    extension: &str,
    progress: Option<&RemoteRecognitionProgress>,
) -> Result<Vec<String>, String> {
    recognize_paddleocr_aistudio_at_with_progress(
        client,
        configuration,
        access_token,
        image_bytes,
        mime,
        extension,
        PADDLEOCR_JOBS_URL,
        progress,
    )
    .await
}

#[cfg_attr(not(test), allow(dead_code))]
async fn recognize_paddleocr_aistudio_at(
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
async fn recognize_paddleocr_aistudio_at_with_progress(
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
    let optional_payload = json!({
        "useDocOrientationClassify": false,
        "useDocUnwarping": false,
        "useLayoutDetection": true,
        "useChartRecognition": false,
        "showFormulaNumber": false,
        "prettifyMarkdown": false
    });
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
        .text("model", model.to_string())
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
        return Err("PaddleOCR AI Studio daily quota has been exhausted. The official quota is currently 3000 pages per user per model per day; try again after the quota resets or select another OCR provider.".to_string());
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
                return Err(format!(
                    "PaddleOCR AI Studio returned an unknown job state: {other}"
                ));
            }
        }
    }
    Err("PaddleOCR AI Studio did not finish within 120 seconds".to_string())
}

fn ensure_paddleocr_api_code(value: &Value, operation: &str) -> Result<(), String> {
    if value.get("code").and_then(Value::as_i64).unwrap_or(0) == 0 {
        return Ok(());
    }
    Err(format!(
        "{operation} failed: {}",
        extract_api_error(value).unwrap_or_else(|| "unknown API error".to_string())
    ))
}

async fn download_and_parse_paddleocr_result_with_client(
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
    let response =
        read_http_response(response, PADDLEOCR_RESULT_MAX_BYTES, "PaddleOCR result").await?;
    ensure_success(&response, "PaddleOCR result download")?;
    parse_paddleocr_result_text(&response.text)
}

fn validate_paddleocr_result_url(value: &str) -> Result<(), String> {
    let url = reqwest::Url::parse(value)
        .map_err(|error| format!("Invalid PaddleOCR result URL: {error}"))?;
    if url.username() != "" || url.password().is_some() {
        return Err("PaddleOCR result URL must not contain embedded credentials".to_string());
    }
    if url.scheme() == "https" {
        return Ok(());
    }
    #[cfg(test)]
    if url.scheme() == "http" && is_loopback_url(value)? {
        return Ok(());
    }
    Err("PaddleOCR result URL must use HTTPS".to_string())
}

fn parse_paddleocr_result_text(text: &str) -> Result<Vec<String>, String> {
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

async fn recognize_mathpix(
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

fn parse_paddleocr_formula_output(value: &Value) -> Result<Vec<String>, String> {
    let pages = value
        .pointer("/result/layoutParsingResults")
        .and_then(Value::as_array)
        .ok_or_else(|| {
            "PaddleOCR AI Studio response contains no layoutParsingResults".to_string()
        })?;

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

fn extract_markdown_math(markdown: &str) -> Vec<String> {
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

fn collect_paddleocr_structured_formulas(value: &Value, output: &mut Vec<String>) {
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

fn push_unique_formula_candidate(output: &mut Vec<String>, candidate: String) {
    let normalized = candidate.trim();
    if normalized.is_empty() || output.iter().any(|item| item.trim() == normalized) {
        return;
    }
    output.push(normalized.to_string());
}

fn append_endpoint(base_url: &str, route: &str) -> String {
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

struct HttpJsonResponse {
    status: u16,
    text: String,
    json: Option<Value>,
}

async fn read_http_response(
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
    let mut bytes =
        Vec::with_capacity(response.content_length().unwrap_or(0).min(max_bytes as u64) as usize);
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

async fn post_json(
    client: &reqwest::Client,
    endpoint: &str,
    bearer_token: Option<&str>,
    headers: &[(&str, &str)],
    body: &Value,
) -> Result<HttpJsonResponse, String> {
    let payload = serde_json::to_vec(body)
        .map_err(|error| format!("Unable to encode OCR API request: {error}"))?;
    let mut request = client
        .post(endpoint)
        .header("content-type", "application/json")
        .body(payload);
    if let Some(token) = bearer_token.filter(|token| !token.trim().is_empty()) {
        request = request.bearer_auth(token);
    }
    for (name, value) in headers {
        request = request.header(*name, *value);
    }
    let mut response = request
        .send()
        .await
        .map_err(|error| format!("OCR API request failed: {error}"))?;
    let status = response.status().as_u16();
    if response
        .content_length()
        .is_some_and(|length| length > MAX_RESPONSE_BYTES as u64)
    {
        return Err("OCR API response is larger than the 4 MB limit".to_string());
    }
    let mut bytes = Vec::with_capacity(
        response
            .content_length()
            .unwrap_or(0)
            .min(MAX_RESPONSE_BYTES as u64) as usize,
    );
    while let Some(chunk) = response
        .chunk()
        .await
        .map_err(|error| format!("Unable to read OCR API response: {error}"))?
    {
        if bytes.len().saturating_add(chunk.len()) > MAX_RESPONSE_BYTES {
            return Err("OCR API response is larger than the 4 MB limit".to_string());
        }
        bytes.extend_from_slice(&chunk);
    }
    let text = String::from_utf8_lossy(&bytes).trim().to_string();
    let json = serde_json::from_slice::<Value>(&bytes).ok();
    Ok(HttpJsonResponse { status, text, json })
}

fn ensure_success(response: &HttpJsonResponse, provider: &str) -> Result<(), String> {
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

fn extract_api_error(value: &Value) -> Option<String> {
    value
        .pointer("/error/message")
        .and_then(Value::as_str)
        .or_else(|| value.get("error").and_then(Value::as_str))
        .or_else(|| value.get("message").and_then(Value::as_str))
        .or_else(|| value.get("errorMsg").and_then(Value::as_str))
        .or_else(|| value.get("msg").and_then(Value::as_str))
        .map(|message| truncate_for_error(message, 800))
}

fn structured_output_rejected(text: &str) -> bool {
    let normalized = text.to_ascii_lowercase();
    normalized.contains("response_format")
        || normalized.contains("json_schema")
        || normalized.contains("structured output")
        || normalized.contains("unknown field") && normalized.contains("format")
}

fn extract_responses_text(value: &Value) -> Option<String> {
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

fn extract_chat_completions_text(value: &Value) -> Option<String> {
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

fn formula_json_schema() -> Value {
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

fn parse_formula_output(raw: &str) -> Result<Vec<String>, String> {
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

fn formula_candidates_from_json(value: &Value) -> Vec<String> {
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

fn normalize_formula_candidates(candidates: Vec<String>) -> Result<Vec<String>, String> {
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

fn strip_markdown_fence(value: &str) -> &str {
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

fn extract_embedded_json(value: &str) -> Option<Value> {
    let start = value.find('{')?;
    let end = value.rfind('}')?;
    (end > start)
        .then(|| serde_json::from_str::<Value>(&value[start..=end]).ok())
        .flatten()
}

fn strip_formula_wrapper(value: &str) -> &str {
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

fn truncate_for_error(value: &str, max_chars: usize) -> String {
    let mut result = value.chars().take(max_chars).collect::<String>();
    if value.chars().count() > max_chars {
        result.push('…');
    }
    result
}

fn mime_for_extension(extension: &str) -> Result<&'static str, String> {
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

fn image_dimensions(bytes: &[u8]) -> (u32, u32) {
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

fn jpeg_dimensions(bytes: &[u8]) -> Option<(u32, u32)> {
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

#[cfg(test)]
mod tests {
    use super::*;
    use axum::{
        body::Bytes,
        extract::{Json, Path},
        http::{HeaderMap, StatusCode},
        routing::{get, post},
        Router,
    };
    use tokio::net::TcpListener;

    async fn spawn_mock_api() -> String {
        async fn openai(Json(body): Json<Value>) -> Json<Value> {
            assert_eq!(
                body.get("model").and_then(Value::as_str),
                Some("vision-model")
            );
            assert!(body
                .pointer("/input/0/content/1/image_url")
                .and_then(Value::as_str)
                .is_some_and(|url| url.starts_with("data:image/png;base64,")));
            Json(json!({
                "output_text": "{\"formulas\":[{\"latex\":\"x^2+y^2\"},{\"latex\":\"z=3\"}]}"
            }))
        }

        async fn ollama(Json(body): Json<Value>) -> Json<Value> {
            assert_eq!(
                body.get("model").and_then(Value::as_str),
                Some("vision-model")
            );
            assert_eq!(body.get("stream").and_then(Value::as_bool), Some(false));
            assert!(body
                .pointer("/messages/0/images/0")
                .and_then(Value::as_str)
                .is_some_and(|image| !image.is_empty()));
            Json(json!({
                "message": {
                    "content": "{\"formulas\":[{\"latex\":\"a=b\"}]}"
                }
            }))
        }

        fn assert_paddle_bearer(headers: &HeaderMap) {
            assert_eq!(
                headers
                    .get("authorization")
                    .and_then(|value| value.to_str().ok()),
                Some("Bearer paddle-token")
            );
        }

        async fn paddle_vl_submit(headers: HeaderMap, body: Bytes) -> Json<Value> {
            assert_paddle_bearer(&headers);
            let body = String::from_utf8_lossy(&body);
            assert!(body.contains("PaddleOCR-VL-1.6"));
            assert!(body.contains("optionalPayload"));
            assert!(body.contains("showFormulaNumber"));
            assert!(body.contains("prettifyMarkdown"));
            assert!(body.contains("visualtex-formula.png"));
            Json(json!({"code": 0, "msg": "Success", "data": {"jobId": "vl-job"}}))
        }

        async fn simpletex(headers: HeaderMap, body: Bytes) -> Json<Value> {
            assert_eq!(
                headers.get("token").and_then(|value| value.to_str().ok()),
                Some("simpletex-token")
            );
            let body = String::from_utf8_lossy(&body);
            assert!(body.contains("visualtex-formula.png"));
            Json(json!({
                "status": true,
                "res": {"latex": "\\sqrt{x^2+y^2}", "conf": 0.99},
                "request_id": "tr_visualtex_mock"
            }))
        }

        async fn simpletex_slow() -> Json<Value> {
            tokio::time::sleep(Duration::from_millis(200)).await;
            Json(json!({"status": true, "res": {"latex": "x"}}))
        }

        async fn paddle_status(Path(job_id): Path<String>, headers: HeaderMap) -> Json<Value> {
            assert_paddle_bearer(&headers);
            let host = headers
                .get("host")
                .and_then(|value| value.to_str().ok())
                .unwrap();
            Json(json!({
                "code": 0,
                "msg": "Success",
                "data": {
                    "jobId": job_id,
                    "state": "done",
                    "extractProgress": {"totalPages": 1, "extractedPages": 1},
                    "resultUrl": {
                        "jsonUrl": format!("http://{host}/paddle-result/{job_id}")
                    }
                }
            }))
        }

        async fn paddle_result(Path(_job_id): Path<String>) -> String {
            let result = json!({
                "result": {
                    "layoutParsingResults": [{
                        "prunedResult": {
                            "parsing_res_list": [
                                {"block_label": "text", "block_content": "ignored"},
                                {"block_label": "formula", "block_content": "$$x&=1\\\\y&=2$$"},
                                {"block_label": "equation", "block_content": "\\frac{a}{b}"}
                            ]
                        },
                        "markdown": {"text": "ignored markdown", "images": {}}
                    }]
                }
            });
            format!("{}\n", serde_json::to_string(&result).unwrap())
        }

        async fn paddle_quota() -> (StatusCode, Json<Value>) {
            (
                StatusCode::TOO_MANY_REQUESTS,
                Json(json!({"code": 429, "msg": "quota exceeded"})),
            )
        }

        async fn mathpix(headers: HeaderMap, Json(body): Json<Value>) -> Json<Value> {
            assert_eq!(
                headers.get("app_id").and_then(|value| value.to_str().ok()),
                Some("app-id")
            );
            assert_eq!(
                headers.get("app_key").and_then(|value| value.to_str().ok()),
                Some("app-key")
            );
            assert!(body
                .get("src")
                .and_then(Value::as_str)
                .is_some_and(|src| src.starts_with("data:image/png;base64,")));
            Json(json!({"latex_styled": "\\frac{1}{2}"}))
        }

        let router = Router::new()
            .route("/responses", post(openai))
            .route("/api/chat", post(ollama))
            .route("/simpletex", post(simpletex))
            .route("/simpletex-slow", post(simpletex_slow))
            .route("/jobs-vl", post(paddle_vl_submit))
            .route("/jobs-vl/{job_id}", get(paddle_status))
            .route("/paddle-result/{job_id}", get(paddle_result))
            .route("/quota", post(paddle_quota))
            .route("/v3/text", post(mathpix));
        let listener = TcpListener::bind("127.0.0.1:0").await.unwrap();
        let address = listener.local_addr().unwrap();
        tokio::spawn(async move {
            axum::serve(listener, router).await.unwrap();
        });
        format!("http://{address}")
    }

    #[test]
    fn default_configuration_keeps_local_ocr_active() {
        let configuration = StoredOcrProviderConfiguration::default();
        assert_eq!(configuration.active_provider, LOCAL_PROVIDER);
        assert_eq!(configuration.open_ai_compatible.protocol, "responses");
        assert!(!configuration.open_ai_compatible.prompt.is_empty());
    }

    #[test]
    fn legacy_provider_configuration_deserializes_without_paddle_fields() {
        let stored = serde_json::from_value::<StoredOcrProviderConfiguration>(json!({
            "schemaVersion": 1,
            "activeProvider": "local",
            "openAiCompatible": {
                "protocol": "responses",
                "baseUrl": "https://api.openai.com/v1",
                "model": "",
                "prompt": "prompt"
            },
            "ollama": {
                "baseUrl": "http://127.0.0.1:11434",
                "model": "",
                "prompt": "prompt"
            },
            "mathpix": {
                "baseUrl": "https://api.mathpix.com",
                "appId": ""
            }
        }))
        .unwrap();
        assert_eq!(stored.paddle_ocr.model, "PaddleOCR-VL-1.6");
        assert_eq!(stored.simple_tex.model, "standard");

        let migrated = migrate_legacy_paddleocr_configuration(
            serde_json::from_value::<StoredOcrProviderConfiguration>(json!({
                "schemaVersion": 1,
                "activeProvider": "paddleocr",
                "openAiCompatible": {
                    "protocol": "responses",
                    "baseUrl": "https://api.openai.com/v1",
                    "model": "",
                    "prompt": "prompt"
                },
                "ollama": {
                    "baseUrl": "http://127.0.0.1:11434",
                    "model": "",
                    "prompt": "prompt"
                },
                "mathpix": {
                    "baseUrl": "https://api.mathpix.com",
                    "appId": ""
                },
                "paddleOcr": {
                    "apiUrl": "https://example.test/legacy-v15",
                    "model": "PaddleOCR-VL-1.5"
                }
            }))
            .unwrap(),
        );
        assert_eq!(migrated.paddle_ocr.model, "PaddleOCR-VL-1.6");

        let update = serde_json::from_value::<OcrProviderConfigurationUpdate>(json!({
            "activeProvider": "local",
            "openAiCompatible": {
                "protocol": "responses",
                "baseUrl": "https://api.openai.com/v1",
                "model": "",
                "prompt": "prompt"
            },
            "ollama": {
                "baseUrl": "http://127.0.0.1:11434",
                "model": "",
                "prompt": "prompt"
            },
            "mathpix": {
                "baseUrl": "https://api.mathpix.com",
                "appId": ""
            }
        }))
        .unwrap();
        assert_eq!(update.paddle_ocr.model, "PaddleOCR-VL-1.6");
        assert_eq!(update.simple_tex.model, "standard");
    }

    #[test]
    fn parses_structured_and_fenced_formula_outputs() {
        assert_eq!(
            parse_formula_output(r#"{"formulas":[{"latex":"x^2+y^2"},{"latex":"\\frac{a}{b}"}]}"#)
                .unwrap(),
            vec!["x^2+y^2", "\\frac{a}{b}"]
        );
        assert_eq!(
            parse_formula_output("```json\n{\"latex\":\"$$x+y$$\"}\n```").unwrap(),
            vec!["x+y"]
        );
    }

    #[test]
    fn endpoint_append_is_idempotent() {
        assert_eq!(
            append_endpoint("https://api.openai.com/v1", "responses"),
            "https://api.openai.com/v1/responses"
        );
        assert_eq!(
            append_endpoint("http://127.0.0.1:11434/api/chat", "api/chat"),
            "http://127.0.0.1:11434/api/chat"
        );
    }

    #[test]
    fn png_dimensions_are_read_without_decoding_the_image() {
        let mut bytes = vec![0u8; 24];
        bytes[..8].copy_from_slice(b"\x89PNG\r\n\x1a\n");
        bytes[16..20].copy_from_slice(&640u32.to_be_bytes());
        bytes[20..24].copy_from_slice(&480u32.to_be_bytes());
        assert_eq!(image_dimensions(&bytes), (640, 480));
    }

    #[test]
    fn wrappers_are_removed_only_around_the_complete_formula() {
        assert_eq!(strip_formula_wrapper("$$x+y$$"), "x+y");
        assert_eq!(strip_formula_wrapper("$x$+$y$"), "x$+$y");
        assert_eq!(strip_formula_wrapper("\\[x+y\\]"), "x+y");
    }

    #[test]
    fn paddleocr_formula_parser_prefers_structured_blocks_and_falls_back_to_markdown() {
        let structured = json!({
            "result": {
                "layoutParsingResults": [{
                    "prunedResult": {
                        "parsing_res_list": [
                            {"block_label": "text", "block_content": "ignore me"},
                            {"block_label": "formula", "block_content": "$$x+y$$"},
                            {"rec_formula": "\\frac{1}{2}"}
                        ]
                    },
                    "markdown": {"text": "text $$z=3$$"}
                }]
            }
        });
        assert_eq!(
            parse_paddleocr_formula_output(&structured).unwrap(),
            vec!["x+y", "\\frac{1}{2}"]
        );

        let markdown_only = json!({
            "result": {
                "layoutParsingResults": [{
                    "prunedResult": {"parsing_res_list": [{"block_label": "text", "block_content": "plain"}]},
                    "markdown": {"text": "before $$a=1$$ and \\[b=2\\] after"}
                }]
            }
        });
        assert_eq!(
            parse_paddleocr_formula_output(&markdown_only).unwrap(),
            vec!["a=1", "b=2"]
        );
    }

    #[test]
    fn simpletex_formula_parser_validates_status_and_empty_results() {
        assert_eq!(
            simpletex_endpoint("standard").unwrap(),
            SIMPLETEX_STANDARD_URL
        );
        assert_eq!(simpletex_endpoint("turbo").unwrap(), SIMPLETEX_TURBO_URL);
        assert_eq!(
            parse_simpletex_formula_output(&json!({
                "status": true,
                "res": {"latex": "$$\\frac{1}{2}$$", "conf": 0.98}
            }))
            .unwrap(),
            vec!["\\frac{1}{2}"]
        );
        assert!(parse_simpletex_formula_output(&json!({
            "status": true,
            "res": {"latex": "[EMPTY]"}
        }))
        .unwrap_err()
        .contains("no usable formula"));
        assert!(parse_simpletex_formula_output(&json!({
            "status": false,
            "errInfo": {"msg": "invalid token"}
        }))
        .unwrap_err()
        .contains("failed"));
    }

    #[test]
    fn paddleocr_remote_states_map_to_visible_progress_stages() {
        assert_eq!(
            paddleocr_progress_for_state("pending"),
            Some(("api-queued", "PaddleOCR 任务正在排队…"))
        );
        assert_eq!(
            paddleocr_progress_for_state("running"),
            Some(("api-inference", "PaddleOCR 正在识别公式…"))
        );
        assert_eq!(
            paddleocr_progress_for_state("done"),
            Some(("api-result", "PaddleOCR 识别完成，正在读取结果…"))
        );
        assert_eq!(paddleocr_progress_for_state("failed"), None);
    }

    #[test]
    fn every_remote_api_has_visible_progress_messages() {
        for provider in [
            OPENAI_COMPATIBLE_PROVIDER,
            OLLAMA_PROVIDER,
            MATHPIX_PROVIDER,
            SIMPLETEX_PROVIDER,
        ] {
            for stage in ["api-submit", "api-inference", "api-result"] {
                assert!(
                    remote_provider_progress_message(provider, stage).is_some(),
                    "missing {stage} progress for {provider}"
                );
            }
        }
        assert!(
            remote_provider_progress_message(PADDLEOCR_PROVIDER, "api-submit").is_some()
        );
        assert_eq!(
            remote_provider_progress_message(LOCAL_PROVIDER, "api-submit"),
            None
        );
    }

    #[test]
    fn macos_system_https_proxy_is_parsed_from_scutil_output() {
        let settings = parse_macos_system_proxy_settings(
            r#"<dictionary> {
  HTTPEnable : 1
  HTTPPort : 7890
  HTTPProxy : 127.0.0.1
  HTTPSEnable : 1
  HTTPSPort : 7890
  HTTPSProxy : 127.0.0.1
  SOCKSEnable : 1
}"#,
        );
        assert_eq!(
            settings.https_proxy.as_deref(),
            Some("http://127.0.0.1:7890")
        );
        assert_eq!(
            parse_macos_system_proxy_settings(
                "HTTPSEnable : 0\nHTTPSPort : 7890\nHTTPSProxy : 127.0.0.1"
            )
            .https_proxy,
            None
        );
    }

    #[test]
    fn provider_urls_enforce_transport_safety() {
        assert!(validate_url("https://api.openai.com/v1").is_ok());
        assert!(validate_url("http://127.0.0.1:11434").is_ok());
        assert!(validate_url("file:///tmp/ocr").is_err());
        assert!(validate_url("https://user:pass@example.com/v1").is_err());
        assert!(validate_url("https://example.com/v1?token=secret").is_err());
        assert!(validate_secret_transport("http://127.0.0.1:11434", true, "API key").is_ok());
        assert!(validate_secret_transport("http://example.com", true, "API key").is_err());
    }

    #[tokio::test]
    async fn remote_provider_http_protocols_work_on_macos() {
        let base_url = spawn_mock_api().await;
        let client = reqwest::Client::builder()
            .redirect(reqwest::redirect::Policy::none())
            .build()
            .unwrap();
        let data_url = "data:image/png;base64,iVBORw0KGgo=";

        let openai_configuration = StoredOpenAiCompatibleConfiguration {
            protocol: "responses".to_string(),
            base_url: base_url.clone(),
            model: "vision-model".to_string(),
            prompt: "Return formula JSON".to_string(),
        };
        assert_eq!(
            recognize_openai_compatible(&client, &openai_configuration, None, data_url,)
                .await
                .unwrap(),
            vec!["x^2+y^2", "z=3"]
        );

        let ollama_configuration = StoredOllamaConfiguration {
            base_url: base_url.clone(),
            model: "vision-model".to_string(),
            prompt: "Return formula JSON".to_string(),
        };
        assert_eq!(
            recognize_ollama(&client, &ollama_configuration, "iVBORw0KGgo=")
                .await
                .unwrap(),
            vec!["a=b"]
        );

        let simpletex_configuration = StoredSimpleTexConfiguration {
            model: "standard".to_string(),
        };
        assert_eq!(
            recognize_simpletex_at(
                &client,
                &simpletex_configuration,
                "simpletex-token",
                b"fake-png",
                "image/png",
                "png",
                &format!("{base_url}/simpletex"),
                Duration::from_secs(2),
            )
            .await
            .unwrap(),
            vec!["\\sqrt{x^2+y^2}"]
        );
        let timeout_error = recognize_simpletex_at(
            &client,
            &simpletex_configuration,
            "simpletex-token",
            b"fake-png",
            "image/png",
            "png",
            &format!("{base_url}/simpletex-slow"),
            Duration::from_millis(25),
        )
        .await
        .expect_err("SimpleTex timeout mock must fail");
        assert!(timeout_error.to_ascii_lowercase().contains("timed out"));

        let paddle_configuration = StoredPaddleOcrConfiguration {
            model: "PaddleOCR-VL-1.6".to_string(),
        };
        assert_eq!(
            recognize_paddleocr_aistudio_at(
                &client,
                &paddle_configuration,
                "paddle-token",
                b"fake-png",
                "image/png",
                "png",
                &format!("{base_url}/jobs-vl"),
            )
            .await
            .unwrap(),
            vec!["x&=1\\\\y&=2", "\\frac{a}{b}"]
        );

        assert!(normalize_paddleocr_model("PP-StructureV3").is_err());

        let quota_configuration = StoredPaddleOcrConfiguration {
            model: "PaddleOCR-VL-1.6".to_string(),
        };
        let quota_error = recognize_paddleocr_aistudio_at(
            &client,
            &quota_configuration,
            "paddle-token",
            b"fake-png",
            "image/png",
            "png",
            &format!("{base_url}/quota"),
        )
        .await
        .unwrap_err();
        assert!(quota_error.contains("quota"));

        let mathpix_configuration = StoredMathpixConfiguration {
            base_url: base_url.clone(),
            app_id: "app-id".to_string(),
        };
        assert_eq!(
            recognize_mathpix(&client, &mathpix_configuration, "app-key", data_url,)
                .await
                .unwrap(),
            vec!["\\frac{1}{2}"]
        );
    }
}
