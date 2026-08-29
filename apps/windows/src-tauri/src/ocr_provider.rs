use super::{OcrFormulaResult, OcrImageRequest, OcrRecognitionResult, MAX_IMAGE_BYTES};
use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::fs;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::{Arc, Mutex};
use std::time::{Duration, Instant};
use tauri::{AppHandle, Manager};

pub(crate) const LOCAL_PROVIDER: &str = "local";
pub(crate) const OPENAI_COMPATIBLE_PROVIDER: &str = "openai-compatible";
pub(crate) const OLLAMA_PROVIDER: &str = "ollama";
pub(crate) const MATHPIX_PROVIDER: &str = "mathpix";

const CONFIGURATION_SCHEMA_VERSION: u32 = 1;
const CONFIGURATION_FILE: &str = "providers.json";
const MAX_RESPONSE_BYTES: usize = 4 * 1024 * 1024;
const MATHPIX_MAX_BASE64_IMAGE_BYTES: usize = 2 * 1024 * 1024;
const MAX_FORMULAS: usize = 64;
const MAX_FORMULA_CHARS: usize = 200_000;
const DEFAULT_PROMPT: &str = "Read every mathematical formula in this image in visual order. Return JSON only in the exact form {\"formulas\":[{\"latex\":\"...\"}]}. Return each independent visual formula row as a separate formulas-array item, while keeping a matrix or cases construction as one item. Use valid LaTeX without markdown fences or surrounding dollar delimiters. Preserve symbols, superscripts, subscripts, matrices, cases, alignment, and line structure. Do not explain the result.";

#[derive(Clone, Default)]
pub(crate) struct OcrProviderState {
    configuration: Arc<Mutex<Option<StoredOcrProviderConfiguration>>>,
    remote_recognition_running: Arc<AtomicBool>,
}

struct RemoteRecognitionLease {
    running: Arc<AtomicBool>,
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

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrProviderConfigurationUpdate {
    active_provider: String,
    open_ai_compatible: OpenAiCompatibleConfigurationUpdate,
    ollama: OllamaConfigurationUpdate,
    mathpix: MathpixConfigurationUpdate,
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

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredOcrProviderConfiguration {
    schema_version: u32,
    active_provider: String,
    open_ai_compatible: StoredOpenAiCompatibleConfiguration,
    ollama: StoredOllamaConfiguration,
    mathpix: StoredMathpixConfiguration,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredOpenAiCompatibleConfiguration {
    protocol: String,
    base_url: String,
    model: String,
    prompt: String,
    encrypted_api_key: Option<String>,
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
    encrypted_app_key: Option<String>,
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
                encrypted_api_key: None,
            },
            ollama: StoredOllamaConfiguration {
                base_url: "http://127.0.0.1:11434".to_string(),
                model: String::new(),
                prompt: DEFAULT_PROMPT.to_string(),
            },
            mathpix: StoredMathpixConfiguration {
                base_url: "https://api.mathpix.com".to_string(),
                app_id: String::new(),
                encrypted_app_key: None,
            },
        }
    }
}

impl StoredOcrProviderConfiguration {
    fn view(&self) -> OcrProviderConfigurationView {
        OcrProviderConfigurationView {
            active_provider: self.active_provider.clone(),
            open_ai_compatible: OpenAiCompatibleConfigurationView {
                protocol: self.open_ai_compatible.protocol.clone(),
                base_url: self.open_ai_compatible.base_url.clone(),
                model: self.open_ai_compatible.model.clone(),
                prompt: self.open_ai_compatible.prompt.clone(),
                has_api_key: self.open_ai_compatible.encrypted_api_key.is_some(),
            },
            ollama: OllamaConfigurationView {
                base_url: self.ollama.base_url.clone(),
                model: self.ollama.model.clone(),
                prompt: self.ollama.prompt.clone(),
            },
            mathpix: MathpixConfigurationView {
                base_url: self.mathpix.base_url.clone(),
                app_id: self.mathpix.app_id.clone(),
                has_app_key: self.mathpix.encrypted_app_key.is_some(),
            },
        }
    }
}

impl OcrProviderState {
    pub(crate) fn configuration(
        &self,
        app: &AppHandle,
    ) -> Result<OcrProviderConfigurationView, String> {
        Ok(self.load(app)?.view())
    }

    pub(crate) fn active_provider(&self, app: &AppHandle) -> Result<String, String> {
        Ok(self.load(app)?.active_provider)
    }

    pub(crate) fn save_configuration(
        &self,
        app: &AppHandle,
        update: OcrProviderConfigurationUpdate,
    ) -> Result<OcrProviderConfigurationView, String> {
        // A malformed on-disk JSON file must not permanently lock the user out of
        // the provider screen. The update contains the complete non-secret
        // configuration, so allow it to repair a damaged file. Existing secrets
        // can only be retained when the previous configuration was read and
        // validated successfully.
        let previous = self.load_for_update(app)?;
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

        let encrypted_api_key = update_secret(
            previous.open_ai_compatible.encrypted_api_key,
            update.open_ai_compatible.api_key,
            update.open_ai_compatible.clear_api_key,
        )?;
        let encrypted_app_key = update_secret(
            previous.mathpix.encrypted_app_key,
            update.mathpix.app_key,
            update.mathpix.clear_app_key,
        )?;

        validate_secret_transport(&openai_base_url, encrypted_api_key.is_some(), "API key")?;
        validate_secret_transport(&mathpix_base_url, encrypted_app_key.is_some(), "app_key")?;

        let next = StoredOcrProviderConfiguration {
            schema_version: CONFIGURATION_SCHEMA_VERSION,
            active_provider,
            open_ai_compatible: StoredOpenAiCompatibleConfiguration {
                protocol,
                base_url: openai_base_url,
                model: openai_model,
                prompt: openai_prompt,
                encrypted_api_key,
            },
            ollama: StoredOllamaConfiguration {
                base_url: ollama_base_url,
                model: ollama_model,
                prompt: ollama_prompt,
            },
            mathpix: StoredMathpixConfiguration {
                base_url: mathpix_base_url,
                app_id: mathpix_app_id,
                encrypted_app_key,
            },
        };
        validate_active_provider_configuration(&next)?;
        persist_configuration(app, &next)?;
        let mut cache = self
            .configuration
            .lock()
            .map_err(|_| "OCR provider configuration state is unavailable".to_string())?;
        *cache = Some(next.clone());
        Ok(next.view())
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
        validate_active_provider_configuration(&configuration)?;
        recognize_remote_inner(&configuration, request).await
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
    let configuration = match fs::read(&path) {
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
    };
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
    })
}

fn normalize_provider(value: &str) -> Result<&'static str, String> {
    match value.trim().to_ascii_lowercase().as_str() {
        LOCAL_PROVIDER => Ok(LOCAL_PROVIDER),
        OPENAI_COMPATIBLE_PROVIDER | "openai" => Ok(OPENAI_COMPATIBLE_PROVIDER),
        OLLAMA_PROVIDER => Ok(OLLAMA_PROVIDER),
        MATHPIX_PROVIDER => Ok(MATHPIX_PROVIDER),
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
    validate_secret_transport(
        &configuration.open_ai_compatible.base_url,
        configuration.open_ai_compatible.encrypted_api_key.is_some(),
        "API key",
    )?;
    validate_secret_transport(
        &configuration.mathpix.base_url,
        configuration.mathpix.encrypted_app_key.is_some(),
        "app_key",
    )?;
    Ok(())
}

fn validate_active_provider_configuration(
    configuration: &StoredOcrProviderConfiguration,
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
            // Compatible self-hosted services may intentionally use no API key,
            // including endpoints on another machine in a trusted LAN. The HTTPS
            // transport rule below applies only when a secret is actually present.
            Ok(())
        }
        OLLAMA_PROVIDER => {
            normalize_short_text(&configuration.ollama.model, "Ollama model", 160, true)?;
            Ok(())
        }
        MATHPIX_PROVIDER => {
            normalize_short_text(&configuration.mathpix.app_id, "Mathpix app_id", 240, true)?;
            if configuration.mathpix.encrypted_app_key.is_none() {
                return Err("Mathpix app_key is required".to_string());
            }
            Ok(())
        }
        _ => Err("Unsupported OCR provider".to_string()),
    }
}

#[cfg(windows)]
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

#[cfg(not(windows))]
fn validate_url(value: &str) -> Result<(), String> {
    if value.starts_with("http://") || value.starts_with("https://") {
        Ok(())
    } else {
        Err("OCR provider URL must use http or https".to_string())
    }
}

#[cfg(windows)]
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

#[cfg(not(windows))]
fn is_loopback_url(value: &str) -> Result<bool, String> {
    Ok(value.contains("localhost") || value.contains("127.0.0.1"))
}

fn validate_secret_transport(value: &str, has_secret: bool, label: &str) -> Result<(), String> {
    if has_secret && value.starts_with("http://") && !is_loopback_url(value)? {
        return Err(format!(
            "Refusing to send {label} over plain HTTP to a non-local OCR endpoint"
        ));
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
        // Rust does not guarantee that String storage is cleared on drop. Wipe the
        // owned plaintext after the HTTP request has finished. The HTTP library may
        // make its own short-lived header copy, but VisualTeX retains no plaintext
        // secret in provider state or on disk.
        unsafe {
            self.0.as_mut_vec().fill(0);
        }
    }
}

fn update_secret(
    previous: Option<String>,
    replacement: Option<String>,
    clear: bool,
) -> Result<Option<String>, String> {
    if clear {
        return Ok(None);
    }
    let replacement = replacement.unwrap_or_default();
    let replacement = replacement.trim();
    if replacement.is_empty() {
        return Ok(previous);
    }
    if replacement.chars().count() > 8_192 {
        return Err("OCR provider secret is too long".to_string());
    }
    protect_secret(replacement).map(Some)
}

#[cfg(windows)]
#[repr(C)]
struct DataBlob {
    cb_data: u32,
    pb_data: *mut u8,
}

#[cfg(windows)]
#[link(name = "crypt32")]
unsafe extern "system" {
    fn CryptProtectData(
        data_in: *const DataBlob,
        data_description: *const u16,
        optional_entropy: *const DataBlob,
        reserved: *mut std::ffi::c_void,
        prompt_struct: *mut std::ffi::c_void,
        flags: u32,
        data_out: *mut DataBlob,
    ) -> i32;
    fn CryptUnprotectData(
        data_in: *const DataBlob,
        data_description: *mut *mut u16,
        optional_entropy: *const DataBlob,
        reserved: *mut std::ffi::c_void,
        prompt_struct: *mut std::ffi::c_void,
        flags: u32,
        data_out: *mut DataBlob,
    ) -> i32;
}

#[cfg(windows)]
#[link(name = "kernel32")]
unsafe extern "system" {
    fn LocalFree(memory: *mut std::ffi::c_void) -> *mut std::ffi::c_void;
}

#[cfg(windows)]
fn protect_secret(secret: &str) -> Result<String, String> {
    const CRYPTPROTECT_UI_FORBIDDEN: u32 = 0x1;
    let mut input = secret.as_bytes().to_vec();
    let mut entropy = b"VisualTeX OCR provider secrets v1".to_vec();
    let input_blob = DataBlob {
        cb_data: input.len() as u32,
        pb_data: input.as_mut_ptr(),
    };
    let entropy_blob = DataBlob {
        cb_data: entropy.len() as u32,
        pb_data: entropy.as_mut_ptr(),
    };
    let mut output = DataBlob {
        cb_data: 0,
        pb_data: std::ptr::null_mut(),
    };
    let succeeded = unsafe {
        CryptProtectData(
            &input_blob,
            std::ptr::null(),
            &entropy_blob,
            std::ptr::null_mut(),
            std::ptr::null_mut(),
            CRYPTPROTECT_UI_FORBIDDEN,
            &mut output,
        )
    };
    input.fill(0);
    entropy.fill(0);
    if succeeded == 0 || output.pb_data.is_null() {
        return Err(format!(
            "Windows could not protect the OCR provider secret: {}",
            std::io::Error::last_os_error()
        ));
    }
    let protected =
        unsafe { std::slice::from_raw_parts(output.pb_data, output.cb_data as usize).to_vec() };
    unsafe {
        let _ = LocalFree(output.pb_data.cast());
    }
    Ok(BASE64_STANDARD.encode(protected))
}

#[cfg(windows)]
fn unprotect_secret(protected: &str) -> Result<SensitiveString, String> {
    const CRYPTPROTECT_UI_FORBIDDEN: u32 = 0x1;
    let mut input = BASE64_STANDARD
        .decode(protected.as_bytes())
        .map_err(|error| format!("OCR provider secret storage is damaged: {error}"))?;
    let mut entropy = b"VisualTeX OCR provider secrets v1".to_vec();
    let input_blob = DataBlob {
        cb_data: input.len() as u32,
        pb_data: input.as_mut_ptr(),
    };
    let entropy_blob = DataBlob {
        cb_data: entropy.len() as u32,
        pb_data: entropy.as_mut_ptr(),
    };
    let mut output = DataBlob {
        cb_data: 0,
        pb_data: std::ptr::null_mut(),
    };
    let succeeded = unsafe {
        CryptUnprotectData(
            &input_blob,
            std::ptr::null_mut(),
            &entropy_blob,
            std::ptr::null_mut(),
            std::ptr::null_mut(),
            CRYPTPROTECT_UI_FORBIDDEN,
            &mut output,
        )
    };
    input.fill(0);
    entropy.fill(0);
    if succeeded == 0 || output.pb_data.is_null() {
        return Err(format!(
            "Windows could not unlock the OCR provider secret for the current user: {}",
            std::io::Error::last_os_error()
        ));
    }
    let mut plaintext = unsafe {
        let native_plaintext =
            std::slice::from_raw_parts_mut(output.pb_data, output.cb_data as usize);
        let copy = native_plaintext.to_vec();
        native_plaintext.fill(0);
        let _ = LocalFree(output.pb_data.cast());
        copy
    };
    match String::from_utf8(plaintext) {
        Ok(secret) => Ok(SensitiveString(secret)),
        Err(error) => {
            plaintext = error.into_bytes();
            plaintext.fill(0);
            Err("OCR provider secret is not valid UTF-8".to_string())
        }
    }
}

#[cfg(not(windows))]
fn protect_secret(_secret: &str) -> Result<String, String> {
    Err("Remote OCR secrets are supported by the Windows application only".to_string())
}

#[cfg(not(windows))]
fn unprotect_secret(_protected: &str) -> Result<SensitiveString, String> {
    Err("Remote OCR secrets are supported by the Windows application only".to_string())
}

#[cfg(windows)]
async fn recognize_remote_inner(
    configuration: &StoredOcrProviderConfiguration,
    request: &OcrImageRequest,
) -> Result<OcrRecognitionResult, String> {
    let started = Instant::now();
    let mime = mime_for_extension(&request.extension)?;
    let data_base64 = BASE64_STANDARD.encode(&request.bytes);
    let data_url = format!("data:{mime};base64,{data_base64}");
    let (processed_width, processed_height) = image_dimensions(&request.bytes);
    let client = reqwest::Client::builder()
        .connect_timeout(Duration::from_secs(15))
        .timeout(Duration::from_secs(120))
        // Never forward an Authorization/app_key request through an HTTP redirect
        // to a different endpoint. Users must configure the final API base URL.
        .redirect(reqwest::redirect::Policy::none())
        .user_agent(format!("VisualTeX/{} OCR", env!("CARGO_PKG_VERSION")))
        .build()
        .map_err(|error| format!("Unable to initialize the OCR API client: {error}"))?;

    let (provider, model, formulas) = match configuration.active_provider.as_str() {
        OPENAI_COMPATIBLE_PROVIDER => {
            let api_key = configuration
                .open_ai_compatible
                .encrypted_api_key
                .as_deref()
                .map(unprotect_secret)
                .transpose()?;
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
            let formulas = recognize_ollama(&client, &configuration.ollama, &data_base64).await?;
            (
                OLLAMA_PROVIDER,
                configuration.ollama.model.clone(),
                formulas,
            )
        }
        MATHPIX_PROVIDER => {
            if data_url.len() > MATHPIX_MAX_BASE64_IMAGE_BYTES {
                return Err(
                    "Mathpix accepts at most 2 MB for a base64-encoded image. Crop the formula more tightly or use a smaller PNG/JPEG image."
                        .to_string(),
                );
            }
            let app_key = unprotect_secret(
                configuration
                    .mathpix
                    .encrypted_app_key
                    .as_deref()
                    .ok_or_else(|| "Mathpix app_key is required".to_string())?,
            )?;
            let formulas =
                recognize_mathpix(&client, &configuration.mathpix, app_key.as_str(), &data_url)
                    .await?;
            (MATHPIX_PROVIDER, "Mathpix Text API".to_string(), formulas)
        }
        _ => return Err("The active OCR provider is local, not remote".to_string()),
    };
    if formulas.is_empty() {
        return Err("OCR API returned no usable formula".to_string());
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

#[cfg(not(windows))]
async fn recognize_remote_inner(
    _configuration: &StoredOcrProviderConfiguration,
    _request: &OcrImageRequest,
) -> Result<OcrRecognitionResult, String> {
    Err("Remote OCR providers are supported by the Windows application only".to_string())
}

#[cfg(windows)]
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

#[cfg(windows)]
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

#[cfg(windows)]
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

#[cfg(windows)]
struct HttpJsonResponse {
    status: u16,
    text: String,
    json: Option<Value>,
}

#[cfg(windows)]
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

#[cfg(windows)]
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

    #[test]
    fn default_configuration_keeps_local_ocr_active() {
        let configuration = StoredOcrProviderConfiguration::default();
        assert_eq!(configuration.active_provider, LOCAL_PROVIDER);
        assert_eq!(configuration.open_ai_compatible.protocol, "responses");
        assert!(!configuration.open_ai_compatible.prompt.is_empty());
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
}
