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

#[path = "../../../../shared/visualtex-backend/src/ocr_provider_common.rs"]
mod common;
pub(crate) use common::*;

const PADDLEOCR_ROUTE_PROBE_TIMEOUT: Duration = Duration::from_secs(3);

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
struct StoredMathpixConfiguration {
    base_url: String,
    app_id: String,
    encrypted_app_key: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredPaddleOcrConfiguration {
    model: String,
    #[serde(default)]
    encrypted_access_token: Option<String>,
}

impl Default for StoredPaddleOcrConfiguration {
    fn default() -> Self {
        Self {
            model: "PaddleOCR-VL-1.6".to_string(),
            encrypted_access_token: None,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredSimpleTexConfiguration {
    model: String,
    #[serde(default)]
    encrypted_access_token: Option<String>,
}

impl Default for StoredSimpleTexConfiguration {
    fn default() -> Self {
        Self {
            model: "standard".to_string(),
            encrypted_access_token: None,
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
            paddle_ocr: StoredPaddleOcrConfiguration::default(),
            simple_tex: StoredSimpleTexConfiguration::default(),
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
            paddle_ocr: PaddleOcrConfigurationView {
                model: self.paddle_ocr.model.clone(),
                has_access_token: self.paddle_ocr.encrypted_access_token.is_some(),
            },
            simple_tex: SimpleTexConfigurationView {
                model: self.simple_tex.model.clone(),
                has_access_token: self.simple_tex.encrypted_access_token.is_some(),
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
        let paddle_ocr_model = normalize_paddleocr_model(&update.paddle_ocr.model)?.to_string();
        let simpletex_model = normalize_simpletex_model(&update.simple_tex.model)?.to_string();

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
        let encrypted_access_token = update_secret(
            previous.paddle_ocr.encrypted_access_token,
            update.paddle_ocr.access_token,
            update.paddle_ocr.clear_access_token,
        )?;
        let encrypted_simpletex_access_token = update_secret(
            previous.simple_tex.encrypted_access_token,
            update.simple_tex.access_token,
            update.simple_tex.clear_access_token,
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
            paddle_ocr: StoredPaddleOcrConfiguration {
                model: paddle_ocr_model,
                encrypted_access_token,
            },
            simple_tex: StoredSimpleTexConfiguration {
                model: simpletex_model,
                encrypted_access_token: encrypted_simpletex_access_token,
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
        #[cfg(windows)]
        {
            let progress = RemoteRecognitionProgress::new(app, &request.model);
            // Paddle selects its direct/system-proxy route before submitting and
            // therefore emits a more specific submit message from its adapter.
            if configuration.active_provider != PADDLEOCR_PROVIDER {
                emit_remote_provider_progress(
                    Some(&progress),
                    &configuration.active_provider,
                    "api-submit",
                );
            }
            recognize_remote_inner(&configuration, request, Some(&progress)).await
        }
        #[cfg(not(windows))]
        {
            recognize_remote_inner(&configuration, request, None).await
        }
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

fn simpletex_endpoint(model: &str) -> Result<&'static str, String> {
    match normalize_simpletex_model(model)? {
        "standard" => Ok(SIMPLETEX_STANDARD_URL),
        "turbo" => Ok(SIMPLETEX_TURBO_URL),
        _ => unreachable!("normalize_simpletex_model returns a fixed model"),
    }
}

fn migrate_legacy_paddleocr_configuration(
    mut configuration: StoredOcrProviderConfiguration,
) -> StoredOcrProviderConfiguration {
    if configuration.schema_version == 1 {
        configuration.paddle_ocr.model = match configuration.paddle_ocr.model.as_str() {
            // Every experimental Paddle model now converges on the validated
            // VL-1.6 path. PP-StructureV3 remains disabled until it reliably
            // returns usable formulas for real VisualTeX samples.
            "PaddleOCR-VL-1.5" | "PaddleOCR-VL" | "PP-StructureV3" => {
                "PaddleOCR-VL-1.6".to_string()
            }
            // Preserve the reliable VL-1.6 default for existing installations.
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
        PADDLEOCR_PROVIDER => {
            normalize_paddleocr_model(&configuration.paddle_ocr.model)?;
            if configuration.paddle_ocr.encrypted_access_token.is_none() {
                return Err("PaddleOCR AI Studio access token is required".to_string());
            }
            Ok(())
        }
        SIMPLETEX_PROVIDER => {
            normalize_simpletex_model(&configuration.simple_tex.model)?;
            if configuration.simple_tex.encrypted_access_token.is_none() {
                return Err("SimpleTex user access token is required".to_string());
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
fn build_ocr_http_client(
    timeout: Duration,
    redirect: reqwest::redirect::Policy,
    user_agent: String,
) -> Result<reqwest::Client, String> {
    // The Windows reqwest dependency enables `system-proxy`, so this builder
    // honors HTTPS_PROXY/HTTP_PROXY/NO_PROXY first and then the current Windows
    // Internet Settings proxy configuration (including its bypass list).
    reqwest::Client::builder()
        .connect_timeout(Duration::from_secs(15))
        .timeout(timeout)
        .redirect(redirect)
        .user_agent(user_agent)
        .build()
        .map_err(|error| format!("Unable to initialize the OCR API client: {error}"))
}

#[cfg(windows)]
fn build_direct_ocr_http_client(
    timeout: Duration,
    redirect: reqwest::redirect::Policy,
    user_agent: String,
) -> Result<reqwest::Client, String> {
    reqwest::Client::builder()
        .no_proxy()
        .connect_timeout(Duration::from_secs(10))
        .timeout(timeout)
        .redirect(redirect)
        .user_agent(user_agent)
        .build()
        .map_err(|error| format!("Unable to initialize the direct OCR API client: {error}"))
}

#[cfg(windows)]
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
        // Never forward an Authorization/app_key/access-token request through an
        // HTTP redirect to a different endpoint. Users must configure the final
        // API base URL; Paddle result downloads use a separate limited client.
        reqwest::redirect::Policy::none(),
        format!("VisualTeX/{} OCR", env!("CARGO_PKG_VERSION")),
    )?;

    let (provider, model, formulas) = match configuration.active_provider.as_str() {
        OPENAI_COMPATIBLE_PROVIDER => {
            let api_key = configuration
                .open_ai_compatible
                .encrypted_api_key
                .as_deref()
                .map(unprotect_secret)
                .transpose()?;
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
            emit_remote_provider_progress(progress, MATHPIX_PROVIDER, "api-inference");
            let formulas =
                recognize_mathpix(&client, &configuration.mathpix, app_key.as_str(), &data_url)
                    .await?;
            (MATHPIX_PROVIDER, "Mathpix Text API".to_string(), formulas)
        }
        PADDLEOCR_PROVIDER => {
            let access_token = unprotect_secret(
                configuration
                    .paddle_ocr
                    .encrypted_access_token
                    .as_deref()
                    .ok_or_else(|| "PaddleOCR AI Studio access token is required".to_string())?,
            )?;
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
            let access_token = unprotect_secret(
                configuration
                    .simple_tex
                    .encrypted_access_token
                    .as_deref()
                    .ok_or_else(|| "SimpleTex user access token is required".to_string())?,
            )?;
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

#[cfg(not(windows))]
async fn recognize_remote_inner(
    _configuration: &StoredOcrProviderConfiguration,
    _request: &OcrImageRequest,
    _progress: Option<&()>,
) -> Result<OcrRecognitionResult, String> {
    Err("Remote OCR providers are supported by the Windows application only".to_string())
}

#[cfg(windows)]
async fn recognize_paddleocr_aistudio(
    system_proxy_client: &reqwest::Client,
    configuration: &StoredPaddleOcrConfiguration,
    access_token: &str,
    image_bytes: &[u8],
    mime: &str,
    extension: &str,
    progress: Option<&RemoteRecognitionProgress>,
) -> Result<Vec<String>, String> {
    // Paddle's AI Studio endpoint is hosted in mainland China. On Windows a
    // user-level proxy may be required for OpenAI/Mathpix yet add several seconds
    // of TLS latency (or intermittent failures) to Paddle. Probe the official
    // endpoint directly before submitting the non-idempotent job; any HTTP
    // response (normally 405 for GET) proves direct routing works. Only fall back
    // to the system-proxy client when the direct probe itself cannot connect.
    let direct_client = build_direct_ocr_http_client(
        Duration::from_secs(30),
        reqwest::redirect::Policy::none(),
        format!("VisualTeX/{} PaddleOCR", env!("CARGO_PKG_VERSION")),
    )?;
    let direct_available = direct_client
        .get(PADDLEOCR_JOBS_URL)
        .timeout(PADDLEOCR_ROUTE_PROBE_TIMEOUT)
        .send()
        .await
        .is_ok();
    let client = if direct_available {
        &direct_client
    } else {
        system_proxy_client
    };
    if let Some(progress) = progress {
        progress.emit(
            "api-submit",
            if direct_available {
                "正在直连 PaddleOCR 并提交图片…"
            } else {
                "PaddleOCR 直连不可用，正在通过系统代理提交图片…"
            },
        );
    }
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

#[cfg(windows)]
async fn download_and_parse_paddleocr_result(result_url: &str) -> Result<Vec<String>, String> {
    let client = build_ocr_http_client(
        PADDLEOCR_RESULT_TIMEOUT,
        reqwest::redirect::Policy::limited(5),
        format!("VisualTeX/{} PaddleOCR", env!("CARGO_PKG_VERSION")),
    )?;
    download_and_parse_paddleocr_result_with_client(&client, result_url).await
}

#[cfg(windows)]
fn validate_paddleocr_result_url(value: &str) -> Result<(), String> {
    let url = reqwest::Url::parse(value)
        .map_err(|error| format!("Invalid PaddleOCR result URL: {error}"))?;
    if !url.username().is_empty() || url.password().is_some() {
        return Err("PaddleOCR result URL must not contain embedded credentials".to_string());
    }
    if url.scheme() == "https" {
        return Ok(());
    }
    #[cfg(any(test, feature = "ocr-provider-connectivity-acceptance"))]
    if url.scheme() == "http" && is_loopback_url(value)? {
        return Ok(());
    }
    Err("PaddleOCR result URL must use HTTPS".to_string())
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
    let response = request
        .send()
        .await
        .map_err(|error| format!("OCR API request failed: {error}"))?;
    read_http_response(response, MAX_RESPONSE_BYTES, "OCR API response").await
}

fn extract_api_error(value: &Value) -> Option<String> {
    value
        .pointer("/error/message")
        .and_then(Value::as_str)
        .or_else(|| value.get("error").and_then(Value::as_str))
        .or_else(|| value.get("message").and_then(Value::as_str))
        .map(|message| truncate_for_error(message, 800))
}

#[cfg(all(windows, feature = "ocr-provider-connectivity-acceptance"))]
pub(crate) async fn run_provider_connectivity_acceptance(public_probe: bool) -> Result<(), String> {
    use axum::{
        body::Bytes,
        extract::{Json, Path},
        http::{HeaderMap, StatusCode},
        routing::{get, post},
        Router,
    };
    use tokio::net::TcpListener;

    let legacy = serde_json::from_value::<StoredOcrProviderConfiguration>(json!({
        "schemaVersion": 1,
        "activeProvider": "local",
        "openAiCompatible": {
            "protocol": "responses",
            "baseUrl": "https://api.openai.com/v1",
            "model": "",
            "prompt": "prompt",
            "encryptedApiKey": null
        },
        "ollama": {
            "baseUrl": "http://127.0.0.1:11434",
            "model": "",
            "prompt": "prompt"
        },
        "mathpix": {
            "baseUrl": "https://api.mathpix.com",
            "appId": "",
            "encryptedAppKey": null
        }
    }))
    .map_err(|error| format!("Legacy OCR provider configuration did not deserialize: {error}"))?;
    if legacy.paddle_ocr.model != "PaddleOCR-VL-1.6"
        || legacy.paddle_ocr.encrypted_access_token.is_some()
    {
        return Err("Legacy OCR provider configuration did not receive safe Paddle defaults".to_string());
    }

    let migrated = migrate_legacy_paddleocr_configuration(
        serde_json::from_value::<StoredOcrProviderConfiguration>(json!({
            "schemaVersion": 1,
            "activeProvider": "paddleocr",
            "openAiCompatible": {
                "protocol": "responses",
                "baseUrl": "https://api.openai.com/v1",
                "model": "",
                "prompt": "prompt",
                "encryptedApiKey": null
            },
            "ollama": {
                "baseUrl": "http://127.0.0.1:11434",
                "model": "",
                "prompt": "prompt"
            },
            "mathpix": {
                "baseUrl": "https://api.mathpix.com",
                "appId": "",
                "encryptedAppKey": null
            },
            "paddleOcr": {
                "model": "PaddleOCR-VL-1.5",
                "encryptedAccessToken": null
            }
        }))
        .map_err(|error| format!("Legacy PaddleOCR configuration did not deserialize: {error}"))?,
    );
    if migrated.schema_version != CONFIGURATION_SCHEMA_VERSION
        || migrated.paddle_ocr.model != "PaddleOCR-VL-1.6"
    {
        return Err("Legacy PaddleOCR-VL model/schema migration failed".to_string());
    }

    let previous_default = migrate_legacy_paddleocr_configuration(
        serde_json::from_value::<StoredOcrProviderConfiguration>(json!({
            "schemaVersion": 1,
            "activeProvider": "paddleocr",
            "openAiCompatible": {
                "protocol": "responses",
                "baseUrl": "https://api.openai.com/v1",
                "model": "",
                "prompt": "prompt",
                "encryptedApiKey": null
            },
            "ollama": {
                "baseUrl": "http://127.0.0.1:11434",
                "model": "",
                "prompt": "prompt"
            },
            "mathpix": {
                "baseUrl": "https://api.mathpix.com",
                "appId": "",
                "encryptedAppKey": null
            },
            "paddleOcr": {
                "model": "PaddleOCR-VL-1.6",
                "encryptedAccessToken": null
            }
        }))
        .map_err(|error| format!("Previous Paddle default did not deserialize: {error}"))?,
    );
    if previous_default.schema_version != CONFIGURATION_SCHEMA_VERSION
        || previous_default.paddle_ocr.model != "PaddleOCR-VL-1.6"
    {
        return Err("Previous Paddle VL default was not preserved".to_string());
    }

    let protected_token = protect_secret("paddle-token")?;
    if protected_token.contains("paddle-token") {
        return Err("Windows DPAPI unexpectedly retained the Paddle token as plaintext".to_string());
    }
    let unprotected_token = unprotect_secret(&protected_token)?;
    if unprotected_token.as_str() != "paddle-token" {
        return Err("Windows DPAPI Paddle token round-trip changed the secret".to_string());
    }
    let mut paddle_active = StoredOcrProviderConfiguration::default();
    paddle_active.active_provider = PADDLEOCR_PROVIDER.to_string();
    paddle_active.paddle_ocr.encrypted_access_token = Some(protected_token);
    validate_active_provider_configuration(&paddle_active)?;

    async fn openai(Json(body): Json<Value>) -> Json<Value> {
        assert_eq!(body.get("model").and_then(Value::as_str), Some("vision-model"));
        Json(json!({
            "output_text": "{\"formulas\":[{\"latex\":\"x^2+y^2\"},{\"latex\":\"z=3\"}]}"
        }))
    }

    async fn ollama(Json(body): Json<Value>) -> Json<Value> {
        assert_eq!(body.get("model").and_then(Value::as_str), Some("vision-model"));
        Json(json!({
            "message": {"content": "{\"formulas\":[{\"latex\":\"a=b\"}]}"}
        }))
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

    fn require_paddle_bearer(headers: &HeaderMap) {
        assert_eq!(
            headers
                .get("authorization")
                .and_then(|value| value.to_str().ok()),
            Some("Bearer paddle-token")
        );
    }

    async fn paddle_vl_submit(headers: HeaderMap, body: Bytes) -> Json<Value> {
        require_paddle_bearer(&headers);
        let body = String::from_utf8_lossy(&body);
        assert!(body.contains("PaddleOCR-VL-1.6"));
        assert!(body.contains("showFormulaNumber"));
        assert!(body.contains("prettifyMarkdown"));
        Json(json!({"code": 0, "msg": "Success", "data": {"jobId": "vl-job"}}))
    }

    async fn paddle_structure_submit(headers: HeaderMap, body: Bytes) -> Json<Value> {
        require_paddle_bearer(&headers);
        let body = String::from_utf8_lossy(&body);
        assert!(body.contains("PP-StructureV3"));
        assert!(body.contains("useFormulaRecognition"));
        assert!(body.contains("useTableRecognition"));
        Json(json!({"code": 0, "msg": "Success", "data": {"jobId": "structure-job"}}))
    }

    async fn paddle_status(Path(job_id): Path<String>, headers: HeaderMap) -> Json<Value> {
        require_paddle_bearer(&headers);
        let host = headers
            .get("host")
            .and_then(|value| value.to_str().ok())
            .unwrap_or_default();
        Json(json!({
            "code": 0,
            "msg": "Success",
            "data": {
                "jobId": job_id,
                "state": "done",
                "resultUrl": {"jsonUrl": format!("http://{host}/paddle-result/{job_id}")}
            }
        }))
    }

    async fn paddle_pending_status(Path(job_id): Path<String>, headers: HeaderMap) -> Json<Value> {
        require_paddle_bearer(&headers);
        Json(json!({
            "code": 0,
            "msg": "Success",
            "data": {
                "jobId": job_id,
                "state": "pending"
            }
        }))
    }

    async fn paddle_result(Path(job_id): Path<String>) -> String {
        let value = if job_id == "structure-job" {
            json!({
                "result": {"layoutParsingResults": [{
                    "prunedResult": {"parsing_res_list": [
                        {"block_label": "formula", "block_content": "$$E=mc^2$$"}
                    ]}
                }]}
            })
        } else {
            json!({
                "result": {"layoutParsingResults": [{
                    "prunedResult": {"parsing_res_list": [
                        {"block_label": "formula", "block_content": "$$x&=1\\\\y&=2$$"},
                        {"block_label": "equation", "block_content": "\\frac{a}{b}"}
                    ]}
                }]}
            })
        };
        format!("{}\n", serde_json::to_string(&value).unwrap())
    }

    async fn paddle_quota() -> (StatusCode, Json<Value>) {
        (
            StatusCode::TOO_MANY_REQUESTS,
            Json(json!({"code": 429, "msg": "quota exceeded"})),
        )
    }

    async fn mathpix(headers: HeaderMap, Json(body): Json<Value>) -> Json<Value> {
        assert_eq!(headers.get("app_id").and_then(|value| value.to_str().ok()), Some("app-id"));
        assert_eq!(headers.get("app_key").and_then(|value| value.to_str().ok()), Some("app-key"));
        assert!(body.get("src").and_then(Value::as_str).is_some());
        Json(json!({"latex_styled": "\\frac{1}{2}"}))
    }

    let router = Router::new()
        .route("/responses", post(openai))
        .route("/api/chat", post(ollama))
        .route("/simpletex", post(simpletex))
        .route("/simpletex-slow", post(simpletex_slow))
        .route("/jobs-vl", post(paddle_vl_submit))
        .route("/jobs-vl/{job_id}", get(paddle_status))
        .route("/jobs-structure", post(paddle_structure_submit))
        .route("/jobs-structure/{job_id}", get(paddle_status))
        .route("/jobs-pending", post(paddle_structure_submit))
        .route("/jobs-pending/{job_id}", get(paddle_pending_status))
        .route("/paddle-result/{job_id}", get(paddle_result))
        .route("/quota", post(paddle_quota))
        .route("/v3/text", post(mathpix));
    let listener = TcpListener::bind("127.0.0.1:0")
        .await
        .map_err(|error| format!("Unable to start OCR mock API: {error}"))?;
    let address = listener
        .local_addr()
        .map_err(|error| format!("Unable to inspect OCR mock API address: {error}"))?;
    tokio::spawn(async move {
        axum::serve(listener, router).await.unwrap();
    });
    let base_url = format!("http://{address}");
    let direct_client = reqwest::Client::builder()
        .no_proxy()
        .redirect(reqwest::redirect::Policy::none())
        .build()
        .map_err(|error| format!("Unable to initialize direct OCR mock client: {error}"))?;
    let data_url = "data:image/png;base64,iVBORw0KGgo=";

    let openai_configuration = StoredOpenAiCompatibleConfiguration {
        protocol: "responses".to_string(),
        base_url: base_url.clone(),
        model: "vision-model".to_string(),
        prompt: "Return formula JSON".to_string(),
        encrypted_api_key: None,
    };
    let openai = recognize_openai_compatible(
        &direct_client,
        &openai_configuration,
        None,
        data_url,
    )
    .await?;
    if openai != ["x^2+y^2", "z=3"] {
        return Err(format!("OpenAI-compatible mock result mismatch: {openai:?}"));
    }

    let ollama_configuration = StoredOllamaConfiguration {
        base_url: base_url.clone(),
        model: "vision-model".to_string(),
        prompt: "Return formula JSON".to_string(),
    };
    let ollama = recognize_ollama(&direct_client, &ollama_configuration, "iVBORw0KGgo=").await?;
    if ollama != ["a=b"] {
        return Err(format!("Ollama mock result mismatch: {ollama:?}"));
    }

    let simpletex_configuration = StoredSimpleTexConfiguration {
        model: "standard".to_string(),
        encrypted_access_token: None,
    };
    let simpletex = recognize_simpletex_at(
        &direct_client,
        &simpletex_configuration,
        "simpletex-token",
        b"fake-png",
        "image/png",
        "png",
        &format!("{base_url}/simpletex"),
        Duration::from_secs(2),
    )
    .await?;
    if simpletex != ["\\sqrt{x^2+y^2}"] {
        return Err(format!("SimpleTex mock result mismatch: {simpletex:?}"));
    }
    let timeout_error = recognize_simpletex_at(
        &direct_client,
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
    if !timeout_error.to_ascii_lowercase().contains("timed out") {
        return Err(format!("Unexpected SimpleTex timeout error: {timeout_error}"));
    }

    let paddle_vl = StoredPaddleOcrConfiguration {
        model: "PaddleOCR-VL-1.6".to_string(),
        encrypted_access_token: None,
    };
    let vl_result = recognize_paddleocr_aistudio_at(
        &direct_client,
        &paddle_vl,
        "paddle-token",
        b"fake-png",
        "image/png",
        "png",
        &format!("{base_url}/jobs-vl"),
    )
    .await?;
    if vl_result != ["x&=1\\\\y&=2", "\\frac{a}{b}"] {
        return Err(format!("PaddleOCR-VL mock result mismatch: {vl_result:?}"));
    }

    if normalize_paddleocr_model("PP-StructureV3").is_ok() {
        return Err("PP-StructureV3 must remain disabled".to_string());
    }

    let quota_error = recognize_paddleocr_aistudio_at(
        &direct_client,
        &paddle_vl,
        "paddle-token",
        b"fake-png",
        "image/png",
        "png",
        &format!("{base_url}/quota"),
    )
    .await
    .expect_err("Paddle quota mock must fail");
    if !quota_error.to_ascii_lowercase().contains("quota") {
        return Err(format!("Unexpected Paddle quota error: {quota_error}"));
    }

    let mathpix_configuration = StoredMathpixConfiguration {
        base_url: base_url,
        app_id: "app-id".to_string(),
        encrypted_app_key: None,
    };
    let mathpix = recognize_mathpix(
        &direct_client,
        &mathpix_configuration,
        "app-key",
        data_url,
    )
    .await?;
    if mathpix != ["\\frac{1}{2}"] {
        return Err(format!("Mathpix mock result mismatch: {mathpix:?}"));
    }

    println!(
        "OCR provider mock protocols passed: OpenAI-compatible, Ollama, Mathpix, PaddleOCR-VL-1.6, SimpleTex standard."
    );

    if public_probe {
        let client = build_ocr_http_client(
            Duration::from_secs(20),
            reqwest::redirect::Policy::none(),
            format!("VisualTeX/{} OCR-connectivity", env!("CARGO_PKG_VERSION")),
        )?;
        for (name, url) in [
            ("OpenAI", "https://api.openai.com/v1/models"),
            ("Mathpix", "https://api.mathpix.com/v3/text"),
            ("PaddleOCR", PADDLEOCR_JOBS_URL),
            ("SimpleTex standard", SIMPLETEX_STANDARD_URL),
            ("SimpleTex turbo", SIMPLETEX_TURBO_URL),
        ] {
            let response = client.get(url).send().await.map_err(|error| {
                format!(
                    "{name} endpoint was not reachable through the VisualTeX OCR client: {error}"
                )
            })?;
            println!(
                "{name} public endpoint reachable: HTTP {}",
                response.status().as_u16()
            );
        }
        match client.get("http://127.0.0.1:11434/api/tags").send().await {
            Ok(response) => println!("Ollama local endpoint reachable: HTTP {}", response.status().as_u16()),
            Err(error) => println!("Ollama local endpoint not running/reachable on 127.0.0.1:11434: {error}"),
        }
    }

    Ok(())
}

#[cfg(all(windows, feature = "ocr-provider-connectivity-acceptance"))]
pub(crate) async fn run_saved_paddle_latency_diagnostic(
    image_path: &std::path::Path,
    repeats: usize,
    model_override: Option<&str>,
) -> Result<(), String> {
    let app_data = std::env::var_os("APPDATA")
        .ok_or_else(|| "APPDATA is unavailable for the saved Paddle diagnostic".to_string())?;
    let configuration_path = PathBuf::from(app_data)
        .join("com.visualtex.studio")
        .join("ocr")
        .join(CONFIGURATION_FILE);
    let configuration = migrate_legacy_paddleocr_configuration(
        serde_json::from_slice::<StoredOcrProviderConfiguration>(
            &fs::read(&configuration_path).map_err(|error| {
                format!(
                    "Unable to read saved OCR provider configuration ({}): {error}",
                    configuration_path.display()
                )
            })?,
        )
        .map_err(|error| format!("Unable to parse saved OCR provider configuration: {error}"))?,
    );
    if configuration.active_provider != PADDLEOCR_PROVIDER {
        return Err(format!(
            "Saved OCR provider is {}, not PaddleOCR",
            configuration.active_provider
        ));
    }
    let token = unprotect_secret(
        configuration
            .paddle_ocr
            .encrypted_access_token
            .as_deref()
            .ok_or_else(|| "Saved PaddleOCR access token is missing".to_string())?,
    )?;
    let image_bytes = fs::read(image_path)
        .map_err(|error| format!("Unable to read diagnostic image ({}): {error}", image_path.display()))?;
    let extension = image_path
        .extension()
        .and_then(|value| value.to_str())
        .unwrap_or("png");
    let mime = mime_for_extension(extension)?;
    let model = match model_override.map(str::trim).filter(|value| !value.is_empty()) {
        Some("PaddleOCR-VL-1.6") => "PaddleOCR-VL-1.6",
        Some("PaddleOCR-VL-1.5") => "PaddleOCR-VL-1.5",
        Some("PaddleOCR-VL") => "PaddleOCR-VL",
        Some("PP-StructureV3") => "PP-StructureV3",
        Some(other) => {
            return Err(format!(
                "Unsupported Paddle diagnostic model override: {other}"
            ))
        }
        None => normalize_paddleocr_model(&configuration.paddle_ocr.model)?,
    };
    let api_model = model;

    let diagnostic_direct = std::env::var("VISUALTEX_PADDLE_DIAGNOSTIC_DIRECT")
        .ok()
        .is_some_and(|value| value == "1");
    println!(
        "Saved Paddle diagnostic: model={model}, bytes={}, repeats={}, route={}",
        image_bytes.len(),
        repeats.max(1),
        if diagnostic_direct { "direct" } else { "system-proxy" }
    );
    for run in 1..=repeats.max(1) {
        let total_started = Instant::now();
        let client = if diagnostic_direct {
            reqwest::Client::builder()
                .no_proxy()
                .connect_timeout(Duration::from_secs(10))
                .timeout(Duration::from_secs(120))
                .redirect(reqwest::redirect::Policy::none())
                .user_agent(format!(
                    "VisualTeX/{} Paddle-latency-diagnostic",
                    env!("CARGO_PKG_VERSION")
                ))
                .build()
                .map_err(|error| format!("Unable to initialize direct Paddle diagnostic client: {error}"))?
        } else {
            build_ocr_http_client(
                Duration::from_secs(120),
                reqwest::redirect::Policy::none(),
                format!("VisualTeX/{} Paddle-latency-diagnostic", env!("CARGO_PKG_VERSION")),
            )?
        };
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
            .map_err(|error| format!("Unable to encode Paddle diagnostic options: {error}"))?;
        let part = reqwest::multipart::Part::bytes(image_bytes.clone())
            .file_name(format!("visualtex-formula.{extension}"))
            .mime_str(mime)
            .map_err(|error| format!("Unable to prepare Paddle diagnostic image: {error}"))?;
        let form = reqwest::multipart::Form::new()
            .text("model", api_model.to_string())
            .text("optionalPayload", options)
            .part("file", part);

        let submit_started = Instant::now();
        let submit_response = client
            .post(PADDLEOCR_JOBS_URL)
            .bearer_auth(token.as_str())
            .multipart(form)
            .send()
            .await
            .map_err(|error| format!("Paddle diagnostic submit failed: {error}"))?;
        let submit_network_ms = submit_started.elapsed().as_millis();
        let submit = read_http_response(
            submit_response,
            MAX_RESPONSE_BYTES,
            "Paddle diagnostic submission",
        )
        .await?;
        ensure_success(&submit, "PaddleOCR AI Studio")?;
        let submit_json = submit
            .json
            .as_ref()
            .ok_or_else(|| "Paddle diagnostic submit returned non-JSON output".to_string())?;
        ensure_paddleocr_api_code(submit_json, "Paddle diagnostic submission")?;
        let job_id = submit_json
            .pointer("/data/jobId")
            .and_then(Value::as_str)
            .ok_or_else(|| "Paddle diagnostic submission returned no jobId".to_string())?;
        println!("run {run}: submit={submit_network_ms}ms");

        let status_url = format!("{}/{job_id}", PADDLEOCR_JOBS_URL.trim_end_matches('/'));
        let mut poll_index = 0usize;
        let mut first_running_ms: Option<u128> = None;
        let result_url = loop {
            poll_index += 1;
            let poll_started = Instant::now();
            let response = client
                .get(&status_url)
                .bearer_auth(token.as_str())
                .send()
                .await
                .map_err(|error| format!("Paddle diagnostic poll #{poll_index} failed: {error}"))?;
            let status = read_http_response(
                response,
                MAX_RESPONSE_BYTES,
                "Paddle diagnostic status",
            )
            .await?;
            ensure_success(&status, "PaddleOCR AI Studio")?;
            let value = status
                .json
                .as_ref()
                .ok_or_else(|| "Paddle diagnostic status returned non-JSON output".to_string())?;
            ensure_paddleocr_api_code(value, "Paddle diagnostic status")?;
            let state = value
                .pointer("/data/state")
                .and_then(Value::as_str)
                .unwrap_or("missing");
            println!(
                "run {run}: poll#{poll_index} state={state} request={}ms total={}ms",
                poll_started.elapsed().as_millis(),
                total_started.elapsed().as_millis()
            );
            if state == "running" && first_running_ms.is_none() {
                first_running_ms = Some(total_started.elapsed().as_millis());
            }
            match state {
                "done" => {
                    let done_ms = total_started.elapsed().as_millis();
                    println!(
                        "run {run}: queue={}ms inference={}ms",
                        first_running_ms.unwrap_or(done_ms),
                        first_running_ms
                            .map(|running_ms| done_ms.saturating_sub(running_ms))
                            .unwrap_or(0)
                    );
                    break value
                        .pointer("/data/resultUrl/jsonUrl")
                        .and_then(Value::as_str)
                        .filter(|value| !value.trim().is_empty())
                        .ok_or_else(|| "Paddle diagnostic completed without result URL".to_string())?
                        .to_string();
                }
                "failed" => {
                    return Err(format!(
                        "Paddle diagnostic remote job failed: {}",
                        value
                            .pointer("/data/errorMsg")
                            .and_then(Value::as_str)
                            .unwrap_or("unknown error")
                    ));
                }
                "pending" | "running" => tokio::time::sleep(PADDLEOCR_POLL_INTERVAL).await,
                other => return Err(format!("Paddle diagnostic unknown job state: {other}")),
            }
        };

        let download_started = Instant::now();
        let formulas = if diagnostic_direct {
            validate_paddleocr_result_url(&result_url)?;
            let response = client
                .get(&result_url)
                .send()
                .await
                .map_err(|error| format!("Direct Paddle diagnostic result download failed: {error}"))?;
            let response = read_http_response(
                response,
                PADDLEOCR_RESULT_MAX_BYTES,
                "Direct Paddle diagnostic result",
            )
            .await?;
            ensure_success(&response, "PaddleOCR result download")?;
            match parse_paddleocr_result_text(&response.text) {
                Ok(formulas) => formulas,
                Err(error) => {
                    println!("run {run}: result parse warning: {error}");
                    Vec::new()
                }
            }
        } else {
            download_and_parse_paddleocr_result(&result_url).await?
        };
        println!(
            "run {run}: result-download+parse={}ms total={}ms formulas={}",
            download_started.elapsed().as_millis(),
            total_started.elapsed().as_millis(),
            formulas.len()
        );
    }
    Ok(())
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
            assert_eq!(body.get("model").and_then(Value::as_str), Some("vision-model"));
            assert!(body
                .pointer("/input/0/content/1/image_url")
                .and_then(Value::as_str)
                .is_some_and(|url| url.starts_with("data:image/png;base64,")));
            Json(json!({
                "output_text": "{\"formulas\":[{\"latex\":\"x^2+y^2\"},{\"latex\":\"z=3\"}]}"
            }))
        }

        async fn ollama(Json(body): Json<Value>) -> Json<Value> {
            assert_eq!(body.get("model").and_then(Value::as_str), Some("vision-model"));
            assert_eq!(body.get("stream").and_then(Value::as_bool), Some(false));
            assert!(body
                .pointer("/messages/0/images/0")
                .and_then(Value::as_str)
                .is_some_and(|image| !image.is_empty()));
            Json(json!({
                "message": {"content": "{\"formulas\":[{\"latex\":\"a=b\"}]}"}
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

        async fn paddle_structure_submit(headers: HeaderMap, body: Bytes) -> Json<Value> {
            assert_paddle_bearer(&headers);
            let body = String::from_utf8_lossy(&body);
            assert!(body.contains("PP-StructureV3"));
            assert!(body.contains("useFormulaRecognition"));
            assert!(body.contains("useTableRecognition"));
            assert!(body.contains("useSealRecognition"));
            assert!(body.contains("visualtex-formula.png"));
            Json(json!({"code": 0, "msg": "Success", "data": {"jobId": "structure-job"}}))
        }

        async fn paddle_status(
            Path(job_id): Path<String>,
            headers: HeaderMap,
        ) -> Json<Value> {
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
                    "resultUrl": {"jsonUrl": format!("http://{host}/paddle-result/{job_id}")}
                }
            }))
        }

        async fn paddle_result(Path(job_id): Path<String>) -> String {
            let result = if job_id == "structure-job" {
                json!({
                    "result": {"layoutParsingResults": [{
                        "prunedResult": {"parsing_res_list": [
                            {"block_label": "formula", "block_content": "$$E=mc^2$$"}
                        ]}
                    }]}
                })
            } else {
                json!({
                    "result": {"layoutParsingResults": [{
                        "prunedResult": {"parsing_res_list": [
                            {"block_label": "text", "block_content": "ignored"},
                            {"block_label": "formula", "block_content": "$$x&=1\\\\y&=2$$"},
                            {"block_label": "equation", "block_content": "\\frac{a}{b}"}
                        ]},
                        "markdown": {"text": "ignored markdown", "images": {}}
                    }]}
                })
            };
            format!("{}\n", serde_json::to_string(&result).unwrap())
        }

        async fn paddle_quota() -> (StatusCode, Json<Value>) {
            (
                StatusCode::TOO_MANY_REQUESTS,
                Json(json!({"code": 429, "msg": "quota exceeded"})),
            )
        }

        async fn mathpix(headers: HeaderMap, Json(body): Json<Value>) -> Json<Value> {
            assert_eq!(headers.get("app_id").and_then(|value| value.to_str().ok()), Some("app-id"));
            assert_eq!(headers.get("app_key").and_then(|value| value.to_str().ok()), Some("app-key"));
            assert!(body
                .get("src")
                .and_then(Value::as_str)
                .is_some_and(|src| src.starts_with("data:image/png;base64,")));
            Json(json!({"latex_styled": "\\frac{1}{2}"}))
        }

        let router = Router::new()
            .route("/responses", post(openai))
            .route("/api/chat", post(ollama))
            .route("/jobs-vl", post(paddle_vl_submit))
            .route("/jobs-vl/{job_id}", get(paddle_status))
            .route("/jobs-structure", post(paddle_structure_submit))
            .route("/jobs-structure/{job_id}", get(paddle_status))
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
                "prompt": "prompt",
                "encryptedApiKey": null
            },
            "ollama": {
                "baseUrl": "http://127.0.0.1:11434",
                "model": "",
                "prompt": "prompt"
            },
            "mathpix": {
                "baseUrl": "https://api.mathpix.com",
                "appId": "",
                "encryptedAppKey": null
            }
        }))
        .unwrap();
        assert_eq!(stored.paddle_ocr.model, "PaddleOCR-VL-1.6");
        assert!(stored.paddle_ocr.encrypted_access_token.is_none());
        assert_eq!(stored.simple_tex.model, "standard");
        assert!(stored.simple_tex.encrypted_access_token.is_none());

        let migrated = migrate_legacy_paddleocr_configuration(
            serde_json::from_value::<StoredOcrProviderConfiguration>(json!({
                "schemaVersion": 1,
                "activeProvider": "paddleocr",
                "openAiCompatible": {
                    "protocol": "responses",
                    "baseUrl": "https://api.openai.com/v1",
                    "model": "",
                    "prompt": "prompt",
                    "encryptedApiKey": null
                },
                "ollama": {
                    "baseUrl": "http://127.0.0.1:11434",
                    "model": "",
                    "prompt": "prompt"
                },
                "mathpix": {
                    "baseUrl": "https://api.mathpix.com",
                    "appId": "",
                    "encryptedAppKey": null
                },
                "paddleOcr": {
                    "model": "PaddleOCR-VL-1.5",
                    "encryptedAccessToken": null
                }
            }))
            .unwrap(),
        );
        assert_eq!(migrated.schema_version, CONFIGURATION_SCHEMA_VERSION);
        assert_eq!(migrated.paddle_ocr.model, "PaddleOCR-VL-1.6");

        let previous_default = migrate_legacy_paddleocr_configuration(
            serde_json::from_value::<StoredOcrProviderConfiguration>(json!({
                "schemaVersion": 1,
                "activeProvider": "paddleocr",
                "openAiCompatible": {
                    "protocol": "responses",
                    "baseUrl": "https://api.openai.com/v1",
                    "model": "",
                    "prompt": "prompt",
                    "encryptedApiKey": null
                },
                "ollama": {
                    "baseUrl": "http://127.0.0.1:11434",
                    "model": "",
                    "prompt": "prompt"
                },
                "mathpix": {
                    "baseUrl": "https://api.mathpix.com",
                    "appId": "",
                    "encryptedAppKey": null
                },
                "paddleOcr": {
                    "model": "PaddleOCR-VL-1.6",
                    "encryptedAccessToken": null
                }
            }))
            .unwrap(),
        );
        assert_eq!(previous_default.schema_version, CONFIGURATION_SCHEMA_VERSION);
        assert_eq!(previous_default.paddle_ocr.model, "PaddleOCR-VL-1.6");

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
                "res": {"latex": "$$\\frac{1}{2}$$", "conf": 0.98},
                "request_id": "tr_test"
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

    #[tokio::test]
    async fn remote_provider_http_protocols_include_paddle_models() {
        let base_url = spawn_mock_api().await;
        let client = reqwest::Client::builder()
            .no_proxy()
            .redirect(reqwest::redirect::Policy::none())
            .build()
            .unwrap();
        let data_url = "data:image/png;base64,iVBORw0KGgo=";

        let openai_configuration = StoredOpenAiCompatibleConfiguration {
            protocol: "responses".to_string(),
            base_url: base_url.clone(),
            model: "vision-model".to_string(),
            prompt: "Return formula JSON".to_string(),
            encrypted_api_key: None,
        };
        assert_eq!(
            recognize_openai_compatible(&client, &openai_configuration, None, data_url)
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

        let paddle_configuration = StoredPaddleOcrConfiguration {
            model: "PaddleOCR-VL-1.6".to_string(),
            encrypted_access_token: None,
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

        let quota_error = recognize_paddleocr_aistudio_at(
            &client,
            &paddle_configuration,
            "paddle-token",
            b"fake-png",
            "image/png",
            "png",
            &format!("{base_url}/quota"),
        )
        .await
        .unwrap_err();
        assert!(quota_error.to_ascii_lowercase().contains("quota"));

        let mathpix_configuration = StoredMathpixConfiguration {
            base_url: base_url,
            app_id: "app-id".to_string(),
            encrypted_app_key: None,
        };
        assert_eq!(
            recognize_mathpix(&client, &mathpix_configuration, "app-key", data_url)
                .await
                .unwrap(),
            vec!["\\frac{1}{2}"]
        );
    }

    #[tokio::test]
    #[ignore = "manual public OCR API connectivity probe"]
    async fn public_ocr_api_endpoints_are_reachable_with_app_proxy_client() {
        let client = build_ocr_http_client(
            Duration::from_secs(20),
            reqwest::redirect::Policy::none(),
            format!("VisualTeX/{} OCR-connectivity", env!("CARGO_PKG_VERSION")),
        )
        .unwrap();
        for (name, url) in [
            ("OpenAI", "https://api.openai.com/v1/models"),
            ("Mathpix", "https://api.mathpix.com/v3/text"),
            ("PaddleOCR", PADDLEOCR_JOBS_URL),
        ] {
            let response = client
                .get(url)
                .send()
                .await
                .unwrap_or_else(|error| panic!("{name} endpoint was not reachable through the VisualTeX OCR client: {error}"));
            println!("{name} connectivity status={}", response.status().as_u16());
        }

        match client.get("http://127.0.0.1:11434/api/tags").send().await {
            Ok(response) => println!("Ollama local connectivity status={}", response.status().as_u16()),
            Err(error) => println!("Ollama local service is not running/reachable: {error}"),
        }
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
