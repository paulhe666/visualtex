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

#[path = "../../../../shared/visualtex-backend/src/ocr_provider_common.rs"]
mod common;
pub(crate) use common::*;

const KEYCHAIN_SERVICE: &str = "com.visualtex.studio.ocr-providers";
const OPENAI_API_KEY_ACCOUNT: &str = "openai-compatible-api-key";
const MATHPIX_APP_KEY_ACCOUNT: &str = "mathpix-app-key";
const PADDLEOCR_ACCESS_TOKEN_ACCOUNT: &str = "paddleocr-aistudio-access-token";
const SIMPLETEX_ACCESS_TOKEN_ACCOUNT: &str = "simpletex-user-access-token";
const MAX_SECRET_CHARS: usize = 8_192;

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct OpenAiCompatibleConfigurationView {
    protocol: String,
    base_url: String,
    model: String,
    prompt: String,
    has_api_key: Option<bool>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct MathpixConfigurationView {
    base_url: String,
    app_id: String,
    has_app_key: Option<bool>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct PaddleOcrConfigurationView {
    model: String,
    has_access_token: Option<bool>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
struct SimpleTexConfigurationView {
    model: String,
    has_access_token: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredOpenAiCompatibleConfiguration {
    protocol: String,
    base_url: String,
    model: String,
    prompt: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    has_api_key: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredMathpixConfiguration {
    base_url: String,
    app_id: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    has_app_key: Option<bool>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredPaddleOcrConfiguration {
    model: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    has_access_token: Option<bool>,
}

impl Default for StoredPaddleOcrConfiguration {
    fn default() -> Self {
        Self {
            model: "PaddleOCR-VL-1.6".to_string(),
            has_access_token: Some(false),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct StoredSimpleTexConfiguration {
    model: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    has_access_token: Option<bool>,
}

impl Default for StoredSimpleTexConfiguration {
    fn default() -> Self {
        Self {
            model: "standard".to_string(),
            has_access_token: Some(false),
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
                has_api_key: Some(false),
            },
            ollama: StoredOllamaConfiguration {
                base_url: "http://127.0.0.1:11434".to_string(),
                model: String::new(),
                prompt: DEFAULT_PROMPT.to_string(),
            },
            mathpix: StoredMathpixConfiguration {
                base_url: "https://api.mathpix.com".to_string(),
                app_id: String::new(),
                has_app_key: Some(false),
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

        let openai_replacement = normalize_secret_update(update.open_ai_compatible.api_key)?;
        let mathpix_replacement = normalize_secret_update(update.mathpix.app_key)?;
        let paddle_ocr_replacement = normalize_secret_update(update.paddle_ocr.access_token)?;
        let simpletex_replacement = normalize_secret_update(update.simple_tex.access_token)?;
        let has_openai_secret = updated_secret_status(
            previous.open_ai_compatible.has_api_key,
            openai_replacement.as_deref(),
            update.open_ai_compatible.clear_api_key,
        );
        let has_mathpix_secret = updated_secret_status(
            previous.mathpix.has_app_key,
            mathpix_replacement.as_deref(),
            update.mathpix.clear_app_key,
        );
        let has_paddle_ocr_secret = updated_secret_status(
            previous.paddle_ocr.has_access_token,
            paddle_ocr_replacement.as_deref(),
            update.paddle_ocr.clear_access_token,
        );
        let has_simpletex_secret = updated_secret_status(
            previous.simple_tex.has_access_token,
            simpletex_replacement.as_deref(),
            update.simple_tex.clear_access_token,
        );

        // Unknown means a pre-migration Keychain item may exist. Preserve the
        // transport guard without reading any credential while saving settings.
        validate_secret_transport(&openai_base_url, has_openai_secret != Some(false), "API key")?;
        validate_secret_transport(&mathpix_base_url, has_mathpix_secret != Some(false), "app_key")?;

        let next = StoredOcrProviderConfiguration {
            schema_version: CONFIGURATION_SCHEMA_VERSION,
            active_provider,
            open_ai_compatible: StoredOpenAiCompatibleConfiguration {
                protocol,
                base_url: openai_base_url,
                model: openai_model,
                prompt: openai_prompt,
                has_api_key: has_openai_secret,
            },
            ollama: StoredOllamaConfiguration {
                base_url: ollama_base_url,
                model: ollama_model,
                prompt: ollama_prompt,
            },
            mathpix: StoredMathpixConfiguration {
                base_url: mathpix_base_url,
                app_id: mathpix_app_id,
                has_app_key: has_mathpix_secret,
            },
            paddle_ocr: StoredPaddleOcrConfiguration {
                model: paddle_ocr_model,
                has_access_token: has_paddle_ocr_secret,
            },
            simple_tex: StoredSimpleTexConfiguration {
                model: simpletex_model,
                has_access_token: has_simpletex_secret,
            },
        };
        validate_active_provider_configuration(
            &next,
            has_mathpix_secret != Some(false),
            has_paddle_ocr_secret != Some(false),
            has_simpletex_secret != Some(false),
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
        validate_active_provider_configuration(
            &configuration,
            configuration.mathpix.has_app_key != Some(false),
            configuration.paddle_ocr.has_access_token != Some(false),
            configuration.simple_tex.has_access_token != Some(false),
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
            has_api_key: configuration.open_ai_compatible.has_api_key,
        },
        ollama: OllamaConfigurationView {
            base_url: configuration.ollama.base_url.clone(),
            model: configuration.ollama.model.clone(),
            prompt: configuration.ollama.prompt.clone(),
        },
        mathpix: MathpixConfigurationView {
            base_url: configuration.mathpix.base_url.clone(),
            app_id: configuration.mathpix.app_id.clone(),
            has_app_key: configuration.mathpix.has_app_key,
        },
        paddle_ocr: PaddleOcrConfigurationView {
            model: normalize_paddleocr_model(&configuration.paddle_ocr.model)?.to_string(),
            has_access_token: configuration.paddle_ocr.has_access_token,
        },
        simple_tex: SimpleTexConfigurationView {
            model: normalize_simpletex_model(&configuration.simple_tex.model)?.to_string(),
            has_access_token: configuration.simple_tex.has_access_token,
        },
    })
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

fn updated_secret_status(
    previous: Option<bool>,
    replacement: Option<&str>,
    clear: bool,
) -> Option<bool> {
    if clear {
        Some(false)
    } else if replacement.is_some() {
        Some(true)
    } else {
        previous
    }
}

fn apply_secret_update(
    account: &str,
    replacement: Option<&str>,
    clear: bool,
) -> Result<(), String> {
    if !clear && replacement.is_none() {
        return Ok(());
    }
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

impl Drop for SensitiveString {
    fn drop(&mut self) {
        unsafe { self.0.as_mut_vec().fill(0) };
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
            let api_key = if configuration.open_ai_compatible.has_api_key == Some(false) {
                None
            } else {
                read_secret(OPENAI_API_KEY_ACCOUNT)?
            };
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
        assert_eq!(configuration.open_ai_compatible.has_api_key, Some(false));
    }

    #[test]
    fn provider_configuration_view_does_not_read_legacy_keychain_entries() {
        let mut configuration = StoredOcrProviderConfiguration::default();
        configuration.open_ai_compatible.has_api_key = None;
        configuration.mathpix.has_app_key = None;
        configuration.paddle_ocr.has_access_token = None;
        configuration.simple_tex.has_access_token = None;
        let view = configuration_view(&configuration).unwrap();
        assert_eq!(view.open_ai_compatible.has_api_key, None);
        assert_eq!(view.mathpix.has_app_key, None);
        assert_eq!(view.paddle_ocr.has_access_token, None);
        assert_eq!(view.simple_tex.has_access_token, None);

        assert_eq!(updated_secret_status(None, None, false), None);
        assert_eq!(updated_secret_status(None, Some("new-key"), false), Some(true));
        assert_eq!(updated_secret_status(Some(true), None, true), Some(false));
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
            has_api_key: None,
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
            has_access_token: None,
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
            has_access_token: None,
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
            has_access_token: None,
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
            has_app_key: None,
        };
        assert_eq!(
            recognize_mathpix(&client, &mathpix_configuration, "app-key", data_url,)
                .await
                .unwrap(),
            vec!["\\frac{1}{2}"]
        );
    }
}
