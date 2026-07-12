use std::fs;
use std::io::Read;
use std::path::{Path, PathBuf};
use std::process::Stdio;
use std::time::Duration;

use image::{ImageFormat, ImageReader, Limits};
use serde::de::DeserializeOwned;
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
use tokio::io::{AsyncBufReadExt, AsyncWriteExt, BufReader, Lines};
use tokio::process::{Child, ChildStdin, ChildStdout, Command};
use tokio::time::timeout;
use vt_protocol::{DocumentOcrResult, FormulaOcrResult, OcrWorkerHealth};

#[derive(Clone, Debug)]
pub struct OcrWorkerConfig {
    pub python_executable: PathBuf,
    pub worker_script: PathBuf,
    pub project_root: PathBuf,
    pub request_timeout: Duration,
    pub formula_model_dir: Option<PathBuf>,
    pub formula_model_name: Option<String>,
    pub document_pipeline_config: Option<PathBuf>,
    pub document_package_root: Option<PathBuf>,
    pub document_model_name: Option<String>,
    pub device: Option<String>,
    pub mock: bool,
}

impl OcrWorkerConfig {
    pub fn bundled(project_root: impl AsRef<Path>) -> Self {
        Self {
            python_executable: PathBuf::from(if cfg!(windows) { "python" } else { "python3" }),
            worker_script: PathBuf::from(env!("CARGO_MANIFEST_DIR"))
                .join("../../services/ocr-python/worker.py"),
            project_root: project_root.as_ref().to_path_buf(),
            request_timeout: Duration::from_secs(120),
            formula_model_dir: None,
            formula_model_name: None,
            document_pipeline_config: None,
            document_package_root: None,
            document_model_name: None,
            device: None,
            mock: false,
        }
    }
}

#[derive(Debug, thiserror::Error)]
pub enum OcrError {
    #[error("OCR input escapes the project root: {0}")]
    PathEscape(PathBuf),
    #[error("invalid OCR image input: {0}")]
    InvalidInput(String),
    #[error("failed to decode or normalize OCR image: {0}")]
    Image(String),
    #[error("OCR worker did not provide stdin or stdout")]
    MissingPipe,
    #[error("OCR request timed out after {0:?}")]
    Timeout(Duration),
    #[error("OCR worker closed its output")]
    WorkerClosed,
    #[error("OCR worker returned request id {received}, expected {expected}")]
    ResponseMismatch { expected: u64, received: Value },
    #[error("OCR worker error {code}: {message}")]
    WorkerError {
        code: i64,
        message: String,
        data: Option<Value>,
    },
    #[error(transparent)]
    Io(#[from] std::io::Error),
    #[error(transparent)]
    Json(#[from] serde_json::Error),
}

pub fn import_image(
    project_root: impl AsRef<Path>,
    source: impl AsRef<Path>,
) -> Result<PathBuf, OcrError> {
    const MAX_SOURCE_BYTES: u64 = 64 * 1024 * 1024;
    const MAX_DIMENSION: u32 = 12_000;
    const MAX_ALLOC_BYTES: u64 = 256 * 1024 * 1024;

    let project_root = project_root.as_ref().canonicalize()?;
    let source = source.as_ref();
    let metadata = fs::symlink_metadata(source)?;
    if metadata.file_type().is_symlink() || !metadata.is_file() {
        return Err(OcrError::InvalidInput(
            "input must be a regular image file, not a symbolic link".to_owned(),
        ));
    }
    if metadata.len() > MAX_SOURCE_BYTES {
        return Err(OcrError::InvalidInput(format!(
            "image exceeds the {} MiB input limit",
            MAX_SOURCE_BYTES / 1024 / 1024
        )));
    }
    let source = source.canonicalize()?;
    let mut reader = ImageReader::open(&source)?
        .with_guessed_format()
        .map_err(|error| OcrError::Image(error.to_string()))?;
    let mut limits = Limits::default();
    limits.max_image_width = Some(MAX_DIMENSION);
    limits.max_image_height = Some(MAX_DIMENSION);
    limits.max_alloc = Some(MAX_ALLOC_BYTES);
    reader.limits(limits);
    let image = reader
        .decode()
        .map_err(|error| OcrError::Image(error.to_string()))?;

    let mut file = fs::File::open(&source)?;
    let mut hasher = Sha256::new();
    let mut buffer = [0_u8; 1024 * 1024];
    loop {
        let read = file.read(&mut buffer)?;
        if read == 0 {
            break;
        }
        hasher.update(&buffer[..read]);
    }
    let digest = format!("{:x}", hasher.finalize());
    let output_dir = project_root.join(".visualtex/ocr-input");
    fs::create_dir_all(&output_dir)?;
    let destination = output_dir.join(format!("{digest}.png"));
    if destination.is_file() {
        return Ok(destination.canonicalize()?);
    }
    let temporary = output_dir.join(format!(".{digest}.{}.tmp.png", std::process::id()));
    if temporary.exists() {
        fs::remove_file(&temporary)?;
    }
    image
        .save_with_format(&temporary, ImageFormat::Png)
        .map_err(|error| OcrError::Image(error.to_string()))?;
    match fs::rename(&temporary, &destination) {
        Ok(()) => {}
        Err(_) if destination.is_file() => {
            let _ = fs::remove_file(&temporary);
        }
        Err(error) => return Err(error.into()),
    }
    Ok(destination.canonicalize()?)
}

pub fn import_project_image(
    project_root: impl AsRef<Path>,
    source: impl AsRef<Path>,
) -> Result<PathBuf, OcrError> {
    let project_root = project_root.as_ref().canonicalize()?;
    let source = source.as_ref();
    let candidate = if source.is_absolute() {
        source.to_path_buf()
    } else {
        project_root.join(source)
    };
    let candidate = candidate.canonicalize()?;
    if !candidate.starts_with(&project_root) {
        return Err(OcrError::PathEscape(candidate));
    }
    import_image(project_root, candidate)
}

pub struct OcrWorker {
    config: OcrWorkerConfig,
    child: Child,
    stdin: ChildStdin,
    stdout: Lines<BufReader<ChildStdout>>,
    next_id: u64,
}

impl OcrWorker {
    pub async fn spawn(config: OcrWorkerConfig) -> Result<Self, OcrError> {
        let project_root = config.project_root.canonicalize()?;
        let worker_script = config.worker_script.canonicalize()?;
        let mut command = Command::new(&config.python_executable);
        command
            .arg(worker_script)
            .current_dir(&project_root)
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::inherit())
            .kill_on_drop(true)
            .env("PYTHONUNBUFFERED", "1")
            .env_remove("VISUALTEX_FORMULA_MODEL_DIR")
            .env_remove("VISUALTEX_OCR_MODEL_DIR")
            .env_remove("VISUALTEX_FORMULA_MODEL_NAME")
            .env_remove("VISUALTEX_DOCUMENT_PIPELINE_CONFIG")
            .env_remove("VISUALTEX_DOCUMENT_PACKAGE_ROOT")
            .env_remove("VISUALTEX_DOCUMENT_MODEL_NAME")
            .env_remove("VISUALTEX_OCR_DEVICE");
        if let Some(model_dir) = &config.formula_model_dir {
            command.env("VISUALTEX_FORMULA_MODEL_DIR", model_dir);
        }
        if let Some(model_name) = &config.formula_model_name {
            command.env("VISUALTEX_FORMULA_MODEL_NAME", model_name);
        }
        if let Some(config_path) = &config.document_pipeline_config {
            command.env("VISUALTEX_DOCUMENT_PIPELINE_CONFIG", config_path);
        }
        if let Some(package_root) = &config.document_package_root {
            command.env("VISUALTEX_DOCUMENT_PACKAGE_ROOT", package_root);
        }
        if let Some(model_name) = &config.document_model_name {
            command.env("VISUALTEX_DOCUMENT_MODEL_NAME", model_name);
        }
        if let Some(device) = &config.device {
            command.env("VISUALTEX_OCR_DEVICE", device);
        }
        if config.mock {
            command.arg("--mock");
        }
        let mut child = command.spawn()?;
        let stdin = child.stdin.take().ok_or(OcrError::MissingPipe)?;
        let stdout = child.stdout.take().ok_or(OcrError::MissingPipe)?;
        Ok(Self {
            config: OcrWorkerConfig {
                project_root,
                ..config
            },
            child,
            stdin,
            stdout: BufReader::new(stdout).lines(),
            next_id: 1,
        })
    }

    pub async fn health(&mut self) -> Result<OcrWorkerHealth, OcrError> {
        self.request("health", json!({})).await
    }

    pub async fn capabilities(&mut self) -> Result<Value, OcrError> {
        self.request("capabilities", json!({})).await
    }

    pub async fn document_health(&mut self) -> Result<OcrWorkerHealth, OcrError> {
        self.request("document.health", json!({})).await
    }

    pub async fn recognize_formula(
        &mut self,
        image_path: impl AsRef<Path>,
    ) -> Result<FormulaOcrResult, OcrError> {
        let image_path = self.validate_input(image_path.as_ref())?;
        self.request(
            "recognizeFormula",
            json!({ "imagePath": image_path, "maxCandidates": 5 }),
        )
        .await
    }

    pub async fn recognize_document(
        &mut self,
        image_path: impl AsRef<Path>,
    ) -> Result<DocumentOcrResult, OcrError> {
        let image_path = self.validate_input(image_path.as_ref())?;
        self.request(
            "recognizeDocument",
            json!({ "imagePath": image_path, "preserveReadingOrder": true }),
        )
        .await
    }

    pub async fn shutdown(mut self) -> Result<(), OcrError> {
        let _: Value = self.request("shutdown", json!({})).await?;
        let _ = timeout(Duration::from_secs(2), self.child.wait()).await;
        Ok(())
    }

    pub async fn restart(&mut self) -> Result<(), OcrError> {
        let _ = self.child.kill().await;
        let replacement = Self::spawn(self.config.clone()).await?;
        *self = replacement;
        Ok(())
    }

    async fn request<T: DeserializeOwned>(
        &mut self,
        method: &str,
        params: Value,
    ) -> Result<T, OcrError> {
        let id = self.next_id;
        self.next_id = self.next_id.saturating_add(1);
        let request = RpcRequest {
            jsonrpc: "2.0",
            id,
            method,
            params,
        };
        let mut bytes = serde_json::to_vec(&request)?;
        bytes.push(b'\n');
        self.stdin.write_all(&bytes).await?;
        self.stdin.flush().await?;

        let line = timeout(self.config.request_timeout, self.stdout.next_line())
            .await
            .map_err(|_| OcrError::Timeout(self.config.request_timeout))??
            .ok_or(OcrError::WorkerClosed)?;
        let response: RpcResponse = serde_json::from_str(&line)?;
        if response.id.as_u64() != Some(id) {
            return Err(OcrError::ResponseMismatch {
                expected: id,
                received: response.id,
            });
        }
        if let Some(error) = response.error {
            return Err(OcrError::WorkerError {
                code: error.code,
                message: error.message,
                data: error.data,
            });
        }
        Ok(serde_json::from_value(
            response.result.unwrap_or(Value::Null),
        )?)
    }

    fn validate_input(&self, input: &Path) -> Result<PathBuf, OcrError> {
        const MAX_INPUT_BYTES: u64 = 64 * 1024 * 1024;
        let candidate = if input.is_absolute() {
            input.to_path_buf()
        } else {
            self.config.project_root.join(input)
        };
        let metadata = fs::symlink_metadata(&candidate)?;
        if metadata.file_type().is_symlink() || !metadata.is_file() {
            return Err(OcrError::InvalidInput(
                "worker input must be a regular non-symlink file".to_owned(),
            ));
        }
        if metadata.len() > MAX_INPUT_BYTES {
            return Err(OcrError::InvalidInput(format!(
                "worker input exceeds the {} MiB limit",
                MAX_INPUT_BYTES / 1024 / 1024
            )));
        }
        let candidate = candidate.canonicalize()?;
        if !candidate.starts_with(&self.config.project_root) {
            return Err(OcrError::PathEscape(candidate));
        }
        Ok(candidate)
    }
}

#[derive(Serialize)]
struct RpcRequest<'a> {
    jsonrpc: &'static str,
    id: u64,
    method: &'a str,
    params: Value,
}

#[derive(Deserialize)]
struct RpcResponse {
    id: Value,
    #[serde(default)]
    result: Option<Value>,
    #[serde(default)]
    error: Option<RpcError>,
}

#[derive(Deserialize)]
struct RpcError {
    code: i64,
    message: String,
    #[serde(default)]
    data: Option<Value>,
}

#[cfg(test)]
mod tests {
    use super::*;
    use image::{Rgb, RgbImage, Rgba, RgbaImage};
    use tempfile::tempdir;

    #[test]
    fn image_import_normalizes_and_deduplicates_png() {
        let project = tempdir().unwrap();
        let input = tempdir().unwrap();
        let source = input.path().join("formula.jpg");
        RgbImage::from_pixel(24, 12, Rgb([255, 255, 255]))
            .save(&source)
            .unwrap();
        let first = import_image(project.path(), &source).unwrap();
        let second = import_image(project.path(), &source).unwrap();
        assert_eq!(first, second);
        assert!(first.starts_with(project.path().canonicalize().unwrap()));
        assert_eq!(
            first.extension().and_then(|value| value.to_str()),
            Some("png")
        );
        let imported = image::open(&first).unwrap();
        assert_eq!((imported.width(), imported.height()), (24, 12));
    }

    #[test]
    fn project_image_import_rejects_outside_path() {
        let project = tempdir().unwrap();
        let outside = tempdir().unwrap();
        let source = outside.path().join("outside.png");
        RgbaImage::from_pixel(4, 4, Rgba([0, 0, 0, 255]))
            .save(&source)
            .unwrap();
        assert!(matches!(
            import_project_image(project.path(), &source),
            Err(OcrError::PathEscape(_))
        ));
    }

    #[test]
    fn image_import_rejects_oversized_file_before_decode() {
        let project = tempdir().unwrap();
        let input = tempdir().unwrap();
        let source = input.path().join("oversized.png");
        let file = fs::File::create(&source).unwrap();
        file.set_len(64 * 1024 * 1024 + 1).unwrap();
        assert!(matches!(
            import_image(project.path(), &source),
            Err(OcrError::InvalidInput(message)) if message.contains("64 MiB")
        ));
    }

    #[test]
    fn image_import_rejects_dimensions_over_limit() {
        let project = tempdir().unwrap();
        let input = tempdir().unwrap();
        let source = input.path().join("too-wide.png");
        RgbaImage::from_pixel(12_001, 1, Rgba([0, 0, 0, 255]))
            .save(&source)
            .unwrap();
        assert!(matches!(
            import_image(project.path(), &source),
            Err(OcrError::Image(_))
        ));
    }

    #[cfg(unix)]
    #[test]
    fn image_import_rejects_symbolic_links() {
        use std::os::unix::fs::symlink;

        let project = tempdir().unwrap();
        let input = tempdir().unwrap();
        let target = input.path().join("target.png");
        let link = input.path().join("link.png");
        RgbaImage::from_pixel(4, 4, Rgba([0, 0, 0, 255]))
            .save(&target)
            .unwrap();
        symlink(&target, &link).unwrap();
        assert!(matches!(
            import_image(project.path(), &link),
            Err(OcrError::InvalidInput(message)) if message.contains("symbolic link")
        ));
    }

    #[tokio::test]
    async fn mock_worker_round_trip_is_local_and_typed() {
        let temp = tempdir().unwrap();
        let image = temp.path().join("formula.png");
        std::fs::write(&image, b"not decoded in mock mode").unwrap();
        let mut config = OcrWorkerConfig::bundled(temp.path());
        config.mock = true;
        config.request_timeout = Duration::from_secs(5);
        let mut worker = OcrWorker::spawn(config).await.unwrap();
        let health = worker.health().await.unwrap();
        assert!(health.available);
        assert_eq!(health.backend, "mock");
        let capabilities = worker.capabilities().await.unwrap();
        assert_eq!(capabilities["formulaRecognition"]["available"], true);
        assert_eq!(capabilities["documentRecognition"]["available"], true);
        assert_eq!(capabilities["offlineOnly"], true);
        let result = worker.recognize_formula(&image).await.unwrap();
        assert!(!result.candidates.is_empty());
        assert!(result.candidates[0].latex.contains("\\int"));
        let document = worker.recognize_document(&image).await.unwrap();
        assert_eq!(document.reading_order, vec![0, 1, 2]);
        assert_eq!(document.regions[0].kind, "document_title");
        assert_eq!(document.regions[2].latex.as_deref(), Some("E=mc^2"));
        worker.shutdown().await.unwrap();
    }

    #[tokio::test]
    async fn local_formula_model_configuration_reaches_worker_health() {
        let temp = tempdir().unwrap();
        let model = temp.path().join("model");
        std::fs::create_dir(&model).unwrap();
        std::fs::write(
            model.join("visualtex-model.json"),
            br#"{"version":"formula-test-2026"}"#,
        )
        .unwrap();
        let mut config = OcrWorkerConfig::bundled(temp.path());
        config.formula_model_dir = Some(model);
        config.formula_model_name = Some("formula-test".to_owned());
        config.request_timeout = Duration::from_secs(5);
        let mut worker = OcrWorker::spawn(config).await.unwrap();
        let health = worker.health().await.unwrap();
        assert!(health.available);
        assert_eq!(health.backend, "paddleocr-formula");
        assert_eq!(health.model_version.as_deref(), Some("formula-test-2026"));
        worker.shutdown().await.unwrap();
    }

    #[tokio::test]
    async fn missing_model_returns_structured_worker_error() {
        let temp = tempdir().unwrap();
        let image = temp.path().join("formula.png");
        std::fs::write(&image, b"not decoded because health fails first").unwrap();
        let mut config = OcrWorkerConfig::bundled(temp.path());
        config.request_timeout = Duration::from_secs(5);
        let mut worker = OcrWorker::spawn(config).await.unwrap();
        let health = worker.health().await.unwrap();
        assert!(!health.available);
        let error = worker.recognize_formula(&image).await.unwrap_err();
        assert!(matches!(
            error,
            OcrError::WorkerError { code: -32001, message, .. }
                if message.contains("VISUALTEX_FORMULA_MODEL_DIR")
        ));
        worker.shutdown().await.unwrap();
    }

    #[tokio::test]
    async fn missing_document_pipeline_returns_structured_worker_error() {
        let temp = tempdir().unwrap();
        let image = temp.path().join("page.png");
        std::fs::write(&image, b"not decoded because health fails first").unwrap();
        let mut config = OcrWorkerConfig::bundled(temp.path());
        config.request_timeout = Duration::from_secs(5);
        let mut worker = OcrWorker::spawn(config).await.unwrap();
        let health = worker.document_health().await.unwrap();
        assert!(!health.available);
        let error = worker.recognize_document(&image).await.unwrap_err();
        assert!(matches!(
            error,
            OcrError::WorkerError { code: -32001, message, .. }
                if message.contains("layout_ocr")
        ));
        worker.shutdown().await.unwrap();
    }

    #[cfg(unix)]
    #[tokio::test]
    async fn rejects_image_outside_project_scope() {
        let project = tempdir().unwrap();
        let outside = tempdir().unwrap();
        let image = outside.path().join("outside.png");
        std::fs::write(&image, b"x").unwrap();
        let mut config = OcrWorkerConfig::bundled(project.path());
        config.mock = true;
        let mut worker = OcrWorker::spawn(config).await.unwrap();
        assert!(matches!(
            worker.recognize_formula(&image).await,
            Err(OcrError::PathEscape(_))
        ));
        worker.shutdown().await.unwrap();
    }
}
