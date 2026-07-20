use std::{
    env,
    path::{Path, PathBuf},
    process::Stdio,
    time::Duration,
};

use serde_json::Value;
use tauri::{AppHandle, Manager};
use tokio::{
    io::{AsyncBufReadExt, AsyncWriteExt, BufReader},
    process::{Child, ChildStdin, ChildStdout, Command},
    time::timeout,
};

use super::protocol::{OcrError, SidecarResponse};

#[derive(Clone, Debug)]
pub struct RuntimePaths {
    pub python: PathBuf,
    pub script: PathBuf,
}

struct OcrProcess {
    child: Child,
    stdin: ChildStdin,
    stdout: BufReader<ChildStdout>,
}

pub struct OcrManager {
    process: Option<OcrProcess>,
}

impl Default for OcrManager {
    fn default() -> Self {
        Self { process: None }
    }
}

impl OcrManager {
    pub fn is_running(&mut self) -> bool {
        if let Some(process) = self.process.as_mut() {
            match process.child.try_wait() {
                Ok(None) => true,
                Ok(Some(_)) | Err(_) => {
                    self.process = None;
                    false
                }
            }
        } else {
            false
        }
    }

    pub async fn request(
        &mut self,
        app: &AppHandle,
        payload: Value,
        timeout_duration: Duration,
    ) -> Result<SidecarResponse, OcrError> {
        let request_id = payload
            .get("id")
            .and_then(Value::as_str)
            .unwrap_or("unknown")
            .to_owned();

        let mut last_transport_error: Option<OcrError> = None;
        for attempt in 0..2 {
            if self.process.is_none() {
                self.process = Some(spawn_process(app).await?);
            }

            let result = self
                .send_once(&request_id, &payload, timeout_duration)
                .await;
            match result {
                Ok(response) => return decode_response(response),
                Err(error) => {
                    last_transport_error = Some(error);
                    self.stop().await;
                    if attempt == 1 {
                        break;
                    }
                }
            }
        }

        Err(last_transport_error.unwrap_or_else(|| {
            OcrError::new("SIDECAR_CRASHED", "The OCR process stopped unexpectedly")
        }))
    }

    async fn send_once(
        &mut self,
        request_id: &str,
        payload: &Value,
        timeout_duration: Duration,
    ) -> Result<Value, OcrError> {
        let process = self.process.as_mut().ok_or_else(|| {
            OcrError::new("SIDECAR_CRASHED", "The OCR process is not running")
        })?;

        let mut message = serde_json::to_vec(payload).map_err(|error| {
            OcrError::with_detail(
                "INVALID_REQUEST",
                "Unable to serialize OCR request",
                error.to_string(),
            )
        })?;
        message.push(b'\n');

        process.stdin.write_all(&message).await.map_err(|error| {
            OcrError::with_detail(
                "SIDECAR_CRASHED",
                "Unable to send request to OCR process",
                error.to_string(),
            )
        })?;
        process.stdin.flush().await.map_err(|error| {
            OcrError::with_detail(
                "SIDECAR_CRASHED",
                "Unable to flush OCR request",
                error.to_string(),
            )
        })?;

        let mut line = String::new();
        let bytes_read = timeout(timeout_duration, process.stdout.read_line(&mut line))
            .await
            .map_err(|_| OcrError::new("OCR_TIMEOUT", "Formula recognition timed out"))?
            .map_err(|error| {
                OcrError::with_detail(
                    "SIDECAR_CRASHED",
                    "Unable to read OCR response",
                    error.to_string(),
                )
            })?;

        if bytes_read == 0 {
            return Err(OcrError::new(
                "SIDECAR_CRASHED",
                "The OCR process closed its output stream",
            ));
        }

        let response: Value = serde_json::from_str(line.trim()).map_err(|error| {
            OcrError::with_detail(
                "SIDECAR_CRASHED",
                "The OCR process returned invalid JSON",
                format!("{error}: {line}"),
            )
        })?;
        let response_id = response
            .get("id")
            .and_then(Value::as_str)
            .unwrap_or("unknown");
        if response_id != request_id {
            return Err(OcrError::with_detail(
                "SIDECAR_CRASHED",
                "OCR response id did not match the request",
                format!("expected {request_id}, received {response_id}"),
            ));
        }

        Ok(response)
    }

    pub async fn stop(&mut self) {
        if let Some(mut process) = self.process.take() {
            let _ = process.child.kill().await;
            let _ = process.child.wait().await;
        }
    }
}

fn decode_response(value: Value) -> Result<SidecarResponse, OcrError> {
    let response: SidecarResponse = serde_json::from_value(value).map_err(|error| {
        OcrError::with_detail(
            "SIDECAR_CRASHED",
            "Unable to decode OCR response",
            error.to_string(),
        )
    })?;

    if response.ok {
        return Ok(response);
    }

    let error = response.error.unwrap_or_else(|| super::protocol::SidecarErrorResponse {
        code: "INFERENCE_FAILED".to_owned(),
        message: "Formula recognition failed".to_owned(),
        detail: None,
    });
    Err(OcrError {
        code: error.code,
        message: error.message,
        detail: error.detail,
    })
}

async fn spawn_process(app: &AppHandle) -> Result<OcrProcess, OcrError> {
    let runtime = resolve_runtime(app)?;
    let script_parent = runtime.script.parent().unwrap_or_else(|| Path::new("/"));

    let mut command = Command::new(&runtime.python);
    command
        .arg(&runtime.script)
        .current_dir(script_parent)
        .env("PYTHONUNBUFFERED", "1")
        .env("PADDLE_PDX_MODEL_SOURCE", "BOS")
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .kill_on_drop(true);

    let mut child = command.spawn().map_err(|error| {
        OcrError::with_detail(
            "PYTHON_ENV_MISSING",
            "Unable to start the VisualTeX OCR runtime",
            error.to_string(),
        )
    })?;
    let stdin = child.stdin.take().ok_or_else(|| {
        OcrError::new("SIDECAR_CRASHED", "Unable to open OCR stdin")
    })?;
    let stdout = child.stdout.take().ok_or_else(|| {
        OcrError::new("SIDECAR_CRASHED", "Unable to open OCR stdout")
    })?;

    if let Some(stderr) = child.stderr.take() {
        tokio::spawn(async move {
            let mut lines = BufReader::new(stderr).lines();
            while let Ok(Some(line)) = lines.next_line().await {
                eprintln!("[VisualTeX OCR] {line}");
            }
        });
    }

    Ok(OcrProcess {
        child,
        stdin,
        stdout: BufReader::new(stdout),
    })
}

pub fn resolve_runtime(app: &AppHandle) -> Result<RuntimePaths, OcrError> {
    #[cfg(debug_assertions)]
    let project_root = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .map(Path::to_path_buf);
    #[cfg(not(debug_assertions))]
    let project_root: Option<PathBuf> = None;
    let resource_dir = app.path().resource_dir().ok();

    let mut script_candidates = Vec::new();
    if let Ok(custom) = env::var("VISUALTEX_OCR_SERVER") {
        script_candidates.push(PathBuf::from(custom));
    }
    if let Some(resource) = resource_dir.as_ref() {
        script_candidates.push(resource.join("ocr/formula_ocr_server.py"));
        script_candidates.push(resource.join("_up_/ocr/formula_ocr_server.py"));
    }
    if let Some(root) = project_root.as_ref() {
        script_candidates.push(root.join("ocr/formula_ocr_server.py"));
    }
    let script = first_existing(script_candidates).ok_or_else(|| {
        OcrError::new(
            "PYTHON_ENV_MISSING",
            "VisualTeX OCR server script was not found",
        )
    })?;

    let mut python_candidates = Vec::new();
    if let Ok(custom) = env::var("VISUALTEX_OCR_PYTHON") {
        python_candidates.push(PathBuf::from(custom));
    }
    if let Some(resource) = resource_dir.as_ref() {
        python_candidates.push(resource.join("ocr-runtime/bin/python3"));
        python_candidates.push(resource.join("ocr-runtime/bin/python"));
    }
    if let Some(root) = project_root.as_ref() {
        python_candidates.push(root.join("ocr/.venv/bin/python"));
    }
    python_candidates.push(PathBuf::from("/usr/bin/python3"));

    let python = first_existing(python_candidates).ok_or_else(|| {
        OcrError::new(
            "PYTHON_ENV_MISSING",
            "No compatible Python runtime was found for formula OCR",
        )
    })?;

    Ok(RuntimePaths { python, script })
}

fn first_existing(paths: Vec<PathBuf>) -> Option<PathBuf> {
    paths.into_iter().find(|path| path.is_file())
}
