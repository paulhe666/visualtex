//! The parts of the local OCR runtime both backends have word for word: the PaddleOCR worker process and its
//! JSON protocol, the request and result types, the event bus the Office editor polls, the cached runtime status
//! and small helpers. Each backend includes this file as its ocr::common module and re-exports it, so `super`
//! is that backend's ocr module (its OcrState, RuntimePaths and imports).

use super::*;

pub(crate) const PADDLE_VERSION: &str = "3.3.1";
pub(crate) const PADDLEOCR_VERSION: &str = "3.7.0";
pub(crate) const MAX_IMAGE_BYTES: usize = 20 * 1024 * 1024;
pub(crate) const OCR_CANCELLED: &str = "OCR_CANCELLED";
pub(crate) const ALLOWED_MODELS: &[&str] = &[
    "PP-FormulaNet_plus-S",
    "PP-FormulaNet_plus-M",
    "PP-FormulaNet_plus-L",
    "PP-FormulaNet-S",
    "PP-FormulaNet-L",
];
pub(crate) const MAX_OCR_EVENTS: usize = 256;

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrEventRecord {
    pub(crate) id: u64,
    pub(crate) event: String,
    pub(crate) payload: Value,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct OcrEventEnvelope {
    pub(crate) cursor: u64,
    pub(crate) events: Vec<OcrEventRecord>,
}

#[derive(Clone, Default)]
pub(crate) struct OcrEventBus {
    pub(crate) next_id: Arc<AtomicU64>,
    pub(crate) events: Arc<Mutex<VecDeque<OcrEventRecord>>>,
}

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

#[derive(Debug, Deserialize)]
pub(crate) struct WorkerRecognitionResponse {
    pub(crate) ok: bool,
    pub(crate) model: Option<String>,
    pub(crate) elapsed_ms: Option<u64>,
    pub(crate) processed_width: Option<u32>,
    pub(crate) processed_height: Option<u32>,
    pub(crate) background_inverted: Option<bool>,
    pub(crate) background_luminance: Option<f64>,
    pub(crate) formulas: Option<Vec<OcrFormulaResult>>,
    pub(crate) error: Option<String>,
    pub(crate) details: Option<String>,
}

pub(crate) struct OcrWorker {
    pub(crate) child: Child,
    pub(crate) stdin: BufWriter<ChildStdin>,
    pub(crate) stdout: BufReader<ChildStdout>,
    pub(crate) pid_state: Arc<AtomicU32>,
    pub(crate) log_path: PathBuf,
    pub(crate) loaded_model: Option<String>,
}

impl OcrEventBus {
    pub(crate) fn publish<T: Serialize>(&self, event: &str, payload: &T) {
        let Ok(payload) = serde_json::to_value(payload) else {
            return;
        };
        let id = self.next_id.fetch_add(1, Ordering::SeqCst) + 1;
        if let Ok(mut events) = self.events.lock() {
            events.push_back(OcrEventRecord {
                id,
                event: event.to_string(),
                payload,
            });
            while events.len() > MAX_OCR_EVENTS {
                events.pop_front();
            }
        }
    }

    pub(crate) fn poll(&self, cursor: u64, event: Option<&str>) -> OcrEventEnvelope {
        let events = self
            .events
            .lock()
            .map(|events| {
                events
                    .iter()
                    .filter(|item| item.id > cursor)
                    .filter(|item| event.is_none_or(|name| item.event == name))
                    .cloned()
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        OcrEventEnvelope {
            cursor: self.next_id.load(Ordering::SeqCst),
            events,
        }
    }
}

impl Drop for OcrWorker {
    fn drop(&mut self) {
        let pid = self.child.id();
        let _ = self.child.kill();
        let _ = self.child.wait();
        let _ = self
            .pid_state
            .compare_exchange(pid, 0, Ordering::SeqCst, Ordering::SeqCst);
    }
}

pub(crate) fn read_worker_json<R: BufRead>(
    reader: &mut R,
    closed_message: &str,
    response_name: &str,
) -> Result<Value, String> {
    loop {
        let mut bytes = Vec::new();
        let count = reader
            .read_until(b'\n', &mut bytes)
            .map_err(|error| format!("Unable to read {response_name}: {error}"))?;
        if count == 0 {
            return Err(closed_message.to_string());
        }

        while bytes
            .last()
            .is_some_and(|byte| matches!(*byte, b'\n' | b'\r'))
        {
            bytes.pop();
        }
        if bytes.iter().all(u8::is_ascii_whitespace) {
            continue;
        }

        let first_non_whitespace = bytes
            .iter()
            .copied()
            .find(|byte| !byte.is_ascii_whitespace());
        if first_non_whitespace != Some(b'{') {
            // Native dependencies occasionally write diagnostics to stdout.
            // Ignore those lines so they cannot corrupt the JSON protocol.
            continue;
        }

        match serde_json::from_slice(&bytes) {
            Ok(value) => return Ok(value),
            Err(error) => {
                // This fallback prevents a legacy Windows code page from
                // crashing the reader. New workers always emit ASCII-safe JSON.
                let output = String::from_utf8_lossy(&bytes);
                return serde_json::from_str(output.trim()).map_err(|lossy_error| {
                    format!(
                        "{response_name} returned invalid JSON: {error}; UTF-8-lossy parse: {lossy_error}; output={output:?}"
                    )
                });
            }
        }
    }
}

pub(crate) fn tail_text(value: &str, max_chars: usize) -> String {
    let total = value.chars().count();
    if total <= max_chars {
        return value.to_string();
    }
    value.chars().skip(total - max_chars).collect()
}

pub(crate) fn is_final_ocr_state_owner(worker: &Arc<Mutex<Option<OcrWorker>>>) -> bool {
    Arc::strong_count(worker) == 1
}

pub(crate) async fn wait_for_ocr_cancellation(generation: Arc<AtomicU64>, expected: u64) {
    loop {
        if generation.load(Ordering::SeqCst) != expected {
            return;
        }
        tokio::time::sleep(Duration::from_millis(80)).await;
    }
}

pub(super) fn cleanup_worker_temp(paths: &RuntimePaths) -> Result<(), String> {
    if paths.temp.exists() {
        fs::remove_dir_all(&paths.temp)
            .map_err(|error| format!("Unable to clean OCR temporary files: {error}"))?;
    }
    fs::create_dir_all(&paths.temp)
        .map_err(|error| format!("Unable to create OCR temporary directory: {error}"))
}

pub(crate) fn stop_worker(
    worker: &Arc<Mutex<Option<OcrWorker>>>,
    worker_pid: &Arc<AtomicU32>,
) -> Result<(), String> {
    let terminate_result = terminate_worker_process(worker_pid);
    if let Ok(mut guard) = worker.lock() {
        guard.take();
    }
    terminate_result.map(|_| ())
}

pub(crate) fn worker_script_path(app: &AppHandle) -> Result<PathBuf, String> {
    if let Ok(path) = app.path().resolve("ocr/worker.py", BaseDirectory::Resource) {
        if path.exists() {
            return Ok(path);
        }
    }

    #[cfg(debug_assertions)]
    {
        let development_path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("ocr")
            .join("worker.py");
        if development_path.exists() {
            return Ok(development_path);
        }
    }

    Err("Unable to locate bundled OCR worker.py".to_string())
}

pub(crate) fn read_cached_runtime_status(
    cache: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
) -> Result<Option<OcrRuntimeStatus>, String> {
    cache
        .lock()
        .map(|status| status.clone())
        .map_err(|_| "OCR runtime status cache is unavailable".to_string())
}

pub(crate) fn write_cached_runtime_status(
    cache: &Arc<Mutex<Option<OcrRuntimeStatus>>>,
    status: Option<OcrRuntimeStatus>,
) -> Result<(), String> {
    let mut guard = cache
        .lock()
        .map_err(|_| "OCR runtime status cache is unavailable".to_string())?;
    *guard = status;
    Ok(())
}
