//! The offline Python runtime: how its processes start and probing what is installed.

use super::*;

pub(super) fn command_output(command: &mut Command, label: &str) -> Result<String, String> {
    let output = command
        .output()
        .map_err(|error| format!("Unable to start {label}: {error}"))?;
    if output.status.success() {
        return Ok(String::from_utf8_lossy(&output.stdout).trim().to_string());
    }

    let stdout = String::from_utf8_lossy(&output.stdout);
    let stderr = String::from_utf8_lossy(&output.stderr);
    let combined = format!("{stdout}\n{stderr}");
    let tail = tail_text(&combined, 6000);
    Err(format!(
        "{label} failed with {}:\n{}",
        output.status,
        tail.trim()
    ))
}

pub(super) fn validate_runtime_probe(probe: &RuntimeProbe) -> Result<(), String> {
    if probe.paddle_version != PADDLE_VERSION || probe.paddleocr_version != PADDLEOCR_VERSION {
        return Err(format!(
            "OCR runtime version mismatch: PaddlePaddle {}, PaddleOCR {}; expected {} and {}",
            probe.paddle_version, probe.paddleocr_version, PADDLE_VERSION, PADDLEOCR_VERSION
        ));
    }
    Ok(())
}

pub(super) fn python_runtime_signature(path: &PathBuf) -> Result<(u64, u64), String> {
    let metadata = fs::metadata(path)
        .map_err(|error| format!("Unable to inspect OCR Python runtime: {error}"))?;
    let modified_ms = metadata
        .modified()
        .ok()
        .and_then(|value| value.duration_since(UNIX_EPOCH).ok())
        .map(|duration| duration.as_millis() as u64)
        .unwrap_or_default();
    Ok((metadata.len(), modified_ms))
}

pub(super) fn probe_runtime(paths: &RuntimePaths) -> Result<RuntimeProbe, String> {
    if !paths.python.exists() {
        return Err("OCR offline runtime is not installed".to_string());
    }
    let script = r#"import json, platform; import paddle; import paddleocr; import tokenizers, imagesize, ftfy, wand; from importlib.metadata import version; from paddleocr import FormulaRecognition; print(json.dumps({'python_version': platform.python_version(), 'paddle_version': paddle.__version__, 'paddleocr_version': version('paddleocr')}))"#;
    let output = command_output(
        Command::new(&paths.python).arg("-c").arg(script),
        "OCR runtime verification",
    )?;
    let probe: RuntimeProbe = serde_json::from_str(&output)
        .map_err(|error| format!("OCR runtime returned invalid version information: {error}"))?;
    validate_runtime_probe(&probe)?;
    if let Err(error) = write_persisted_runtime_probe(paths, &probe) {
        eprintln!("Unable to persist OCR runtime verification: {error}");
    }
    Ok(probe)
}
