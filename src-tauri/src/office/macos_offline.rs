use crate::office::server::metadata_from_session;
use crate::office::sessions::{
    valid_uuid, CreateOfficeSessionInput, FormulaLine, OfficeFormulaSession, OfficeHost,
    OfficeSessionMode, OfficeSessionStatus, SessionError, VisualTeXFormulaMetadata,
};
use crate::office::state::OfficeCompanionState;
use base64::{
    engine::general_purpose::{STANDARD as BASE64_STANDARD, URL_SAFE_NO_PAD},
    Engine as _,
};
use flate2::{read::DeflateDecoder, write::DeflateEncoder, Compression};
use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::fs::{self, OpenOptions};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};
use std::process::Command;
use std::sync::Mutex;
use tauri::{AppHandle, Manager, WebviewUrl, WebviewWindowBuilder};
use uuid::Uuid;

const OFFLINE_PROTOCOL_VERSION: u32 = 1;
const REQUEST_FILE: &str = "request.json";
const DISPATCH_FILE: &str = "dispatch.txt";
const RESULT_IMAGE_FILE: &str = "formula.png";
const WORD_POINTER_FILE: &str = "word-active-session.txt";
const POWERPOINT_POINTER_FILE: &str = "powerpoint-active-session.txt";
const OFFICE_GROUP_CONTAINER: &str = "Library/Group Containers/UBF8T346G9.Office";
const METADATA_PREFIX: &str = "visualtex:v1:deflate:";
const PENDING_PREFIX: &str = "visualtex:pending:v1:";
const MAX_REQUEST_BYTES: u64 = 256 * 1024;
const MAX_METADATA_BYTES: usize = 2 * 1024 * 1024;
const MAX_IDENTITY_CHARS: usize = 2048;
const MAX_SHAPE_NAME_CHARS: usize = 128;
const MAX_WORD_WIDTH_PT: f64 = 500.0;
static WORD_DISPATCH_LOCK: Mutex<()> = Mutex::new(());
static POWERPOINT_DISPATCH_LOCK: Mutex<()> = Mutex::new(());

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct MacOfflinePowerPointRequest {
    presentation_identity: String,
    slide_index: u32,
    slide_id: u32,
    shape_name: String,
    left: f64,
    top: f64,
    width: f64,
    height: f64,
    rotation: f64,
    z_order: u32,
}

#[derive(Debug, Clone, Deserialize, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct MacOfflineSessionRequest {
    protocol_version: u32,
    session_id: String,
    host: String,
    mode: String,
    formula_id: Option<String>,
    display_mode: String,
    numbered: bool,
    source_document_id: Option<String>,
    source_object_id: Option<String>,
    encoded_metadata: Option<String>,
    pending_marker: Option<String>,
    power_point: Option<MacOfflinePowerPointRequest>,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct MacOfflinePluginHealth {
    loaded: bool,
    plugin_version: Option<String>,
    host: String,
    timestamp: Option<String>,
    status_path: String,
}

fn user_home() -> Result<PathBuf, String> {
    std::env::var_os("HOME")
        .map(PathBuf::from)
        .filter(|path| path.is_absolute())
        .ok_or_else(|| "Unable to resolve the current user's home directory".to_string())
}

pub(crate) fn offline_root() -> Result<PathBuf, String> {
    Ok(user_home()?.join(OFFICE_GROUP_CONTAINER).join("VisualTeX"))
}

fn sessions_root() -> Result<PathBuf, String> {
    Ok(offline_root()?.join("OfficeSessions"))
}

fn session_directory(session_id: &str) -> Result<PathBuf, String> {
    validate_uuid(session_id, "Session id")?;
    Ok(sessions_root()?.join(session_id))
}

fn request_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(session_id)?.join(REQUEST_FILE))
}

fn dispatch_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(session_id)?.join(DISPATCH_FILE))
}

fn result_image_path(session_id: &str) -> Result<PathBuf, String> {
    Ok(session_directory(session_id)?.join(RESULT_IMAGE_FILE))
}

fn cleanup_session_files_at(directory: &Path) -> Result<(), String> {
    for name in [REQUEST_FILE, DISPATCH_FILE, RESULT_IMAGE_FILE] {
        let path = directory.join(name);
        match fs::remove_file(&path) {
            Ok(()) => {}
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {}
            Err(error) => return Err(format!("Unable to remove {}: {error}", path.display())),
        }
    }
    match fs::remove_dir(directory) {
        Ok(()) => Ok(()),
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => Ok(()),
        Err(error) if error.kind() == std::io::ErrorKind::DirectoryNotEmpty => Ok(()),
        Err(error) => Err(format!(
            "Unable to remove offline Office Session directory {}: {error}",
            directory.display()
        )),
    }
}

fn cleanup_session_files(session_id: &str) -> Result<(), String> {
    cleanup_session_files_at(&session_directory(session_id)?)
}

fn pointer_path(host: OfficeHost) -> Result<PathBuf, String> {
    Ok(sessions_root()?.join(match host {
        OfficeHost::Word => WORD_POINTER_FILE,
        OfficeHost::Powerpoint => POWERPOINT_POINTER_FILE,
    }))
}

fn validate_uuid(value: &str, label: &str) -> Result<(), String> {
    if valid_uuid(value) {
        Ok(())
    } else {
        Err(format!("{label} must be a canonical UUID"))
    }
}

fn validate_bounded_text(value: &str, maximum: usize, label: &str) -> Result<(), String> {
    if value.chars().count() > maximum || value.chars().any(char::is_control) {
        return Err(format!("{label} contains unsupported characters or is too long"));
    }
    Ok(())
}

fn validate_finite_geometry(value: f64, label: &str) -> Result<(), String> {
    if !value.is_finite() || value.abs() > 10_000_000.0 {
        Err(format!("PowerPoint {label} is invalid"))
    } else {
        Ok(())
    }
}

fn validate_request(request: &MacOfflineSessionRequest, session_id: &str) -> Result<(), String> {
    if request.protocol_version != OFFLINE_PROTOCOL_VERSION {
        return Err("Unsupported VisualTeX macOS offline protocol version".to_string());
    }
    validate_uuid(&request.session_id, "Request Session id")?;
    if request.session_id != session_id {
        return Err("Request Session id does not match the custom URL".to_string());
    }
    if !matches!(request.host.as_str(), "word" | "powerpoint") {
        return Err("Offline Office request host must be word or powerpoint".to_string());
    }
    if !matches!(request.mode.as_str(), "create" | "edit") {
        return Err("Offline Office request mode must be create or edit".to_string());
    }
    if !matches!(request.display_mode.as_str(), "inline" | "block") {
        return Err("Offline Office displayMode must be inline or block".to_string());
    }
    if request.numbered && (request.host != "word" || request.display_mode != "block") {
        return Err("Only Word display formulas can be numbered".to_string());
    }
    if let Some(formula_id) = request.formula_id.as_deref() {
        validate_uuid(formula_id, "Formula id")?;
    }
    for (value, label) in [
        (request.source_document_id.as_deref(), "sourceDocumentId"),
        (request.source_object_id.as_deref(), "sourceObjectId"),
        (request.pending_marker.as_deref(), "pendingMarker"),
    ] {
        if let Some(value) = value {
            validate_bounded_text(value, MAX_IDENTITY_CHARS, label)?;
        }
    }
    if let Some(marker) = request.pending_marker.as_deref() {
        if !marker.starts_with(PENDING_PREFIX) {
            return Err("Offline Office pending marker is invalid".to_string());
        }
    }
    if let Some(encoded) = request.encoded_metadata.as_deref() {
        if encoded.len() > MAX_METADATA_BYTES || !encoded.starts_with(METADATA_PREFIX) {
            return Err("Offline Office metadata envelope is invalid".to_string());
        }
    }

    match (request.host.as_str(), request.power_point.as_ref()) {
        ("word", None) => {}
        ("word", Some(_)) => return Err("Word request must not contain PowerPoint geometry".to_string()),
        ("powerpoint", None) => return Err("PowerPoint request requires geometry".to_string()),
        ("powerpoint", Some(powerpoint)) => {
            validate_bounded_text(
                &powerpoint.presentation_identity,
                MAX_IDENTITY_CHARS,
                "PowerPoint presentation identity",
            )?;
            validate_bounded_text(
                &powerpoint.shape_name,
                MAX_SHAPE_NAME_CHARS,
                "PowerPoint shape name",
            )?;
            if powerpoint.slide_index == 0 || powerpoint.slide_id == 0 || powerpoint.z_order == 0 {
                return Err("PowerPoint slide and z-order references must be positive".to_string());
            }
            for (value, label) in [
                (powerpoint.left, "left"),
                (powerpoint.top, "top"),
                (powerpoint.width, "width"),
                (powerpoint.height, "height"),
                (powerpoint.rotation, "rotation"),
            ] {
                validate_finite_geometry(value, label)?;
            }
            if powerpoint.width <= 0.0 || powerpoint.height <= 0.0 {
                return Err("PowerPoint formula geometry must have positive dimensions".to_string());
            }
        }
        _ => unreachable!(),
    }
    Ok(())
}

fn read_request(session_id: &str) -> Result<MacOfflineSessionRequest, String> {
    let path = request_path(session_id)?;
    let metadata = fs::metadata(&path)
        .map_err(|error| format!("Unable to read offline Office request metadata: {error}"))?;
    if !metadata.is_file() || metadata.len() == 0 || metadata.len() > MAX_REQUEST_BYTES {
        return Err("Offline Office request has an invalid size".to_string());
    }
    let bytes = fs::read(&path)
        .map_err(|error| format!("Unable to read offline Office request: {error}"))?;
    let request: MacOfflineSessionRequest = serde_json::from_slice(&bytes)
        .map_err(|error| format!("Offline Office request contains invalid JSON: {error}"))?;
    validate_request(&request, session_id)?;
    Ok(request)
}

fn decode_metadata(encoded: &str) -> Result<VisualTeXFormulaMetadata, String> {
    let payload = encoded
        .strip_prefix(METADATA_PREFIX)
        .ok_or_else(|| "VisualTeX formula metadata prefix is invalid".to_string())?;
    if payload.is_empty() || payload.len() > MAX_METADATA_BYTES {
        return Err("VisualTeX formula metadata payload is invalid".to_string());
    }
    let compressed = URL_SAFE_NO_PAD
        .decode(payload)
        .map_err(|error| format!("Unable to decode VisualTeX formula metadata: {error}"))?;
    let decoder = DeflateDecoder::new(compressed.as_slice());
    let mut json = Vec::new();
    decoder
        .take((MAX_METADATA_BYTES + 1) as u64)
        .read_to_end(&mut json)
        .map_err(|error| format!("Unable to inflate VisualTeX formula metadata: {error}"))?;
    if json.len() > MAX_METADATA_BYTES {
        return Err("VisualTeX formula metadata expands beyond the allowed size".to_string());
    }
    let metadata: VisualTeXFormulaMetadata = serde_json::from_slice(&json)
        .map_err(|error| format!("VisualTeX formula metadata JSON is invalid: {error}"))?;
    validate_metadata(&metadata)?;
    Ok(metadata)
}

fn validate_metadata(metadata: &VisualTeXFormulaMetadata) -> Result<(), String> {
    if metadata.schema != "visualtex-formula" || metadata.schema_version != 1 {
        return Err("Unsupported VisualTeX formula metadata schema".to_string());
    }
    validate_uuid(&metadata.formula_id, "Metadata formulaId")?;
    if metadata.lines.is_empty() || metadata.lines.len() > 512 {
        return Err("VisualTeX formula metadata must contain 1 to 512 lines".to_string());
    }
    for line in &metadata.lines {
        validate_uuid(&line.id, "Metadata line id")?;
        if line.latex.len() > 1_000_000 {
            return Err("A VisualTeX formula line exceeds the 1 MB limit".to_string());
        }
    }
    if !matches!(metadata.display_mode.as_str(), "inline" | "block") {
        return Err("VisualTeX metadata displayMode is invalid".to_string());
    }
    Ok(())
}

fn encode_metadata(metadata: &VisualTeXFormulaMetadata) -> Result<String, String> {
    validate_metadata(metadata)?;
    let json = serde_json::to_vec(metadata)
        .map_err(|error| format!("Unable to encode VisualTeX formula metadata: {error}"))?;
    let mut encoder = DeflateEncoder::new(Vec::new(), Compression::best());
    encoder
        .write_all(&json)
        .map_err(|error| format!("Unable to compress VisualTeX formula metadata: {error}"))?;
    let compressed = encoder
        .finish()
        .map_err(|error| format!("Unable to finish VisualTeX formula metadata: {error}"))?;
    Ok(format!("{METADATA_PREFIX}{}", URL_SAFE_NO_PAD.encode(compressed)))
}

fn hex_encode(value: &str) -> String {
    value.as_bytes().iter().map(|byte| format!("{byte:02x}")).collect()
}

fn import_request(
    state: &OfficeCompanionState,
    request: MacOfflineSessionRequest,
) -> Result<OfficeFormulaSession, String> {
    match state.session_store.get(&request.session_id) {
        Ok(existing) => return Ok(existing),
        Err(SessionError::NotFound) => {}
        Err(error) => return Err(error.to_string()),
    }

    let original_metadata = request
        .encoded_metadata
        .as_deref()
        .map(decode_metadata)
        .transpose()?;
    let metadata_formula_id = original_metadata.as_ref().map(|value| value.formula_id.clone());
    let formula_id = match (request.formula_id.clone(), metadata_formula_id) {
        (Some(request_id), Some(metadata_id)) if request_id != metadata_id => {
            return Err("Request formulaId does not match encoded metadata".to_string())
        }
        (Some(request_id), _) => request_id,
        (None, Some(metadata_id)) => metadata_id,
        (None, None) if request.mode == "create" => Uuid::new_v4().to_string(),
        (None, None) => return Err("Edit request does not contain a formulaId".to_string()),
    };
    validate_uuid(&formula_id, "Imported formula id")?;

    let host = match request.host.as_str() {
        "word" => OfficeHost::Word,
        "powerpoint" => OfficeHost::Powerpoint,
        _ => unreachable!(),
    };
    let mode = match request.mode.as_str() {
        "create" => OfficeSessionMode::Create,
        "edit" => OfficeSessionMode::Edit,
        _ => unreachable!(),
    };
    let lines = original_metadata
        .as_ref()
        .map(|metadata| {
            metadata
                .lines
                .iter()
                .map(|line| FormulaLine {
                    id: line.id.clone(),
                    latex: line.latex.clone(),
                })
                .collect::<Vec<_>>()
        })
        .unwrap_or_else(|| {
            vec![FormulaLine {
                id: Uuid::new_v4().to_string(),
                latex: String::new(),
            }]
        });
    let source_document_id = match host {
        OfficeHost::Word => request.source_document_id.clone(),
        OfficeHost::Powerpoint => request.power_point.as_ref().map(|powerpoint| {
            format!(
                "visualtex-ppt-native-presentation:{}",
                powerpoint.presentation_identity
            )
        }),
    };
    let source_object_id = match host {
        OfficeHost::Word => match mode {
            OfficeSessionMode::Create => request.pending_marker.clone(),
            OfficeSessionMode::Edit => request.encoded_metadata.clone(),
        },
        OfficeHost::Powerpoint => request.power_point.as_ref().map(|powerpoint| {
            format!(
                "visualtex-ppt-native-edit:{}:{}",
                powerpoint.slide_index,
                hex_encode(&powerpoint.shape_name)
            )
        }),
    };
    let title = original_metadata
        .as_ref()
        .map(|metadata| metadata.title.clone())
        .unwrap_or_else(|| match host {
            OfficeHost::Word => "Word Formula".to_string(),
            OfficeHost::Powerpoint => "PowerPoint Formula".to_string(),
        });
    let code_format = original_metadata
        .as_ref()
        .map(|metadata| metadata.code_format.clone())
        .unwrap_or_else(|| "latex".to_string());

    let session_id = request.session_id.clone();
    match state
        .session_store
        .create_external(
            session_id.clone(),
            CreateOfficeSessionInput {
                mode,
                host,
                formula_id: Some(formula_id),
                source_document_id,
                source_object_id,
                title: Some(title),
                lines: Some(lines),
                active_line_id: None,
                code_format: Some(code_format),
                display_mode: Some(request.display_mode),
                numbered: Some(request.numbered),
                export_width: None,
                export_height: None,
                original_metadata,
                auto_commit_on_close: Some(true),
            },
        )
    {
        Ok(session) => Ok(session),
        Err(SessionError::Conflict(_)) => state
            .session_store
            .get(&session_id)
            .map_err(|error| error.to_string()),
        Err(error) => Err(error.to_string()),
    }
}

pub(crate) fn parse_office_url(value: &str) -> Result<String, String> {
    const PREFIX: &str = "visualtex://office/open?session=";
    let session_id = value
        .strip_prefix(PREFIX)
        .ok_or_else(|| "VisualTeX URL must use visualtex://office/open".to_string())?;
    if session_id.contains(['&', '#', '?', '/', '%']) {
        return Err("VisualTeX URL contains unsupported query data".to_string());
    }
    validate_uuid(session_id, "VisualTeX URL Session id")?;
    Ok(session_id.to_string())
}

fn open_editor_window(app: &AppHandle, session_id: &str) -> Result<(), String> {
    let label = format!("office-native-{}", session_id.replace('-', ""));
    if let Some(window) = app.get_webview_window(&label) {
        window.show().map_err(|error| error.to_string())?;
        window.set_focus().map_err(|error| error.to_string())?;
        return Ok(());
    }
    let path = format!(
        "office-native-dialog.html?sessionId={session_id}&transport=tauri"
    );
    let window = WebviewWindowBuilder::new(app, label, WebviewUrl::App(path.into()))
        .title("VisualTeX Office Formula")
        .inner_size(1180.0, 820.0)
        .min_inner_size(720.0, 560.0)
        .center()
        .build()
        .map_err(|error| format!("Unable to open the VisualTeX Office editor: {error}"))?;
    window.show().map_err(|error| error.to_string())?;
    window.set_focus().map_err(|error| error.to_string())?;
    Ok(())
}

pub(crate) fn focus_open_office_editor(app: &AppHandle) -> bool {
    for (label, window) in app.webview_windows() {
        if label.starts_with("office-native-") {
            let _ = window.show();
            let _ = window.set_focus();
            return true;
        }
    }
    false
}

pub(crate) fn handle_open_url(app: &AppHandle, value: &str) -> Result<(), String> {
    let session_id = parse_office_url(value)?;
    let state = app
        .try_state::<OfficeCompanionState>()
        .ok_or_else(|| "VisualTeX Office state is not initialized".to_string())?;
    let request = read_request(&session_id)?;
    import_request(state.inner(), request)?;

    // Office formula requests must open only the dedicated formula editor.
    // Keeping the main VisualTeX workspace visible makes Word/PowerPoint
    // insertion look like a jump into the full application instead of a
    // focused Office editing transaction.
    crate::office::background::hide_main_window(app)?;
    open_editor_window(app, &session_id)
}

fn decode_png(value: &str) -> Result<Vec<u8>, String> {
    let payload = value
        .split_once(',')
        .filter(|(prefix, _)| prefix.starts_with("data:image/png;base64"))
        .map(|(_, payload)| payload)
        .unwrap_or(value);
    let bytes = BASE64_STANDARD
        .decode(payload.trim())
        .map_err(|error| format!("Unable to decode the Office PNG export: {error}"))?;
    if bytes.len() < 8 || &bytes[..8] != b"\x89PNG\r\n\x1a\n" {
        return Err("Office formula export is not a valid PNG image".to_string());
    }
    Ok(bytes)
}

#[cfg(unix)]
fn set_mode(path: &Path, mode: u32) -> Result<(), String> {
    use std::os::unix::fs::PermissionsExt;
    fs::set_permissions(path, fs::Permissions::from_mode(mode))
        .map_err(|error| format!("Unable to set permissions on {}: {error}", path.display()))
}

#[cfg(not(unix))]
fn set_mode(_path: &Path, _mode: u32) -> Result<(), String> {
    Ok(())
}

fn atomic_write(path: &Path, bytes: &[u8], mode: u32) -> Result<(), String> {
    let parent = path
        .parent()
        .ok_or_else(|| format!("Path has no parent: {}", path.display()))?;
    fs::create_dir_all(parent)
        .map_err(|error| format!("Unable to create {}: {error}", parent.display()))?;
    set_mode(parent, 0o700)?;
    let temporary = parent.join(format!(
        ".{}.{}.tmp",
        path.file_name().and_then(|value| value.to_str()).unwrap_or("visualtex"),
        Uuid::new_v4()
    ));
    let mut file = OpenOptions::new()
        .create_new(true)
        .write(true)
        .open(&temporary)
        .map_err(|error| format!("Unable to create {}: {error}", temporary.display()))?;
    file.write_all(bytes)
        .map_err(|error| format!("Unable to write {}: {error}", temporary.display()))?;
    file.sync_all()
        .map_err(|error| format!("Unable to sync {}: {error}", temporary.display()))?;
    set_mode(&temporary, mode)?;
    fs::rename(&temporary, path).map_err(|error| {
        let _ = fs::remove_file(&temporary);
        format!("Unable to replace {}: {error}", path.display())
    })?;
    set_mode(path, mode)
}

fn sanitize_dispatch_value(value: &str, label: &str) -> Result<String, String> {
    if value.contains(['\r', '\n', '\0']) {
        return Err(format!("{label} contains unsupported control characters"));
    }
    Ok(value.to_string())
}

fn dispatch_text(entries: &[(&str, String)]) -> Result<String, String> {
    let mut seen = std::collections::HashSet::new();
    let mut output = String::new();
    for (key, value) in entries {
        if !seen.insert(*key)
            || key.is_empty()
            || !key.bytes().all(|byte| byte.is_ascii_alphanumeric())
        {
            return Err("VisualTeX dispatch contains an invalid key".to_string());
        }
        output.push_str(key);
        output.push('=');
        output.push_str(&sanitize_dispatch_value(value, key)?);
        output.push('\n');
    }
    Ok(output)
}

fn run_vba_callback(host: OfficeHost) -> Result<(), String> {
    let script = match host {
        OfficeHost::Word => r#"tell application "Microsoft Word"
if not (exists active document) then error "Microsoft Word has no active document"
run VB macro macro name "VisualTeX_ApplyPendingResult"
end tell"#,
        OfficeHost::Powerpoint => r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "Microsoft PowerPoint has no active presentation"
run VB macro macro name "VisualTeX_ApplyPendingResult" list of parameters {}
end tell"#,
    };
    let output = Command::new("/usr/bin/osascript")
        .arg("-e")
        .arg(script)
        .output()
        .map_err(|error| format!("Unable to launch the Office VBA callback: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
        Err(if detail.is_empty() {
            "The Office VBA callback failed".to_string()
        } else {
            format!("The Office VBA callback failed: {detail}")
        })
    }
}

fn with_dispatch_pointer<T>(
    host: OfficeHost,
    session_id: &str,
    operation: impl FnOnce() -> Result<T, String>,
) -> Result<T, String> {
    let lock = match host {
        OfficeHost::Word => &WORD_DISPATCH_LOCK,
        OfficeHost::Powerpoint => &POWERPOINT_DISPATCH_LOCK,
    };
    let _guard = lock
        .lock()
        .map_err(|_| "VisualTeX Office dispatch lock is unavailable".to_string())?;
    let pointer = pointer_path(host)?;
    atomic_write(&pointer, session_id.as_bytes(), 0o600)?;
    let result = operation();
    let _ = fs::remove_file(pointer);
    result
}

fn calculate_word_geometry(session: &OfficeFormulaSession) -> Result<(f64, f64, i32), String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Word Session has no formula export".to_string())?;
    if !export.width.is_finite()
        || !export.height.is_finite()
        || export.width <= 0.0
        || export.height <= 0.0
    {
        return Err("Word formula export has invalid dimensions".to_string());
    }
    let natural_width = export.width * 0.75;
    let natural_height = export.height * 0.75;
    let scale = f64::min(1.0, MAX_WORD_WIDTH_PT / natural_width);
    let width = natural_width * scale;
    let height = natural_height * scale;
    let baseline = export
        .baseline
        .filter(|value| value.is_finite() && *value >= 0.0 && *value <= export.height)
        .map(|value| {
            let descent_ratio = (export.height - value) / export.height;
            -(height * descent_ratio).round().max(0.0) as i32
        })
        .unwrap_or(0)
        .clamp(-256, 0);
    Ok((width, height, baseline))
}

fn calculate_powerpoint_geometry(
    request: &MacOfflinePowerPointRequest,
    session: &OfficeFormulaSession,
) -> Result<(f64, f64, f64, f64), String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "PowerPoint Session has no formula export".to_string())?;
    if !export.width.is_finite()
        || !export.height.is_finite()
        || export.width <= 0.0
        || export.height <= 0.0
    {
        return Err("PowerPoint formula export has invalid dimensions".to_string());
    }
    let height_ratio = session
        .original_metadata
        .as_ref()
        .and_then(|metadata| metadata.render_height_px)
        .filter(|value| value.is_finite() && *value > 0.0)
        .map(|previous| export.height / previous)
        .unwrap_or(1.0);
    let target_height = request.height * height_ratio;
    let target_width = target_height * (export.width / export.height);
    let center_x = request.left + request.width / 2.0;
    let center_y = request.top + request.height / 2.0;
    let left = center_x - target_width / 2.0;
    let top = center_y - target_height / 2.0;
    for (value, label) in [
        (left, "target left"),
        (top, "target top"),
        (target_width, "target width"),
        (target_height, "target height"),
    ] {
        validate_finite_geometry(value, label)?;
    }
    if target_width <= 0.0 || target_height <= 0.0 {
        return Err("PowerPoint target formula dimensions are invalid".to_string());
    }
    Ok((left, top, target_width, target_height))
}

fn materialize_result_image(session: &OfficeFormulaSession) -> Result<PathBuf, String> {
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Office Session has no formula export".to_string())?;
    let png = export
        .png_base64
        .as_deref()
        .ok_or_else(|| "Offline Office Session requires a PNG export".to_string())
        .and_then(decode_png)?;
    let path = result_image_path(&session.id)?;
    atomic_write(&path, &png, 0o600)?;
    Ok(path)
}

fn commit_word(
    state: &OfficeCompanionState,
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
    metadata: &str,
) -> Result<(), String> {
    let image_path = materialize_result_image(session)?;
    let (width, height, baseline) = calculate_word_geometry(session)?;
    let source_marker = request.encoded_metadata.clone().unwrap_or_default();
    let pending_marker = request.pending_marker.clone().unwrap_or_default();
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session.id.clone()),
        ("action", "commit".to_string()),
        ("host", "word".to_string()),
        ("mode", request.mode.clone()),
        ("formulaId", session.formula_id.clone()),
        ("displayMode", session.display_mode.clone()),
        ("numbered", if session.numbered { "1" } else { "0" }.to_string()),
        ("imagePath", image_path.to_string_lossy().to_string()),
        ("metadata", metadata.to_string()),
        ("pendingMarker", pending_marker),
        ("sourceMarker", source_marker),
        (
            "sourceDocumentId",
            request.source_document_id.clone().unwrap_or_default(),
        ),
        ("widthPoints", format!("{width:.6}")),
        ("heightPoints", format!("{height:.6}")),
        ("baseline", baseline.to_string()),
    ])?;
    atomic_write(&dispatch_path(&session.id)?, dispatch.as_bytes(), 0o600)?;
    with_dispatch_pointer(OfficeHost::Word, &session.id, || {
        run_vba_callback(OfficeHost::Word)
    })?;
    state
        .formula_cache
        .put(&session.formula_id, metadata_from_session(session))
        .map_err(|error| format!("Formula metadata could not be saved: {error}"))?;
    Ok(())
}

fn commit_powerpoint(
    state: &OfficeCompanionState,
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
    metadata: &str,
) -> Result<(), String> {
    let powerpoint = request
        .power_point
        .as_ref()
        .ok_or_else(|| "PowerPoint request geometry is missing".to_string())?;
    let image_path = materialize_result_image(session)?;
    let (left, top, width, height) = calculate_powerpoint_geometry(powerpoint, session)?;
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session.id.clone()),
        ("action", "commit".to_string()),
        ("host", "powerpoint".to_string()),
        ("mode", request.mode.clone()),
        ("formulaId", session.formula_id.clone()),
        ("displayMode", "block".to_string()),
        ("numbered", "0".to_string()),
        ("imagePath", image_path.to_string_lossy().to_string()),
        ("metadata", metadata.to_string()),
        (
            "pendingMarker",
            request.pending_marker.clone().unwrap_or_default(),
        ),
        ("sourceMarker", request.encoded_metadata.clone().unwrap_or_default()),
        ("sourceShapeName", powerpoint.shape_name.clone()),
        ("shapeName", format!("VisualTeX_{}", session.formula_id)),
        ("targetLeft", format!("{left:.6}")),
        ("targetTop", format!("{top:.6}")),
        ("targetWidth", format!("{width:.6}")),
        ("targetHeight", format!("{height:.6}")),
        ("rotation", format!("{:.6}", powerpoint.rotation)),
        ("zOrder", powerpoint.z_order.to_string()),
        ("presentationIdentity", powerpoint.presentation_identity.clone()),
        ("slideIndex", powerpoint.slide_index.to_string()),
        ("slideId", powerpoint.slide_id.to_string()),
    ])?;
    atomic_write(&dispatch_path(&session.id)?, dispatch.as_bytes(), 0o600)?;
    with_dispatch_pointer(OfficeHost::Powerpoint, &session.id, || {
        run_vba_callback(OfficeHost::Powerpoint)
    })?;
    state
        .formula_cache
        .put(&session.formula_id, metadata_from_session(session))
        .map_err(|error| format!("Formula metadata could not be saved: {error}"))?;
    Ok(())
}

fn cancel_host(request: &MacOfflineSessionRequest) -> Result<(), String> {
    let host = if request.host == "word" {
        OfficeHost::Word
    } else {
        OfficeHost::Powerpoint
    };
    let entries = vec![
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", request.session_id.clone()),
        ("action", "cancel".to_string()),
        ("host", request.host.clone()),
        ("mode", request.mode.clone()),
        (
            "pendingMarker",
            request.pending_marker.clone().unwrap_or_default(),
        ),
        (
            "sourceDocumentId",
            request.source_document_id.clone().unwrap_or_default(),
        ),
    ];
    let dispatch = dispatch_text(&entries)?;
    atomic_write(
        &dispatch_path(&request.session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    if request.mode == "create" {
        with_dispatch_pointer(host, &request.session_id, || run_vba_callback(host))?;
    }
    Ok(())
}

fn complete_session(
    state: &OfficeCompanionState,
    session_id: &str,
) -> Result<OfficeFormulaSession, String> {
    state
        .session_store
        .patch(
            session_id,
            json!({ "status": "completed", "error": null }),
        )
        .map_err(|error| error.to_string())
}

fn fail_session(state: &OfficeCompanionState, session_id: &str, error: &str) {
    let _ = state.session_store.patch(
        session_id,
        json!({ "status": "failed", "error": error }),
    );
}

fn commit_session_blocking(
    state: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    let session = state
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())?;
    if session.status == OfficeSessionStatus::Completed {
        let _ = cleanup_session_files(&session_id);
        return Ok(session);
    }
    if session.status != OfficeSessionStatus::Committing {
        return Err("Offline Office Session is not ready to commit".to_string());
    }
    if session.mode == OfficeSessionMode::Edit && !session.dirty {
        let completed = complete_session(&state, &session_id)?;
        let _ = cleanup_session_files(&session_id);
        return Ok(completed);
    }
    let request = read_request(&session_id)?;
    let metadata = metadata_from_session(&session);
    let encoded = encode_metadata(&metadata)?;
    let result = match session.host {
        OfficeHost::Word => commit_word(&state, &request, &session, &encoded),
        OfficeHost::Powerpoint => commit_powerpoint(&state, &request, &session, &encoded),
    };
    if let Err(error) = result {
        fail_session(&state, &session_id, &error);
        return Err(error);
    }
    let completed = complete_session(&state, &session_id)?;
    let _ = cleanup_session_files(&session_id);
    Ok(completed)
}

fn cancel_session_blocking(
    state: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    if let Err(error) = cancel_host(&request) {
        fail_session(&state, &session_id, &error);
        return Err(error);
    }
    let cancelled = state
        .session_store
        .patch(
            &session_id,
            json!({
                "status": "cancelled",
                "explicitCancel": true,
                "error": null
            }),
        )
        .map_err(|error| error.to_string())?;
    let _ = cleanup_session_files(&session_id);
    Ok(cancelled)
}

#[tauri::command]
pub fn get_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    state
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn update_macos_offline_office_session(
    session_id: String,
    patch: Value,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    validate_uuid(&session_id, "Session id")?;
    state
        .session_store
        .patch(&session_id, patch)
        .map_err(|error| error.to_string())
}

#[tauri::command]
pub fn delete_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<(), String> {
    validate_uuid(&session_id, "Session id")?;
    state
        .session_store
        .delete(&session_id)
        .map_err(|error| error.to_string())?;
    cleanup_session_files(&session_id)
}

#[tauri::command]
pub async fn commit_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || commit_session_blocking(state, session_id))
        .await
        .map_err(|error| format!("Offline Office commit task failed: {error}"))?
}

#[tauri::command]
pub async fn cancel_macos_offline_office_session(
    session_id: String,
    state: tauri::State<'_, OfficeCompanionState>,
) -> Result<OfficeFormulaSession, String> {
    let state = state.inner().clone();
    tokio::task::spawn_blocking(move || cancel_session_blocking(state, session_id))
        .await
        .map_err(|error| format!("Offline Office cancel task failed: {error}"))?
}

#[cfg(target_os = "macos")]
fn refresh_health_signal(host: &str) {
    let (process_name, script) = match host {
        "word" => (
            "Microsoft Word",
            r#"tell application "Microsoft Word" to run VB macro macro name "AutoExec""#,
        ),
        "powerpoint" => (
            "Microsoft PowerPoint",
            r#"tell application "Microsoft PowerPoint" to run VB macro macro name "Auto_Open" list of parameters {}"#,
        ),
        _ => return,
    };
    let running = Command::new("/usr/bin/pgrep")
        .args(["-x", process_name])
        .output()
        .is_ok_and(|output| output.status.success());
    if !running {
        return;
    }
    let _ = Command::new("/usr/bin/osascript")
        .arg("-e")
        .arg(script)
        .output();
}

#[cfg(not(target_os = "macos"))]
fn refresh_health_signal(_host: &str) {}

fn read_health(host: &str) -> Result<MacOfflinePluginHealth, String> {
    let path = offline_root()?
        .join("OfficePluginStatus")
        .join(format!("{host}.json"));
    let fallback = || MacOfflinePluginHealth {
        loaded: false,
        plugin_version: None,
        host: host.to_string(),
        timestamp: None,
        status_path: path.display().to_string(),
    };
    let bytes = match fs::read(&path) {
        Ok(bytes) => bytes,
        Err(error) if error.kind() == std::io::ErrorKind::NotFound => return Ok(fallback()),
        Err(error) => return Err(format!("Unable to read {} health: {error}", host)),
    };
    let value: Value = serde_json::from_slice(&bytes)
        .map_err(|error| format!("{host} health file contains invalid JSON: {error}"))?;
    Ok(MacOfflinePluginHealth {
        loaded: value.get("loaded").and_then(Value::as_bool).unwrap_or(false),
        plugin_version: value
            .get("pluginVersion")
            .and_then(Value::as_str)
            .map(str::to_string),
        host: value
            .get("host")
            .and_then(Value::as_str)
            .unwrap_or(host)
            .to_string(),
        timestamp: value
            .get("timestamp")
            .and_then(Value::as_str)
            .map(str::to_string),
        status_path: path.display().to_string(),
    })
}

#[tauri::command]
pub fn get_macos_offline_plugin_health() -> Result<Vec<MacOfflinePluginHealth>, String> {
    refresh_health_signal("word");
    refresh_health_signal("powerpoint");
    Ok(vec![read_health("word")?, read_health("powerpoint")?])
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn offline_root_uses_the_office_application_group() {
        let root = offline_root().expect("offline Office root should resolve");
        assert!(root.ends_with("Library/Group Containers/UBF8T346G9.Office/VisualTeX"));
        assert!(!root.to_string_lossy().contains("Library/Application Support/VisualTeX"));
    }

    #[test]
    fn completed_session_cleanup_removes_only_known_ephemeral_files() {
        let directory = std::env::temp_dir().join(format!(
            "visualtex-offline-cleanup-test-{}",
            Uuid::new_v4()
        ));
        fs::create_dir_all(&directory).expect("test Session directory should be created");
        for name in [REQUEST_FILE, DISPATCH_FILE, RESULT_IMAGE_FILE] {
            fs::write(directory.join(name), b"temporary").expect("temporary file should exist");
        }
        fs::write(directory.join("keep.txt"), b"keep").expect("unknown file should exist");

        cleanup_session_files_at(&directory).expect("known files should be cleaned");
        for name in [REQUEST_FILE, DISPATCH_FILE, RESULT_IMAGE_FILE] {
            assert!(!directory.join(name).exists());
        }
        assert!(directory.join("keep.txt").is_file());
        assert!(directory.is_dir());

        fs::remove_file(directory.join("keep.txt")).unwrap();
        cleanup_session_files_at(&directory).expect("empty Session directory should be removed");
        assert!(!directory.exists());
    }

    #[test]
    fn office_url_accepts_only_the_exact_canonical_form() {
        let id = "12345678-1234-4234-9234-123456789abc";
        assert_eq!(
            parse_office_url(&format!("visualtex://office/open?session={id}")),
            Ok(id.to_string())
        );
        assert!(parse_office_url(&format!("https://office/open?session={id}")).is_err());
        assert!(parse_office_url(&format!("visualtex://office/open?session={id}&x=1")).is_err());
        assert!(parse_office_url("visualtex://office/open?session=not-a-uuid").is_err());
    }

    #[test]
    fn offline_request_json_accepts_utf8_office_identities() {
        let session_id = "32345678-1234-4234-9234-123456789abc".to_string();
        let request = MacOfflineSessionRequest {
            protocol_version: OFFLINE_PROTOCOL_VERSION,
            session_id: session_id.clone(),
            host: "word".to_string(),
            mode: "create".to_string(),
            formula_id: Some("12345678-1234-4234-9234-123456789abc".to_string()),
            display_mode: "inline".to_string(),
            numbered: false,
            source_document_id: Some("/Users/测试/公式😀.docx".to_string()),
            source_object_id: Some("书签-公式".to_string()),
            encoded_metadata: None,
            pending_marker: Some(
                "visualtex:pending:v1:32345678-1234-4234-9234-123456789abc:12345678-1234-4234-9234-123456789abc"
                    .to_string(),
            ),
            power_point: None,
        };
        let json = serde_json::to_vec(&request).expect("UTF-8 request should encode");
        let decoded: MacOfflineSessionRequest =
            serde_json::from_slice(&json).expect("UTF-8 request should decode");
        validate_request(&decoded, &session_id).expect("UTF-8 request should validate");
        assert_eq!(
            decoded.source_document_id.as_deref(),
            Some("/Users/测试/公式😀.docx")
        );
    }

    #[test]
    fn metadata_codec_round_trips_the_shared_schema() {
        let metadata = VisualTeXFormulaMetadata {
            schema: "visualtex-formula".to_string(),
            schema_version: 1,
            formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
            title: "Formula".to_string(),
            latex: "x^2".to_string(),
            lines: vec![crate::office::sessions::MetadataLine {
                id: "22345678-1234-4234-9234-123456789abc".to_string(),
                latex: "x^2".to_string(),
            }],
            code_format: "latex".to_string(),
            display_mode: "inline".to_string(),
            numbered: false,
            render_width_px: Some(50.0),
            render_height_px: Some(20.0),
            created_with_version: "1.1.0".to_string(),
            updated_with_version: "1.1.0".to_string(),
            created_at: "unix-ms:1".to_string(),
            updated_at: "unix-ms:1".to_string(),
        };
        let encoded = encode_metadata(&metadata).expect("metadata should encode");
        let decoded = decode_metadata(&encoded).expect("metadata should decode");
        assert_eq!(decoded.formula_id, metadata.formula_id);
        assert_eq!(decoded.lines[0].latex, "x^2");
    }

    #[test]
    fn dispatch_rejects_newlines_and_duplicate_keys() {
        assert!(dispatch_text(&[("sessionId", "a\nb".to_string())]).is_err());
        assert!(dispatch_text(&[
            ("sessionId", "a".to_string()),
            ("sessionId", "b".to_string())
        ])
        .is_err());
    }

    #[test]
    fn powerpoint_geometry_preserves_center_and_visual_height_ratio() {
        let request = MacOfflinePowerPointRequest {
            presentation_identity: "Deck".to_string(),
            slide_index: 1,
            slide_id: 2,
            shape_name: "VisualTeX_12345678-1234-4234-9234-123456789abc".to_string(),
            left: 100.0,
            top: 200.0,
            width: 120.0,
            height: 40.0,
            rotation: 0.0,
            z_order: 2,
        };
        let session = OfficeFormulaSession {
            id: "32345678-1234-4234-9234-123456789abc".to_string(),
            mode: OfficeSessionMode::Edit,
            host: OfficeHost::Powerpoint,
            formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
            source_document_id: None,
            source_object_id: None,
            title: "Formula".to_string(),
            lines: vec![],
            active_line_id: None,
            code_format: "latex".to_string(),
            display_mode: "block".to_string(),
            numbered: false,
            export_width: 0.0,
            export_height: 0.0,
            export_result: Some(crate::office::sessions::OfficeExportResult {
                svg: "<svg/>".to_string(),
                svg_base64: String::new(),
                png_base64: None,
                width: 300.0,
                height: 50.0,
                baseline: None,
            }),
            original_metadata: Some(VisualTeXFormulaMetadata {
                schema: "visualtex-formula".to_string(),
                schema_version: 1,
                formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
                title: "Formula".to_string(),
                latex: String::new(),
                lines: vec![],
                code_format: "latex".to_string(),
                display_mode: "block".to_string(),
                numbered: false,
                render_width_px: Some(120.0),
                render_height_px: Some(40.0),
                created_with_version: "1".to_string(),
                updated_with_version: "1".to_string(),
                created_at: "1".to_string(),
                updated_at: "1".to_string(),
            }),
            dirty: true,
            status: OfficeSessionStatus::Committing,
            auto_commit_on_close: true,
            explicit_cancel: false,
            error: None,
            created_at: 1,
            updated_at: 1,
            expires_at: 2,
        };
        let (left, top, width, height) =
            calculate_powerpoint_geometry(&request, &session).expect("geometry should scale");
        assert!((height - 50.0).abs() < 0.001);
        assert!((width - 300.0).abs() < 0.001);
        assert!((left + width / 2.0 - 160.0).abs() < 0.001);
        assert!((top + height / 2.0 - 220.0).abs() < 0.001);
    }
}
