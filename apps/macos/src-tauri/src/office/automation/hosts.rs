//! Word and PowerPoint for Mac as agent hosts (docs/office-mcp/MACOS.md). There is no add-in
//! channel to poll: an operation is one call of the add-in's `VisualTeX_AgentRun` macro through
//! AppleScript (`run VB macro`), with the request and the answer as files in the host's VisualTeX
//! runtime folder (the add-in reads that folder directly and writes it through its AppleScriptTask
//! file bridge). WPS Writer and WPS Presentation cannot be scripted from outside: their add-ins ask
//! VisualTeX for the operations over the WPS connection (office::wps::bridge) and answer the same way.
use crate::office::macos_offline;
use crate::office::sessions::OfficeHost;
use base64::{engine::general_purpose::URL_SAFE_NO_PAD as BASE64_URL, Engine as _};
use serde::{Deserialize, Serialize};
use serde_json::{Map, Value};
use std::fs;
use std::path::PathBuf;
use std::sync::{Mutex, OnceLock};
use std::time::{Duration, Instant};
use tauri::AppHandle;

pub const REQUEST_FILE: &str = "agent-request.txt";
pub const RESPONSE_FILE: &str = "agent-response.txt";

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct HostInfo {
    pub host_id: String,
    pub app: String,
    pub vendor: String,
    pub product_version: String,
    pub addin_version: String,
    pub pid: u32,
    pub capabilities: Vec<String>,
}

#[derive(Debug, Clone, Deserialize)]
pub struct CommandError {
    pub code: String,
    pub message: String,
    #[serde(default)]
    pub details: Option<Value>,
}

impl CommandError {
    pub fn new(code: &str, message: impl Into<String>) -> Self {
        Self { code: code.into(), message: message.into(), details: None }
    }
}

/// The add-in's answer: `{ok: true, data}` or `{ok: false, error: {code, message}}`.
#[derive(Debug, Deserialize)]
struct AgentResponse {
    ok: bool,
    #[serde(default)]
    id: Option<String>,
    #[serde(default)]
    data: Option<Value>,
    #[serde(default)]
    error: Option<CommandError>,
}

struct HostKind {
    host: OfficeHost,
    id: &'static str,
    app: &'static str,
    process: &'static str,
    bundle: &'static str,
    /// The WPS add-in ("writer", "presentation") that runs the operations, for WPS.
    wps: Option<&'static str>,
}

const WPS_BUNDLE: &str = "/Applications/WPS Office.app";

const HOSTS: [HostKind; 4] = [
    HostKind { host: OfficeHost::Word, id: "w", app: "word", process: "Microsoft Word", bundle: "/Applications/Microsoft Word.app", wps: None },
    HostKind { host: OfficeHost::Powerpoint, id: "p", app: "powerpoint", process: "Microsoft PowerPoint", bundle: "/Applications/Microsoft PowerPoint.app", wps: None },
    HostKind { host: OfficeHost::Word, id: "ww", app: "word", process: "wpsoffice", bundle: WPS_BUNDLE, wps: Some("writer") },
    HostKind { host: OfficeHost::Powerpoint, id: "wp", app: "powerpoint", process: "wpsoffice", bundle: WPS_BUNDLE, wps: Some("presentation") },
];

static APP: OnceLock<AppHandle> = OnceLock::new();
static WORD_LOCK: Mutex<()> = Mutex::new(());
static POWERPOINT_LOCK: Mutex<()> = Mutex::new(());

/// The app the agents' formulas are drawn in.
pub fn set_app(app: AppHandle) {
    let _ = APP.set(app);
}

pub fn app() -> Option<AppHandle> {
    APP.get().cloned()
}

fn pid_of(process: &str) -> Option<u32> {
    let output = std::process::Command::new("/usr/bin/pgrep").args(["-x", process]).output().ok()?;
    String::from_utf8_lossy(&output.stdout).lines().next()?.trim().parse().ok()
}

fn bundle_version(bundle: &str) -> String {
    let plist = PathBuf::from(bundle).join("Contents/Info.plist");
    let Ok(text) = fs::read_to_string(plist) else { return String::new() };
    let key = "<key>CFBundleShortVersionString</key>";
    text.find(key)
        .and_then(|at| {
            let rest = &text[at + key.len()..];
            let start = rest.find("<string>")? + "<string>".len();
            let end = rest[start..].find("</string>")?;
            Some(rest[start..start + end].trim().to_string())
        })
        .unwrap_or_default()
}

/// Word and PowerPoint while they run (the add-in is checked by the first call); WPS Writer and WPS
/// Presentation while their add-in asks for operations.
pub fn hosts() -> Vec<HostInfo> {
    let wps = crate::office::wps::bridge::agent_hosts();
    HOSTS
        .iter()
        .filter_map(|kind| {
            if let Some(addin) = kind.wps {
                wps.iter().find(|(name, _)| name == addin)?;
            }
            let pid = pid_of(kind.process)?;
            Some(HostInfo {
                host_id: kind.id.to_string(),
                app: kind.app.to_string(),
                vendor: if kind.wps.is_some() { "kingsoft" } else { "microsoft" }.to_string(),
                product_version: bundle_version(kind.bundle),
                addin_version: crate::office::state::OFFICE_UI_VERSION.to_string(),
                pid,
                capabilities: vec!["read".into(), "write".into()],
            })
        })
        .collect()
}

pub fn host(host_id: &str) -> Option<HostInfo> {
    hosts().into_iter().find(|host| host.host_id == host_id)
}

fn kind(host_id: &str) -> Option<&'static HostKind> {
    HOSTS.iter().find(|kind| kind.id == host_id)
}

/// `key=value` lines the add-in's dispatch reader understands: strings `s.<name>` (base64url
/// UTF-8), numbers `n.<name>`, booleans `b.<name>` (1/0), nested objects `<kind>.<name>.<field>`,
/// arrays `<kind>.<name>.<index>` with `c.<name>` the count.
fn flatten(prefix: &str, value: &Value, lines: &mut Vec<String>) {
    match value {
        Value::Null => {}
        Value::Bool(flag) => lines.push(format!("b.{prefix}={}", u8::from(*flag))),
        Value::Number(number) => lines.push(format!("n.{prefix}={number}")),
        Value::String(text) => lines.push(format!("s.{prefix}={}", BASE64_URL.encode(text.as_bytes()))),
        Value::Array(items) => {
            lines.push(format!("c.{prefix}={}", items.len()));
            for (index, item) in items.iter().enumerate() {
                flatten(&format!("{prefix}.{index}"), item, lines);
            }
        }
        Value::Object(map) => {
            for (key, item) in map {
                flatten(&format!("{prefix}.{key}"), item, lines);
            }
        }
    }
}

pub fn request_text(id: &str, operation: &str, arguments: &Map<String, Value>) -> String {
    let mut lines = vec![format!("id={id}"), format!("op={operation}")];
    for (key, value) in arguments {
        flatten(key, value, &mut lines);
    }
    lines.join("\n") + "\n"
}

fn script(kind: &HostKind, seconds: u64) -> String {
    match kind.host {
        OfficeHost::Word => format!(
            "with timeout of {seconds} seconds\ntell application \"Microsoft Word\"\nrun VB macro macro name \"VisualTeX_AgentRun\"\nend tell\nend timeout"
        ),
        OfficeHost::Powerpoint => format!(
            "with timeout of {seconds} seconds\ntell application \"Microsoft PowerPoint\"\nrun VB macro macro name \"VisualTeX_AgentRun\" list of parameters {{}}\nend tell\nend timeout"
        ),
    }
}

fn run_blocking(kind: &'static HostKind, operation: &str, arguments: &Map<String, Value>, deadline: Duration)
    -> Result<Value, CommandError> {
    let lock = match kind.host {
        OfficeHost::Word => &WORD_LOCK,
        OfficeHost::Powerpoint => &POWERPOINT_LOCK,
    };
    let started = Instant::now();
    let _guard = loop {
        match lock.try_lock() {
            Ok(guard) => break guard,
            Err(_) if started.elapsed() < deadline => std::thread::sleep(Duration::from_millis(50)),
            Err(_) => return Err(CommandError::new("host_busy", "Another agent operation is still running in this application")),
        }
    };
    let root = macos_offline::runtime_root(kind.host).map_err(|error| CommandError::new("internal", error))?;
    let folder = root.join("OfficeSessions");
    fs::create_dir_all(&folder).map_err(|error| CommandError::new("internal", error.to_string()))?;
    let request_path = folder.join(REQUEST_FILE);
    let response_path = folder.join(RESPONSE_FILE);
    let _ = fs::remove_file(&response_path);
    let id = uuid::Uuid::new_v4().to_string();
    macos_offline::atomic_write_runtime(&request_path, request_text(&id, operation, arguments).as_bytes(), 0o600)
        .map_err(|error| CommandError::new("internal", error))?;
    let seconds = deadline.as_secs().clamp(5, 1800);
    // osascript, not the app's main thread: a write's command may need the app while it runs
    let outcome = macos_offline::run_office_vba_script_subprocess(&script(kind, seconds), "VisualTeX agent operation");
    let response = fs::read_to_string(&response_path);
    let _ = fs::remove_file(&response_path);
    let answer = answer(kind, outcome, response, &id);
    // PowerPoint's native equations, copied while the request still names the presentation
    let answer = answer.map(|mut data| {
        if kind.host == OfficeHost::Powerpoint {
            copy_native_equations(operation, &mut data, &folder);
        }
        data
    });
    let _ = fs::remove_file(&request_path);
    answer
}

fn answer(kind: &HostKind, outcome: Result<(), String>, response: std::io::Result<String>, id: &str) -> Result<Value, CommandError> {
    if let Err(error) = outcome {
        let name = if kind.host == OfficeHost::Word { "Word" } else { "PowerPoint" };
        let lower = error.to_ascii_lowercase();
        return Err(if lower.contains("-1712") || lower.contains("timed out") {
            CommandError::new("host_unresponsive", format!("{name} did not finish in time (a dialog may be open): {error}"))
        } else if lower.contains("visualtex_agentrun") || lower.contains("macro") {
            CommandError::new("addin_not_loaded", format!("The VisualTeX add-in in {name} does not answer; update it (VisualTeX: Settings, Office integration) and restart {name}: {error}"))
        } else {
            CommandError::new("host_error", error)
        });
    }
    let text = response.map_err(|_| CommandError::new("host_error", "The add-in did not answer"))?;
    let answer: AgentResponse = serde_json::from_str(&text)
        .map_err(|error| CommandError::new("host_error", format!("The add-in's answer is not valid JSON: {error}")))?;
    if answer.id.as_deref().is_some_and(|answered| answered != id) {
        return Err(CommandError::new("host_error", "The add-in answered another request"));
    }
    if answer.ok {
        Ok(answer.data.unwrap_or(Value::Null))
    } else {
        Err(answer.error.unwrap_or_else(|| CommandError::new("operation_failed", "The add-in reported a failure")))
    }
}

/// The formulas of an answer: a list's, a read's, a selection's, or the one a write leaves behind.
fn for_each_formula(operation: &str, data: &mut Value, visit: &mut dyn FnMut(&mut Value)) {
    match operation {
        "list_formulas" => {
            if let Some(formulas) = data.get_mut("formulas").and_then(Value::as_array_mut) {
                formulas.iter_mut().for_each(|formula| visit(formula));
            }
        }
        "read_formula" => visit(data),
        _ => {
            for path in [&["formula"][..], &["result", "formula"], &["selection", "formula"]] {
                let mut target = Some(&mut *data);
                for key in path {
                    target = target.and_then(|value| value.get_mut(*key));
                }
                if let Some(formula) = target.filter(|formula| formula.is_object()) {
                    visit(formula);
                }
            }
        }
    }
}

/// PowerPoint's object model has no OMML: each native equation is copied by the add-in
/// (`VisualTeX_AgentCopyEquation`) and its clipboard package kept, with the user's clipboard put
/// back afterwards, as a restore reads them (presentation_redraw::copy_powerpoint_equations).
/// The equation's OMML goes into the formula as `omml`, or why it could not as `omml_error`.
fn copy_native_equations(operation: &str, data: &mut Value, folder: &std::path::Path) {
    let mut ids: Vec<String> = Vec::new();
    for_each_formula(operation, data, &mut |formula| {
        if formula.get("kind").and_then(Value::as_str) == Some("office_equation") {
            if let Some(id) = formula.get("formula_id").and_then(Value::as_str) {
                if id.starts_with("ppt:") && id[4..].chars().all(|c| c.is_ascii_digit() || c == ':') && !ids.iter().any(|known| known == id) {
                    ids.push(id.to_string());
                }
            }
        }
    });
    if ids.is_empty() {
        return;
    }
    let file = |index: usize| folder.join(format!("agent-equation-{index}.bin"));
    for index in 0..ids.len() {
        let _ = fs::remove_file(file(index));
    }
    let list = ids.iter().map(|id| format!("\"{id}\"")).collect::<Vec<_>>().join(", ");
    let directory = serde_json::to_string(folder.to_string_lossy().as_ref()).unwrap_or_default();
    let script = COPY_SCRIPT.replace("__IDS__", &list).replace("__DIRECTORY__", &directory);
    let copied = macos_offline::run_office_vba_script_subprocess(&script, "PowerPoint equation copy");
    let mut equations = std::collections::HashMap::new();
    for (index, id) in ids.iter().enumerate() {
        let equation = fs::read(file(index))
            .map_err(|_| match &copied {
                Err(error) => error.clone(),
                Ok(()) => "PowerPoint did not copy this equation".to_string(),
            })
            .and_then(|bytes| crate::office::powerpoint_omml::native_equation_from_clipboard(&bytes));
        let _ = fs::remove_file(file(index));
        equations.insert(id.clone(), equation);
    }
    for_each_formula(operation, data, &mut |formula| {
        let Some(id) = formula.get("formula_id").and_then(Value::as_str) else { return };
        match equations.get(id) {
            Some(Ok(omml)) => formula["omml"] = Value::String(omml.clone()),
            Some(Err(error)) => formula["omml_error"] = Value::String(error.clone()),
            None => {}
        }
    });
}

/// Copies the equations one after another with the user's clipboard saved around them; each
/// package is written as agent-equation-<n>.bin. One that cannot be copied is left out.
const COPY_SCRIPT: &str = r#"use framework "AppKit"
use scripting additions
set pb to current application's NSPasteboard's generalPasteboard()
set savedItems to current application's NSMutableArray's array()
repeat with sourceItem in (pb's pasteboardItems())
    set savedItem to current application's NSPasteboardItem's alloc()'s init()
    repeat with t in sourceItem's types()
        savedItem's setData:(sourceItem's dataForType:t) forType:t
    end repeat
    savedItems's addObject:savedItem
end repeat
set initialCount to (pb's changeCount()) as integer
set idList to {__IDS__}
set failureText to ""
try
    repeat with k from 1 to (count of idList)
        set countBefore to (pb's changeCount()) as integer
        with timeout of 60 seconds
            tell application "Microsoft PowerPoint" to run VB macro macro name "VisualTeX_AgentCopyEquation" list of parameters {item k of idList}
        end timeout
        if ((pb's changeCount()) as integer) is not countBefore then
            set nativeData to pb's dataForType:"com.microsoft.Art--Text-ClipFormat"
            if nativeData is not missing value then
                nativeData's writeToFile:(__DIRECTORY__ & "/agent-equation-" & ((k - 1) as text) & ".bin") atomically:true
            end if
        end if
    end repeat
on error errorText
    set failureText to errorText
end try
if ((pb's changeCount()) as integer) is not initialCount then
    pb's clearContents()
    pb's writeObjects:savedItems
end if
if failureText is not "" then error failureText"#;

/// Runs one operation in the host's add-in and returns its data, in the form the Windows add-ins give.
pub async fn dispatch(host_id: &str, operation: &str, arguments: Value, deadline: Duration) -> Result<Value, CommandError> {
    let kind = kind(host_id).ok_or_else(|| CommandError::new("document_not_found", "Unknown application"))?;
    let arguments = match arguments {
        Value::Object(map) => map,
        Value::Null => Map::new(),
        _ => return Err(CommandError::new("invalid_arguments", "arguments must be an object")),
    };
    let operation = operation.to_string();
    if let Some(addin) = kind.wps {
        let mut data = crate::office::wps::bridge::agent_run(addin, &operation, Value::Object(arguments), deadline)
            .await
            .map_err(|(code, message)| CommandError::new(&code, message))?;
        // WPS Presentation's equations come as the a14:m zone its saved copy holds
        for_each_formula(&operation, &mut data, &mut |formula| {
            if let Some(Value::String(zone)) = formula.as_object_mut().and_then(|item| item.remove("zone")) {
                match crate::office::wps::presentation::zone_equation(&zone) {
                    Ok(omml) => formula["omml"] = Value::String(omml),
                    Err(error) => formula["omml_error"] = Value::String(error),
                }
            }
        });
        enrich(kind, &operation, &mut data);
        return Ok(data);
    }
    tokio::task::spawn_blocking(move || {
        let mut data = run_blocking(kind, &operation, &arguments, deadline)?;
        enrich(kind, &operation, &mut data);
        Ok(data)
    })
    .await
    .map_err(|error| CommandError::new("internal", error.to_string()))?
}

// What the add-in reports, read the way Windows reads it.

fn decode_payload(value: Option<Value>) -> Option<Result<String, String>> {
    let text = value?.as_str()?.to_string();
    Some(
        BASE64_URL
            .decode(text.trim_end_matches('='))
            .map_err(|error| error.to_string())
            .and_then(|bytes| String::from_utf8(bytes).map_err(|error| error.to_string())),
    )
}

/// A formula of the add-in's answer: its payload (a VisualTeX formula's encoded metadata, a native Word
/// equation's XML) or a PowerPoint equation's copied OMML becomes LaTeX; its place becomes `location`.
fn enrich_formula(value: &mut Value) {
    let Some(item) = value.as_object_mut() else { return };
    let kind = item.get("kind").and_then(Value::as_str).unwrap_or_default().to_string();
    let display = item.get("display").and_then(Value::as_str).unwrap_or("inline").to_string();
    let payload = decode_payload(item.remove("payload_b64"));
    let copied = match (item.remove("omml"), item.remove("omml_error")) {
        (Some(Value::String(omml)), _) => {
            Some(crate::office::powerpoint_omml::to_latex::from_omml(&omml).map(|native| native.latex))
        }
        (_, Some(Value::String(error))) => Some(Err(error)),
        _ => None,
    };
    let mut lines: Option<Vec<String>> = None;
    let mut windows_size: Option<f64> = None;
    let converted: Option<Result<String, String>> = copied.or_else(|| payload.map(|payload| {
        let payload = payload?;
        match kind.as_str() {
            "picture" => {
                let metadata = macos_offline::decode_metadata(&payload)?;
                windows_size = windows_picture_size(&metadata, item.get("frame"), item.get("font_size_pt"));
                lines = Some(metadata.lines.iter().map(|line| line.latex.clone()).filter(|line| !line.is_empty()).collect());
                macos_offline::canonical_document_formula_latex(&metadata)
            }
            "office_equation" => {
                let math_ml = macos_offline::word_omml_to_mathml(&payload, display == "block")?;
                crate::office::omml_batch::mathml_to_latex(&math_ml)
            }
            _ => Err(format!("Unknown formula kind {kind}")),
        }
    }));
    match converted {
        Some(Ok(latex)) => {
            let lines = lines.filter(|lines| !lines.is_empty()).unwrap_or_else(|| vec![latex.clone()]);
            item.insert("latex".into(), Value::String(latex));
            item.insert("lines".into(), serde_json::json!(lines));
        }
        Some(Err(error)) => {
            item.insert("latex".into(), Value::Null);
            item.insert("error".into(), Value::String(error));
        }
        None => {}
    }
    if let Some(size) = windows_size {
        item.insert("font_size_pt".into(), serde_json::json!(size));
    }
    let mut location = Map::new();
    for key in ["start", "end", "slide", "slide_id", "shape_id", "shape_name", "equation"] {
        if let Some(value) = item.remove(key) {
            location.insert(key.into(), value);
        }
    }
    if let Some(id) = item.remove("vtx_formula_id").filter(|id| !id.is_null()) {
        location.insert("formula_id".into(), id);
    }
    item.insert("location".into(), Value::Object(location));
}

/// A picture VisualTeX for Windows drew, without the Mac's scale tags (no size, or the add-in's 18 pt default): its
/// size is the one its render size was drawn at (renderFontSizePt), scaled by the picture's height.
fn windows_picture_size(
    metadata: &crate::office::sessions::VisualTeXFormulaMetadata,
    frame: Option<&Value>,
    reported: Option<&Value>,
) -> Option<f64> {
    let reported = reported.and_then(Value::as_f64);
    if reported.is_some_and(|size| (size - 18.0).abs() > 0.001) {
        return None;
    }
    let render = metadata.render_font_size_pt.filter(|value| value.is_finite() && *value > 0.0)?;
    let natural = metadata.render_height_px.filter(|value| value.is_finite() && *value > 0.0)? * 0.75;
    let height = frame?.get("height_pt")?.as_f64()?;
    let size = (render * height / natural * 100.0).round() / 100.0;
    (size.is_finite() && size > 0.0).then_some(size)
}

/// The bundle id of the application the user is working in (`lsappinfo`; no permission prompt).
fn frontmost_bundle() -> Option<String> {
    let front = std::process::Command::new("/usr/bin/lsappinfo").arg("front").output().ok()?;
    let asn = String::from_utf8_lossy(&front.stdout).trim().to_string();
    if asn.is_empty() {
        return None;
    }
    let info = std::process::Command::new("/usr/bin/lsappinfo").args(["info", "-only", "bundleid", &asn]).output().ok()?;
    let text = String::from_utf8_lossy(&info.stdout);
    text.split('=').nth(1).map(|value| value.trim().trim_matches('"').to_string())
}

fn fnv_hash8(text: &str) -> String {
    let mut hash: u32 = 0x811c_9dc5;
    for byte in text.bytes() {
        hash ^= u32::from(byte);
        hash = hash.wrapping_mul(0x0100_0193);
    }
    format!("{hash:08x}")
}

fn enrich(kind: &HostKind, operation: &str, data: &mut Value) {
    match operation {
        "list_documents" => {
            let bundle = match (kind.wps, kind.host) {
                (Some(_), _) => "com.kingsoft.wpsoffice.mac",
                (None, OfficeHost::Word) => "com.microsoft.Word",
                (None, OfficeHost::Powerpoint) => "com.microsoft.Powerpoint",
            };
            let foreground = frontmost_bundle().is_some_and(|front| front.eq_ignore_ascii_case(bundle));
            if let Some(documents) = data.get_mut("documents").and_then(Value::as_array_mut) {
                for document in documents {
                    let active = document.get("active").and_then(Value::as_bool).unwrap_or(false);
                    document["foreground"] = Value::Bool(foreground && active);
                }
            }
        }
        "list_formulas" => {
            if let Some(revision) = data.get("revision").and_then(Value::as_str).map(fnv_hash8) {
                data["revision"] = Value::String(revision);
            }
        }
        _ => {}
    }
    if operation != "list_documents" && operation != "get_slide_layout" {
        for_each_formula(operation, data, &mut enrich_formula);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn picture_formulas_read_their_latex_from_the_metadata() {
        let metadata: crate::office::sessions::VisualTeXFormulaMetadata = serde_json::from_value(json!({
            "schema": "visualtex-formula", "schemaVersion": 1, "formulaId": "3f6c0f0a-1b2c-4d5e-8f90-123456789abc", "title": "",
            "latex": "E=mc^2", "lines": [{ "id": "6a0d6c55-2f0e-4c8b-9e3a-0b1c2d3e4f50", "latex": "E=mc^2" }],
            "codeFormat": "raw", "displayMode": "block", "createdWithVersion": "1.3.0", "updatedWithVersion": "1.3.0",
            "createdAt": "2026-11-02T00:00:00Z", "updatedAt": "2026-11-02T00:00:00Z" }))
        .unwrap();
        let encoded = macos_offline::encode_metadata(&metadata).unwrap();
        let mut item = json!({ "formula_id": "vtx:3f6c0f0a-1b2c-4d5e-8f90-123456789abc", "kind": "picture", "display": "block", "start": 3, "end": 9,
            "vtx_formula_id": "3f6c0f0a-1b2c-4d5e-8f90-123456789abc", "payload_b64": BASE64_URL.encode(encoded.as_bytes()) });
        enrich_formula(&mut item);
        assert_eq!(item["latex"], "E=mc^2", "{item}");
        assert_eq!(item["lines"], json!(["E=mc^2"]));
        assert_eq!(item["location"], json!({ "start": 3, "end": 9, "formula_id": "3f6c0f0a-1b2c-4d5e-8f90-123456789abc" }));
        assert!(item.get("payload_b64").is_none());
    }

    #[test]
    fn powerpoint_equations_read_their_latex_from_the_copied_omml() {
        let omml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:oMath>"#;
        let frame = json!({ "left_pt": 10, "top_pt": 20, "width_pt": 300, "height_pt": 40, "rotation": 0, "z_order": 2 });
        let mut data = json!({ "revision": "256.4.1;", "total": 2, "offset": 0, "formulas": [
            { "formula_id": "ppt:256:4:1", "kind": "office_equation", "display": "block", "slide": 1, "slide_id": 256, "shape_id": 4,
              "shape_name": "TextBox 3", "equation": 1, "frame": frame, "frame_of": "text_box", "payload_b64": null, "omml": omml },
            { "formula_id": "ppt:256:4:2", "kind": "office_equation", "display": "inline", "slide": 1, "slide_id": 256, "shape_id": 4,
              "equation": 2, "payload_b64": null, "omml_error": "PowerPoint did not copy this equation" } ] });
        enrich(&HOSTS[1], "list_formulas", &mut data);
        let first = &data["formulas"][0];
        assert_eq!(first["latex"], r"\frac{a}{b}", "{first}");
        assert_eq!(first["frame"], frame);
        assert_eq!(first["location"], json!({ "slide": 1, "slide_id": 256, "shape_id": 4, "shape_name": "TextBox 3", "equation": 1 }));
        assert!(first.get("omml").is_none());
        let second = &data["formulas"][1];
        assert_eq!(second["latex"], Value::Null);
        assert_eq!(second["error"], "PowerPoint did not copy this equation");
        assert_eq!(data["revision"].as_str().unwrap().len(), 8);
    }

    #[test]
    fn requests_flatten_to_dispatch_lines() {
        let arguments = json!({ "document": "d1", "limit": 50, "numbered": true,
            "location": { "slide": 2, "left_pt": 10.5 }, "lines": ["a", "b"] });
        let text = request_text("x", "list_formulas", arguments.as_object().unwrap());
        assert!(text.starts_with("id=x\nop=list_formulas\n"));
        assert!(text.contains(&format!("s.document={}\n", BASE64_URL.encode("d1"))));
        assert!(text.contains("n.limit=50\n"));
        assert!(text.contains("b.numbered=1\n"));
        assert!(text.contains("n.location.slide=2\n"));
        assert!(text.contains("n.location.left_pt=10.5\n"));
        assert!(text.contains("c.lines=2\n"));
        assert!(text.contains(&format!("s.lines.1={}\n", BASE64_URL.encode("b"))));
    }
}
