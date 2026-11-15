//! Opening what an add-in asks for: the visualtex: URL, document operations, WPS sessions, and sessions opened
//! by the Office MCP agent.

use super::*;

pub(super) fn handle_open_url(app: &AppHandle, value: &str) -> Result<(), String> {
    let received_at = Instant::now();
    let received_epoch_ms = epoch_ms();
    let session_id = parse_office_url(value)?;
    let state = app
        .try_state::<OfficeCompanionState>()
        .ok_or_else(|| "VisualTeX Office state is not initialized".to_string())?;
    // The host is not trusted until request validation completes, so queue the
    // first stage immediately after read_request resolves it below.
    let request = read_request(&session_id)?;
    let host = host_from_request_name(&request.host)?;
    queue_editor_performance(
        host,
        &session_id,
        "url-received",
        0.0,
        None,
        json!({ "receivedEpochMs": received_epoch_ms }),
    );
    queue_editor_performance(
        host,
        &session_id,
        "request-read",
        received_at.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    ensure_runtime_root(host)?;

    if matches!(
        request.operation.as_deref(),
        Some("documentImport" | "latexRedraw" | "formulaRestore")
    ) {
        return open_document_operation_window(app, &request, &session_id);
    }

    let silent = matches!(
        request.operation.as_deref(),
        Some("nativeToImage" | "imageToNative")
    );
    import_request(state.inner(), request)?;
    queue_editor_performance(
        host,
        &session_id,
        "request-imported",
        received_at.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    open_editor_window(
        app,
        host,
        &session_id,
        received_epoch_ms,
        received_at,
        silent,
    )
}

/// A document operation (import, LaTeX redraw, formula restore) runs in its own
/// window: any formula editor Session and any other document window ends.
pub(super) fn open_document_operation_window(
    app: &AppHandle,
    request: &MacOfflineSessionRequest,
    session_id: &str,
) -> Result<(), String> {
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        if let Some(window) = app.get_webview_window(editor_window_label(host)) {
            let _ = set_resident_editor_content_visible(&window, false);
            if let Some(active) = clear_any_editor_session(host) {
                let _ = window.emit(
                    OFFICE_EDITOR_CLEAR_EVENT,
                    json!({
                        "sessionId": active.session_id,
                        "generation": active.generation,
                    }),
                );
            }
            let _ = set_resident_editor_parked(&window, true);
        }
    }
    for (label, window) in app.webview_windows() {
        if label.starts_with("office-native-document-")
            && label != document_import_window_label(session_id)
        {
            let _ = window.destroy();
        }
    }
    if matches!(
        request.operation.as_deref(),
        Some("latexRedraw" | "formulaRestore")
    ) {
        if presentation_redraw::needs_copied_equations(request) {
            // PowerPoint copies the equations to restore first, off the main
            // thread; the window opens when they are there (or with the reason
            // they are not).
            let app = app.clone();
            let session_id = session_id.to_string();
            std::thread::spawn(move || {
                let _ = presentation_redraw::copy_powerpoint_equations(&session_id);
                let opener = app.clone();
                let _ = app.run_on_main_thread(move || {
                    let _ = open_word_latex_redraw_window(&opener, &session_id);
                });
            });
            return Ok(());
        }
        open_word_latex_redraw_window(app, session_id)
    } else {
        let office = if request.client.as_deref() == Some("wps") { "WPS" } else { "Word" };
        open_document_import_window(app, session_id, office)
    }
}

/// Opens a document operation for the WPS add-in (office::wps::bridge): the
/// request and its source (the LaTeX redraw text or the formula restore
/// manifest, as the Word add-in writes them) are stored where Word's live, so
/// the same validation, windows and renderer run; the result goes back to the
/// add-in (deliver_wps_document_commit).
pub(crate) fn open_wps_document_session(
    app: &AppHandle,
    request: Value,
    source: Option<String>,
) -> Result<(), String> {
    let request: MacOfflineSessionRequest = serde_json::from_value(request)
        .map_err(|error| format!("The WPS document request is invalid: {error}"))?;
    if request.client.as_deref() != Some("wps") {
        return Err("The WPS document request is not marked as WPS".to_string());
    }
    let operation = request.operation.clone().unwrap_or_default();
    if !matches!(operation.as_str(), "documentImport" | "latexRedraw" | "formulaRestore") {
        return Err("The WPS document request has no document operation".to_string());
    }
    let session_id = request.session_id.clone();
    validate_uuid(&session_id, "Request Session id")?;
    // WPS Writer's operations live where Word's do, WPS Presentation's where
    // PowerPoint's do.
    let host = host_from_request_name(&request.host)?;
    ensure_runtime_root(host)?;
    let directory = session_directory(host, &session_id)?;
    fs::create_dir_all(&directory)
        .map_err(|error| format!("Unable to create {}: {error}", directory.display()))?;
    set_mode(&directory, 0o700)?;
    let opened = (|| {
        let source_path = match (host, operation.as_str()) {
            (OfficeHost::Word, "latexRedraw") => Some(directory.join(LATEX_REDRAW_SOURCE_FILE)),
            (OfficeHost::Word, "formulaRestore") => Some(formula_restore_source_path(&session_id)?),
            (OfficeHost::Powerpoint, "latexRedraw") => Some(directory.join(presentation_redraw::REDRAW_SOURCE_FILE)),
            (OfficeHost::Powerpoint, "formulaRestore") => Some(directory.join(presentation_redraw::RESTORE_SOURCE_FILE)),
            _ => None,
        };
        match (source_path, source) {
            (Some(path), Some(text)) => atomic_write_runtime(&path, text.as_bytes(), 0o600)?,
            (None, None) => {}
            (Some(_), None) => return Err("The WPS document request is missing its source".to_string()),
            (None, Some(_)) => return Err("A WPS document import takes no source".to_string()),
        }
        validate_request(&request, &session_id)?;
        let bytes = serde_json::to_vec(&request)
            .map_err(|error| format!("Unable to encode the WPS document request: {error}"))?;
        atomic_write_runtime(&request_path(host, &session_id)?, &bytes, 0o600)?;
        let request = read_request(&session_id)?;
        open_document_operation_window(app, &request, &session_id)
    })();
    if opened.is_err() {
        let _ = cleanup_session_files(host, &session_id);
    }
    opened
}

/// The request of a document window's Session.
pub(super) fn document_window_request(window: &WebviewWindow) -> Option<MacOfflineSessionRequest> {
    let compact = window.label().strip_prefix("office-native-document-")?;
    if compact.len() != 32 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return None;
    }
    let session_id = format!(
        "{}-{}-{}-{}-{}",
        &compact[..8],
        &compact[8..12],
        &compact[12..16],
        &compact[16..20],
        &compact[20..]
    );
    read_request(&session_id).ok()
}

/// Whether a document window belongs to a WPS Session: WPS, not Word, is
/// the application to bring forward and to tell about progress.
pub(super) fn document_window_is_wps(window: &WebviewWindow) -> bool {
    document_window_request(window).is_some_and(|request| request.client.as_deref() == Some("wps"))
}

/// Whether a document window belongs to a PowerPoint (not WPS) Session.
pub(super) fn document_window_is_powerpoint(window: &WebviewWindow) -> bool {
    document_window_request(window)
        .is_some_and(|request| request.host == "powerpoint" && request.client.as_deref() != Some("wps"))
}

/// Opens a formula Session for a WPS add-in (office::wps::bridge). The
/// request is stored where a Word (or, from WPS Presentation, a PowerPoint)
/// request lives and then read back, so it passes the same validation,
/// import and editor window as a Word or PowerPoint formula.
pub(crate) fn open_wps_session(app: &AppHandle, request: Value) -> Result<(), String> {
    let received_at = Instant::now();
    let received_epoch_ms = epoch_ms();
    let request: MacOfflineSessionRequest = serde_json::from_value(request)
        .map_err(|error| format!("The WPS formula request is invalid: {error}"))?;
    if request.client.as_deref() != Some("wps") {
        return Err("The WPS formula request is not marked as WPS".to_string());
    }
    validate_request(&request, &request.session_id)?;
    let host = host_from_request_name(&request.host)?;
    let session_id = request.session_id.clone();
    ensure_runtime_root(host)?;
    let directory = session_directory(host, &session_id)?;
    fs::create_dir_all(&directory)
        .map_err(|error| format!("Unable to create {}: {error}", directory.display()))?;
    set_mode(&directory, 0o700)?;
    let bytes = serde_json::to_vec(&request)
        .map_err(|error| format!("Unable to encode the WPS formula request: {error}"))?;
    atomic_write_runtime(&request_path(host, &session_id)?, &bytes, 0o600)?;
    let state = app
        .try_state::<OfficeCompanionState>()
        .ok_or_else(|| "VisualTeX Office state is not initialized".to_string())?;
    let request = read_request(&session_id)?;
    // A conversion renders in the hidden editor and commits by itself.
    let silent = matches!(request.operation.as_deref(), Some("nativeToImage" | "imageToNative"));
    import_request(state.inner(), request)?;
    open_editor_window(
        app,
        host,
        &session_id,
        received_epoch_ms,
        received_at,
        silent,
    )
}

// AI agents' writes (office::automation::writes): the add-in's command prepares
// the formula's request as its Ribbon button does; the app draws the agent's
// LaTeX in the hidden editor, which applies it like Insert.
/// How the sessions agents wait for ended: None while one runs.
pub(super) fn agent_sessions() -> &'static Mutex<std::collections::HashMap<String, Option<Result<(), String>>>> {
    static SESSIONS: OnceLock<Mutex<std::collections::HashMap<String, Option<Result<(), String>>>>> = OnceLock::new();
    SESSIONS.get_or_init(|| Mutex::new(std::collections::HashMap::new()))
}

/// An agent waits for this session (a formula's, or a document operation's window).
pub(crate) fn watch_agent_session(session_id: &str) {
    if let Ok(mut sessions) = agent_sessions().lock() {
        sessions.entry(session_id.to_string()).or_insert(None);
    }
}

/// The first outcome of a watched session counts (a failed commit is then cancelled).
pub(super) fn finish_agent_session(session_id: &str, outcome: Result<(), String>) {
    if let Ok(mut sessions) = agent_sessions().lock() {
        if let Some(slot @ None) = sessions.get_mut(session_id) {
            *slot = Some(outcome);
        }
    }
}

/// The outcome of a watched session once it has one (and the session is no longer watched).
pub(crate) fn take_agent_session_outcome(session_id: &str) -> Option<Result<(), String>> {
    let mut sessions = agent_sessions().lock().ok()?;
    if matches!(sessions.get(session_id), Some(Some(_))) {
        return sessions.remove(session_id).flatten();
    }
    None
}

pub(crate) fn forget_agent_session(session_id: &str) {
    if let Ok(mut sessions) = agent_sessions().lock() {
        sessions.remove(session_id);
    }
    take_agent_session_notes(session_id);
}

/// What a watched session's window would have shown the user (formulas it skipped, why it failed).
pub(super) fn agent_notes() -> &'static Mutex<std::collections::HashMap<String, Vec<String>>> {
    static NOTES: OnceLock<Mutex<std::collections::HashMap<String, Vec<String>>>> = OnceLock::new();
    NOTES.get_or_init(|| Mutex::new(std::collections::HashMap::new()))
}

pub(crate) fn take_agent_session_notes(session_id: &str) -> Vec<String> {
    agent_notes().lock().ok().and_then(|mut notes| notes.remove(session_id)).unwrap_or_default()
}

/// Opens an agent's formula: `request` is what the add-in's command wrote for
/// the editor (create, edit or conversion), `lines` the agent's LaTeX rows
/// (none: a conversion keeps the formula's). The hidden editor draws and
/// applies it; watch the session for the outcome.
pub(crate) fn open_agent_session(
    app: &AppHandle,
    request: Value,
    lines: Option<Vec<String>>,
    font_size_pt: Option<f64>,
) -> Result<String, String> {
    let mut request: MacOfflineSessionRequest = serde_json::from_value(request)
        .map_err(|error| format!("The add-in's formula request is invalid: {error}"))?;
    request.agent = true;
    if let Some(size) = font_size_pt {
        // a slide's formula takes its size from the PowerPoint geometry
        match request.power_point.as_mut() {
            Some(powerpoint) => powerpoint.font_size_pt = Some(size),
            None => request.font_size_pt = Some(size),
        }
    }
    let session_id = request.session_id.clone();
    validate_request(&request, &session_id)?;
    let host = host_from_request_name(&request.host)?;
    if active_editor_session(host).is_some_and(|active| active.activation.session_id != session_id) {
        return Err("The VisualTeX formula editor is open; try again when the user has closed it".to_string());
    }
    ensure_runtime_root(host)?;
    let directory = session_directory(host, &session_id)?;
    fs::create_dir_all(&directory)
        .map_err(|error| format!("Unable to create {}: {error}", directory.display()))?;
    set_mode(&directory, 0o700)?;
    let bytes = serde_json::to_vec(&request)
        .map_err(|error| format!("Unable to encode the formula request: {error}"))?;
    atomic_write_runtime(&request_path(host, &session_id)?, &bytes, 0o600)?;
    let state = app
        .try_state::<OfficeCompanionState>()
        .ok_or_else(|| "VisualTeX Office state is not initialized".to_string())?;
    import_request(state.inner(), read_request(&session_id)?)?;
    if let Some(lines) = lines {
        let lines: Vec<FormulaLine> = lines
            .into_iter()
            .map(|latex| FormulaLine { id: Uuid::new_v4().to_string(), latex })
            .collect();
        state
            .session_store
            .patch(&session_id, json!({ "lines": lines, "activeLineId": null, "dirty": true }))
            .map_err(|error| error.to_string())?;
    }
    watch_agent_session(&session_id);
    let (sender, receiver) = mpsc::channel();
    let opener = app.clone();
    let id = session_id.clone();
    app.run_on_main_thread(move || {
        let _ = sender.send(open_editor_window(&opener, host, &id, epoch_ms(), Instant::now(), true));
    })
    .map_err(|error| error.to_string())?;
    let opened = receiver
        .recv_timeout(Duration::from_secs(30))
        .map_err(|_| "The VisualTeX editor did not start".to_string())
        .and_then(|opened| opened);
    if let Err(error) = opened {
        forget_agent_session(&session_id);
        let _ = cancel_host(&read_request(&session_id)?);
        let _ = cleanup_session_files(host, &session_id);
        return Err(error);
    }
    Ok(session_id)
}

/// Opens an agent's Markdown/LaTeX import (the add-in's LaTeX/Markdown import
/// request): the import window takes the agent's text and inserts it.
pub(crate) fn open_agent_import(app: &AppHandle, request: Value, agent: Value) -> Result<String, String> {
    let request: MacOfflineSessionRequest = serde_json::from_value(request)
        .map_err(|error| format!("The add-in's import request is invalid: {error}"))?;
    let session_id = request.session_id.clone();
    if request.operation.as_deref() != Some("documentImport") {
        return Err("The add-in did not start an import".to_string());
    }
    validate_request(&request, &session_id)?;
    let host = host_from_request_name(&request.host)?;
    ensure_runtime_root(host)?;
    let directory = session_directory(host, &session_id)?;
    fs::create_dir_all(&directory)
        .map_err(|error| format!("Unable to create {}: {error}", directory.display()))?;
    set_mode(&directory, 0o700)?;
    let bytes = serde_json::to_vec(&request)
        .map_err(|error| format!("Unable to encode the import request: {error}"))?;
    atomic_write_runtime(&request_path(host, &session_id)?, &bytes, 0o600)?;
    atomic_write_runtime(&directory.join(AGENT_IMPORT_FILE), agent.to_string().as_bytes(), 0o600)?;
    watch_agent_session(&session_id);
    let (sender, receiver) = mpsc::channel();
    let opener = app.clone();
    let id = session_id.clone();
    app.run_on_main_thread(move || {
        let opened = read_request(&id).and_then(|request| open_document_operation_window(&opener, &request, &id));
        let _ = sender.send(opened);
    })
    .map_err(|error| error.to_string())?;
    let opened = receiver
        .recv_timeout(Duration::from_secs(30))
        .map_err(|_| "The VisualTeX import window did not open".to_string())
        .and_then(|opened| opened);
    if let Err(error) = opened {
        forget_agent_session(&session_id);
        let _ = cancel_document_import_blocking(session_id.clone());
        return Err(error);
    }
    Ok(session_id)
}

pub(super) fn agent_import(request: &MacOfflineSessionRequest) -> Option<Value> {
    let host = host_from_request_name(&request.host).ok()?;
    let path = session_directory(host, &request.session_id).ok()?.join(AGENT_IMPORT_FILE);
    serde_json::from_slice(&fs::read(path).ok()?).ok()
}

pub(super) fn protect_office_open<T>(
    operation: impl FnOnce() -> Result<T, String>,
) -> Result<T, String> {
    match std::panic::catch_unwind(std::panic::AssertUnwindSafe(operation)) {
        Ok(result) => result,
        Err(payload) => Err(format!(
            "The VisualTeX Office editor recovered from an internal error: {}",
            crate::panic_payload_detail(payload.as_ref())
        )),
    }
}

pub(super) fn record_office_open_failure(value: &str, error: &str) {
    let Ok(session_id) = parse_office_url(value) else {
        return;
    };
    let detail = error
        .chars()
        .take(4_096)
        .map(|character| {
            if matches!(character, '\r' | '\n' | '\0') {
                ' '
            } else {
                character
            }
        })
        .collect::<String>();
    let record = format!(
        "visualtex-office-open-error-v1\nepochMs={}\nsessionId={session_id}\nerror={detail}\n",
        epoch_ms()
    );
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        let Ok(directory) = session_directory(host, &session_id) else {
            continue;
        };
        if !directory.is_dir() || !directory.join(REQUEST_FILE).is_file() {
            continue;
        }
        if let Err(write_error) = atomic_write_runtime(
            &directory.join(EDITOR_OPEN_ERROR_FILE),
            record.as_bytes(),
            0o600,
        ) {
            eprintln!("Unable to record the recoverable Office open error: {write_error}");
        }
    }
}

pub(crate) fn handle_open_url_safely(app: &AppHandle, value: &str) -> Result<(), String> {
    let result = protect_office_open(|| handle_open_url(app, value));
    if let Err(error) = &result {
        record_office_open_failure(value, error);
    }
    result
}
