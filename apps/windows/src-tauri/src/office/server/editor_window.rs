//! The Office editor window: its saved size, the session it shows, opening, showing, hiding and prewarming
//! it, and the bulk import and conversion windows.

use super::*;

#[cfg(target_os = "windows")]
#[derive(Debug, Clone, Copy, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub(super) struct OfficeEditorWindowSizePreference {
    pub(super) width: f64,
    pub(super) height: f64,
}

#[cfg(target_os = "windows")]
pub(super) fn normalize_office_editor_window_size(width: f64, height: f64) -> (f64, f64) {
    let width = if width.is_finite() {
        width.clamp(MIN_OFFICE_EDITOR_WIDTH, MAX_OFFICE_EDITOR_WIDTH)
    } else {
        DEFAULT_OFFICE_EDITOR_WIDTH
    };
    let height = if height.is_finite() {
        height.clamp(MIN_OFFICE_EDITOR_HEIGHT, MAX_OFFICE_EDITOR_HEIGHT)
    } else {
        DEFAULT_OFFICE_EDITOR_HEIGHT
    };
    (width, height)
}

#[cfg(target_os = "windows")]
pub(super) fn office_editor_window_size_path(app: &tauri::AppHandle) -> Result<std::path::PathBuf, String> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX app-data directory: {error}"))?;
    Ok(app_data.join("office").join(OFFICE_EDITOR_WINDOW_SIZE_FILE))
}

#[cfg(target_os = "windows")]
pub(super) fn load_office_editor_window_size(app: &tauri::AppHandle) -> (f64, f64) {
    let preference = office_editor_window_size_path(app)
        .ok()
        .and_then(|path| fs::read(path).ok())
        .and_then(|bytes| serde_json::from_slice::<OfficeEditorWindowSizePreference>(&bytes).ok());
    preference
        .map(|value| normalize_office_editor_window_size(value.width, value.height))
        .unwrap_or((DEFAULT_OFFICE_EDITOR_WIDTH, DEFAULT_OFFICE_EDITOR_HEIGHT))
}

#[cfg(target_os = "windows")]
pub(super) fn write_office_editor_window_size(
    app: &tauri::AppHandle,
    width: f64,
    height: f64,
) -> Result<(f64, f64), String> {
    let (width, height) = normalize_office_editor_window_size(width, height);
    let path = office_editor_window_size_path(app)?;
    let parent = path
        .parent()
        .ok_or_else(|| "Unable to resolve the Office editor size directory.".to_string())?;
    fs::create_dir_all(parent).map_err(|error| error.to_string())?;
    let preference = OfficeEditorWindowSizePreference { width, height };
    let bytes = serde_json::to_vec_pretty(&preference).map_err(|error| error.to_string())?;
    fs::write(path, bytes).map_err(|error| error.to_string())?;
    Ok((width, height))
}

#[cfg(target_os = "windows")]
pub(crate) fn configuration_office_editor_window_size(
    app: &tauri::AppHandle,
) -> Option<(f64, f64)> {
    if let Some(window) = app.get_webview_window("office-session-editor") {
        if let (Ok(physical), Ok(scale_factor)) = (window.inner_size(), window.scale_factor()) {
            let scale_factor = scale_factor.max(0.1);
            let size = normalize_office_editor_window_size(
                f64::from(physical.width) / scale_factor,
                f64::from(physical.height) / scale_factor,
            );
            return Some(size);
        }
    }
    Some(load_office_editor_window_size(app))
}

#[cfg(target_os = "windows")]
pub(crate) fn apply_configuration_office_editor_window_size(
    app: &tauri::AppHandle,
    width: f64,
    height: f64,
) -> Result<(), String> {
    let (width, height) = write_office_editor_window_size(app, width, height)?;
    if let Some(window) = app.get_webview_window("office-session-editor") {
        window
            .set_size(tauri::LogicalSize::new(width, height))
            .map_err(|error| error.to_string())?;
    }
    Ok(())
}

#[cfg(target_os = "windows")]
pub(crate) fn schedule_persist_office_editor_window_size(
    app: tauri::AppHandle,
    width: f64,
    height: f64,
) {
    use std::sync::atomic::Ordering;

    let (width, height) = normalize_office_editor_window_size(width, height);
    let generation = OFFICE_EDITOR_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst) + 1;
    tauri::async_runtime::spawn(async move {
        tokio::time::sleep(std::time::Duration::from_millis(280)).await;
        if OFFICE_EDITOR_SIZE_WRITE_GENERATION.load(Ordering::SeqCst) != generation {
            return;
        }
        if let Err(error) = write_office_editor_window_size(&app, width, height) {
            eprintln!("Unable to persist the Office editor window size: {error}");
        }
    });
}

pub(super) async fn reveal_desktop_app(State(context): State<ServerContext>) -> Response {
    let Some(app) = context.companion.app.clone() else {
        return (
            StatusCode::SERVICE_UNAVAILABLE,
            Json(serde_json::json!({
                "error": "The VisualTeX desktop window is unavailable in this runtime"
            })),
        )
            .into_response();
    };

    match crate::app_lifecycle::ensure_main_window(&app) {
        Ok(_) => StatusCode::NO_CONTENT.into_response(),
        Err(error) => {
            crate::app_lifecycle::append_lifecycle_log(format!(
                "Office companion reveal request could not ensure main window: {error}"
            ));
            (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({ "error": error })),
            )
                .into_response()
        }
    }
}

#[cfg(target_os = "windows")]
pub(super) fn office_editor_session_is_active(session_id: &str) -> bool {
    OFFICE_EDITOR_ACTIVE_SESSION
        .get_or_init(|| Mutex::new(None))
        .lock()
        .map(|current| current.as_deref() == Some(session_id))
        .unwrap_or(false)
}

#[cfg(target_os = "windows")]
pub(super) fn office_editor_page_is_ready() -> bool {
    OFFICE_EDITOR_PAGE_READY.load(std::sync::atomic::Ordering::Acquire)
}

#[cfg(target_os = "windows")]
pub(super) fn office_editor_session_id_from_url(url: &tauri::Url) -> Option<String> {
    if let Some(mut segments) = url.path_segments() {
        if segments.next() == Some("dialog") {
            if let Some(session_id) = segments.next() {
                if valid_session_id(session_id) {
                    return Some(session_id.to_string());
                }
            }
        }
    }
    url.query_pairs()
        .find_map(|(key, value)| {
            (key == "sessionId" && valid_session_id(value.as_ref()))
                .then(|| value.into_owned())
        })
}

#[cfg(target_os = "windows")]
pub(super) fn handle_office_editor_page_load(
    window: WebviewWindow,
    payload: PageLoadPayload<'_>,
) {
    crate::app_lifecycle::append_lifecycle_log(format!(
        "Office editor page-load event={:?} url={}",
        payload.event(),
        payload.url()
    ));
    if matches!(payload.event(), PageLoadEvent::Started) {
        OFFICE_EDITOR_PAGE_READY.store(false, std::sync::atomic::Ordering::Release);
        return;
    }
    if !matches!(payload.event(), PageLoadEvent::Finished) {
        return;
    }

    let Some(loaded_session_id) = office_editor_session_id_from_url(payload.url()) else {
        let editor_is_inactive = OFFICE_EDITOR_ACTIVE_SESSION
            .get_or_init(|| Mutex::new(None))
            .lock()
            .map(|current| current.is_none())
            .unwrap_or(false);
        if editor_is_inactive {
            OFFICE_EDITOR_PAGE_READY.store(true, std::sync::atomic::Ordering::Release);
            crate::app_lifecycle::append_lifecycle_log(
                "hidden Office editor WebView prewarm shell is ready",
            );
        } else {
            crate::app_lifecycle::append_lifecycle_log(
                "Office editor ignored shell page-load completion after a Session became active",
            );
        }
        return;
    };
    if !office_editor_session_is_active(&loaded_session_id) {
        crate::app_lifecycle::append_lifecycle_log(format!(
            "Office editor ignored stale page-load completion for session={loaded_session_id}"
        ));
        return;
    }
    OFFICE_EDITOR_PAGE_READY.store(true, std::sync::atomic::Ordering::Release);
    if let Err(error) = bring_session_window_to_front(&window, Some(loaded_session_id.clone())) {
        crate::app_lifecycle::append_lifecycle_log(format!(
            "Office editor could not reveal loaded session={loaded_session_id}: {error}"
        ));
    }
}

#[cfg(target_os = "windows")]
pub(super) fn switch_office_editor_session(window: &WebviewWindow, session_id: &str) -> Result<(), String> {
    let encoded = serde_json::to_string(session_id)
        .map_err(|error| format!("Unable to encode the Office Session id: {error}"))?;
    let script = format!(
        "window.__VISUALTEX_OFFICE_SESSION_ID__={encoded};history.replaceState(null,'','/dialog?runtime=vsto-desktop&sessionId='+encodeURIComponent({encoded}));window.dispatchEvent(new CustomEvent('visualtex-office-session',{{detail:{{sessionId:{encoded}}}}}));"
    );
    window
        .eval(&script)
        .map_err(|error| format!("Unable to switch the VisualTeX Session window: {error}"))
}

#[cfg(target_os = "windows")]
pub(super) fn clear_office_editor_session(window: &WebviewWindow) {
    let _ = window.eval(
        "window.__VISUALTEX_OFFICE_SESSION_ID__=undefined;history.replaceState(null,'','/dialog?runtime=vsto-desktop');window.dispatchEvent(new CustomEvent('visualtex-office-session',{detail:{sessionId:''}}));",
    );
}

#[cfg(target_os = "windows")]
pub(super) fn schedule_office_editor_reveal_fallback(app: tauri::AppHandle, session_id: String) {
    tauri::async_runtime::spawn(async move {
        tokio::time::sleep(std::time::Duration::from_secs(4)).await;
        let Ok(_transition) = OFFICE_EDITOR_WINDOW_TRANSITION_LOCK
            .get_or_init(|| Mutex::new(()))
            .lock()
        else {
            return;
        };
        if !office_editor_session_is_active(&session_id) {
            return;
        }
        if !office_editor_page_is_ready() {
            crate::app_lifecycle::append_lifecycle_log(format!(
                "Office editor reveal fallback deferred because page is not ready for session={session_id}"
            ));
            return;
        }
        let Some(window) = app.get_webview_window(OFFICE_EDITOR_WINDOW_LABEL) else {
            return;
        };
        if window.is_visible().unwrap_or(false) {
            return;
        }
        crate::app_lifecycle::append_lifecycle_log(format!(
            "Office editor page-load reveal fallback used for session={session_id}"
        ));
        if let Err(error) = bring_session_window_to_front(&window, Some(session_id.clone())) {
            crate::app_lifecycle::append_lifecycle_log(format!(
                "Office editor reveal fallback failed for session={session_id}: {error}"
            ));
        }
    });
}

#[cfg(target_os = "windows")]
pub(super) fn office_session_export_is_committable(session: &OfficeFormulaSession) -> bool {
    if session.mode == OfficeSessionMode::Edit && !session.dirty {
        return true;
    }
    if !session.lines.iter().any(|line| !line.latex.trim().is_empty()) {
        return false;
    }
    let Some(export) = session.export_result.as_ref() else {
        return false;
    };
    let has_svg = !export.svg.trim().is_empty() || !export.svg_base64.trim().is_empty();
    if session.host == OfficeHost::Powerpoint {
        return match session.object_mode.as_str() {
            "wordOmml" | "mathTypeOle" => export
                .math_ml
                .as_deref()
                .map(str::trim)
                .is_some_and(|math_ml| math_ml.starts_with("<math")),
            "nativeOle" | "crossPlatformPicture" => has_svg,
            _ => false,
        };
    }
    match session.object_mode.as_str() {
        "wordOmml" | "mathTypeOle" => export
            .math_ml
            .as_deref()
            .map(str::trim)
            .is_some_and(|math_ml| math_ml.starts_with("<math")),
        "nativeOle" => has_svg
            && export
                .png_base64
                .as_deref()
                .is_some_and(|png| !png.trim().is_empty()),
        "crossPlatformPicture" => export
            .png_base64
            .as_deref()
            .is_some_and(|png| !png.trim().is_empty()),
        _ => false,
    }
}

#[cfg(target_os = "windows")]
pub(super) fn close_patch_for_office_editor(session: &OfficeFormulaSession) -> Option<serde_json::Value> {
    if !matches!(
        session.status,
        OfficeSessionStatus::Created | OfficeSessionStatus::Editing
    ) {
        return None;
    }
    if session.auto_commit_on_close && office_session_export_is_committable(session) {
        return Some(serde_json::json!({
            "status": "committing",
            "explicitCancel": false,
            "error": null
        }));
    }
    Some(serde_json::json!({
        "status": "cancelled",
        "explicitCancel": true,
        "error": null
    }))
}

#[cfg(target_os = "windows")]
pub(crate) fn request_desktop_editor_window_close(app: &tauri::AppHandle) -> Result<(), String> {
    let transition = OFFICE_EDITOR_WINDOW_TRANSITION_LOCK
        .get_or_init(|| Mutex::new(()))
        .lock()
        .map_err(|_| "VisualTeX Office editor window transition state is unavailable.".to_string())?;
    let Some(window) = app.get_webview_window(OFFICE_EDITOR_WINDOW_LABEL) else {
        return Ok(());
    };
    let active_session = OFFICE_EDITOR_ACTIVE_SESSION
        .get_or_init(|| Mutex::new(None))
        .lock()
        .map_err(|_| "VisualTeX Office editor Session state is unavailable.".to_string())?
        .clone();
    let Some(session_id) = active_session else {
        let _ = window.hide();
        return Ok(());
    };

    crate::app_lifecycle::append_lifecycle_log(format!(
        "native Office editor close requested for session={session_id}"
    ));
    let page_dispatch_error = window
        .eval(
            "window.dispatchEvent(new CustomEvent('visualtex-office-close-requested'));",
        )
        .err()
        .map(|error| error.to_string());
    drop(transition);

    let dispatch_session_id = session_id.clone();
    let companion = app
        .try_state::<OfficeCompanionState>()
        .map(|state| state.inner().clone())
        .ok_or_else(|| "VisualTeX Office companion state is unavailable.".to_string())?;
    let app = app.clone();
    tauri::async_runtime::spawn(async move {
        // Give the page's normal awaited commit/cancel path the first chance to
        // finish. WebView2 can occasionally lose a synthetic event while its
        // renderer is busy or navigating, so the persistent Session store is
        // the authoritative fallback rather than leaving Word blocked forever.
        tokio::time::sleep(std::time::Duration::from_secs(5)).await;
        let store = companion.session_store.clone();
        let fallback_session_id = session_id.clone();
        let fallback = tokio::task::spawn_blocking(move || {
            let session = store.get(&fallback_session_id)?;
            if let Some(patch) = close_patch_for_office_editor(&session) {
                let next = store.patch(&fallback_session_id, patch)?;
                Ok::<_, SessionError>(Some(next.status))
            } else {
                Ok::<_, SessionError>(None)
            }
        })
        .await;
        match fallback {
            Ok(Ok(Some(status))) => crate::app_lifecycle::append_lifecycle_log(format!(
                "native Office editor close fallback finalized session={session_id} status={status:?}"
            )),
            Ok(Ok(None)) => {}
            Ok(Err(error)) => crate::app_lifecycle::append_lifecycle_log(format!(
                "native Office editor close fallback could not finalize session={session_id}: {error}"
            )),
            Err(error) => crate::app_lifecycle::append_lifecycle_log(format!(
                "native Office editor close fallback task failed for session={session_id}: {error}"
            )),
        }

        let deadline = std::time::Instant::now() + std::time::Duration::from_secs(45);
        loop {
            let store = companion.session_store.clone();
            let poll_session_id = session_id.clone();
            let status = tokio::task::spawn_blocking(move || {
                store.get(&poll_session_id).map(|session| session.status)
            })
            .await;
            let terminal = matches!(
                status,
                Ok(Ok(
                    OfficeSessionStatus::Completed
                        | OfficeSessionStatus::Cancelled
                        | OfficeSessionStatus::Failed
                ))
            );
            if terminal || std::time::Instant::now() >= deadline {
                let cleanup_owned_editor =
                    hide_owned_office_editor_on_main_thread(app.clone(), session_id.clone())
                        .await
                        .unwrap_or(false);
                if !cleanup_owned_editor {
                    crate::app_lifecycle::append_lifecycle_log(format!(
                        "native Office editor stale close cleanup skipped for session={session_id}"
                    ));
                }
                if !terminal {
                    crate::app_lifecycle::append_lifecycle_log(format!(
                        "native Office editor close timed out waiting for terminal session={session_id}"
                    ));
                }
                break;
            }
            tokio::time::sleep(std::time::Duration::from_millis(100)).await;
        }
    });

    if let Some(error) = page_dispatch_error {
        crate::app_lifecycle::append_lifecycle_log(format!(
            "native Office editor close page dispatch failed for session={dispatch_session_id}: {error}"
        ));
    }
    Ok(())
}

#[cfg(target_os = "windows")]
pub(super) fn prime_office_editor_native_presentation(window: &WebviewWindow) -> Result<(), String> {
    let hwnd = window
        .hwnd()
        .map_err(|error| format!("Unable to resolve the VisualTeX Session HWND: {error}"))?;
    let original_position = window.outer_position().ok();
    if let Err(error) = window.set_position(tauri::PhysicalPosition::new(-32_000, -32_000)) {
        crate::app_lifecycle::append_lifecycle_log(format!(
            "Office editor prewarm could not move HWND off-screen: {error}"
        ));
    }
    unsafe {
        let _ = ShowWindow(hwnd.0, SW_SHOWNOACTIVATE);
        let _ = UpdateWindow(hwnd.0);
        if IsWindowVisible(hwnd.0) == 0 {
            return Err(
                "VisualTeX Session HWND remained hidden during native presentation prewarm."
                    .to_string(),
            );
        }
        let _ = ShowWindow(hwnd.0, SW_HIDE);
        if IsWindowVisible(hwnd.0) != 0 {
            return Err(
                "VisualTeX Session HWND remained visible after native presentation prewarm."
                    .to_string(),
            );
        }
    }
    if let Some(position) = original_position {
        let _ = window.set_position(position);
    }
    crate::app_lifecycle::append_lifecycle_log(
        "hidden Office editor HWND native presentation prewarm completed",
    );
    Ok(())
}

#[cfg(target_os = "windows")]
pub(super) fn force_office_editor_window_visible(window: &WebviewWindow) -> Result<(), String> {
    let hwnd = window
        .hwnd()
        .map_err(|error| format!("Unable to resolve the VisualTeX Session HWND: {error}"))?;
    unsafe {
        let _ = ShowWindow(hwnd.0, SW_RESTORE);
        let _ = BringWindowToTop(hwnd.0);
        let _ = SetForegroundWindow(hwnd.0);
        if IsWindowVisible(hwnd.0) == 0 {
            return Err("VisualTeX Session HWND remained hidden after synchronous ShowWindow.".to_string());
        }
    }
    Ok(())
}

#[cfg(target_os = "windows")]
pub(super) fn force_office_editor_window_hidden(window: &WebviewWindow) -> Result<(), String> {
    let hwnd = window
        .hwnd()
        .map_err(|error| format!("Unable to resolve the VisualTeX Session HWND: {error}"))?;
    unsafe {
        let _ = ShowWindow(hwnd.0, SW_HIDE);
        if IsWindowVisible(hwnd.0) != 0 {
            return Err("VisualTeX Session HWND remained visible after synchronous hide.".to_string());
        }
    }
    Ok(())
}

#[cfg(target_os = "windows")]
pub(super) fn bring_session_window_to_front(
    window: &WebviewWindow,
    guarded_editor_session: Option<String>,
) -> Result<(), String> {
    window
        .show()
        .map_err(|error| format!("Unable to show the VisualTeX Session window: {error}"))?;
    let _ = window.unminimize();
    force_office_editor_window_visible(window)?;
    // Give the editor one immediate foreground attempt. Office can briefly
    // reactivate itself while a Ribbon/double-click callback unwinds, so keep
    // the editor top-most for a short interval and make at most one delayed
    // focus repair. Repeated show/focus cycles and user-attention requests can
    // visibly flash the editor or taskbar on some Windows/Office combinations.
    let _ = window.set_always_on_top(true);
    let _ = window.set_focus();

    let repair_window = window.clone();
    tauri::async_runtime::spawn(async move {
        tokio::time::sleep(std::time::Duration::from_millis(220)).await;
        if let Some(session_id) = guarded_editor_session.as_deref() {
            if !office_editor_session_is_active(session_id) {
                let _ = repair_window.set_always_on_top(false);
                return;
            }
        }
        if !repair_window.is_focused().unwrap_or(false) {
            let _ = repair_window.set_focus();
        }
        tokio::time::sleep(std::time::Duration::from_millis(120)).await;
        let _ = repair_window.set_always_on_top(false);
    });
    Ok(())
}

#[cfg(target_os = "windows")]
pub(super) async fn hide_owned_office_editor_on_main_thread(
    app: tauri::AppHandle,
    session_id: String,
) -> Result<bool, String> {
    let (sender, receiver) = tokio::sync::oneshot::channel();
    let scheduled_app = app.clone();
    app.run_on_main_thread(move || {
        let result = (|| -> Result<bool, String> {
            let _transition = OFFICE_EDITOR_WINDOW_TRANSITION_LOCK
                .get_or_init(|| Mutex::new(()))
                .lock()
                .map_err(|_| {
                    "VisualTeX Office editor window transition state is unavailable.".to_string()
                })?;
            let owns_editor = OFFICE_EDITOR_ACTIVE_SESSION
                .get_or_init(|| Mutex::new(None))
                .lock()
                .map_err(|_| {
                    "VisualTeX Office editor Session state is unavailable.".to_string()
                })
                .map(|mut current| {
                    if current.as_deref() != Some(session_id.as_str()) {
                        return false;
                    }
                    *current = None;
                    true
                })?;
            if owns_editor {
                if let Some(editor) =
                    scheduled_app.get_webview_window(OFFICE_EDITOR_WINDOW_LABEL)
                {
                    let _ = editor.set_always_on_top(false);
                    let _ = editor.request_user_attention(None);
                    let _ = editor.hide();
                    force_office_editor_window_hidden(&editor)?;
                    clear_office_editor_session(&editor);
                }
            }
            Ok(owns_editor)
        })();
        let _ = sender.send(result);
    })
    .map_err(|error| {
        format!("Unable to schedule Office editor cleanup on the UI thread: {error}")
    })?;
    receiver.await.map_err(|error| {
        format!("Office editor cleanup UI-thread task was cancelled: {error}")
    })?
}

#[cfg(target_os = "windows")]
pub(super) async fn hide_inactive_office_editor_on_main_thread(
    app: tauri::AppHandle,
) -> Result<bool, String> {
    let (sender, receiver) = tokio::sync::oneshot::channel();
    let scheduled_app = app.clone();
    app.run_on_main_thread(move || {
        let result = (|| -> Result<bool, String> {
            let _transition = OFFICE_EDITOR_WINDOW_TRANSITION_LOCK
                .get_or_init(|| Mutex::new(()))
                .lock()
                .map_err(|_| {
                    "VisualTeX Office editor window transition state is unavailable.".to_string()
                })?;
            let editor_is_inactive = OFFICE_EDITOR_ACTIVE_SESSION
                .get_or_init(|| Mutex::new(None))
                .lock()
                .map_err(|_| {
                    "VisualTeX Office editor Session state is unavailable.".to_string()
                })
                .map(|current| current.is_none())?;
            if editor_is_inactive {
                if let Some(editor) =
                    scheduled_app.get_webview_window(OFFICE_EDITOR_WINDOW_LABEL)
                {
                    let _ = editor.set_always_on_top(false);
                    let _ = editor.request_user_attention(None);
                    let _ = editor.hide();
                    force_office_editor_window_hidden(&editor)?;
                    clear_office_editor_session(&editor);
                }
            }
            Ok(editor_is_inactive)
        })();
        let _ = sender.send(result);
    })
    .map_err(|error| {
        format!("Unable to schedule inactive Office editor cleanup on the UI thread: {error}")
    })?;
    receiver.await.map_err(|error| {
        format!("Inactive Office editor cleanup UI-thread task was cancelled: {error}")
    })?
}

#[cfg(target_os = "windows")]
pub(super) fn prewarm_desktop_editor_window(app: tauri::AppHandle) -> Result<(), String> {
    let _transition = OFFICE_EDITOR_WINDOW_TRANSITION_LOCK
        .get_or_init(|| Mutex::new(()))
        .lock()
        .map_err(|_| "VisualTeX Office editor window transition state is unavailable.".to_string())?;
    if app.get_webview_window(OFFICE_EDITOR_WINDOW_LABEL).is_some() {
        return Ok(());
    }
    let active_session = OFFICE_EDITOR_ACTIVE_SESSION
        .get_or_init(|| Mutex::new(None))
        .lock()
        .map_err(|_| "VisualTeX Office editor Session state is unavailable.".to_string())?
        .clone();
    if active_session.is_some() {
        return Ok(());
    }

    let url = tauri::Url::parse(&format!(
        "https://127.0.0.1:{}/dialog?runtime=vsto-desktop",
        crate::office::state::office_port(),
    ))
    .map_err(|error| format!("Unable to construct the VisualTeX editor prewarm URL: {error}"))?;
    let (saved_editor_width, saved_editor_height) = load_office_editor_window_size(&app);
    OFFICE_EDITOR_PAGE_READY.store(false, std::sync::atomic::Ordering::Release);
    WebviewWindowBuilder::new(
        &app,
        OFFICE_EDITOR_WINDOW_LABEL,
        WebviewUrl::External(url),
    )
    .title("Office 公式编辑器")
    .inner_size(saved_editor_width, saved_editor_height)
    .min_inner_size(640.0, 480.0)
    .resizable(true)
    .center()
    .focused(false)
    .visible(false)
    .initialization_script(&office_install_token_script(&app))
    .on_page_load(handle_office_editor_page_load)
    .build()
    .map(|_| {
        crate::app_lifecycle::append_lifecycle_log(
            "created hidden Office editor WebView prewarm shell",
        );
    })
    .map_err(|error| format!("Unable to prewarm the VisualTeX Session window: {error}"))
}

#[cfg(target_os = "windows")]
pub(super) async fn prime_office_editor_native_presentation_on_main_thread(
    app: tauri::AppHandle,
) -> Result<(), String> {
    let (sender, receiver) = tokio::sync::oneshot::channel();
    let scheduled_app = app.clone();
    app.run_on_main_thread(move || {
        let result = (|| -> Result<(), String> {
            let _transition = OFFICE_EDITOR_WINDOW_TRANSITION_LOCK
                .get_or_init(|| Mutex::new(()))
                .lock()
                .map_err(|_| {
                    "VisualTeX Office editor window transition state is unavailable.".to_string()
                })?;
            let editor_is_inactive = OFFICE_EDITOR_ACTIVE_SESSION
                .get_or_init(|| Mutex::new(None))
                .lock()
                .map_err(|_| {
                    "VisualTeX Office editor Session state is unavailable.".to_string()
                })?
                .is_none();
            if !editor_is_inactive {
                return Ok(());
            }
            let window = scheduled_app
                .get_webview_window(OFFICE_EDITOR_WINDOW_LABEL)
                .ok_or_else(|| {
                    "VisualTeX Office editor prewarm window is unavailable.".to_string()
                })?;
            prime_office_editor_native_presentation(&window)
        })();
        let _ = sender.send(result);
    })
    .map_err(|error| {
        format!("Unable to schedule Office editor native prewarm on the UI thread: {error}")
    })?;
    receiver.await.map_err(|error| {
        format!("Office editor native prewarm UI-thread task was cancelled: {error}")
    })?
}

#[cfg(target_os = "windows")]
pub(super) async fn prewarm_desktop_editor_window_on_main_thread(
    app: tauri::AppHandle,
) -> Result<(), String> {
    let (sender, receiver) = tokio::sync::oneshot::channel();
    let scheduled_app = app.clone();
    app.run_on_main_thread(move || {
        let result = prewarm_desktop_editor_window(scheduled_app);
        let _ = sender.send(result);
    })
    .map_err(|error| {
        format!("Unable to schedule Office editor prewarm on the UI thread: {error}")
    })?;
    receiver.await.map_err(|error| {
        format!("Office editor prewarm UI-thread task was cancelled: {error}")
    })??;

    let deadline = std::time::Instant::now() + std::time::Duration::from_secs(5);
    while std::time::Instant::now() < deadline {
        if office_editor_page_is_ready() {
            return prime_office_editor_native_presentation_on_main_thread(app.clone()).await;
        }
        let editor_is_active = OFFICE_EDITOR_ACTIVE_SESSION
            .get_or_init(|| Mutex::new(None))
            .lock()
            .map(|current| current.is_some())
            .unwrap_or(false);
        if editor_is_active {
            return Ok(());
        }
        tokio::time::sleep(std::time::Duration::from_millis(20)).await;
    }
    Err("VisualTeX Office editor prewarm shell did not become ready in time.".to_string())
}

#[cfg(target_os = "windows")]
pub(super) fn open_desktop_session_window(app: tauri::AppHandle, session_id: String) -> Result<(), String> {
    let label = OFFICE_EDITOR_WINDOW_LABEL.to_string();
    let url = tauri::Url::parse(&format!(
        "https://127.0.0.1:{}/dialog/{}?runtime=vsto-desktop&sessionId={}",
        crate::office::state::office_port(),
        session_id,
        session_id
    ))
    .map_err(|error| format!("Unable to construct the VisualTeX Session URL: {error}"))?;
    let _transition = OFFICE_EDITOR_WINDOW_TRANSITION_LOCK
        .get_or_init(|| Mutex::new(()))
        .lock()
        .map_err(|_| "VisualTeX Office editor window transition state is unavailable.".to_string())?;

    if let Some(window) = app.get_webview_window(&label) {
        let active_session = OFFICE_EDITOR_ACTIVE_SESSION
            .get_or_init(|| Mutex::new(None))
            .lock()
            .map_err(|_| "VisualTeX Office editor Session state is unavailable.".to_string())?
            .clone();
        if active_session.as_deref() == Some(session_id.as_str()) {
            if !office_editor_page_is_ready() {
                crate::app_lifecycle::append_lifecycle_log(format!(
                    "active Office editor WebView is still loading for session={session_id}; reveal deferred"
                ));
                schedule_office_editor_reveal_fallback(app, session_id);
                return Ok(());
            }
            crate::app_lifecycle::append_lifecycle_log(format!(
                "reusing active Office editor WebView for session={session_id}"
            ));
            return bring_session_window_to_front(&window, Some(session_id));
        }

        let _ = window.set_always_on_top(false);
        let _ = window.request_user_attention(None);
        let _ = window.hide();
        if let Ok(mut current) = OFFICE_EDITOR_ACTIVE_SESSION
            .get_or_init(|| Mutex::new(None))
            .lock()
        {
            *current = Some(session_id.clone());
        }

        // Once the reusable dialog shell has finished loading, switch Sessions
        // in-place instead of navigating WebView2 again. This keeps React alive,
        // avoids a blank /dialog shell, and removes the load race that could
        // expose a window before the new Session id existed in the page.
        if office_editor_page_is_ready() {
            match switch_office_editor_session(&window, &session_id) {
                Ok(()) => {
                    crate::app_lifecycle::append_lifecycle_log(format!(
                        "switched loaded reused Office editor WebView to session={session_id}"
                    ));
                    return bring_session_window_to_front(&window, Some(session_id));
                }
                Err(error) => {
                    crate::app_lifecycle::append_lifecycle_log(format!(
                        "loaded Office editor in-place switch failed for session={session_id}; falling back to navigation: {error}"
                    ));
                }
            }
        }

        crate::app_lifecycle::append_lifecycle_log(format!(
            "navigating hidden reused Office editor WebView to session={session_id}"
        ));
        OFFICE_EDITOR_PAGE_READY.store(false, std::sync::atomic::Ordering::Release);
        if let Err(error) = window.navigate(url.clone()) {
            if let Ok(mut current) = OFFICE_EDITOR_ACTIVE_SESSION
                .get_or_init(|| Mutex::new(None))
                .lock()
            {
                if current.as_deref() == Some(session_id.as_str()) {
                    *current = None;
                }
            }
            return Err(format!(
                "Unable to navigate the VisualTeX Session window: {error}"
            ));
        }
        schedule_office_editor_reveal_fallback(app, session_id);
        return Ok(());
    }

    crate::app_lifecycle::append_lifecycle_log(format!(
        "creating hidden Office editor WebView on demand for session={session_id}"
    ));
    let (saved_editor_width, saved_editor_height) = load_office_editor_window_size(&app);
    if let Ok(mut current) = OFFICE_EDITOR_ACTIVE_SESSION
        .get_or_init(|| Mutex::new(None))
        .lock()
    {
        *current = Some(session_id.clone());
    }
    OFFICE_EDITOR_PAGE_READY.store(false, std::sync::atomic::Ordering::Release);
    let build_result = WebviewWindowBuilder::new(&app, label, WebviewUrl::External(url))
        .title("Office 公式编辑器")
        .inner_size(saved_editor_width, saved_editor_height)
        .min_inner_size(640.0, 480.0)
        .resizable(true)
        .center()
        .focused(false)
        .visible(false)
        .initialization_script(&office_install_token_script(&app))
        .on_page_load(handle_office_editor_page_load)
        .build();
    if let Err(error) = build_result {
        OFFICE_EDITOR_PAGE_READY.store(false, std::sync::atomic::Ordering::Release);
        if let Ok(mut current) = OFFICE_EDITOR_ACTIVE_SESSION
            .get_or_init(|| Mutex::new(None))
            .lock()
        {
            if current.as_deref() == Some(session_id.as_str()) {
                *current = None;
            }
        }
        return Err(format!(
            "Unable to create the VisualTeX Session window: {error}"
        ));
    }
    schedule_office_editor_reveal_fallback(app, session_id);
    Ok(())
}

#[cfg(target_os = "windows")]
pub(super) async fn open_desktop_session_window_on_main_thread(
    app: tauri::AppHandle,
    session_id: String,
) -> Result<(), String> {
    let (sender, receiver) = tokio::sync::oneshot::channel();
    let scheduled_app = app.clone();
    app.run_on_main_thread(move || {
        let result = open_desktop_session_window(scheduled_app, session_id);
        let _ = sender.send(result);
    })
    .map_err(|error| {
        format!("Unable to schedule Office editor open on the UI thread: {error}")
    })?;
    receiver.await.map_err(|error| {
        format!("Office editor open UI-thread task was cancelled: {error}")
    })?
}

#[cfg(target_os = "windows")]
pub(super) fn open_desktop_bulk_import_window(
    app: tauri::AppHandle,
    session_id: String,
) -> Result<(), String> {
    let label = format!("office-import-{}", session_id.replace('-', ""));
    let url = tauri::Url::parse(&format!(
        "https://127.0.0.1:{}/dialog/{}?runtime=vsto-bulk-import",
        crate::office::state::office_port(),
        session_id
    ))
    .map_err(|error| format!("Unable to construct the VisualTeX import URL: {error}"))?;

    if let Some(window) = app.get_webview_window(&label) {
        window
            .navigate(url)
            .map_err(|error| format!("Unable to navigate the VisualTeX import window: {error}"))?;
        return bring_session_window_to_front(&window, None);
    }

    let window = WebviewWindowBuilder::new(&app, label, WebviewUrl::External(url))
        .title("VisualTeX · Word 文档批量导入")
        .inner_size(1380.0, 880.0)
        .min_inner_size(900.0, 620.0)
        .resizable(true)
        .center()
        .focused(true)
        .visible(true)
        .initialization_script(&office_install_token_script(&app))
        .build()
        .map_err(|error| format!("Unable to create the VisualTeX import window: {error}"))?;
    bring_session_window_to_front(&window, None)
}

#[cfg(target_os = "windows")]
pub(super) fn ensure_desktop_conversion_window(app: &tauri::AppHandle) -> Result<WebviewWindow, String> {
    if let Some(window) = app.get_webview_window(OFFICE_CONVERTER_WINDOW_LABEL) {
        return Ok(window);
    }
    let url = tauri::Url::parse(&format!(
        "https://127.0.0.1:{}/dialog?runtime=vsto-convert",
        crate::office::state::office_port(),
    ))
    .map_err(|error| format!("Unable to construct the VisualTeX conversion URL: {error}"))?;
    WebviewWindowBuilder::new(
        app,
        OFFICE_CONVERTER_WINDOW_LABEL,
        WebviewUrl::External(url),
    )
    .title("VisualTeX · Office 格式转换")
    .inner_size(1240.0, 820.0)
    .resizable(false)
    .focused(false)
    .visible(false)
    .initialization_script(&office_install_token_script(app))
    .build()
    .map_err(|error| format!("Unable to create the reusable VisualTeX conversion window: {error}"))
}

#[cfg(target_os = "windows")]
pub(super) fn open_desktop_conversion_window(
    app: tauri::AppHandle,
    session_id: String,
) -> Result<(), String> {
    if let Some(window) = app.get_webview_window(OFFICE_CONVERTER_WINDOW_LABEL) {
        let encoded = serde_json::to_string(&session_id)
            .map_err(|error| format!("Unable to encode the conversion Session id: {error}"))?;
        let script = format!(
            "window.__VISUALTEX_OFFICE_SESSION_ID__={encoded};history.replaceState(null,'','/dialog?runtime=vsto-convert&sessionId='+encodeURIComponent({encoded}));window.dispatchEvent(new CustomEvent('visualtex-office-session',{{detail:{{sessionId:{encoded}}}}}));"
        );
        return window
            .eval(&script)
            .map_err(|error| format!("Unable to switch the VisualTeX conversion Session: {error}"));
    }

    let url = tauri::Url::parse(&format!(
        "https://127.0.0.1:{}/dialog?runtime=vsto-convert&sessionId={}",
        crate::office::state::office_port(),
        session_id,
    ))
    .map_err(|error| format!("Unable to construct the VisualTeX conversion URL: {error}"))?;
    WebviewWindowBuilder::new(
        &app,
        OFFICE_CONVERTER_WINDOW_LABEL,
        WebviewUrl::External(url),
    )
    .title("VisualTeX · Office 格式转换")
    .inner_size(1240.0, 820.0)
    .resizable(false)
    .focused(false)
    .visible(false)
    .initialization_script(&office_install_token_script(&app))
    .build()
    .map(|_| ())
    .map_err(|error| format!("Unable to create the reusable VisualTeX conversion window: {error}"))
}

#[cfg(target_os = "windows")]
pub(super) fn queue_desktop_batch_conversion(
    app: tauri::AppHandle,
    session_ids: Vec<String>,
) -> Result<(), String> {
    let window = ensure_desktop_conversion_window(&app)?;
    let mut queue = office_batch_conversion_queue()
        .lock()
        .map_err(|_| "The VisualTeX batch conversion queue is unavailable".to_string())?;
    for chunk in chunk_office_batch_conversion(session_ids) {
        queue.push_back(chunk);
    }
    drop(queue);
    // The queue is authoritative. This event only avoids waiting for the next
    // frontend poll and may safely be missed while the hidden WebView starts.
    let _ = window.eval(
        "window.dispatchEvent(new CustomEvent('visualtex-office-batch-conversion'));",
    );
    Ok(())
}

#[cfg(target_os = "windows")]
pub(super) fn prewarm_desktop_conversion_window(app: tauri::AppHandle) -> Result<(), String> {
    ensure_desktop_conversion_window(&app).map(|_| ())
}

pub(super) async fn close_desktop_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    if !valid_session_id(&session_id) {
        return StatusCode::NOT_FOUND.into_response();
    }

    #[cfg(not(target_os = "windows"))]
    {
        return StatusCode::NO_CONTENT.into_response();
    }

    #[cfg(target_os = "windows")]
    {
        // The desktop Office editor is reusable and is hidden rather than
        // destroyed after a formula finishes. Finalize every still-active
        // Session before hiding it: a committable auto-apply draft becomes
        // `committing`; an empty, incomplete, or explicitly non-auto-apply
        // draft becomes `cancelled`. Leaving `created`/`editing` here strands
        // the Word add-in's operation gate even though the native window is gone.
        let store = context.companion.session_store.clone();
        let closing_session_id = session_id.clone();
        if let Err(error) = run_session_operation(move || {
            let session = store.get(&closing_session_id)?;
            if let Some(patch) = close_patch_for_office_editor(&session) {
                store.patch(&closing_session_id, patch)?;
            }
            Ok(())
        })
        .await
        {
            return session_error_response(error);
        }

        let editor_was_active = OFFICE_EDITOR_WINDOW_TRANSITION_LOCK
            .get_or_init(|| Mutex::new(()))
            .lock()
            .ok()
            .and_then(|_transition| {
                OFFICE_EDITOR_ACTIVE_SESSION
                    .get_or_init(|| Mutex::new(None))
                    .lock()
                    .ok()
                    .map(|mut current| {
                        if current.as_deref() == Some(session_id.as_str()) {
                            *current = None;
                            true
                        } else {
                            false
                        }
                    })
            })
            .unwrap_or(false);
        let Some(app) = context.companion.app.clone() else {
            return desktop_app_unavailable();
        };
        let suffix = session_id.replace('-', "");
        let disposable_labels = [format!("office-import-{suffix}")];
        tauri::async_runtime::spawn(async move {
            // Return the HTTP response before closing the WebView that issued
            // the request; otherwise WebView2 can cancel the fetch and leave a
            // blank native window behind. The hide/clear itself must execute on
            // Tauri's main thread while holding the transition lock. Calling
            // WebviewWindow::hide/eval from this Tokio worker only queues UI
            // messages; releasing the lock before those messages execute lets a
            // later Session show first and then get hidden by this stale cleanup.
            tokio::time::sleep(std::time::Duration::from_millis(120)).await;
            if editor_was_active {
                let _ = hide_inactive_office_editor_on_main_thread(app.clone()).await;
            }
            for label in disposable_labels {
                if let Some(window) = app.get_webview_window(&label) {
                    let _ = window.close();
                }
            }
        });
        StatusCode::NO_CONTENT.into_response()
    }
}

pub(super) async fn open_desktop_session(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    if !valid_session_id(&session_id) {
        return StatusCode::NOT_FOUND.into_response();
    }
    let store = context.companion.session_store.clone();
    let lookup_id = session_id.clone();
    if let Err(error) = run_session_operation(move || store.get(&lookup_id)).await {
        return session_error_response(error);
    }

    #[cfg(not(target_os = "windows"))]
    {
        return (
            StatusCode::NOT_IMPLEMENTED,
            Json(serde_json::json!({
                "error": "Desktop VSTO Session windows are available only on Windows"
            })),
        )
            .into_response();
    }

    #[cfg(target_os = "windows")]
    {
        let Some(app) = context.companion.app.clone() else {
            return (
                StatusCode::SERVICE_UNAVAILABLE,
                Json(serde_json::json!({
                    "error": "The VisualTeX desktop application is unavailable"
                })),
            )
                .into_response();
        };
        match open_desktop_session_window_on_main_thread(app, session_id).await {
            Ok(()) => StatusCode::NO_CONTENT.into_response(),
            Err(error) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({ "error": error })),
            )
                .into_response(),
        }
    }
}

pub(super) async fn open_desktop_bulk_import(
    AxumPath(session_id): AxumPath<String>,
    State(context): State<ServerContext>,
) -> Response {
    if !valid_session_id(&session_id) {
        return StatusCode::NOT_FOUND.into_response();
    }
    let store = context.companion.session_store.clone();
    let lookup_id = session_id.clone();
    if let Err(error) = run_session_operation(move || store.get(&lookup_id)).await {
        return session_error_response(error);
    }

    #[cfg(not(target_os = "windows"))]
    {
        return (
            StatusCode::NOT_IMPLEMENTED,
            Json(serde_json::json!({
                "error": "Word document import is available only on Windows"
            })),
        )
            .into_response();
    }

    #[cfg(target_os = "windows")]
    {
        let Some(app) = context.companion.app.clone() else {
            return (
                StatusCode::SERVICE_UNAVAILABLE,
                Json(serde_json::json!({
                    "error": "The VisualTeX desktop application is unavailable"
                })),
            )
                .into_response();
        };
        match tokio::task::spawn_blocking(move || {
            open_desktop_bulk_import_window(app, session_id)
        })
        .await
        {
            Ok(Ok(())) => StatusCode::NO_CONTENT.into_response(),
            Ok(Err(error)) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({ "error": error })),
            )
                .into_response(),
            Err(error) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({
                    "error": format!("VisualTeX import window task failed: {error}")
                })),
            )
                .into_response(),
        }
    }
}

pub(super) async fn prewarm_desktop_editor(State(context): State<ServerContext>) -> Response {
    #[cfg(not(target_os = "windows"))]
    {
        return StatusCode::NO_CONTENT.into_response();
    }

    #[cfg(target_os = "windows")]
    {
        let Some(app) = context.companion.app.clone() else {
            return desktop_app_unavailable();
        };
        match prewarm_desktop_editor_window_on_main_thread(app).await {
            Ok(()) => StatusCode::NO_CONTENT.into_response(),
            Err(error) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({ "error": error })),
            )
                .into_response(),
        }
    }
}

/// The companion runs without the VisualTeX desktop app (it cannot open windows): 503 with the reason.
#[allow(dead_code)]
pub(super) fn desktop_app_unavailable() -> Response {
    (
        StatusCode::SERVICE_UNAVAILABLE,
        Json(serde_json::json!({
            "error": "The VisualTeX desktop app is not running, so no VisualTeX window (editor or conversion) can be opened"
        })),
    )
        .into_response()
}

pub(super) async fn prewarm_desktop_conversion(State(context): State<ServerContext>) -> Response {
    #[cfg(not(target_os = "windows"))]
    {
        return StatusCode::NO_CONTENT.into_response();
    }

    #[cfg(target_os = "windows")]
    {
        let Some(app) = context.companion.app.clone() else {
            return desktop_app_unavailable();
        };
        match tokio::task::spawn_blocking(move || prewarm_desktop_conversion_window(app)).await {
            Ok(Ok(())) => StatusCode::NO_CONTENT.into_response(),
            Ok(Err(error)) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({ "error": error })),
            )
                .into_response(),
            Err(error) => (
                StatusCode::INTERNAL_SERVER_ERROR,
                Json(serde_json::json!({
                    "error": format!("VisualTeX conversion prewarm task failed: {error}")
                })),
            )
                .into_response(),
        }
    }
}
