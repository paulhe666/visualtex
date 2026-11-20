//! The Office editor window: its URL and label, size per screen, the resident window kept parked between edits,
//! and the document import and LaTeX redraw windows.

use super::*;

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

pub(super) fn office_host_name(host: OfficeHost) -> &'static str {
    match host {
        OfficeHost::Word => "word",
        OfficeHost::Powerpoint => "powerpoint",
    }
}

pub(super) fn editor_window_label(host: OfficeHost) -> &'static str {
    match host {
        OfficeHost::Word => "office-native-word-editor",
        OfficeHost::Powerpoint => "office-native-powerpoint-editor",
    }
}

pub(super) fn editor_prewarmed_marker_path(app: &AppHandle, host: OfficeHost) -> Result<PathBuf, String> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX application data: {error}"))?;
    Ok(app_data
        .join("office")
        .join(format!("{}-editor-prewarmed.json", office_host_name(host))))
}

pub(super) fn editor_prewarm_diagnostic_path(app: &AppHandle, host: OfficeHost) -> Result<PathBuf, String> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX application data: {error}"))?;
    Ok(app_data
        .join("office")
        .join(format!("{}-editor-prewarm-diagnostic.json", office_host_name(host))))
}

pub(super) fn editor_window_host(label: &str) -> Option<OfficeHost> {
    match label {
        "office-native-word-editor" => Some(OfficeHost::Word),
        "office-native-powerpoint-editor" => Some(OfficeHost::Powerpoint),
        _ => None,
    }
}

pub(super) fn office_editor_monitor_geometry(monitor: &Monitor) -> Option<OfficeEditorMonitorGeometry> {
    let scale_factor = monitor.scale_factor();
    if !scale_factor.is_finite() || scale_factor <= 0.0 {
        return None;
    }
    let screen_width = monitor.size().width as f64 / scale_factor;
    let screen_height = monitor.size().height as f64 / scale_factor;
    let work_width = monitor.work_area().size.width as f64 / scale_factor;
    let work_height = monitor.work_area().size.height as f64 / scale_factor;
    if screen_width <= 0.0 || screen_height <= 0.0 {
        return None;
    }
    Some(OfficeEditorMonitorGeometry {
        screen_width,
        screen_height,
        maximum_inner_width: work_width.max(MIN_OFFICE_EDITOR_WIDTH),
        maximum_inner_height: (work_height - OFFICE_EDITOR_TITLE_BAR_ALLOWANCE)
            .max(MIN_OFFICE_EDITOR_HEIGHT),
    })
}

pub(super) fn primary_office_editor_monitor_geometry(
    app: &AppHandle,
) -> Option<OfficeEditorMonitorGeometry> {
    app.primary_monitor()
        .ok()
        .flatten()
        .as_ref()
        .and_then(office_editor_monitor_geometry)
}

pub(super) fn current_office_editor_monitor_geometry(
    app: &AppHandle,
    window: &WebviewWindow,
) -> Option<OfficeEditorMonitorGeometry> {
    window
        .current_monitor()
        .ok()
        .flatten()
        .as_ref()
        .and_then(office_editor_monitor_geometry)
        .or_else(|| primary_office_editor_monitor_geometry(app))
}

pub(super) fn normalize_office_editor_ratio(value: Option<f64>, fallback: f64) -> f64 {
    value
        .filter(|value| value.is_finite() && *value >= 0.1 && *value <= 1.0)
        .unwrap_or(fallback)
}

pub(super) fn normalize_office_editor_window_size(
    width: f64,
    height: f64,
    monitor: Option<OfficeEditorMonitorGeometry>,
) -> OfficeEditorWindowSize {
    let width = if width.is_finite() {
        width
    } else {
        DEFAULT_OFFICE_EDITOR_FALLBACK_WIDTH
    };
    let height = if height.is_finite() {
        height
    } else {
        DEFAULT_OFFICE_EDITOR_FALLBACK_HEIGHT
    };
    let maximum_width = monitor
        .map(|geometry| geometry.maximum_inner_width)
        .unwrap_or(MAX_OFFICE_EDITOR_WIDTH)
        .min(MAX_OFFICE_EDITOR_WIDTH)
        .max(MIN_OFFICE_EDITOR_WIDTH);
    let maximum_height = monitor
        .map(|geometry| geometry.maximum_inner_height)
        .unwrap_or(MAX_OFFICE_EDITOR_HEIGHT)
        .min(MAX_OFFICE_EDITOR_HEIGHT)
        .max(MIN_OFFICE_EDITOR_HEIGHT);
    OfficeEditorWindowSize {
        width: width.clamp(MIN_OFFICE_EDITOR_WIDTH, maximum_width),
        height: height.clamp(MIN_OFFICE_EDITOR_HEIGHT, maximum_height),
    }
}

pub(super) fn resolve_office_editor_window_size(
    preference: Option<OfficeEditorWindowSizePreference>,
    monitor: Option<OfficeEditorMonitorGeometry>,
) -> OfficeEditorWindowSize {
    let preference = preference.unwrap_or_default();
    let fallback_width = preference
        .fallback_width
        .or(preference.width)
        .filter(|value| value.is_finite() && *value > 0.0)
        .unwrap_or(DEFAULT_OFFICE_EDITOR_FALLBACK_WIDTH);
    let fallback_height = preference
        .fallback_height
        .or(preference.height)
        .filter(|value| value.is_finite() && *value > 0.0)
        .unwrap_or(DEFAULT_OFFICE_EDITOR_FALLBACK_HEIGHT);

    let (width, height) = if let Some(geometry) = monitor {
        let migrated_width_ratio = preference
            .width
            .filter(|value| value.is_finite() && *value > 0.0)
            .map(|value| value / geometry.screen_width);
        let migrated_height_ratio = preference
            .height
            .filter(|value| value.is_finite() && *value > 0.0)
            .map(|value| value / geometry.screen_height);
        let width_ratio = normalize_office_editor_ratio(
            preference.width_ratio.or(migrated_width_ratio),
            DEFAULT_OFFICE_EDITOR_WIDTH_RATIO,
        );
        let height_ratio = normalize_office_editor_ratio(
            preference.height_ratio.or(migrated_height_ratio),
            DEFAULT_OFFICE_EDITOR_HEIGHT_RATIO,
        );
        (
            geometry.screen_width * width_ratio,
            geometry.screen_height * height_ratio,
        )
    } else {
        (fallback_width, fallback_height)
    };
    normalize_office_editor_window_size(width, height, monitor)
}

pub(super) fn office_editor_window_size_path(app: &AppHandle) -> Result<PathBuf, String> {
    let app_data = app
        .path()
        .app_data_dir()
        .map_err(|error| format!("Unable to resolve VisualTeX application data: {error}"))?;
    Ok(app_data.join("office").join(OFFICE_EDITOR_WINDOW_SIZE_FILE))
}

pub(super) fn read_office_editor_window_size_preference(
    app: &AppHandle,
) -> Option<OfficeEditorWindowSizePreference> {
    let path = office_editor_window_size_path(app).ok()?;
    let bytes = fs::read(path).ok()?;
    serde_json::from_slice(&bytes).ok()
}

pub(super) fn load_office_editor_window_size(app: &AppHandle) -> OfficeEditorWindowSize {
    resolve_office_editor_window_size(
        read_office_editor_window_size_preference(app),
        primary_office_editor_monitor_geometry(app),
    )
}

pub(crate) fn configuration_office_editor_window_size(
    app: &AppHandle,
) -> Option<(f64, f64)> {
    let size = load_office_editor_window_size(app);
    Some((size.width, size.height))
}

pub(crate) fn apply_configuration_office_editor_window_size(
    app: &AppHandle,
    width: f64,
    height: f64,
) -> Result<(), String> {
    let monitor = primary_office_editor_monitor_geometry(app);
    let size = normalize_office_editor_window_size(width, height, monitor);
    let preference = OfficeEditorWindowSizePreference {
        width_ratio: monitor.map(|geometry| {
            (size.width / geometry.screen_width).clamp(0.1, 1.0)
        }),
        height_ratio: monitor.map(|geometry| {
            (size.height / geometry.screen_height).clamp(0.1, 1.0)
        }),
        fallback_width: Some(size.width),
        fallback_height: Some(size.height),
        width: None,
        height: None,
    };
    let path = office_editor_window_size_path(app)?;
    let bytes = serde_json::to_vec_pretty(&preference).map_err(|error| error.to_string())?;
    atomic_write(&path, &bytes, 0o600)?;

    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        if let Some(window) = app.get_webview_window(editor_window_label(host)) {
            window
                .set_size(tauri::LogicalSize::new(size.width, size.height))
                .map_err(|error| error.to_string())?;
        }
    }
    Ok(())
}

pub(crate) fn schedule_persist_office_editor_window_size(
    app: &AppHandle,
    label: &str,
    physical_width: u32,
    physical_height: u32,
) {
    if editor_window_host(label).is_none() || physical_width == 0 || physical_height == 0 {
        return;
    }
    let Some(window) = app.get_webview_window(label) else {
        return;
    };
    let scale_factor = window.scale_factor().unwrap_or(1.0).max(0.1);
    let fallback_width = physical_width as f64 / scale_factor;
    let fallback_height = physical_height as f64 / scale_factor;
    let monitor = current_office_editor_monitor_geometry(app, &window);
    let preference = OfficeEditorWindowSizePreference {
        width_ratio: monitor.map(|geometry| {
            (fallback_width / geometry.screen_width).clamp(0.1, 1.0)
        }),
        height_ratio: monitor.map(|geometry| {
            (fallback_height / geometry.screen_height).clamp(0.1, 1.0)
        }),
        fallback_width: Some(fallback_width),
        fallback_height: Some(fallback_height),
        width: None,
        height: None,
    };
    let generation = OFFICE_EDITOR_SIZE_WRITE_GENERATION.fetch_add(1, Ordering::SeqCst) + 1;
    let app = app.clone();
    crate::spawn_background_task("visualtex-office-window-size", move || {
        std::thread::sleep(Duration::from_millis(250));
        if OFFICE_EDITOR_SIZE_WRITE_GENERATION.load(Ordering::SeqCst) != generation {
            return;
        }
        let Ok(path) = office_editor_window_size_path(&app) else {
            return;
        };
        let Ok(bytes) = serde_json::to_vec_pretty(&preference) else {
            return;
        };
        if let Err(error) = atomic_write(&path, &bytes, 0o600) {
            eprintln!("Unable to persist the Office editor window size: {error}");
        }
    });
}

pub(super) fn document_import_window_label(session_id: &str) -> String {
    format!("office-native-document-{}", session_id.replace('-', ""))
}

pub(super) fn active_editor_session(host: OfficeHost) -> Option<ActiveOfficeEditorSession> {
    office_editor_runtime()
        .lock()
        .ok()
        .and_then(|runtime| runtime.active(host).cloned())
}

pub(super) fn clear_editor_session(host: OfficeHost, session_id: &str, generation: u64) -> Result<(), String> {
    let mut runtime = office_editor_runtime()
        .lock()
        .map_err(|_| "VisualTeX Office editor state is unavailable".to_string())?;
    let matches = runtime.active(host).is_some_and(|active| {
        active.activation.session_id == session_id && active.activation.generation == generation
    });
    if !matches {
        return Err("The Office editor Session is no longer active".to_string());
    }
    *runtime.active_mut(host) = None;
    Ok(())
}

pub(super) fn clear_any_editor_session(host: OfficeHost) -> Option<MacOfflineOfficeEditorActivation> {
    let mut runtime = office_editor_runtime().lock().ok()?;
    runtime
        .active_mut(host)
        .take()
        .map(|active| active.activation)
}

#[cfg(target_os = "macos")]
pub(super) fn configure_resident_editor_space_behavior(window: &WebviewWindow) -> Result<(), String> {
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            let behavior = native_window.collectionBehavior()
                | NSWindowCollectionBehavior::CanJoinAllSpaces
                | NSWindowCollectionBehavior::CanJoinAllApplications
                | NSWindowCollectionBehavior::FullScreenAuxiliary;
            native_window.setCollectionBehavior(behavior);
            native_window.setHidesOnDeactivate(false);
        })
        .map_err(|error| {
            format!("Unable to configure the Office editor fullscreen behavior: {error}")
        })
}

#[cfg(target_os = "macos")]
pub(super) fn center_resident_editor_on_pointer_screen(window: &WebviewWindow) -> Result<(), String> {
    let (sender, receiver) = mpsc::sync_channel(1);
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            let positioned = MainThreadMarker::new().is_some_and(|main_thread| {
                let pointer = NSEvent::mouseLocation();
                let screens = NSScreen::screens(main_thread);
                let target = (0..screens.count()).find_map(|index| {
                    let screen = screens.objectAtIndex(index);
                    let frame = screen.frame();
                    let contains_pointer =
                        pointer.x >= frame.origin.x
                            && pointer.x < frame.origin.x + frame.size.width
                            && pointer.y >= frame.origin.y
                            && pointer.y < frame.origin.y + frame.size.height;
                    contains_pointer.then(|| screen.visibleFrame())
                });
                let Some(screen_frame) = target else {
                    return false;
                };
                let window_frame = native_window.frame();
                let x = screen_frame.origin.x
                    + (screen_frame.size.width - window_frame.size.width) / 2.0;
                let y = screen_frame.origin.y
                    + (screen_frame.size.height - window_frame.size.height) / 2.0;
                native_window.setFrameOrigin(objc2_foundation::NSPoint::new(x, y));
                true
            });
            let _ = sender.send(positioned);
        })
        .map_err(|error| {
            format!("Unable to position the Office editor on the active display: {error}")
        })?;
    let _ = receiver.recv_timeout(Duration::from_millis(250));
    Ok(())
}

/// The resident editor stays ordered in at a near-zero alpha so WebKit keeps
/// it running, which made Mission Control show it as an empty "VisualTeX
/// Office Formula" tile. While it is not presented it is Transient (hidden by
/// Mission Control and App Exposé) and left out of Cmd-` cycling; presenting
/// it restores the default managed behaviour.
#[cfg(target_os = "macos")]
pub(super) fn set_resident_editor_listed(native_window: &objc2_app_kit::NSWindow, listed: bool) {
    let unlisted = NSWindowCollectionBehavior::Transient | NSWindowCollectionBehavior::IgnoresCycle;
    let mut behavior = native_window.collectionBehavior().difference(
        unlisted | NSWindowCollectionBehavior::Managed | NSWindowCollectionBehavior::Stationary,
    );
    if !listed {
        behavior |= unlisted;
    }
    native_window.setCollectionBehavior(behavior);
}

#[cfg(target_os = "macos")]
pub(super) fn set_resident_editor_native_state(
    window: &WebviewWindow,
    alpha: f64,
    parked: bool,
) -> Result<(), String> {
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            set_resident_editor_listed(native_window, !parked);
            native_window.setAlphaValue(alpha);
            native_window.setIgnoresMouseEvents(parked);
            native_window.setHidesOnDeactivate(false);
            // Parked renderers are implementation windows, not destinations
            // for Window/Dock navigation while the user opens the workspace.
            native_window.setExcludedFromWindowsMenu(parked);
            native_window.setLevel(if parked {
                objc2_app_kit::NSNormalWindowLevel
            } else {
                objc2_app_kit::NSScreenSaverWindowLevel
            });
        })
        .map_err(|error| format!("Unable to update the resident Office editor window: {error}"))
}

#[cfg(target_os = "macos")]
pub(super) fn set_resident_editor_parked(window: &WebviewWindow, parked: bool) -> Result<(), String> {
    // 0.001 keeps an already-hydrated resident visually imperceptible, but on
    // current macOS/WebKit it is low enough to suspend a freshly mounting
    // WKWebView and a silent conversion. Use the stronger wake state below for
    // those active phases, then return to 0.001 when idle.
    set_resident_editor_native_state(window, if parked { 0.001 } else { 1.0 }, parked)
}

#[cfg(target_os = "macos")]
pub(super) fn wake_resident_editor_invisibly(window: &WebviewWindow) -> Result<(), String> {
    // Keep the editor mouse-inert and at normal window level, but raise native
    // alpha just enough that WebKit continues timers/rendering. Formula content
    // stays CSS-hidden during silent conversions, so this does not expose the
    // editor while preventing the 0.001-alpha suspension observed on macOS 26.
    set_resident_editor_native_state(window, 0.01, true)
}

#[cfg(target_os = "macos")]
pub(super) fn order_resident_editor_behind(window: &WebviewWindow) -> Result<(), String> {
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            // AppKit's orderBack does not resign a key NSPanel. Release it
            // first, then keep the parked renderer ordered for WebKit. This
            // lets the same panel become key again on the next Office edit.
            if native_window.isKeyWindow() {
                native_window.orderOut(None);
            }
            native_window.orderBack(None);
        })
        .map_err(|error| format!("Unable to order the resident Office editor behind: {error}"))
}

#[cfg(not(target_os = "macos"))]
pub(super) fn set_resident_editor_parked(window: &WebviewWindow, parked: bool) -> Result<(), String> {
    if parked {
        window.hide()
    } else {
        window.show()
    }
    .map_err(|error| error.to_string())
}

#[cfg(target_os = "macos")]
pub(super) fn wake_resident_editor_for_hydration(window: &WebviewWindow) -> Result<(), String> {
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            set_resident_editor_listed(native_window, false);
            native_window.setAlphaValue(1.0);
            native_window.setIgnoresMouseEvents(true);
            native_window.setLevel(objc2_app_kit::NSNormalWindowLevel);
        })
        .map_err(|error| format!("Unable to wake the resident Office editor window: {error}"))
}

#[cfg(not(target_os = "macos"))]
pub(super) fn wake_resident_editor_for_hydration(window: &WebviewWindow) -> Result<(), String> {
    window.show().map_err(|error| error.to_string())
}

#[cfg(target_os = "macos")]
pub(super) fn order_main_window_behind_office_editor(app: &AppHandle) -> Result<(), String> {
    let Some(main_window) = app.get_webview_window("main") else {
        return Ok(());
    };
    if !main_window.is_visible().unwrap_or(false) {
        return Ok(());
    }
    main_window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            // Office editor activation targets only the formula window. Keep an
            // already visible desktop workspace explicitly behind Office as an
            // additional guard against LaunchServices or stale reopen events.
            native_window.setLevel(objc2_app_kit::NSNormalWindowLevel - 1);
            native_window.orderBack(None);
        })
        .map_err(|error| format!("Unable to keep the VisualTeX main window behind Office: {error}"))
}

#[cfg(not(target_os = "macos"))]
pub(super) fn order_main_window_behind_office_editor(_app: &AppHandle) -> Result<(), String> {
    Ok(())
}

#[cfg(target_os = "macos")]
pub(super) fn inspect_resident_editor_focus(
    window: &WebviewWindow,
) -> Result<ResidentEditorFocusState, String> {
    let (sender, receiver) = mpsc::sync_channel(1);
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            let state = MainThreadMarker::new()
                .map(|main_thread| {
                    let application = NSApplication::sharedApplication(main_thread);
                    ResidentEditorFocusState {
                        app_active: application.isActive(),
                        window_can_become_key: native_window.canBecomeKeyWindow(),
                        window_is_key: native_window.isKeyWindow(),
                        window_is_main: native_window.isMainWindow(),
                        window_is_visible: native_window.isVisible(),
                        window_on_active_space: native_window.isOnActiveSpace(),
                    }
                })
                .unwrap_or_default();
            let _ = sender.send(state);
        })
        .map_err(|error| format!("Unable to inspect the Office editor window: {error}"))?;
    receiver
        .recv_timeout(Duration::from_millis(250))
        .map_err(|error| format!("Timed out inspecting the Office editor window: {error}"))
}

#[cfg(target_os = "macos")]
pub(super) fn order_resident_editor_to_front(window: &WebviewWindow) -> Result<(), String> {
    let (sender, receiver) = mpsc::sync_channel(1);
    window
        .with_webview(move |webview| unsafe {
            let native_window: &objc2_app_kit::NSWindow = &*webview.ns_window().cast();
            // The Tao window is created as an NSPanel for Office formula
            // editors. Keep its WKWebView in place for the whole lifetime.
            set_resident_editor_listed(native_window, true);
            native_window.setAlphaValue(1.0);
            native_window.setIgnoresMouseEvents(false);
            native_window.setExcludedFromWindowsMenu(false);
            native_window.setLevel(objc2_app_kit::NSScreenSaverWindowLevel);
            native_window.orderFrontRegardless();
            native_window.makeKeyWindow();
            let webview_responder: &objc2_app_kit::NSResponder = &*webview.inner().cast();
            native_window.makeFirstResponder(Some(webview_responder));
            let _ = sender.send(());
        })
        .map_err(|error| format!("Unable to order the Office editor window: {error}"))?;
    receiver
        .recv_timeout(Duration::from_millis(250))
        .map_err(|error| format!("Timed out ordering the Office editor window: {error}"))
}

#[cfg(not(target_os = "macos"))]
pub(super) fn inspect_resident_editor_focus(
    window: &WebviewWindow,
) -> Result<ResidentEditorFocusState, String> {
    Ok(ResidentEditorFocusState {
        app_active: window.is_focused().unwrap_or(false),
        window_can_become_key: true,
        window_is_key: window.is_focused().unwrap_or(false),
        window_is_main: window.is_focused().unwrap_or(false),
        window_is_visible: window.is_visible().unwrap_or(false),
        window_on_active_space: true,
    })
}

#[cfg(target_os = "macos")]
pub(super) fn present_resident_editor_window(_app: &AppHandle, window: &WebviewWindow) -> Result<(), String> {
    configure_resident_editor_space_behavior(window)?;
    center_resident_editor_on_pointer_screen(window)?;
    order_resident_editor_to_front(window)
}

#[cfg(not(target_os = "macos"))]
pub(super) fn present_resident_editor_window(_app: &AppHandle, window: &WebviewWindow) -> Result<(), String> {
    window.show().map_err(|error| error.to_string())?;
    window.set_focus().map_err(|error| error.to_string())
}

pub(super) fn set_resident_editor_content_visible(
    window: &WebviewWindow,
    visible: bool,
) -> Result<(), String> {
    let opacity = if visible { "1" } else { "0" };
    window
        .eval(format!(
            "if (document.body) {{ document.body.style.opacity = '{opacity}'; }}"
        ))
        .map_err(|error| format!("Unable to update resident Office editor content: {error}"))
}

pub(super) fn create_editor_window(app: &AppHandle, host: OfficeHost) -> Result<WebviewWindow, String> {
    let label = editor_window_label(host);
    if let Some(window) = app.get_webview_window(label) {
        #[cfg(target_os = "macos")]
        configure_resident_editor_space_behavior(&window)?;
        return Ok(window);
    }

    let theme = crate::persisted_app_theme(app);
    let saved_size = load_office_editor_window_size(app);
    let path = format!(
        "office-native-dialog.html?transport=tauri&officeHost={}&theme={theme}",
        office_host_name(host),
    );
    let window = WebviewWindowBuilder::new(app, label, WebviewUrl::App(path.into()))
        .title("VisualTeX Office Formula")
        .inner_size(saved_size.width, saved_size.height)
        .min_inner_size(MIN_OFFICE_EDITOR_WIDTH, MIN_OFFICE_EDITOR_HEIGHT)
        .focused(false)
        .skip_taskbar(true)
        .visible(false)
        .background_throttling(BackgroundThrottlingPolicy::Disabled)
        .build()
        .map_err(|error| format!("Unable to initialize the VisualTeX Office editor: {error}"))?;
    #[cfg(target_os = "macos")]
    configure_resident_editor_space_behavior(&window)?;
    // A freshly created WKWebView must stay above WebKit's suspension threshold
    // until React/MathLive report prewarm readiness. It is parked back at 0.001
    // immediately after that handshake.
    wake_resident_editor_invisibly(&window)?;
    #[cfg(target_os = "macos")]
    order_resident_editor_behind(&window)?;
    #[cfg(not(target_os = "macos"))]
    window
        .show()
        .map_err(|error| format!("Unable to prewarm the VisualTeX Office editor: {error}"))?;
    Ok(window)
}

pub(crate) fn prewarm_office_editor_windows(app: &AppHandle) -> Result<(), String> {
    // Both WebKit processes start once with the resident app. Formula
    // interactions only switch a Session in an already initialized WebView.
    crate::office::background::install_application_icon(app)?;
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        create_editor_window(app, host)?;
    }
    Ok(())
}

pub(super) fn open_editor_window(
    app: &AppHandle,
    host: OfficeHost,
    session_id: &str,
    received_epoch_ms: u64,
    received_at: Instant,
    silent: bool,
) -> Result<(), String> {
    let label = editor_window_label(host);
    let reused = app.get_webview_window(label).is_some();
    let window = create_editor_window(app, host)?;
    set_resident_editor_content_visible(&window, false)?;
    let activation = {
        let mut runtime = office_editor_runtime()
            .lock()
            .map_err(|_| "VisualTeX Office editor state is unavailable".to_string())?;
        // Silent conversions need WebKit timers while remaining invisible.
        // Ordinary edits can stay at the idle parked alpha until the explicit
        // full-opacity hydration wake below.
        if silent {
            wake_resident_editor_invisibly(&window)?;
        } else {
            set_resident_editor_parked(&window, true)?;
        }
        runtime.next_generation = runtime.next_generation.saturating_add(1).max(1);
        let activation = MacOfflineOfficeEditorActivation {
            session_id: session_id.to_string(),
            host,
            generation: runtime.next_generation,
            received_epoch_ms,
        };
        *runtime.active_mut(host) = Some(ActiveOfficeEditorSession {
            activation: activation.clone(),
            received_at,
            ready: false,
        });
        activation
    };
    let elapsed_ms = received_at.elapsed().as_secs_f64() * 1000.0;
    queue_editor_performance(
        host,
        session_id,
        if reused {
            "window-reused"
        } else {
            "window-created"
        },
        elapsed_ms,
        Some(activation.generation),
        json!({ "windowLabel": label }),
    );
    if let Err(error) = window.emit(OFFICE_EDITOR_ACTIVATE_EVENT, activation.clone()) {
        let _ = clear_editor_session(host, session_id, activation.generation);
        return Err(format!("Unable to activate the VisualTeX Office editor: {error}"));
    }
    queue_editor_performance(
        host,
        session_id,
        "activation-event-sent",
        received_at.elapsed().as_secs_f64() * 1000.0,
        Some(activation.generation),
        json!({}),
    );
    // Ordinary editing hydrates at full native alpha before foreground
    // promotion. A direct native-to-image conversion uses the same resident
    // renderer, but remains parked and mouse-inert for the whole automatic
    // commit so Word never flashes the formula editor.
    if !silent {
        wake_resident_editor_for_hydration(&window)?;
        #[cfg(target_os = "macos")]
        {
            center_resident_editor_on_pointer_screen(&window)?;
            order_resident_editor_behind(&window)?;
        }
        #[cfg(not(target_os = "macos"))]
        {
            window.show().map_err(|error| error.to_string())?;
            window.unminimize().map_err(|error| error.to_string())?;
        }
    }
    Ok(())
}

pub(super) fn set_word_document_operation_preparing_status(operation: &str) -> Result<(), String> {
    let status = if operation == "latexRedraw" {
        "VisualTeX is rendering LaTeX formulas..."
    } else if operation == "formulaRestore" {
        "VisualTeX is restoring Word formulas..."
    } else {
        "VisualTeX is preparing the Word document import..."
    };
    let script = format!(
        "tell application \"Microsoft Word\" to set status bar to {:?}",
        status
    );
    let output = Command::new("/usr/bin/osascript")
        .args(["-e", "tell application \"Microsoft Word\" to activate", "-e", &script])
        .output()
        .map_err(|error| format!("Unable to activate Microsoft Word: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
        Err(if detail.is_empty() {
            "Unable to activate Microsoft Word".to_string()
        } else {
            format!("Unable to activate Microsoft Word: {detail}")
        })
    }
}

pub(super) fn clear_word_document_import_status() {
    let _ = Command::new("/usr/bin/osascript")
        .args([
            "-e",
            "tell application \"Microsoft Word\" to set status bar to \"\"",
        ])
        .output();
}

pub(super) fn open_document_import_window(app: &AppHandle, session_id: &str, office: &str) -> Result<(), String> {
    crate::office::background::activate_foreground_app(app)?;
    crate::office::background::install_application_icon(app)?;

    let label = document_import_window_label(session_id);
    if let Some(window) = app.get_webview_window(&label) {
        crate::office::background::activate_foreground_app(app)?;
        window.show().map_err(|error| error.to_string())?;
        window.unminimize().map_err(|error| error.to_string())?;
        crate::office::background::activate_foreground_app(app)?;
        window.set_focus().map_err(|error| error.to_string())?;
        order_main_window_behind_office_editor(app)?;
        return Ok(());
    }
    let theme = crate::persisted_app_theme(app);
    let path = format!(
        "index.html?view=office-document-import&sessionId={session_id}&transport=tauri&theme={theme}"
    );
    let window = WebviewWindowBuilder::new(app, label, WebviewUrl::App(path.into()))
        .title(format!("VisualTeX {office} 文档批量导入"))
        .inner_size(1260.0, 840.0)
        .min_inner_size(860.0, 620.0)
        .center()
        .build()
        .map_err(|error| format!("Unable to open the VisualTeX document importer: {error}"))?;
    window.show().map_err(|error| error.to_string())?;
    window.unminimize().map_err(|error| error.to_string())?;
    crate::office::background::activate_foreground_app(app)?;
    window.set_focus().map_err(|error| error.to_string())?;
    order_main_window_behind_office_editor(app)?;
    Ok(())
}

pub(super) fn open_word_latex_redraw_window(app: &AppHandle, session_id: &str) -> Result<(), String> {
    crate::office::background::install_application_icon(app)?;
    let label = document_import_window_label(session_id);
    if let Some(window) = app.get_webview_window(&label) {
        window.hide().map_err(|error| error.to_string())?;
        return Ok(());
    }
    let theme = crate::persisted_app_theme(app);
    let path = format!(
        "index.html?view=office-word-latex-redraw&sessionId={session_id}&transport=tauri&theme={theme}"
    );
    WebviewWindowBuilder::new(app, label, WebviewUrl::App(path.into()))
        .title("VisualTeX Word LaTeX Redraw")
        .inner_size(560.0, 300.0)
        .min_inner_size(420.0, 220.0)
        .center()
        .focused(false)
        .skip_taskbar(true)
        .visible(false)
        .background_throttling(BackgroundThrottlingPolicy::Disabled)
        .build()
        .map_err(|error| format!("Unable to start the VisualTeX Word redraw renderer: {error}"))?;
    Ok(())
}

pub(crate) fn has_open_office_editor(app: &AppHandle) -> bool {
    [OfficeHost::Word, OfficeHost::Powerpoint]
        .into_iter()
        .any(|host| {
            active_editor_session(host).is_some()
                && app.get_webview_window(editor_window_label(host)).is_some()
        })
}

pub(super) fn has_recent_office_editor_request_in_roots(
    now: SystemTime,
    max_age: Duration,
    roots: impl IntoIterator<Item = PathBuf>,
) -> bool {
    roots.into_iter().any(|root| {
        let Ok(entries) = fs::read_dir(&root) else {
            return false;
        };
        entries.flatten().any(|entry| {
            let request = entry.path().join(REQUEST_FILE);
            let Ok(modified) = fs::metadata(request).and_then(|metadata| metadata.modified())
            else {
                return false;
            };
            now.duration_since(modified)
                .map(|age| age <= max_age)
                .unwrap_or(true)
        })
    })
}

pub(crate) fn has_recent_office_editor_request(max_age: Duration) -> bool {
    has_recent_office_editor_request_in_roots(
        SystemTime::now(),
        max_age,
        [OfficeHost::Word, OfficeHost::Powerpoint]
            .into_iter()
            .filter_map(|host| sessions_root(host).ok()),
    )
}

pub(crate) fn focus_open_office_editor(app: &AppHandle) -> bool {
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        if let Some(active) = active_editor_session(host) {
            let Some(window) = app.get_webview_window(editor_window_label(host)) else {
                continue;
            };
            // A parked active window is still hydrating. Treat it as owned so
            // a second native double-click route cannot launch a duplicate.
            // Never touch the desktop main window during formula activation.
            if !active.ready {
                return true;
            }
            let _ = present_resident_editor_window(app, &window);
            return true;
        }
    }
    false
}
