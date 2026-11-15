use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use serde::{Deserialize, Serialize};
use std::fs;
use std::path::PathBuf;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Arc;
use tauri::{AppHandle, Manager};

mod anywhere;
mod app_lifecycle;
mod ocr;
#[path = "../../../shared/visualtex-backend/src/ocr_offline.rs"]
mod ocr_offline;
#[cfg(windows)]
mod system_math_glyphs;

#[cfg(all(windows, feature = "system-font-asset-acceptance"))]
pub fn system_font_asset_acceptance_json(
    font_families: Vec<String>,
) -> Result<String, String> {
    serde_json::to_string(&system_math_glyphs::read_system_font_asset(font_families)?)
        .map_err(|error| format!("VisualTeX could not serialize the system font asset: {error}"))
}
mod document_export;
#[path = "../../../shared/visualtex-backend/src/document_library.rs"]
mod document_library;
mod office;
#[cfg(windows)]
mod windows_quick_ocr;
#[cfg(windows)]
mod windows_silent_ocr_hotkey;



#[derive(Debug, Clone, Copy, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ConfigurationWindowSize {
    width: f64,
    height: f64,
}

#[derive(Debug, Clone, Default, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct AppWindowConfiguration {
    main: Option<ConfigurationWindowSize>,
    keypad: Option<ConfigurationWindowSize>,
    office_editor: Option<ConfigurationWindowSize>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
struct WordNumberingUserConfiguration {
    default_display_equation_numbered: bool,
    default_equation_number_format: String,
}

#[tauri::command]
fn get_word_numbering_user_configuration() -> Result<WordNumberingUserConfiguration, String> {
    let (default_display_equation_numbered, default_equation_number_format) =
        office::windows_backend::word_numbering_user_preferences();
    Ok(WordNumberingUserConfiguration {
        default_display_equation_numbered,
        default_equation_number_format,
    })
}

#[tauri::command]
fn apply_word_numbering_user_configuration(
    configuration: WordNumberingUserConfiguration,
) -> Result<WordNumberingUserConfiguration, String> {
    office::windows_backend::set_word_numbering_user_preferences(
        configuration.default_display_equation_numbered,
        &configuration.default_equation_number_format,
    )?;
    get_word_numbering_user_configuration()
}

#[tauri::command]
fn get_app_window_configuration(app: AppHandle) -> Result<AppWindowConfiguration, String> {
    let main = app_lifecycle::configuration_main_window_size(&app)
        .map(|(width, height)| ConfigurationWindowSize { width, height });
    let keypad = app_lifecycle::configuration_keypad_window_size(&app)
        .map(|(width, height)| ConfigurationWindowSize { width, height });
    #[cfg(target_os = "windows")]
    let office_editor = office::server::configuration_office_editor_window_size(&app)
        .map(|(width, height)| ConfigurationWindowSize { width, height });
    #[cfg(not(target_os = "windows"))]
    let office_editor = None;

    Ok(AppWindowConfiguration {
        main,
        keypad,
        office_editor,
    })
}

#[tauri::command]
fn apply_app_window_configuration(
    app: AppHandle,
    configuration: AppWindowConfiguration,
) -> Result<AppWindowConfiguration, String> {
    if let Some(requested) = configuration.main {
        app_lifecycle::apply_configuration_main_window_size(
            &app,
            requested.width,
            requested.height,
        )?;
    }
    if let Some(requested) = configuration.keypad {
        app_lifecycle::apply_configuration_keypad_window_size(
            &app,
            requested.width,
            requested.height,
        )?;
    }
    #[cfg(target_os = "windows")]
    if let Some(requested) = configuration.office_editor {
        office::server::apply_configuration_office_editor_window_size(
            &app,
            requested.width,
            requested.height,
        )?;
    }
    get_app_window_configuration(app)
}

#[tauri::command]
fn set_main_window_keypad_mode(app: AppHandle, enabled: bool) -> Result<(), String> {
    app_lifecycle::set_main_window_keypad_mode(&app, enabled)
}

#[derive(Debug, Deserialize)]
#[serde(rename_all = "camelCase")]
struct ExportFileRequest {
    path: String,
    text: Option<String>,
    base64: Option<String>,
}

fn write_export_file_request(request: ExportFileRequest) -> Result<(), String> {
    let trimmed_path = request.path.trim();
    if trimmed_path.is_empty() {
        return Err("Export path is empty".to_string());
    }

    let path = PathBuf::from(trimmed_path);
    if path.file_name().is_none() {
        return Err("Export path must include a file name".to_string());
    }
    if let Some(parent) = path.parent().filter(|parent| !parent.as_os_str().is_empty()) {
        fs::create_dir_all(parent)
            .map_err(|error| format!("Unable to create export directory: {error}"))?;
    }

    let bytes = match (request.text, request.base64) {
        (Some(text), None) => text.into_bytes(),
        (None, Some(encoded)) => BASE64_STANDARD
            .decode(encoded.trim())
            .map_err(|error| format!("Unable to decode export data: {error}"))?,
        (Some(_), Some(_)) => {
            return Err("Export request must contain text or base64 data, not both".to_string())
        }
        (None, None) => return Err("Export request contains no data".to_string()),
    };

    fs::write(&path, bytes)
        .map_err(|error| format!("Unable to write export file '{}': {error}", path.display()))
}

#[tauri::command]
fn write_export_file(request: ExportFileRequest) -> Result<(), String> {
    write_export_file_request(request)
}

fn shutdown_runtime(app: &AppHandle, started: &AtomicBool, reason: &str) {
    if started.swap(true, Ordering::SeqCst) {
        return;
    }
    app_lifecycle::append_lifecycle_log(format!("runtime shutdown begin reason={reason}"));

    if let Some(state) = app.try_state::<office::state::OfficeCompanionState>() {
        if let Err(error) = office::server::stop(state.inner()) {
            app_lifecycle::append_lifecycle_log(format!(
                "Unable to stop Office companion server during shutdown: {error}"
            ));
        }
        if office::state::isolated_acceptance_runtime() {
            // No global Office sidecar or OCR runtime was started by this mode.
            app_lifecycle::append_lifecycle_log("isolated Office acceptance shutdown completed");
            return;
        }
        if let Err(error) = state.platform_backend.shutdown() {
            app_lifecycle::append_lifecycle_log(format!(
                "Unable to stop VisualTeX Office platform backend: {error}"
            ));
        }
    }

    if let Some(state) = app.try_state::<ocr::OcrState>() {
        if let Err(error) = state.shutdown(app) {
            app_lifecycle::append_lifecycle_log(format!(
                "Unable to stop OCR worker and child processes: {error}"
            ));
        }
    }

    #[cfg(target_os = "macos")]
    if let Err(error) = office::background::pause_launch_agent_for_quit() {
        app_lifecycle::append_lifecycle_log(format!(
            "Unable to pause VisualTeX Office background service: {error}"
        ));
    }
    app_lifecycle::append_lifecycle_log("runtime shutdown completed");
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let run_mode = app_lifecycle::AppRunMode::current();
    let isolated = match office::state::acceptance_runtime() {
        Ok(Some(runtime)) => {
            if !cfg!(target_os = "windows")
                || run_mode != app_lifecycle::AppRunMode::OfficeBackground
            {
                eprintln!("Isolated Office acceptance requires Windows --office-background");
                return;
            }
            // WebView2 must not reuse the installed application's profile.
            std::env::set_var("WEBVIEW2_USER_DATA_FOLDER", runtime.app_data_root.join("webview2"));
            true
        }
        Ok(None) => false,
        Err(error) => {
            eprintln!("Invalid isolated Office acceptance configuration: {error}");
            return;
        }
    };
    app_lifecycle::append_lifecycle_log(format!(
        "process startup mode={} args={:?}",
        run_mode.label(),
        std::env::args().collect::<Vec<_>>()
    ));

    let ocr_state = ocr::OcrState::default();
    let office_ocr_state = ocr_state.clone();
    let mut builder = tauri::Builder::default();

    if !isolated && run_mode != app_lifecycle::AppRunMode::OfficeBootstrap
        && std::env::var("VISUALTEX_ACCEPTANCE_ALLOW_PARALLEL_INSTANCE")
            .ok()
            .as_deref()
            != Some("1")
    {
        builder = builder.plugin(tauri_plugin_single_instance::init(
            |app, arguments, cwd| {
                app_lifecycle::append_lifecycle_log(format!(
                    "single-instance notification cwd={} args={arguments:?}",
                    cwd
                ));
                if !app_lifecycle::arguments_request_desktop(&arguments) {
                    app_lifecycle::append_lifecycle_log(
                        "single-instance notification does not request a desktop window",
                    );
                    return;
                }
                if let Err(error) = app_lifecycle::ensure_main_window(app) {
                    app_lifecycle::append_lifecycle_log(format!(
                        "single-instance desktop activation failed: {error}"
                    ));
                }
            },
        ));
    }

    let app = builder
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .manage(ocr_state)
        .setup(move |app| {
            // The template is excluded from static ACL loading. Register the
            // same bounded commands/windows for this process's one exact port.
            app.add_capability(office::state::office_capability_for_port(
                office::state::office_port()
            )?)?;
            if run_mode == app_lifecycle::AppRunMode::OfficeBootstrap {
                app_lifecycle::append_lifecycle_log(
                    "office-bootstrap setup skipped all windows, companion startup and warmup tasks",
                );
                return Ok(());
            }

            let office_state = office::initialize(app.handle(), office_ocr_state.clone())
                .map_err(|error| {
                    app_lifecycle::append_lifecycle_log(format!(
                        "Office companion initialization failed: {error}"
                    ));
                    std::io::Error::other(error)
                })?;
            if !isolated {
                if let Err(error) = office::powerpoint_native::start_double_click_monitor(
                    office_state.powerpoint_interactions.clone(),
                ) {
                    app_lifecycle::append_lifecycle_log(format!(
                        "Unable to start PowerPoint double-click monitor: {error}"
                    ));
                }
            }
            app.manage(office_state.clone());
            office::start(office_state);

            #[cfg(windows)]
            if !isolated { windows_silent_ocr_hotkey::initialize(app.handle()).map_err(|error| {
                app_lifecycle::append_lifecycle_log(format!(
                    "Silent OCR hotkey initialization failed: {error}"
                ));
                std::io::Error::other(error)
            })?; }

            // VisualTeX Anywhere: formulas typed in any application through VisualTeX's
            // input method (its shortcut, its link, VisualTeX Input).
            if !isolated {
                if let Err(error) = anywhere::initialize(app.handle()) {
                    app_lifecycle::append_lifecycle_log(format!(
                        "VisualTeX Anywhere initialization failed: {error}"
                    ));
                }
            }

            // MathType's VBA template cannot run in 64-bit WPS (VBA error boxes at every WPS
            // start): keep a placeholder under its name in WPS's startup folder.
            #[cfg(windows)]
            if !isolated {
                office::wps_mathtype_template::start();
            }

            if !isolated && run_mode.schedules_ocr_warmup() {
                // Preserve OCR model prewarming. It is independent from the main
                // embedded index.html and from the removed Office editor WebView prewarm.
                office_ocr_state.schedule_startup_warmup(app.handle().clone());
                app_lifecycle::append_lifecycle_log("OCR startup warmup scheduled");
            }

            if run_mode.creates_main_window() {
                if let Err(error) = office::background::resume_installed_launch_agent() {
                    app_lifecycle::append_lifecycle_log(format!(
                        "Unable to resume installed Office background service: {error}"
                    ));
                }
                app_lifecycle::ensure_main_window(app.handle()).map_err(|error| {
                    app_lifecycle::append_lifecycle_log(format!(
                        "desktop startup could not create main window: {error}"
                    ));
                    std::io::Error::other(error)
                })?;
            } else {
                app_lifecycle::append_lifecycle_log(
                    "office-background mode started companion without requesting main index.html or creating any WebView",
                );
            }
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![
            write_export_file,
            document_export::read_image_file,
            document_export::read_import_file,
            document_export::fetch_image_url,
            document_export::pdf_print_open,
            document_export::pdf_print_save,
            document_export::pdf_print_close,
            document_library::document_library_list,
            document_library::document_library_read,
            document_library::document_library_create,
            document_library::document_library_save,
            document_library::document_library_rename,
            document_library::document_library_duplicate,
            document_library::document_library_move,
            document_library::document_library_trash,
            document_library::document_library_restore,
            document_library::document_library_delete,
            document_library::document_library_empty_trash,
            document_library::document_library_create_folder,
            document_library::document_library_rename_folder,
            document_library::document_library_delete_folder,
            document_library::document_library_search,
            document_library::document_library_reveal,
            ocr::get_ocr_runtime_status,
            ocr::configure_ocr_storage_location,
            ocr::open_ocr_storage_location,
            ocr::install_ocr_runtime,
            ocr::get_ocr_install_status,
            ocr::cancel_ocr_install,
            ocr::open_ocr_install_logs,
            ocr::get_ocr_provider_configuration,
            ocr::save_ocr_provider_configuration,
            ocr::recognize_formula_image,
            ocr::cancel_ocr_recognition,
            ocr::restart_ocr_worker,
            ocr::warmup_ocr_model,
            ocr::reset_ocr_runtime,
            ocr::install_optional_ocr_model,
            ocr::remove_optional_ocr_model,
            ocr::get_ocr_model_catalog,
            ocr::get_ocr_model_download_status,
            ocr::download_ocr_model,
            ocr::cancel_ocr_model_download,
            ocr::configure_silent_ocr,
            windows_silent_ocr_hotkey::get_silent_ocr_hud_status,
            anywhere::anywhere_ready,
            anywhere::anywhere_layout,
            anywhere::anywhere_commit,
            anywhere::anywhere_commit_image,
            anywhere::anywhere_dismiss,
            anywhere::anywhere_marked,
            anywhere::anywhere_mode,
            anywhere::anywhere_log,
            anywhere::anywhere_capture_screenshot,
            anywhere::anywhere_read_clipboard,
            anywhere::anywhere_configure_hotkey,
            anywhere::anywhere_formula_shortcuts,
            anywhere::anywhere_status,
            windows_quick_ocr::capture_windows_quick_ocr,
            windows_quick_ocr::minimize_visualtex_main_window,
            windows_quick_ocr::write_windows_ocr_clipboard_text,
            get_app_window_configuration,
            apply_app_window_configuration,
            set_main_window_keypad_mode,
            get_word_numbering_user_configuration,
            apply_word_numbering_user_configuration,
            system_math_glyphs::probe_system_math_fonts,
            system_math_glyphs::list_system_font_families,
            system_math_glyphs::read_system_font_asset,
            system_math_glyphs::read_system_font_variant_asset,
            system_math_glyphs::extract_system_math_glyph,
            office::lifecycle::set_app_theme,
            office::lifecycle::set_app_editor_layout,
            office::lifecycle::set_app_editor_preferences,
            office::lifecycle::get_word_formula_shortcut_conflicts,
            office::lifecycle::set_powerpoint_default_font_size,
            office::lifecycle::get_powerpoint_default_font_size,
            office::lifecycle::get_mathtype_double_click_edit_enabled,
            office::lifecycle::set_mathtype_double_click_edit_enabled,
            office::lifecycle::get_office_companion_status,
            office::lifecycle::start_office_companion,
            office::lifecycle::stop_office_companion,
            office::lifecycle::get_office_integration_status,
            office::lifecycle::get_office_platform_status,
            office::lifecycle::set_office_background_start,
            office::lifecycle::set_office_integration_mode,
            office::lifecycle::install_windows_ole_integration,
            office::lifecycle::uninstall_windows_ole_integration,
            office::lifecycle::repair_windows_office_integration,
            office::lifecycle::test_windows_office_runtime,
            office::lifecycle::open_windows_office_logs,
            office::lifecycle::get_wps_integration_status,
            office::lifecycle::set_wps_host_enabled,
            office::lifecycle::get_native_render_status,
            office::lifecycle::start_native_render_service,
            office::lifecycle::install_office_integration,
            office::lifecycle::repair_office_integration,
            office::lifecycle::uninstall_office_integration,
            office::lifecycle::regenerate_office_certificate,
            office::lifecycle::open_word,
            office::lifecycle::open_powerpoint,
            office::automation::get_mcp_status,
            office::automation::set_mcp_settings,
            office::automation::reset_mcp_access,
            office::automation::get_mcp_audit
        ])
        .build(tauri::generate_context!())
        .unwrap_or_else(|error| {
            app_lifecycle::append_lifecycle_log(format!(
                "fatal Tauri application build failure: {error}"
            ));
            panic!("error while building VisualTeX: {error}");
        });

    if run_mode == app_lifecycle::AppRunMode::OfficeBootstrap {
        let result = office::bootstrap_configuration(app.handle());
        match result {
            Ok(()) => {
                app_lifecycle::append_lifecycle_log(
                    "office-bootstrap completed successfully and will exit without entering the event loop",
                );
                return;
            }
            Err(error) => {
                app_lifecycle::append_lifecycle_log(format!(
                    "office-bootstrap failed and will exit: {error}"
                ));
                std::process::exit(1);
            }
        }
    }

    let shutdown_started = Arc::new(AtomicBool::new(false));
    let shutdown_for_events = shutdown_started.clone();
    app.run(move |app, event| match event {
        #[cfg(target_os = "windows")]
        tauri::RunEvent::WindowEvent {
            label,
            event: tauri::WindowEvent::Resized(size),
            ..
        } if label == "main" => {
            app_lifecycle::schedule_persist_main_window_size(app, size.width, size.height);
        }
        #[cfg(target_os = "windows")]
        tauri::RunEvent::WindowEvent {
            label,
            event: tauri::WindowEvent::Resized(size),
            ..
        } if label == "office-session-editor" => {
            if let Some(window) = app.get_webview_window("office-session-editor") {
                let scale_factor = window.scale_factor().unwrap_or(1.0).max(f64::EPSILON);
                office::server::schedule_persist_office_editor_window_size(
                    app.clone(),
                    f64::from(size.width) / scale_factor,
                    f64::from(size.height) / scale_factor,
                );
            }
        }
        #[cfg(target_os = "windows")]
        tauri::RunEvent::WindowEvent {
            label,
            event: tauri::WindowEvent::CloseRequested { api, .. },
            ..
        } if label == "office-session-editor" => {
            // Do not let the native title-bar close destroy the reusable editor
            // before its Office Session reaches a terminal state. The page
            // performs an awaited commit/cancel and then calls the companion's
            // close endpoint, which hides the editor only after the Session is
            // safe for Word/PowerPoint to release its operation gate.
            api.prevent_close();
            if let Err(error) = office::server::request_desktop_editor_window_close(app) {
                app_lifecycle::append_lifecycle_log(format!(
                    "Office editor close request could not finalize the active Session: {error}"
                ));
            }
        }
        #[cfg(target_os = "windows")]
        tauri::RunEvent::WindowEvent {
            label,
            event: tauri::WindowEvent::CloseRequested { api, .. },
            ..
        } if label == "main" => {
            if app_lifecycle::background_retention_enabled() {
                api.prevent_close();
                if let Err(error) = app_lifecycle::destroy_main_window_for_background(app) {
                    app_lifecycle::append_lifecycle_log(format!(
                        "main close could not retain background companion: {error}"
                    ));
                }
            } else {
                api.prevent_close();
                app_lifecycle::append_lifecycle_log(
                    "main window close requested with background disabled; stopping companion, OCR worker and child processes",
                );
                shutdown_runtime(app, &shutdown_for_events, "main-window-close");
                app.exit(0);
            }
        }
        #[cfg(target_os = "macos")]
        tauri::RunEvent::WindowEvent {
            label,
            event: tauri::WindowEvent::CloseRequested { api, .. },
            ..
        } if label == "main" => {
            api.prevent_close();
            if let Err(error) = office::background::hide_main_window(app) {
                app_lifecycle::append_lifecycle_log(format!(
                    "Unable to hide macOS main window: {error}"
                ));
            }
        }
        #[cfg(target_os = "macos")]
        tauri::RunEvent::Reopen { .. } => {
            if let Err(error) = app_lifecycle::ensure_main_window(app) {
                app_lifecycle::append_lifecycle_log(format!(
                    "macOS reopen could not ensure main window: {error}"
                ));
            }
        }
        tauri::RunEvent::ExitRequested { code, api, .. } => {
            let automatic_last_window_exit = code.is_none();
            let retain_background = automatic_last_window_exit
                && (run_mode == app_lifecycle::AppRunMode::OfficeBackground
                    || app_lifecycle::background_retention_enabled());
            if retain_background {
                api.prevent_exit();
                app_lifecycle::append_lifecycle_log(
                    "automatic exit prevented because the Office companion is configured to remain in the background",
                );
                return;
            }
            shutdown_runtime(app, &shutdown_for_events, "exit-requested");
        }
        tauri::RunEvent::Exit => {
            shutdown_runtime(app, &shutdown_for_events, "event-loop-exit");
        }
        _ => {}
    });
}
