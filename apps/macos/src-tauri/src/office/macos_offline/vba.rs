//! Calling back into the add-ins: the dispatch file, VBA macros through AppleScript and Apple events, and the hotkey macros.

use super::*;

pub(super) fn sanitize_dispatch_value(value: &str, label: &str) -> Result<String, String> {
    if value.contains(['\r', '\n', '\0']) {
        return Err(format!("{label} contains unsupported control characters"));
    }
    Ok(value.to_string())
}

pub(super) fn dispatch_text(entries: &[(&str, String)]) -> Result<String, String> {
    let dynamic = entries
        .iter()
        .map(|(key, value)| ((*key).to_string(), value.clone()))
        .collect::<Vec<_>>();
    dynamic_dispatch_text(&dynamic)
}

pub(super) fn dynamic_dispatch_text(entries: &[(String, String)]) -> Result<String, String> {
    let mut seen = std::collections::HashSet::new();
    let mut output = String::new();
    for (key, value) in entries {
        if !seen.insert(key.as_str())
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

pub(super) fn vba_callback_script(host: OfficeHost) -> &'static str {
    match host {
        OfficeHost::Word => {
            r#"with timeout of 1800 seconds
tell application "Microsoft Word"
run VB macro macro name "VisualTeX_ApplyPendingResult"
end tell
end timeout"#
        }
        OfficeHost::Powerpoint => {
            r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "Microsoft PowerPoint has no active presentation"
run VB macro macro name "VisualTeX_ApplyPendingResult" list of parameters {}
end tell"#
        }
    }
}

pub(super) fn run_vba_callback(host: OfficeHost) -> Result<(), String> {
    run_office_vba_script(vba_callback_script(host), "Office VBA callback")
}

pub(super) fn run_vba_callback_on_main_thread(
    app: Option<&AppHandle>,
    host: OfficeHost,
) -> Result<(), String> {
    let Some(app) = app else {
        return run_vba_callback(host);
    };
    if host == OfficeHost::Word {
        let (sender, receiver) = mpsc::sync_channel(1);
        app.run_on_main_thread(move || {
            let _ = sender.send(execute_word_vba_apple_event());
        })
        .map_err(|error| format!("Unable to schedule the Office VBA callback: {error}"))?;
        return receiver
            .recv_timeout(Duration::from_secs(1_800))
            .map_err(|error| {
                format!("The Office VBA callback main-thread callback did not finish: {error}")
            })?;
    }
    run_office_vba_script_on_main_thread(
        app,
        vba_callback_script(host),
        "Office VBA callback",
    )
}

pub(crate) fn run_powerpoint_native_double_click_edit(x: f64, y: f64) -> Result<(), String> {
    if !x.is_finite() || !y.is_finite() { return Err("Invalid native equation click coordinates".into()); }
    let script=format!(r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then return
run VB macro macro name "VisualTeX_EditNativeAtScreenPoint" list of parameters {{"{x:.6}", "{y:.6}"}}
end tell"#);
    run_office_vba_script(&script,"PowerPoint native equation double-click")
}

pub(crate) fn run_double_click_edit_macro(host: OfficeHost) -> Result<(), String> {
    let script = match host {
        OfficeHost::Word => {
            r#"tell application "Microsoft Word"
if not (exists active document) then error "Microsoft Word has no active document"
run VB macro macro name "VisualTeX_DoubleClickEditSelected"
end tell"#
        }
        OfficeHost::Powerpoint => {
            r#"tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "Microsoft PowerPoint has no active presentation"
run VB macro macro name "VisualTeX_DoubleClickEditSelected" list of parameters {}
end tell"#
        }
    };
    run_office_vba_script(script, "Office double-click edit macro")
}

pub(crate) fn run_office_hotkey_macro(
    host: OfficeHost,
    macro_name: &'static str,
) -> Result<(), String> {
    let allowed = matches!(
        (host, macro_name),
        (OfficeHost::Word, "VisualTeX_CreateInline")
            | (OfficeHost::Word, "VisualTeX_CreateDisplay")
            | (OfficeHost::Word, "VisualTeX_CreateNativeInline")
            | (OfficeHost::Word, "VisualTeX_CreateNativeDisplay")
            | (OfficeHost::Powerpoint, "VisualTeX_NewFormula")
            | (
                OfficeHost::Powerpoint,
                "VisualTeX_NewInlineNativeEquation"
            )
            | (
                OfficeHost::Powerpoint,
                "VisualTeX_NewDisplayNativeEquation"
            )
    );
    if !allowed {
        return Err("Unsupported Office hotkey action".to_string());
    }
    let script = match host {
        OfficeHost::Word => format!(
            r#"with timeout of 1800 seconds
tell application "Microsoft Word"
if not (exists active document) then error "Microsoft Word has no active document"
run VB macro macro name "{macro_name}"
end tell
end timeout"#,
        ),
        OfficeHost::Powerpoint => format!(
            r#"with timeout of 1800 seconds
tell application "Microsoft PowerPoint"
if not (exists active presentation) then error "Microsoft PowerPoint has no active presentation"
run VB macro macro name "{macro_name}" list of parameters {{}}
end tell
end timeout"#,
        ),
    };
    run_office_vba_script(&script, "Office creation hotkey")
}

pub(crate) fn run_word_image_double_click_edit_macro() -> Result<(), String> {
    run_office_vba_script(
        r#"tell application "Microsoft Word"
if not (exists active document) then error "Microsoft Word has no active document"
run VB macro macro name "VisualTeX_EditSelectedImageFromNativeMonitor"
end tell"#,
        "Word image double-click edit macro",
    )
}

pub(crate) fn run_office_vba_script_subprocess(script: &str, label: &str) -> Result<(), String> {
    let output = Command::new("/usr/bin/osascript")
        .arg("-e")
        .arg(script)
        .output()
        .map_err(|error| format!("Unable to launch the {label}: {error}"))?;
    if output.status.success() {
        Ok(())
    } else {
        let detail = String::from_utf8_lossy(&output.stderr).trim().to_string();
        Err(if detail.is_empty() {
            format!("The {label} failed")
        } else {
            format!("The {label} failed: {detail}")
        })
    }
}

#[cfg(target_os = "macos")]
pub(super) fn execute_office_vba_apple_script(script: &str, label: &str) -> Result<(), String> {
    OFFICE_VBA_APPLESCRIPT_CACHE.with(|cache| {
        let mut cache = cache.borrow_mut();
        let compiled_index = if let Some(index) = cache
            .iter()
            .position(|(cached_source, _)| cached_source == script)
        {
            index
        } else {
            let source = NSString::from_str(script);
            let Some(compiled_script) =
                NSAppleScript::initWithSource(NSAppleScript::alloc(), &source)
            else {
                return Err(format!("Unable to initialize the {label}"));
            };
            let mut compilation_error = None;
            if !unsafe { compiled_script.compileAndReturnError(Some(&mut compilation_error)) } {
                return Err(format!(
                    "Unable to compile the {label}: {compilation_error:?}"
                ));
            }
            cache.push((script.to_string(), compiled_script));
            cache.len() - 1
        };

        let compiled_script = &cache[compiled_index].1;
        let mut execution_error = None;
        let _result = unsafe { compiled_script.executeAndReturnError(Some(&mut execution_error)) };
        if let Some(error) = execution_error {
            Err(format!("The {label} failed: {error:?}"))
        } else {
            Ok(())
        }
    })
}

#[cfg(target_os = "macos")]
pub(super) const fn apple_event_code(code: &[u8; 4]) -> u32 {
    u32::from_be_bytes(*code)
}

#[cfg(target_os = "macos")]
pub(super) fn execute_word_vba_apple_event() -> Result<(), String> {
    // Word publishes this command as sWRD/1149 with the macro-name parameter
    // 5112 in Word.sdef. Sending that event directly avoids the fixed
    // NSAppleScript execution overhead while invoking the identical VBA entry.
    let bundle_id = NSString::from_str("com.microsoft.Word");
    let target = NSAppleEventDescriptor::descriptorWithBundleIdentifier(&bundle_id);
    let event = NSAppleEventDescriptor::appleEventWithEventClass_eventID_targetDescriptor_returnID_transactionID(
        apple_event_code(b"sWRD"),
        apple_event_code(b"1149"),
        Some(&target),
        kAutoGenerateReturnID as i16,
        kAnyTransactionID,
    );
    let macro_name = NSString::from_str("VisualTeX_ApplyPendingResult");
    let macro_descriptor = NSAppleEventDescriptor::descriptorWithString(&macro_name);
    event.setParamDescriptor_forKeyword(&macro_descriptor, apple_event_code(b"5112"));
    let reply = event
        .sendEventWithOptions_timeout_error(NSAppleEventSendOptions::DefaultOptions, 1_800.0)
        .map_err(|error| format!("The Office VBA callback failed: {error}"))?;
    if let Some(error_number) = reply.paramDescriptorForKeyword(keyErrorNumber) {
        let code = error_number.int32Value();
        if code != 0 {
            let detail = reply
                .paramDescriptorForKeyword(keyErrorString)
                .and_then(|descriptor| descriptor.stringValue())
                .map(|value| value.to_string())
                .unwrap_or_else(|| "Microsoft Word returned an Apple Event error".to_string());
            return Err(format!("The Office VBA callback failed ({code}): {detail}"));
        }
    }
    Ok(())
}

#[cfg(target_os = "macos")]
pub(crate) fn run_office_vba_script_on_main_thread(
    app: &AppHandle,
    script: &str,
    label: &str,
) -> Result<(), String> {
    // NSAppleScript may wait indefinitely when instantiated on a Tauri worker
    // thread. Dispatch its complete compile/execute lifetime to AppKit's main
    // thread, then return only the owned Result through a bounded channel.
    let script = script.to_string();
    let label = label.to_string();
    let callback_label = label.clone();
    let (sender, receiver) = mpsc::sync_channel(1);
    app.run_on_main_thread(move || {
        let result = execute_office_vba_apple_script(&script, &callback_label);
        let _ = sender.send(result);
    })
    .map_err(|error| format!("Unable to schedule the {label}: {error}"))?;
    receiver
        .recv_timeout(Duration::from_secs(1_800))
        .map_err(|error| format!("The {label} main-thread callback did not finish: {error}"))?
}

pub(super) fn run_office_vba_script(script: &str, label: &str) -> Result<(), String> {
    run_office_vba_script_subprocess(script, label)
}

#[cfg(not(target_os = "macos"))]
pub(crate) fn run_office_vba_script_on_main_thread(
    _app: &AppHandle,
    script: &str,
    label: &str,
) -> Result<(), String> {
    run_office_vba_script_subprocess(script, label)
}

pub(super) fn with_dispatch_pointer<T>(
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
    atomic_write_runtime(&pointer, session_id.as_bytes(), 0o600)?;
    let result = operation();
    let _ = fs::remove_file(pointer);
    // The add-in's callback reports its failure here instead of raising it
    // (VTReportCallbackFailure): a raised error stops Office in a dialog.
    let reported = session_directory(host, session_id)
        .map(|directory| directory.join(CALLBACK_ERROR_FILE))
        .ok()
        .and_then(|path| {
            let text = fs::read_to_string(&path).ok()?;
            let _ = fs::remove_file(&path);
            Some(text)
        });
    if let Some(text) = reported {
        let message = text.split_once('\n').map_or(text.as_str(), |(_, message)| message).trim();
        return Err(if message.is_empty() { "The Office add-in could not apply the result".to_string() } else { message.to_string() });
    }
    result
}
