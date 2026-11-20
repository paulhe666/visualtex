//! Finishing a formula session: commit or cancel it and tell the host.

use super::*;

pub(super) fn cancel_host(request: &MacOfflineSessionRequest) -> Result<(), String> {
    if request.client.as_deref() == Some("wps") {
        crate::office::wps::bridge::deliver_cancel(&request.session_id);
        return Ok(());
    }
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
        ("nativeEquation", if request.native_equation { "1" } else { "0" }.into()),
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
        &dispatch_path(host, &request.session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    if cancel_needs_host_callback(request) {
        with_dispatch_pointer(host, &request.session_id, || run_vba_callback(host))?;
    }
    Ok(())
}

pub(super) fn complete_session(
    state: &OfficeCompanionState,
    session_id: &str,
) -> Result<OfficeFormulaSession, String> {
    state
        .session_store
        .patch(session_id, json!({ "status": "completed", "error": null }))
        .map_err(|error| error.to_string())
}

pub(super) fn fail_session(state: &OfficeCompanionState, session_id: &str, error: &str) {
    let _ = state
        .session_store
        .patch(session_id, json!({ "status": "failed", "error": error }));
}

pub(super) fn native_powerpoint_edit_source_changed(session: &OfficeFormulaSession) -> Result<bool, String> {
    let original_path = session_directory(OfficeHost::Powerpoint, &session.id)?
        .join("native-edit-original.omml");
    let original_omml = fs::read_to_string(&original_path)
        .map_err(|error| format!("The original native equation snapshot is unavailable: {error}"))?;
    native_powerpoint_edit_latex_changed(&original_omml, &session.lines)
}

pub(super) fn native_powerpoint_edit_latex_changed(original_omml: &str, lines: &[FormulaLine]) -> Result<bool, String> {
    let original = crate::office::powerpoint_omml::to_latex::from_omml(original_omml)?.latex;
    let current = lines.iter().map(|line| line.latex.as_str())
        .collect::<Vec<_>>().join("\n");
    Ok(current != original)
}

pub(super) fn commit_session_blocking(
    state: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    let id = session_id.clone();
    let committed = commit_formula_session(state, session_id);
    finish_agent_session(&id, committed.as_ref().map(|_| ()).map_err(Clone::clone));
    committed
}

pub(super) fn commit_formula_session(
    state: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    let commit_started = Instant::now();
    validate_uuid(&session_id, "Session id")?;
    let session = state
        .session_store
        .get(&session_id)
        .map_err(|error| error.to_string())?;
    queue_editor_performance(
        session.host,
        &session_id,
        "apply-backend-start",
        0.0,
        None,
        json!({ "mode": session.mode, "dirty": session.dirty }),
    );
    if session.status == OfficeSessionStatus::Completed {
        let _ = cleanup_session_files(session.host, &session_id);
        return Ok(session);
    }
    if session.status != OfficeSessionStatus::Committing {
        return Err("Offline Office Session is not ready to commit".to_string());
    }
    let request = read_request(&session_id)?;
    // The native source is captured from Office at open time. Compare against
    // that immutable snapshot even if a stale frontend dirty flag was cleared
    // by autosave; otherwise an edited formula can be silently treated as a
    // no-op and the UI closes without changing PowerPoint.
    let native_powerpoint_source_changed = if session.host == OfficeHost::Powerpoint
        && request.native_equation && session.mode == OfficeSessionMode::Edit {
        native_powerpoint_edit_source_changed(&session)?
    } else { false };
    // An untagged Word equation sent to "convert to image" is unchanged in the
    // editor but still has to become an image.
    let word_format_conversion_requested =
        session.host == OfficeHost::Word && word_format_conversion_requested(&request);
    if session.mode == OfficeSessionMode::Edit
        && !session.dirty
        && !word_format_conversion_requested
        && !native_powerpoint_source_changed
        && !(request.agent && !(request.host == "powerpoint" && request.native_equation))
    {
        // Nothing to apply: release the host's session state (PowerPoint's
        // native edit, Word's session-only anchor) the way Cancel does.
        if (request.host == "powerpoint" && request.native_equation)
            || (request.host == "word" && cancel_needs_host_callback(&request))
        {
            cancel_host(&request)?;
        }
        let completed = complete_session(&state, &session_id)?;
        queue_editor_performance(
            session.host,
            &session_id,
            "apply-noop-complete",
            commit_started.elapsed().as_secs_f64() * 1000.0,
            None,
            json!({}),
        );
        let _ = cleanup_session_files(session.host, &session_id);
        return Ok(completed);
    }
    let mut metadata = metadata_from_session(&session);
    let result = match session.host {
        OfficeHost::Word => {
            let geometry = calculate_word_geometry(&request, &session)?;
            metadata.latex = canonical_document_formula_latex(&metadata)?;
            metadata.font_size_pt = Some(geometry.font_size_pt);
            metadata.reference_width_pt = Some(geometry.reference_width_pt);
            metadata.reference_height_pt = Some(geometry.reference_height_pt);
            metadata.reference_baseline_pt = Some(geometry.reference_baseline_pt);
            let encoded = encode_metadata(&metadata)?;
            commit_word(
                state.app.as_ref(),
                &request,
                &session,
                &encoded,
                &metadata.latex,
                geometry,
            )
        }
        OfficeHost::Powerpoint if request.native_equation => {
            commit_powerpoint_native(state.app.as_ref(), &request, &session)
        }
        OfficeHost::Powerpoint => {
            let powerpoint = request
                .power_point
                .as_ref()
                .ok_or_else(|| "PowerPoint request geometry is missing".to_string())?;
            let geometry = calculate_powerpoint_geometry(powerpoint, &session)?;
            metadata.font_size_pt = Some(geometry.font_size_pt);
            metadata.reference_width_pt = Some(geometry.reference_width_pt);
            metadata.reference_height_pt = Some(geometry.reference_height_pt);
            metadata.reference_baseline_pt = None;
            let encoded = encode_metadata(&metadata)?;
            commit_powerpoint(
                state.app.as_ref(),
                &request,
                &session,
                &encoded,
                geometry,
            )
        }
    };
    if let Err(error) = result {
        fail_session(&state, &session_id, &error);
        return Err(error);
    }
    queue_editor_performance(
        session.host,
        &session_id,
        "apply-office-callback-complete",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    let completed = complete_session(&state, &session_id)?;
    queue_editor_performance(
        session.host,
        &session_id,
        "apply-backend-complete",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );

    // The Office document already contains the durable edit metadata. Cache
    // refresh and ephemeral Session cleanup must not extend the user's Apply
    // wait after the host callback has succeeded.
    let maintenance_state = state.clone();
    let maintenance_session_id = session_id.clone();
    let maintenance_host = session.host;
    let formula_id = session.formula_id.clone();
    crate::spawn_background_task("visualtex-office-apply-maintenance", move || {
        if let Err(error) = maintenance_state.formula_cache.put(&formula_id, metadata) {
            eprintln!("Unable to refresh the VisualTeX formula cache after Apply: {error}");
        }
        let _ = cleanup_session_files(maintenance_host, &maintenance_session_id);
    });
    Ok(completed)
}

pub(super) fn cancel_session_blocking(
    state: OfficeCompanionState,
    session_id: String,
) -> Result<OfficeFormulaSession, String> {
    // an agent's formula is cancelled when the hidden editor could not draw it
    finish_agent_session(&session_id, Err("VisualTeX could not draw the formula; check its LaTeX".to_string()));
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    let host = host_from_request_name(&request.host)?;
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
    let _ = cleanup_session_files(host, &session_id);
    Ok(cancelled)
}
