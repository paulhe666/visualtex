use super::*;

fn wps_request(changes: Value) -> MacOfflineSessionRequest {
    let session_id = "0f8fad5b-d9cb-469f-a165-70867728950e";
    let mut request = json!({
        "protocolVersion": OFFLINE_PROTOCOL_VERSION,
        "sessionId": session_id,
        "host": "word",
        "client": "wps",
        "mode": "create",
        "operation": "formula",
        "displayMode": "inline",
        "numbered": false,
        "nativeEquation": false,
        "sourceDocumentId": "wps:/Users/a/doc.docx",
        "pendingMarker": format!("{PENDING_PREFIX}{session_id}"),
        "fontSizePt": 12.0,
    });
    for (key, value) in changes.as_object().unwrap() {
        if value.is_null() {
            request.as_object_mut().unwrap().remove(key);
        } else {
            request[key] = value.clone();
        }
    }
    serde_json::from_value(request).unwrap()
}

#[test]
fn wps_requests_pass_the_word_rules() {
    let id = "0f8fad5b-d9cb-469f-a165-70867728950e";
    validate_request(&wps_request(json!({})), id).unwrap();
    let image_edit = wps_request(json!({
        "mode": "edit",
        "pendingMarker": null,
        "formulaId": "5028ce4b-71c7-4d06-a165-61151fd64183",
        "sourceObjectId": "visualtex:formula-ref:v1:5028ce4b-71c7-4d06-a165-61151fd64183:inline:0",
        "encodedMetadata": "visualtex:v1:deflate:abc",
    }));
    validate_request(&image_edit, id).unwrap();
    let equation_edit = wps_request(json!({
        "mode": "edit",
        "pendingMarker": null,
        "nativeEquation": true,
        "formulaId": "5028ce4b-71c7-4d06-a165-61151fd64183",
        "sourceObjectId": "VT_E_0f8fad5bd9cb469fa165",
        "nativeEditXml": "<pkg:package/>",
    }));
    assert!(is_untagged_word_edit(&equation_edit));
    validate_request(&equation_edit, id).unwrap();
    assert!(cancel_needs_host_callback(&equation_edit));
    // A numbered formula is edited in WPS and keeps its number.
    let numbered_image_edit = wps_request(json!({
        "mode": "edit",
        "pendingMarker": null,
        "displayMode": "block",
        "numbered": true,
        "formulaId": "5028ce4b-71c7-4d06-a165-61151fd64183",
        "sourceObjectId": "visualtex:formula-ref:v1:5028ce4b-71c7-4d06-a165-61151fd64183:block:1",
        "encodedMetadata": "visualtex:v1:deflate:abc",
    }));
    validate_request(&numbered_image_edit, id).unwrap();
    let numbered_equation_edit = wps_request(json!({
        "mode": "edit",
        "pendingMarker": null,
        "displayMode": "block",
        "numbered": true,
        "nativeEquation": true,
        "formulaId": "5028ce4b-71c7-4d06-a165-61151fd64183",
        "sourceObjectId": "VT_E_0f8fad5bd9cb469fa165",
        "nativeEditXml": "<pkg:package/>",
    }));
    validate_request(&numbered_equation_edit, id).unwrap();
    // WPS numbers new display formulas too.
    validate_request(&wps_request(json!({ "numbered": true, "displayMode": "block" })), id).unwrap();
    // The ribbon's conversions run as hidden edits of the selected formula.
    let to_image = wps_request(json!({
        "mode": "edit",
        "pendingMarker": null,
        "operation": "nativeToImage",
        "formulaId": "5028ce4b-71c7-4d06-a165-61151fd64183",
        "sourceObjectId": "VT_E_0f8fad5bd9cb469fa165",
        "nativeEditXml": "<pkg:package/>",
    }));
    validate_request(&to_image, id).unwrap();
    let to_equation = wps_request(json!({
        "mode": "edit",
        "pendingMarker": null,
        "operation": "imageToNative",
        "nativeEquation": true,
        "formulaId": "5028ce4b-71c7-4d06-a165-61151fd64183",
        "sourceObjectId": "visualtex:formula-ref:v1:5028ce4b-71c7-4d06-a165-61151fd64183:inline:0",
        "encodedMetadata": "visualtex:v1:deflate:abc",
    }));
    validate_request(&to_equation, id).unwrap();
    for rejected in [
        json!({ "numbered": true, "displayMode": "inline" }),
        json!({ "client": "other" }),
        json!({ "operation": "nativeToImage" }),
    ] {
        assert!(validate_request(&wps_request(rejected.clone()), id).is_err(), "{rejected}");
    }
    // The stored request round-trips through the strict reader.
    let bytes = serde_json::to_vec(&image_edit).unwrap();
    let reread: MacOfflineSessionRequest = serde_json::from_slice(&bytes).unwrap();
    assert_eq!(reread.client.as_deref(), Some("wps"));
}

/// What office::wps::bridge sends for the WPS Presentation add-in.
fn wps_presentation_request(changes: Value) -> MacOfflineSessionRequest {
    let session_id = "0f8fad5b-d9cb-469f-a165-70867728950e";
    let formula_id = "5028ce4b-71c7-4d06-a165-61151fd64183";
    let mut request = json!({
        "protocolVersion": OFFLINE_PROTOCOL_VERSION,
        "sessionId": session_id,
        "host": "powerpoint",
        "client": "wps",
        "mode": "create",
        "operation": "formula",
        "displayMode": "block",
        "numbered": false,
        "nativeEquation": false,
        "sourceDocumentId": "/Users/a/Deck.pptx",
        "forkCopiedFormula": false,
        "formulaId": formula_id,
        "sourceObjectId": format!("VisualTeX_{formula_id}"),
        "pendingMarker": format!("{PENDING_PREFIX}{session_id}:{formula_id}"),
        "powerPoint": {
            "presentationIdentity": "/Users/a/Deck.pptx", "slideIndex": 2, "slideId": 257,
            "shapeIndex": 3, "shapeId": 6, "shapeName": format!("VisualTeX_{formula_id}"),
            "left": 390.0, "top": 249.0, "width": 180.0, "height": 42.0, "rotation": 0.0,
            "zOrder": 3, "fontSizePt": 18.0, "referenceWidthPt": null, "referenceHeightPt": null
        },
    });
    for (key, value) in changes.as_object().unwrap() {
        if value.is_null() {
            request.as_object_mut().unwrap().remove(key);
        } else {
            request[key] = value.clone();
        }
    }
    serde_json::from_value(request).unwrap()
}

#[test]
fn wps_presentation_requests_pass_the_powerpoint_rules() {
    let id = "0f8fad5b-d9cb-469f-a165-70867728950e";
    let create = wps_presentation_request(json!({}));
    validate_request(&create, id).unwrap();
    assert!(cancel_needs_host_callback(&create));
    let edit = wps_presentation_request(json!({
        "mode": "edit",
        "pendingMarker": null,
        "encodedMetadata": "visualtex:v1:deflate:abc",
    }));
    validate_request(&edit, id).unwrap();
    assert!(!cancel_needs_host_callback(&edit));
    // A pasted copy is edited as a new formula, as in PowerPoint.
    let copy = wps_presentation_request(json!({
        "mode": "edit",
        "pendingMarker": null,
        "encodedMetadata": "visualtex:v1:deflate:abc",
        "forkCopiedFormula": true,
    }));
    validate_request(&copy, id).unwrap();
    for rejected in [
        json!({ "nativeEquation": true }),
        json!({ "operation": "nativeToImage" }),
        json!({ "powerPoint": null }),
        json!({ "fontSizePt": 18.0 }),
        json!({ "numbered": true }),
        json!({ "client": "other" }),
    ] {
        assert!(validate_request(&wps_presentation_request(rejected.clone()), id).is_err(), "{rejected}");
    }
    assert_eq!(host_from_request_name(&create.host).unwrap(), OfficeHost::Powerpoint);
}

#[test]
fn unmanaged_numbered_omml_restores_as_a_new_numbered_formula() {
    assert!(omml_restore_identity(None, None, "block", true)
        .expect("an equation numbered by Word alone is restorable")
        .is_none());
    assert!(omml_restore_identity(None, None, "inline", false)
        .expect("a plain Word equation is restorable")
        .is_none());
}

#[test]
fn omml_restore_rejects_metadata_without_a_formula_id() {
    let error = omml_restore_identity(None, Some("e30"), "block", true)
        .expect_err("metadata without an id is a broken identity");
    assert!(error.contains("partial VisualTeX identity"));
    assert!(omml_restore_identity(
        Some("00000000-0000-4000-8000-000000000000"),
        None,
        "block",
        false,
    )
    .is_err());
}

#[test]
fn native_powerpoint_edit_detects_changed_source_even_with_stale_dirty_flag() {
    let original = "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>";
    let mut lines = vec![FormulaLine { id: "line".into(), latex: "x".into() }];
    assert!(!native_powerpoint_edit_latex_changed(original, &lines).unwrap());
    lines[0].latex = "y".into();
    assert!(native_powerpoint_edit_latex_changed(original, &lines).unwrap());
}

#[test]
fn office_open_boundary_preserves_success_and_recoverable_errors() {
    assert_eq!(protect_office_open(|| Ok::<_, String>(42)), Ok(42));
    assert_eq!(
        protect_office_open(|| Err::<(), _>("invalid Session".to_string())),
        Err("invalid Session".to_string())
    );
}

#[test]
fn office_open_boundary_converts_panics_to_recoverable_errors() {
    let result: Result<(), String> =
        protect_office_open(|| panic!("malformed formula fixture"));
    let error = result.expect_err("the Office open panic must be contained");
    assert!(error.contains("recovered from an internal error"));
    assert!(error.contains("malformed formula fixture"));
}

#[test]
fn malformed_cached_png_metadata_returns_an_error_instead_of_panicking() {
    let mut truncated = b"\x89PNG\r\n\x1a\n".to_vec();
    truncated.extend_from_slice(&13_u32.to_be_bytes());
    truncated.extend_from_slice(b"IHDR");
    truncated.extend_from_slice(&[0_u8; 4]);
    assert!(png_ink_center_y_ratio_from_bytes(&truncated).is_err());

    let mut overflowing = b"\x89PNG\r\n\x1a\n".to_vec();
    overflowing.extend_from_slice(&u32::MAX.to_be_bytes());
    overflowing.extend_from_slice(b"IDAT");
    overflowing.extend_from_slice(&[0_u8; 4]);
    assert!(png_ink_center_y_ratio_from_bytes(&overflowing).is_err());
}

#[test]
fn omml_document_commits_do_not_require_svg_or_png_artifacts() {
    assert!(!document_output_requires_image_artifacts("omml"));
    assert!(document_output_requires_image_artifacts("image"));
}

#[test]
fn strips_only_visualtex_numbered_equation_array_wrapper() {
    let numbered = concat!(
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ",
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
        "<m:eqArr><m:eqArrPr/><m:e>",
        "<m:r><m:t>x</m:t></m:r>",
        "<m:r><m:t>#</m:t></m:r>",
        "<m:d><m:e><m:r><w:fldChar w:fldCharType=\"begin\"/><m:t> REF VT_N_12345678 </m:t><w:fldChar w:fldCharType=\"end\"/></m:r></m:e></m:d>",
        "</m:e></m:eqArr></m:oMath>"
    );
    let stripped = strip_visualtex_numbered_equation_array(numbered)
        .expect("VisualTeX numbered Equation Array should be stripped");
    assert!(stripped.contains("<m:r><m:t>x</m:t></m:r>"));
    assert!(!stripped.contains("<m:eqArr"));
    assert!(!stripped.contains("<m:t>#</m:t>"));
    assert!(!stripped.contains("VT_N_"));
    assert!(stripped.starts_with("<m:oMath "));
    assert!(stripped.ends_with("</m:oMath>"));

    let ordinary = concat!(
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">",
        "<m:eqArr><m:eqArrPr/><m:e><m:r><m:t>x=y</m:t></m:r></m:e></m:eqArr>",
        "</m:oMath>"
    );
    assert!(strip_visualtex_numbered_equation_array(ordinary).is_none());

    let unrelated_hash = concat!(
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\">",
        "<m:eqArr><m:eqArrPr/><m:e>",
        "<m:r><m:t>x</m:t></m:r><m:r><m:t>#</m:t></m:r>",
        "<m:r><m:t>ordinary text</m:t></m:r>",
        "</m:e></m:eqArr></m:oMath>"
    );
    assert!(strip_visualtex_numbered_equation_array(unrelated_hash).is_none());
}

#[test]
fn strips_native_sequence_number_without_visualtex_identity() {
    let xml = concat!(
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ",
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
        "<m:eqArr><m:e><m:r><m:t>x</m:t></m:r>",
        "<m:r><m:t>#</m:t></m:r><m:d><m:e>",
        "<m:r><w:fldChar w:fldCharType=\"begin\"/></m:r>",
        "<m:r><w:instrText> SEQ Equation \\* ARABIC </w:instrText></m:r>",
        "<m:r><w:fldChar w:fldCharType=\"separate\"/></m:r>",
        "<m:r><m:t>12</m:t></m:r>",
        "<m:r><w:fldChar w:fldCharType=\"end\"/></m:r>",
        "</m:e></m:d></m:e></m:eqArr></m:oMath>"
    );
    let body = strip_visualtex_numbered_equation_array(xml).unwrap();
    assert!(body.contains("<m:t>x</m:t>"));
    assert!(!body.contains("SEQ"));
    assert!(!body.contains("12"));
    assert!(!body.contains("eqArr"));
}

#[test]
fn native_number_stripping_preserves_merged_runs_and_nested_arrays() {
    let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><m:eqArr><m:e><m:eqArr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:eqArr><m:r><m:t>=x#</m:t></m:r><m:d><m:e><m:r><w:fldChar w:fldCharType="begin"/><m:t xml:space="preserve"> SEQ </m:t></m:r><m:r><m:t>公式</m:t></m:r><m:r><m:t> \* ARABIC </m:t><w:fldChar w:fldCharType="separate"/><m:t>4</m:t><w:fldChar w:fldCharType="end"/></m:r></m:e></m:d></m:e></m:eqArr></m:oMath>"#;
    let body = strip_visualtex_numbered_equation_array(xml).unwrap();
    assert!(body.contains("<m:t>=x</m:t>"));
    assert!(body.contains("<m:t>a</m:t>"));
    assert!(body.contains("<m:t>b</m:t>"));
    assert_eq!(body.matches("<m:eqArr>").count(), 1);
    assert!(!body.contains("SEQ"));
    let manual = xml.replace("w:fldChar", "w:other");
    assert!(strip_visualtex_numbered_equation_array(&manual).is_none());
}

#[cfg(target_os = "macos")]
#[test]
fn numbered_equation_array_omml_to_latex_excludes_visualtex_number_field() {
    let numbered = concat!(
        "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" ",
        "xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">",
        "<m:eqArr><m:eqArrPr><m:maxDist m:val=\"1\"/></m:eqArrPr><m:e>",
        "<m:sSup><m:e><m:r><m:t>x</m:t></m:r></m:e>",
        "<m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSup>",
        "<m:r><m:rPr><m:sty m:val=\"p\"/></m:rPr><m:t>#</m:t></m:r>",
        "<m:d><m:e><m:r><w:fldChar w:fldCharType=\"begin\"/></m:r>",
        "<m:r><m:t xml:space=\"preserve\"> REF VT_N_12345678 </m:t></m:r>",
        "<m:r><w:fldChar w:fldCharType=\"end\"/></m:r></m:e></m:d>",
        "</m:e></m:eqArr></m:oMath>"
    );
    let math_ml = word_omml_to_mathml(numbered, true)
        .expect("the stripped formula body should convert to MathML");
    assert!(!math_ml.contains("VT_N_"));
    assert!(!math_ml.contains("REF"));
    let latex = crate::office::omml_batch::mathml_to_latex(&math_ml)
        .expect("the stripped MathML should convert to LaTeX");
    assert!(latex.contains('x'));
    assert!(latex.contains('2'));
    assert!(!latex.contains("VT_N_"));
    assert!(!latex.contains("REF"));
    assert!(!latex.contains("\\#"));
}

#[test]
fn formula_restore_manifest_from_environment_is_valid() {
    let Ok(session_id) = std::env::var("VISUALTEX_FORMULA_RESTORE_SESSION") else {
        return;
    };
    let request = read_request(&session_id)
        .expect("the requested real formula restore Session must be readable");
    let public = document_import_request_data(&request)
        .expect("the requested real formula restore manifest must validate");
    assert_eq!(public.operation, "formulaRestore");
    assert!(public.restore_targets.is_some());
}

#[test]
fn office_formula_editors_have_one_stable_window_label_per_host() {
    assert_eq!(
        editor_window_label(OfficeHost::Word),
        "office-native-word-editor"
    );
    assert_eq!(
        editor_window_label(OfficeHost::Powerpoint),
        "office-native-powerpoint-editor"
    );
    assert_eq!(
        editor_window_host("office-native-word-editor"),
        Some(OfficeHost::Word)
    );
    assert_eq!(
        editor_window_host("office-native-powerpoint-editor"),
        Some(OfficeHost::Powerpoint)
    );
    assert_eq!(editor_window_host("office-native-document-session"), None);
}

#[test]
fn fast_open_inboxes_are_host_sandbox_only_and_uuid_named() {
    let home = Path::new("/Users/visualtex-test");
    let roots = fast_open_inbox_roots_for_home(home);
    assert_eq!(roots.len(), 2);
    assert_eq!(
        roots[0],
        (
            OfficeHost::Word,
            home.join(WORD_FAST_OPEN_INBOX_SUFFIX)
        )
    );
    assert_eq!(
        roots[1],
        (
            OfficeHost::Powerpoint,
            home.join(POWERPOINT_FAST_OPEN_INBOX_SUFFIX)
        )
    );
    assert!(roots
        .iter()
        .all(|(_, root)| root.to_string_lossy().contains("/Library/Containers/com.microsoft.")));

    let session_id = "12345678-1234-4234-9234-123456789abc";
    assert_eq!(
        fast_open_session_id(Path::new(&format!("{session_id}.json"))).as_deref(),
        Some(session_id)
    );
    assert!(fast_open_session_id(Path::new(&format!(".{session_id}.tmp"))).is_none());
    assert!(fast_open_session_id(Path::new("not-a-uuid.json")).is_none());
    assert_eq!(
        fast_open_accept_path(Path::new("/tmp/inbox"), session_id)
            .expect("accepted marker path should be valid"),
        Path::new("/tmp/inbox").join(format!("{session_id}.accepted"))
    );
}

#[test]
fn fast_open_supports_prepared_word_operations_without_accepting_unknown_actions() {
    for operation in [
        "formula",
        "nativeToImage",
        "imageToNative",
        "documentImport",
        "latexRedraw",
        "formulaRestore",
    ] {
        assert!(
            fast_open_operation_supported(operation),
            "fast-open should accept {operation}"
        );
    }
    assert!(!fast_open_operation_supported("unknownOperation"));
}

#[test]
fn fast_open_requests_have_a_short_replay_window() {
    let now = UNIX_EPOCH + Duration::from_secs(100);
    assert!(fast_open_modified_is_recent(
        now - Duration::from_secs(FAST_OPEN_MAX_AGE.as_secs() - 1),
        now,
    ));
    assert!(!fast_open_modified_is_recent(
        now - Duration::from_secs(FAST_OPEN_MAX_AGE.as_secs() + 1),
        now,
    ));
    assert!(!fast_open_modified_is_recent(
        now - Duration::from_millis(FAST_OPEN_MIN_STABLE_AGE.as_millis() as u64 / 2),
        now,
    ));
    assert!(fast_open_modified_is_recent(
        now + FAST_OPEN_FUTURE_TOLERANCE,
        now,
    ));
    assert!(!fast_open_modified_is_recent(
        now + FAST_OPEN_FUTURE_TOLERANCE + Duration::from_secs(1),
        now,
    ));
}

#[test]
fn recent_office_request_suppresses_only_the_matching_reopen_race() {
    let temporary = tempfile::TempDir::new().expect("temporary request root should exist");
    let sessions = temporary.path().join("OfficeSessions");
    let session = sessions.join("12345678-1234-4234-9234-123456789abc");
    fs::create_dir_all(&session).expect("temporary Session directory should exist");
    let request = session.join(REQUEST_FILE);
    fs::write(&request, b"{}\n").expect("temporary request should be writable");
    let modified = fs::metadata(&request)
        .and_then(|metadata| metadata.modified())
        .expect("temporary request modification time should be readable");

    assert!(has_recent_office_editor_request_in_roots(
        modified + Duration::from_secs(1),
        Duration::from_secs(3),
        [sessions.clone()],
    ));
    assert!(!has_recent_office_editor_request_in_roots(
        modified + Duration::from_secs(4),
        Duration::from_secs(3),
        [sessions],
    ));
}

#[test]
fn office_editor_window_size_uses_measured_screen_ratio_and_safe_bounds() {
    let current_monitor = OfficeEditorMonitorGeometry {
        screen_width: 1470.0,
        screen_height: 956.0,
        maximum_inner_width: 1470.0,
        maximum_inner_height: 900.0,
    };
    let default = resolve_office_editor_window_size(None, Some(current_monitor));
    assert!((default.width - 843.0).abs() < 0.001);
    assert!((default.height - 568.0).abs() < 0.001);

    let larger_monitor = OfficeEditorMonitorGeometry {
        screen_width: 1920.0,
        screen_height: 1080.0,
        maximum_inner_width: 1920.0,
        maximum_inner_height: 1020.0,
    };
    let scaled = resolve_office_editor_window_size(None, Some(larger_monitor));
    assert!((scaled.width / 1920.0 - DEFAULT_OFFICE_EDITOR_WIDTH_RATIO).abs() < 0.001);
    assert!((scaled.height / 1080.0 - DEFAULT_OFFICE_EDITOR_HEIGHT_RATIO).abs() < 0.001);

    let legacy = resolve_office_editor_window_size(
        Some(OfficeEditorWindowSizePreference {
            width: Some(930.0),
            height: Some(700.0),
            ..Default::default()
        }),
        Some(current_monitor),
    );
    assert!((legacy.width - 930.0).abs() < 0.001);
    assert!((legacy.height - 700.0).abs() < 0.001);

    let minimum = normalize_office_editor_window_size(1.0, 2.0, Some(current_monitor));
    assert_eq!(minimum.width, MIN_OFFICE_EDITOR_WIDTH);
    assert_eq!(minimum.height, MIN_OFFICE_EDITOR_HEIGHT);

    let maximum = normalize_office_editor_window_size(
        f64::INFINITY,
        99_999.0,
        Some(current_monitor),
    );
    assert_eq!(maximum.width, DEFAULT_OFFICE_EDITOR_FALLBACK_WIDTH);
    assert_eq!(maximum.height, current_monitor.maximum_inner_height);
}

#[test]
fn editor_ready_marker_exposes_machine_readable_frontend_stages() {
    let marker = MacOfflineOfficeEditorReadyMarker {
        schema: "visualtex-office-editor-ready-v1",
        session_id: "12345678-1234-4234-9234-123456789abc".to_string(),
        host: OfficeHost::Word,
        generation: 4,
        epoch_ms: 900,
        url_received_epoch_ms: 100,
        frontend_epoch_ms: 850,
        hydrate_ms: 120.5,
        editor_mounted_ms: 150.25,
        content_ready_ms: 166.75,
        show_focus_ms: 171.0,
        window_focused: true,
        window_visible: true,
        app_active: true,
        window_can_become_key: true,
        window_is_main: true,
    };
    let value = serde_json::to_value(marker).expect("ready marker should serialize");
    assert_eq!(value["sessionId"], "12345678-1234-4234-9234-123456789abc");
    assert_eq!(value["host"], "word");
    assert_eq!(value["hydrateMs"], 120.5);
    assert_eq!(value["contentReadyMs"], 166.75);
    assert_eq!(value["showFocusMs"], 171.0);
}

#[test]
fn runtime_roots_use_each_office_hosts_application_scripts_directory() {
    let word = runtime_root(OfficeHost::Word).expect("Word runtime root should resolve");
    let powerpoint =
        runtime_root(OfficeHost::Powerpoint).expect("PowerPoint runtime root should resolve");
    assert!(word.ends_with(WORD_RUNTIME_SUFFIX));
    assert!(powerpoint.ends_with(POWERPOINT_RUNTIME_SUFFIX));
    assert_ne!(word, powerpoint);
    assert!(!word.to_string_lossy().contains("UBF8T346G9.Office"));
    assert!(!powerpoint.to_string_lossy().contains("UBF8T346G9.Office"));
    assert!(!word.starts_with("/private/tmp"));
    assert!(!powerpoint.starts_with("/private/tmp"));
}

#[test]
fn word_image_geometry_scales_from_the_14_point_reference() {
    let small = scale_word_reference_geometry(140.0, 28.0, -4.0, 10.5)
        .expect("10.5 pt geometry should scale");
    let large = scale_word_reference_geometry(140.0, 28.0, -4.0, 18.0)
        .expect("18 pt geometry should scale");

    assert!((small.width - 105.0).abs() < 0.001);
    assert!((small.height - 21.0).abs() < 0.001);
    assert_eq!(small.baseline, -3);
    assert!((large.width - 180.0).abs() < 0.001);
    assert!((large.height - 36.0).abs() < 0.001);
    assert_eq!(large.baseline, -5);
    assert!((large.width / small.width - 18.0 / 10.5).abs() < 0.001);
}

#[test]
fn word_inline_baseline_rounds_only_after_font_size_scaling() {
    // Real MathJax geometry from the reported L^2 / L_z regression at the
    // canonical 14 pt export size. Their fractional descents are about
    // 0.825 pt and 3.25512 pt respectively.
    let superscript = calculate_word_svg_geometry(
        22.86186666666666,
        17.56613333333333,
        16.56613333333333,
        11.0,
    )
    .expect("L^2 geometry should resolve");
    let subscript = calculate_word_svg_geometry(
        22.39893333333333,
        17.69493333333333,
        13.749333333333333,
        11.0,
    )
    .expect("L_z geometry should resolve");

    assert!((superscript.reference_baseline_pt + 0.825).abs() < 0.001);
    assert!((subscript.reference_baseline_pt + 3.25512).abs() < 0.001);
    assert_eq!(superscript.baseline, -1);
    assert_eq!(subscript.baseline, -3);
    assert_eq!(subscript.baseline - superscript.baseline, -2);
}

#[test]
fn word_times_image_geometry_matches_native_visual_calibration() {
    // Real 11 pt Times New Roman a^2+b^2 geometry from the reported Word
    // document. The old uniform 1.1 calibration produced about 38.245 x
    // 12.379 pt; Word copy-as-picture measured its ink at 36 x 12 px while
    // the same OMML was 35 x 11 px. Times therefore keeps only the small
    // horizontal TeX-layout compensation and uses natural vertical scale.
    // Its -1.49 pt formula-specific descent floors to -2 pt so Word does
    // not render the image baseline above the adjacent native OMath.
    let times = calculate_word_svg_geometry_for_font(
        59.00053333333333,
        19.0968,
        16.566133333333333,
        11.0,
        Some("times"),
    )
    .expect("Times image geometry should resolve");
    let katex = calculate_word_svg_geometry_for_font(
        59.00053333333333,
        19.0968,
        16.566133333333333,
        11.0,
        Some("katex"),
    )
    .expect("KaTeX image geometry should retain its historical calibration");

    assert!((times.width - 37.09763891428571).abs() < 0.001);
    assert!((times.height - 11.25347142857143).abs() < 0.001);
    assert_eq!(times.baseline, -2);
    assert!((times.reference_baseline_pt + 1.898).abs() < 0.001);
    assert!((katex.width - 38.24498857142857).abs() < 0.001);
    assert!((katex.height - 12.378818571428573).abs() < 0.001);
    assert_eq!(katex.baseline, -2);
}

#[test]
fn word_non_times_fonts_preserve_svg_mathematical_descent() {
    for font in ["katex", "cambria", "stix", "palatino", "helvetica"] {
        let shallow = calculate_word_svg_geometry_for_font(
            20.83792,
            17.789648,
            16.5848,
            11.0,
            Some(font),
        )
        .expect("non-Times shallow geometry should resolve");
        let deep = calculate_word_svg_geometry_for_font(
            17.21108,
            21.760533,
            14.1756,
            11.0,
            Some(font),
        )
        .expect("non-Times fraction geometry should resolve");

        assert_eq!(shallow.baseline, -1, "font={font}");
        assert!((shallow.reference_baseline_pt + (17.789648 - 16.5848) * 0.825).abs() < 0.001);
        assert_eq!(deep.baseline, -5, "font={font}");
        assert!(deep.reference_baseline_pt < -6.0);
    }
}

#[test]
fn latex_redraw_source_limits_and_utf16_ranges_are_strict() {
    assert!(validate_latex_redraw_source_size(0).is_err());
    assert!(validate_latex_redraw_source_size(1).is_ok());
    assert!(validate_latex_redraw_source_size(MAX_LATEX_REDRAW_SOURCE_BYTES).is_ok());
    assert!(validate_latex_redraw_source_size(MAX_LATEX_REDRAW_SOURCE_BYTES + 1).is_err());

    let source = "前😀 $x$ 后 $$y$$";
    let source_utf16 = source.encode_utf16().collect::<Vec<_>>();
    let first_start = "前😀 ".encode_utf16().count();
    let first_text = "$x$";
    let first_end = first_start + first_text.encode_utf16().count();
    let second_start = first_end + " 后 ".encode_utf16().count();
    let second_text = "$$y$$";
    let second_end = second_start + second_text.encode_utf16().count();
    let mut previous_end = 0usize;

    validate_latex_redraw_range(
        &source_utf16,
        &mut previous_end,
        first_start,
        first_end,
        first_text,
    )
    .expect("the first formula must use Word-compatible UTF-16 offsets");
    validate_latex_redraw_range(
        &source_utf16,
        &mut previous_end,
        second_start,
        second_end,
        second_text,
    )
    .expect("the second non-overlapping formula must validate");
    assert_eq!(previous_end, second_end);

    let mut overlap_end = first_end;
    assert!(validate_latex_redraw_range(
        &source_utf16,
        &mut overlap_end,
        first_start,
        first_end,
        first_text,
    )
    .is_err());
    let mut mismatch_end = 0usize;
    assert!(validate_latex_redraw_range(
        &source_utf16,
        &mut mismatch_end,
        first_start,
        first_end,
        "$z$",
    )
    .is_err());
}

#[test]
fn latex_redraw_request_rejects_invalid_scope_and_output_before_file_access() {
    let session_id = "32345678-1234-4234-9234-123456789abc".to_string();
    let mut request = MacOfflineSessionRequest {
        protocol_version: OFFLINE_PROTOCOL_VERSION,
        session_id: session_id.clone(),
        host: "word".to_string(),
        mode: "create".to_string(),
        operation: Some("latexRedraw".to_string()),
        fork_copied_formula: false,
        formula_id: None,
        display_mode: "inline".to_string(),
        numbered: false,
        native_equation: false,
        native_edit_xml: None,
        source_document_id: Some("Document".to_string()),
        source_object_id: None,
        encoded_metadata: None,
        pending_marker: None,
        font_size_pt: None,
        reference_width_pt: None,
        reference_height_pt: None,
        formula_letter_font: None,
        word_math_font_name: None,
        word_document_math_font_name: None,
        power_point: None,
        document_import: Some(MacOfflineDocumentImportRequest {
            bookmark_name: "VT_D_323456781234423492341234".to_string(),
            default_font_size_pt: 12.0,
            inline_font_size_pt: None,
            display_font_size_pt: None,
            redraw_scope: Some("page".to_string()),
            output_kind: Some("image".to_string()),
            source_kind: None,
        }),
        client: None,
        agent: false,
    };
    assert!(validate_request(&request, &session_id)
        .unwrap_err()
        .contains("scope"));

    let document_import = request.document_import.as_mut().unwrap();
    document_import.redraw_scope = Some("selection".to_string());
    document_import.output_kind = Some("ole".to_string());
    assert!(validate_request(&request, &session_id)
        .unwrap_err()
        .contains("output kind"));

    request.operation = Some("documentImport".to_string());
    let document_import = request.document_import.as_mut().unwrap();
    document_import.output_kind = Some("image".to_string());
    assert!(validate_request(&request, &session_id)
        .unwrap_err()
        .contains("formula transform fields"));

    // Conversion sizes are checked like the default size.
    let document_import = request.document_import.as_mut().unwrap();
    document_import.display_font_size_pt = Some(600.0);
    assert!(validate_request(&request, &session_id)
        .unwrap_err()
        .contains("conversion font size"));
    let document_import = request.document_import.as_mut().unwrap();
    document_import.inline_font_size_pt = Some(10.5);
    document_import.display_font_size_pt = Some(16.0);
    assert!(validate_request(&request, &session_id)
        .unwrap_err()
        .contains("formula transform fields"));
}

#[test]
fn document_image_requires_a_real_png_compatibility_preview() {
    assert!(decode_document_image_fallback_png(None).is_err());

    let transparent_placeholder = BASE64_STANDARD.encode(
        BASE64_STANDARD
            .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M/wHwAEAQH/8l0Z8QAAAABJRU5ErkJggg==")
            .expect("transparent PNG fixture should decode"),
    );
    assert!(decode_document_image_fallback_png(Some(&transparent_placeholder)).is_err());

    let supplied = BASE64_STANDARD.encode(
        BASE64_STANDARD
            .decode("iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR4nGNkYGD4z8DAwMDEAAUADigBA0dwHFEAAAAASUVORK5CYII=")
            .expect("PNG fixture should decode"),
    );
    let decoded = decode_document_image_fallback_png(Some(&supplied))
        .expect("supplied PNG compatibility preview should decode");
    assert!(decoded.starts_with(b"\x89PNG\r\n\x1a\n"));
    assert_eq!(u32::from_be_bytes(decoded[16..20].try_into().unwrap()), 2);
    assert_eq!(u32::from_be_bytes(decoded[20..24].try_into().unwrap()), 2);
}

#[cfg(target_os = "macos")]
#[test]
fn word_svg_staging_docx_is_a_valid_ooxml_zip() {
    let svg = br#"<svg xmlns="http://www.w3.org/2000/svg" width="16" height="8"><path d="M0 0h16v8H0z"/></svg>"#;
    let png = BASE64_STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
        .expect("PNG fixture should decode");
    let package =
        build_word_svg_docx(svg, &png, 16.0, 8.0).expect("Word SVG package should build");
    assert!(package.starts_with(b"PK\x03\x04"));
    assert!(package
        .windows(b"word/media/formula.svg".len())
        .any(|value| { value == b"word/media/formula.svg" }));
    assert!(package
        .windows(b"drawing/2016/SVG/main".len())
        .any(|value| { value == b"drawing/2016/SVG/main" }));

    let directory = tempfile::tempdir().expect("temporary directory should exist");
    let path = directory.path().join("formula-svg.docx");
    fs::write(&path, &package).expect("Word SVG package should be writable");
    let output = Command::new("/usr/bin/unzip")
        .args(["-tqq"])
        .arg(&path)
        .output()
        .expect("macOS unzip should validate the package");
    assert!(
        output.status.success(),
        "generated Word SVG DOCX is invalid: {}",
        String::from_utf8_lossy(&output.stderr)
    );
}

#[cfg(target_os = "macos")]
#[test]
#[ignore = "writes a real Word SVG probe DOCX to VISUALTEX_WORD_SVG_PROBE_PATH"]
fn write_word_svg_probe_docx() {
    let path = std::env::var("VISUALTEX_WORD_SVG_PROBE_PATH")
        .expect("set VISUALTEX_WORD_SVG_PROBE_PATH to an absolute .docx path");
    let svg = br#"<svg xmlns="http://www.w3.org/2000/svg" width="160" height="80" viewBox="0 0 160 80"><rect width="160" height="80" fill="white"/><path d="M20 55L55 20L90 55L125 20" fill="none" stroke="black" stroke-width="8"/></svg>"#;
    let png = BASE64_STANDARD
        .decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")
        .expect("PNG fixture should decode");
    let package =
        build_word_svg_docx(svg, &png, 160.0, 80.0).expect("Word SVG package should build");
    fs::write(path, package).expect("Word SVG probe should be writable");
}

#[test]
fn native_word_documents_are_formula_scoped_and_outlive_sessions() {
    let formula_id = "12345678-1234-4234-9234-123456789abc";
    let path = native_word_document_path(formula_id)
        .expect("native Word document path should resolve");
    let runtime = runtime_root(OfficeHost::Word).expect("Word runtime root should resolve");

    assert!(path.starts_with(&runtime));
    assert_eq!(
        path.parent()
            .and_then(|value| value.file_name())
            .and_then(|value| value.to_str()),
        Some("NativeDocuments")
    );
    assert_eq!(
        path.file_name().and_then(|value| value.to_str()),
        Some("12345678-1234-4234-9234-123456789abc.docx")
    );
    assert!(!path.to_string_lossy().contains("OfficeSessions"));
}

#[test]
fn completed_session_cleanup_removes_only_known_ephemeral_files() {
    let directory =
        std::env::temp_dir().join(format!("visualtex-offline-cleanup-test-{}", Uuid::new_v4()));
    fs::create_dir_all(&directory).expect("test Session directory should be created");
    for name in [
        REQUEST_FILE,
        DISPATCH_FILE,
        RESULT_PNG_FILE,
        RESULT_SVG_FILE,
        "native-math.bin", "native-paragraph-context.bin",
        "native-context-status.txt", "native-context-paragraph.txt", "native-result.txt",
    ] {
        fs::write(directory.join(name), b"temporary").expect("temporary file should exist");
    }
    fs::write(directory.join("keep.txt"), b"keep").expect("unknown file should exist");

    cleanup_session_files_at(&directory, false).expect("known files should be cleaned");
    for name in [
        REQUEST_FILE,
        DISPATCH_FILE,
        RESULT_PNG_FILE,
        RESULT_SVG_FILE,
        "native-math.bin", "native-paragraph-context.bin",
        "native-context-status.txt", "native-context-paragraph.txt", "native-result.txt",
    ] {
        assert!(!directory.join(name).exists());
    }
    assert!(directory.join("keep.txt").is_file());
    assert!(directory.is_dir());

    fs::remove_file(directory.join("keep.txt")).unwrap();
    cleanup_session_files_at(&directory, false)
        .expect("empty Session directory should be removed");
    assert!(!directory.exists());
}

#[cfg(target_os = "macos")]
#[test]
fn word_svg_decoder_freezes_visualtex_font_text_into_vector_paths() {
    let source = r#"<svg xmlns="http://www.w3.org/2000/svg"><g transform="scale(1,-1)"><text data-c="78" data-visualtex-output-letter-font="helvetica" transform="scale(1,-1)" font-size="1000px" font-family="Helvetica, Arial, sans-serif" font-style="italic">x</text></g></svg>"#;
    let encoded = BASE64_STANDARD.encode(source.as_bytes());
    let stabilized = decode_stabilized_word_svg(&encoded)
        .expect("Word SVG should freeze host-font text before packaging");
    let stabilized = std::str::from_utf8(&stabilized).unwrap();
    assert!(!stabilized.contains("<text"));
    assert!(!stabilized.contains("font-family"));
    assert!(stabilized.contains("data-visualtex-stable-font-outline=\"true\""));
    assert!(stabilized.contains("<path d=\"M"));
}

#[test]
fn powerpoint_svg_decoder_accepts_internal_vector_references_only() {
    let safe = BASE64_STANDARD.encode(
        br##"<svg xmlns="http://www.w3.org/2000/svg"><defs><path id="g" d="M0 0h1v1z"/></defs><use href="#g"/></svg>"##,
    );
    let decoded = decode_svg(&safe).expect("generated SVG should be accepted");
    assert!(std::str::from_utf8(&decoded).unwrap().contains("<use"));

    let external = BASE64_STANDARD.encode(
        br#"<svg xmlns="http://www.w3.org/2000/svg"><image href="https://example.com/a.png"/></svg>"#,
    );
    assert!(decode_svg(&external).is_err());
    let scripted = BASE64_STANDARD
        .encode(br#"<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>"#);
    assert!(decode_svg(&scripted).is_err());
}

#[cfg(target_os = "macos")]
#[test]
#[ignore = "requires target/source PowerPoint Sessions and explicit environment variables"]
fn live_powerpoint_svg_commit_uses_the_real_ppam_transaction() {
    let session_id = std::env::var("VISUALTEX_LIVE_PPT_SESSION")
        .expect("set VISUALTEX_LIVE_PPT_SESSION to an open PowerPoint create Session");
    let request = read_request(&session_id).expect("PowerPoint request should be readable");
    assert_eq!(request.host, "powerpoint");
    assert!(request.mode == "create" || request.mode == "edit");
    let formula_id = request
        .formula_id
        .clone()
        .expect("PowerPoint request should contain a formula id");
    let source_session_id = std::env::var("VISUALTEX_LIVE_PPT_EXPORT_SESSION").expect(
        "set VISUALTEX_LIVE_PPT_EXPORT_SESSION to a completed VisualTeX formula Session",
    );
    validate_uuid(&source_session_id, "Source Session id").unwrap();
    let home = std::env::var("HOME").expect("HOME should be set on macOS");
    let source_session_path = PathBuf::from(home)
        .join("Library/Application Support/com.visualtex.studio/office/sessions")
        .join(&source_session_id)
        .join("session.json");
    let mut session: OfficeFormulaSession = serde_json::from_slice(
        &fs::read(&source_session_path).expect("source VisualTeX Session should be readable"),
    )
    .expect("source VisualTeX Session should decode");
    assert_eq!(session.host, OfficeHost::Powerpoint);
    let source_export = session
        .export_result
        .as_ref()
        .expect("source VisualTeX Session must contain a real formula export");
    decode_svg(&source_export.svg_base64)
        .expect("source VisualTeX Session must contain a validated SVG export");
    session.id = session_id.clone();
    session.mode = if request.mode == "edit" {
        OfficeSessionMode::Edit
    } else {
        OfficeSessionMode::Create
    };
    session.formula_id = formula_id;
    session.source_document_id = request.source_document_id.clone();
    session.source_object_id = request.source_object_id.clone();
    session.original_metadata = request
        .encoded_metadata
        .as_deref()
        .map(decode_metadata)
        .transpose()
        .expect("target PowerPoint metadata should decode");
    session.dirty = true;
    session.status = OfficeSessionStatus::Committing;
    session.explicit_cancel = false;
    session.error = None;

    let root =
        std::env::temp_dir().join(format!("visualtex-live-powerpoint-svg-{}", Uuid::new_v4()));
    let paths = crate::office::state::OfficePaths {
        certificate: root.join("localhost-cert.pem"),
        private_key: root.join("localhost-key.pem"),
        certificate_metadata: root.join("certificate.json"),
        install: root.join("install.json"),
        sessions: root.join("sessions"),
        recovery: root.join("recovery"),
        formula_cache: root.join("formulas"),
        root: root.clone(),
    };
    let session_store = crate::office::sessions::SessionStore::new(&paths)
        .expect("live Session store should initialize");
    let formula_cache = crate::office::formula_cache::FormulaMetadataCache::new(&paths)
        .expect("live formula cache should initialize");
    let state = OfficeCompanionState::new(
        None,
        crate::ocr::OcrState::default(),
        paths,
        "a".repeat(64),
        session_store,
        formula_cache,
        true,
    );
    let metadata =
        encode_metadata(&metadata_from_session(&session)).expect("live metadata should encode");
    let powerpoint = request
        .power_point
        .as_ref()
        .expect("live PowerPoint request should contain geometry");
    let geometry = calculate_powerpoint_geometry(powerpoint, &session)
        .expect("live PowerPoint geometry should resolve");
    commit_powerpoint(None, &request, &session, &metadata, geometry)
        .expect("real PowerPoint PPAM SVG transaction should succeed");
    let svg_path = result_svg_path(OfficeHost::Powerpoint, &session_id)
        .expect("SVG result path should resolve");
    assert_eq!(
        fs::read_to_string(svg_path).unwrap(),
        session.export_result.as_ref().unwrap().svg
    );
    fs::remove_dir_all(root).ok();
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
#[ignore = "requires an explicitly prepared native PowerPoint edit in the owned acceptance presentation"]
fn live_powerpoint_native_edit_source() {
    let id=std::env::var("VISUALTEX_NATIVE_EDIT_SESSION").expect("Native edit session id required");
    let request=read_request(&id).unwrap();
    assert!(request.native_equation && request.host=="powerpoint" && request.mode=="edit");
    assert!(request.formula_id.is_none() && request.encoded_metadata.is_none());
    assert!(request.source_document_id.as_deref().unwrap().ends_with("/NativeOmmlAcceptance.pptm"));
    let recheck=std::env::var("VISUALTEX_NATIVE_EDIT_RECHECK").unwrap_or_default()=="1";
    let result=capture_powerpoint_native_edit_source(None,&id,recheck);
    let output=match result {
        Ok(Some(source)) => {
            if let Ok(path)=std::env::var("VISUALTEX_NATIVE_EDIT_REPLACEMENT_JSON") {
                let samples:Vec<Value>=serde_json::from_slice(&fs::read(path).unwrap()).unwrap();
                let sample=samples.iter().find(|s|s["name"]=="replacement").unwrap();
                let xml=String::from_utf8(URL_SAFE_NO_PAD.decode(sample["ommlBase64"].as_str().unwrap()).unwrap()).unwrap();
                let generated=super::super::powerpoint_omml::clipboard_entries(&xml,24.0,false,false,false).unwrap();
                let entries=super::super::powerpoint_omml::replace_native_edit_context(generated,&source.clipboard).unwrap();
                fs::write(session_directory(OfficeHost::Powerpoint,&id).unwrap().join("native-edit-replacement.bin"),build_stored_zip(&entries).unwrap()).unwrap();
            }
            match super::super::powerpoint_omml::to_latex::from_omml(&source.omml) {
                Ok(converted) => json!({"ok":true,"latex":converted.latex,"font":converted.letter_font,"omml":source.omml,"mathMl":source.math_ml,"displayMode":request.display_mode}),
                Err(error) => json!({"ok":false,"error":error,"omml":source.omml,"mathMl":source.math_ml}),
            }
        }
        Ok(None) => json!({"completed":true}),
        Err(error) => json!({"ok":false,"error":error}),
    };
    fs::write(session_directory(OfficeHost::Powerpoint,&id).unwrap().join("native-edit-test-result.json"),serde_json::to_vec_pretty(&output).unwrap()).unwrap();
}

#[test]
fn untagged_word_edit_requires_its_exact_session_anchor() {
    let mut request: MacOfflineSessionRequest = serde_json::from_value(json!({
        "protocolVersion": OFFLINE_PROTOCOL_VERSION,
        "sessionId": "32345678-1234-4234-9234-123456789abc",
        "host": "word", "mode": "edit", "displayMode": "block",
        "formulaId": "42345678-1234-4234-9234-123456789abc",
        "numbered": true, "nativeEquation": true,
        "sourceDocumentId": "native-word-test",
        "sourceObjectId": "VT_E_32345678123442349234"
    })).unwrap();
    assert!(is_untagged_word_edit(&request));
    validate_request(&request, &request.session_id).unwrap();
    request.source_object_id = Some("VT_E_another-session".into());
    assert!(validate_request(&request, &request.session_id).is_err());
}

#[test]
fn untagged_word_equation_to_image_is_a_conversion_and_its_anchor_is_released() {
    let mut request: MacOfflineSessionRequest = serde_json::from_value(json!({
        "protocolVersion": OFFLINE_PROTOCOL_VERSION,
        "sessionId": "32345678-1234-4234-9234-123456789abc",
        "host": "word", "mode": "edit", "displayMode": "inline",
        "formulaId": "42345678-1234-4234-9234-123456789abc",
        "numbered": false, "nativeEquation": false, "operation": "nativeToImage",
        "sourceDocumentId": "native-word-test",
        "sourceObjectId": "VT_E_32345678123442349234"
    })).unwrap();
    // Selected Word equation -> image: unchanged in the editor, still applied.
    assert!(word_format_conversion_requested(&request));
    assert!(cancel_needs_host_callback(&request));
    // Editing the same Word equation and keeping it one is not a conversion.
    request.operation = None;
    request.native_equation = true;
    assert!(!word_format_conversion_requested(&request));
    assert!(cancel_needs_host_callback(&request));
    // VisualTeX-made OMML (VT_F_) to an image was already a conversion.
    request.source_object_id = Some("VT_F_41345678-1234-4234-9234-123456789abc".into());
    request.native_equation = false;
    assert!(word_format_conversion_requested(&request));
    assert!(!cancel_needs_host_callback(&request));
}

#[cfg(target_os = "macos")]
#[test]
fn untagged_word_edit_reads_current_math_without_metadata() {
    let xml = r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath>"#;
    let source = word_native_edit_latex(xml, false).unwrap();
    assert!(source.latex.contains(r"\sqrt{x}"), "{}", source.latex);
    assert!(!source.latex.contains("VisualTeX"));
}

#[cfg(target_os = "macos")]
#[test]
fn untagged_word_edit_reads_atomic_request_without_a_snapshot_file() {
    let xml = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:body><w:p><m:oMath><m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad></m:oMath></w:p></w:body></w:document>"#;
    let request: MacOfflineSessionRequest = serde_json::from_value(json!({
        "protocolVersion": OFFLINE_PROTOCOL_VERSION,
        "sessionId": "62345678-1234-4234-9234-123456789abc",
        "host": "word", "mode": "edit", "displayMode": "inline",
        "formulaId": "42345678-1234-4234-9234-123456789abc",
        "numbered": false, "nativeEquation": true,
        "sourceDocumentId": "native-word-test",
        "sourceObjectId": "VT_E_62345678123442349234",
        "nativeEditXml": xml
    })).unwrap();
    validate_request(&request, &request.session_id).unwrap();
    let source = word_native_edit_source(&request).unwrap();
    assert_eq!(source.latex, word_native_edit_latex(xml, false).unwrap().latex);
    assert!(source.latex.contains(r"\sqrt{x}"));
    let round_trip: MacOfflineSessionRequest = serde_json::from_slice(&serde_json::to_vec(&request).unwrap()).unwrap();
    assert_eq!(round_trip.native_edit_xml.as_deref(), Some(xml));
    let mut invalid = request.clone();
    invalid.mode = "create".into();
    assert!(validate_request(&invalid, &invalid.session_id).is_err());
    invalid = request.clone();
    invalid.native_edit_xml = Some(String::new());
    assert!(validate_request(&invalid, &invalid.session_id).is_err());
    invalid = request;
    invalid.native_edit_xml = Some("<m:oMath>".into());
    assert!(word_native_edit_source(&invalid).is_err());
}

fn native_powerpoint_request_fixture() -> MacOfflineSessionRequest {
    serde_json::from_value(json!({
        "protocolVersion": OFFLINE_PROTOCOL_VERSION,
        "sessionId": "32345678-1234-4234-9234-123456789abc",
        "host": "powerpoint", "mode": "create", "displayMode": "inline",
        "numbered": false, "nativeEquation": true,
        "sourceDocumentId": "/Users/测试/演示.pptx",
        "sourceObjectId": "visualtex-ppt-native-text:256:3:4:0",
        "powerPoint": {
            "presentationIdentity": "/Users/测试/演示.pptx",
            "slideIndex": 1, "slideId": 256, "shapeIndex": 1, "shapeId": 3,
            "shapeName": "TextBox 3", "left": 0, "top": 0, "width": 360,
            "height": 180, "rotation": 0, "zOrder": 1, "fontSizePt": 24,
            "nativeTextTarget": { "rangeStart": 4, "rangeLength": 0,
                "leadingParagraph": true, "trailingParagraph": true }
        }
    })).unwrap()
}

#[test]
fn native_powerpoint_accepts_exact_text_targets_without_persistent_metadata() {
    let mut request = native_powerpoint_request_fixture();
    for mode in ["inline", "block"] {
        request.display_mode = mode.into();
        validate_request(&request, &request.session_id).unwrap();
        let roundtrip: MacOfflineSessionRequest = serde_json::from_slice(&serde_json::to_vec(&request).unwrap()).unwrap();
        validate_request(&roundtrip, &request.session_id).unwrap();
        assert!(roundtrip.formula_id.is_none() && roundtrip.pending_marker.is_none() && roundtrip.encoded_metadata.is_none());
        assert_eq!(roundtrip.power_point.unwrap().native_text_target.unwrap().range_start, 4);
    }
}

#[test]
fn native_powerpoint_table_cell_target_is_explicit() {
    let mut value = serde_json::to_value(native_powerpoint_request_fixture()).unwrap();
    let target = &mut value["powerPoint"]["nativeTextTarget"];
    assert_eq!(target["tableCell"], json!(false));
    target["tableCell"] = json!(true);
    target["leadingParagraph"] = json!(false);
    target["trailingParagraph"] = json!(false);
    let request: MacOfflineSessionRequest = serde_json::from_value(value).unwrap();
    validate_request(&request, &request.session_id).unwrap();
    assert!(request.power_point.unwrap().native_text_target.unwrap().table_cell);
    // The add-in leaves the field out for text boxes.
    let mut value = serde_json::to_value(native_powerpoint_request_fixture()).unwrap();
    value["powerPoint"]["nativeTextTarget"].as_object_mut().unwrap().remove("tableCell");
    let request: MacOfflineSessionRequest = serde_json::from_value(value).unwrap();
    assert!(!request.power_point.unwrap().native_text_target.unwrap().table_cell);
}

#[test]
fn native_powerpoint_edit_uses_only_current_office_range_not_formula_identity() {
    let mut request=native_powerpoint_request_fixture();
    request.mode="edit".into();
    let text=request.power_point.as_mut().unwrap().native_text_target.as_mut().unwrap();
    text.range_length=8; text.leading_paragraph=false; text.trailing_paragraph=false;
    request.source_object_id=Some("visualtex-ppt-native-text:256:3:4:8".into());
    validate_request(&request,&request.session_id).unwrap();
    assert!(request.formula_id.is_none() && request.encoded_metadata.is_none());
    request.power_point.as_mut().unwrap().native_text_target.as_mut().unwrap().leading_paragraph=true;
    assert!(validate_request(&request,&request.session_id).is_err());
}

#[test]
fn native_powerpoint_standalone_is_explicit_and_has_no_shape_placeholder() {
    let mut request = native_powerpoint_request_fixture();
    request.source_object_id = Some("visualtex-ppt-native-slide:256".into());
    let geometry = request.power_point.as_mut().unwrap();
    geometry.native_text_target = None;
    geometry.shape_index = 0; geometry.shape_id = 0; geometry.z_order = 0;
    geometry.shape_name.clear();
    validate_request(&request, &request.session_id).unwrap();
    request.power_point.as_mut().unwrap().shape_id = 3;
    assert!(validate_request(&request, &request.session_id).is_err());
}

#[test]
fn native_powerpoint_rejects_ambiguous_or_unsupported_requests() {
    for mutation in 0..7 {
        let mut request = native_powerpoint_request_fixture();
        match mutation {
            0 => request.mode = "edit".into(),
            1 => request.power_point.as_mut().unwrap().native_text_target = None,
            2 => request.power_point.as_mut().unwrap().native_text_target.as_mut().unwrap().range_start = 0,
            3 => request.native_equation = false,
            4 => request.numbered = true,
            5 => request.source_object_id = None,
            6 => request.pending_marker = Some("unexpected-image-placeholder".into()),
            _ => unreachable!(),
        }
        assert!(validate_request(&request, &request.session_id).is_err(), "mutation {mutation}");
    }
}

#[test]
fn offline_request_json_accepts_utf8_office_identities() {
    let session_id = "32345678-1234-4234-9234-123456789abc".to_string();
    let request = MacOfflineSessionRequest {
        protocol_version: OFFLINE_PROTOCOL_VERSION,
        session_id: session_id.clone(),
        host: "word".to_string(),
        mode: "create".to_string(),
        operation: None,
        fork_copied_formula: false,
        formula_id: Some("12345678-1234-4234-9234-123456789abc".to_string()),
        display_mode: "inline".to_string(),
        numbered: false,
        native_equation: false,
        native_edit_xml: None,
        source_document_id: Some("/Users/测试/公式😀.docx".to_string()),
        source_object_id: Some("书签-公式".to_string()),
        encoded_metadata: None,
        pending_marker: Some(
            "visualtex:pending:v1:32345678-1234-4234-9234-123456789abc:12345678-1234-4234-9234-123456789abc"
                .to_string(),
        ),
        font_size_pt: Some(10.5),
        reference_width_pt: Some(60.0),
        reference_height_pt: Some(15.0),
        formula_letter_font: None,
        word_math_font_name: None,
        word_document_math_font_name: None,
        power_point: None,
        document_import: None,
        client: None,
        agent: false,
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
fn word_native_commit_uses_the_visualtex_font_after_an_explicit_font_change() {
    assert_eq!(
        resolved_word_math_font_name(
            Some("Cambria Math"),
            "cambria",
            "times",
        ),
        "Times New Roman",
    );
    assert_eq!(
        resolved_word_math_font_name(
            Some("Times New Roman"),
            "times",
            "stix",
        ),
        "STIX Two Math",
    );
    assert_eq!(
        resolved_word_math_font_name(
            Some("Latin Modern Math"),
            "cambria",
            "cambria",
        ),
        "Latin Modern Math",
    );
    assert_eq!(
        resolved_word_math_font_name(None, "cambria", "palatino"),
        "Palatino",
    );
    assert_eq!(
        resolved_word_math_font_name(None, "cambria", "katex"),
        "Cambria Math",
    );
}

#[test]
fn word_edit_accepts_live_math_font_metadata_only_for_word_edits() {
    let session_id = "31345678-1234-4234-9234-123456789abc".to_string();
    let mut request: MacOfflineSessionRequest = serde_json::from_value(json!({
        "protocolVersion": OFFLINE_PROTOCOL_VERSION,
        "sessionId": session_id,
        "host": "word",
        "mode": "edit",
        "operation": "formula",
        "formulaId": "41345678-1234-4234-9234-123456789abc",
        "displayMode": "block",
        "numbered": false,
        "nativeEquation": true,
        "sourceDocumentId": "Document1",
        "sourceObjectId": "VT_F_41345678-1234-4234-9234-123456789abc",
        "encodedMetadata": format!("{METADATA_PREFIX}fixture"),
        "formulaLetterFont": "times",
        "wordMathFontName": "Latin Modern Math"
    }))
    .unwrap();

    validate_request(&request, &request.session_id)
        .expect("Word native edits should accept a live VisualTeX font and raw Word math font");

    request.mode = "create".to_string();
    assert!(validate_request(&request, &request.session_id)
        .unwrap_err()
        .contains("Word edit requests"));

    request.mode = "edit".to_string();
    request.formula_letter_font = Some("unsupported".to_string());
    assert!(validate_request(&request, &request.session_id)
        .unwrap_err()
        .contains("formula font metadata"));

    request.formula_letter_font = Some("times".to_string());
    request.word_math_font_name = Some("x".repeat(129));
    assert!(validate_request(&request, &request.session_id).is_err());
}

#[test]
fn word_equations_use_the_document_math_font() {
    let mut request: MacOfflineSessionRequest = serde_json::from_value(json!({
        "protocolVersion": OFFLINE_PROTOCOL_VERSION,
        "sessionId": "31345678-1234-4234-9234-123456789abd",
        "host": "word",
        "mode": "create",
        "formulaId": "41345678-1234-4234-9234-123456789abd",
        "displayMode": "block",
        "numbered": true,
        "sourceDocumentId": "Document1",
        "sourceObjectId": null,
        "encodedMetadata": null,
        "pendingMarker": null,
        "wordDocumentMathFontName": "STIX Two Math"
    }))
    .unwrap();
    validate_request(&request, &request.session_id)
        .expect("Word creates carry the document math font");
    assert_eq!(word_equation_math_font(&request), "STIX Two Math");

    request.word_document_math_font_name = None;
    assert_eq!(word_equation_math_font(&request), "");
    request.word_math_font_name = Some("Latin Modern Math".to_string());
    assert_eq!(word_equation_math_font(&request), "Latin Modern Math");

    request.word_math_font_name = None;
    request.word_document_math_font_name = Some("x".repeat(129));
    assert!(validate_request(&request, &request.session_id)
        .unwrap_err()
        .contains("wordDocumentMathFontName"));
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
        inline_image_math_style: Some("text".to_string()),
        numbered: false,
        render_width_px: Some(50.0),
        render_height_px: Some(20.0),
        render_font_size_pt: None,
        font_size_pt: Some(10.5),
        formula_letter_font: None,
        formula_chinese_font: None,
        reference_width_pt: Some(37.5),
        reference_height_pt: Some(15.0),
        reference_baseline_pt: Some(-3.0),
        image_ink_center_y_ratio: Some(0.47826087),
        created_with_version: "1.1.0".to_string(),
        updated_with_version: "1.1.0".to_string(),
        created_at: "unix-ms:1".to_string(),
        updated_at: "unix-ms:1".to_string(),
    };
    let encoded = encode_metadata(&metadata).expect("metadata should encode");
    let decoded = decode_metadata(&encoded).expect("metadata should decode");
    assert_eq!(decoded.formula_id, metadata.formula_id);
    assert_eq!(decoded.lines[0].latex, "x^2");
    assert_eq!(decoded.inline_image_math_style.as_deref(), Some("text"));
    assert_eq!(decoded.font_size_pt, Some(10.5));
    assert_eq!(decoded.reference_width_pt, Some(37.5));
    assert_eq!(decoded.reference_height_pt, Some(15.0));
    assert_eq!(decoded.reference_baseline_pt, Some(-3.0));
    assert_eq!(decoded.image_ink_center_y_ratio, Some(0.47826087));
}

#[test]
fn numbered_formula_metadata_changes_only_the_numbered_member() {
    let mut metadata = document_formula_metadata("raw", &["x"], "x", "block", false);
    metadata.reference_width_pt = Some(221.70280000000002);
    let encoded = encode_metadata(&metadata).unwrap();
    let numbered = numbered_formula_metadata(&encoded).unwrap();
    let before = String::from_utf8(metadata_json(&encoded).unwrap()).unwrap();
    let after = String::from_utf8(metadata_json(&numbered).unwrap()).unwrap();
    assert_eq!(after, before.replacen("\"numbered\":false", "\"numbered\":true", 1));
    assert!(decode_metadata(&numbered).unwrap().numbered);
    assert_eq!(numbered_formula_metadata(&numbered).unwrap(), numbered);
    let inline = encode_metadata(&document_formula_metadata("raw", &["x"], "x", "inline", false)).unwrap();
    assert!(numbered_formula_metadata(&inline).is_err());
}

fn document_formula_metadata(
    code_format: &str,
    lines: &[&str],
    latex: &str,
    display_mode: &str,
    numbered: bool,
) -> VisualTeXFormulaMetadata {
    VisualTeXFormulaMetadata {
        schema: "visualtex-formula".to_string(),
        schema_version: 1,
        formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
        title: "Imported formula".to_string(),
        latex: latex.to_string(),
        lines: lines
            .iter()
            .map(|latex| crate::office::sessions::MetadataLine {
                id: Uuid::new_v4().to_string(),
                latex: (*latex).to_string(),
            })
            .collect(),
        code_format: code_format.to_string(),
        display_mode: display_mode.to_string(),
        inline_image_math_style: Some("text".to_string()),
        numbered,
        render_width_px: None,
        render_height_px: None,
        render_font_size_pt: None,
        font_size_pt: Some(14.0),
        formula_letter_font: None,
        formula_chinese_font: None,
        reference_width_pt: None,
        reference_height_pt: None,
        reference_baseline_pt: None,
        image_ink_center_y_ratio: None,
        created_with_version: "1.2.6".to_string(),
        updated_with_version: "1.2.6".to_string(),
        created_at: "unix-ms:1".to_string(),
        updated_at: "unix-ms:1".to_string(),
    }
}

#[test]
fn word_numbering_only_fast_path_requires_unchanged_source_and_appearance() {
    let mut request = native_powerpoint_request_fixture();
    request.host = "word".to_string();
    request.mode = "edit".to_string();
    request.native_equation = false;
    request.operation = None;
    let original = document_formula_metadata("raw", &["x^2"], "x^2", "block", true);
    let mut current = original.clone();
    current.numbered = false;
    assert!(word_image_numbering_only_edit(
        &request,
        Some(&original),
        &current
    ));
    for field in ["latex", "font", "size", "identity", "format", "mode"] {
        let mut changed = current.clone();
        match field {
            "latex" => changed.latex = "y^2".to_string(),
            "font" => changed.formula_letter_font = Some("times".to_string()),
            "size" => changed.font_size_pt = Some(18.0),
            "identity" => changed.formula_id = Uuid::new_v4().to_string(),
            "format" => changed.code_format = "align".to_string(),
            _ => changed.display_mode = "inline".to_string(),
        }
        assert!(
            !word_image_numbering_only_edit(&request, Some(&original), &changed),
            "{field}"
        );
    }
    current.numbered = true;
    assert!(!word_image_numbering_only_edit(
        &request,
        Some(&original),
        &current
    ));
    current.numbered = false;
    assert!(!word_image_numbering_only_edit(&request, None, &current));
    request.native_equation = true;
    assert!(!word_image_numbering_only_edit(
        &request,
        Some(&original),
        &current
    ));
    request.native_equation = false;
    request.operation = Some("nativeToImage".to_string());
    assert!(!word_image_numbering_only_edit(
        &request,
        Some(&original),
        &current
    ));
}

#[test]
fn copied_formula_metadata_gets_independent_formula_and_line_ids() {
    let mut metadata = document_formula_metadata(
        "align",
        &["a=b", "c=d"],
        "a=b\nc=d",
        "block",
        false,
    );
    let source_formula_id = metadata.formula_id.clone();
    let source_line_ids = metadata
        .lines
        .iter()
        .map(|line| line.id.clone())
        .collect::<Vec<_>>();
    let copied_formula_id = "22345678-1234-4234-9234-123456789abc";

    fork_copied_metadata_identity(&mut metadata, copied_formula_id)
        .expect("copied formula metadata should be rekeyed");

    assert_ne!(metadata.formula_id, source_formula_id);
    assert_eq!(metadata.formula_id, copied_formula_id);
    assert_eq!(metadata.lines.len(), source_line_ids.len());
    for (line, source_id) in metadata.lines.iter().zip(source_line_ids) {
        assert_ne!(line.id, source_id);
        validate_uuid(&line.id, "Copied line id").unwrap();
    }
}

#[test]
fn word_formula_edit_requires_explicit_source_target() {
    let session_id = "31345678-1234-4234-9234-123456789abc".to_string();
    let mut request = MacOfflineSessionRequest {
        protocol_version: OFFLINE_PROTOCOL_VERSION,
        session_id: session_id.clone(),
        host: "word".to_string(),
        mode: "edit".to_string(),
        operation: Some("formula".to_string()),
        fork_copied_formula: false,
        formula_id: Some("41345678-1234-4234-9234-123456789abc".to_string()),
        display_mode: "block".to_string(),
        numbered: true,
        native_equation: false,
        native_edit_xml: None,
        source_document_id: Some("Document1".to_string()),
        source_object_id: None,
        encoded_metadata: Some(format!("{METADATA_PREFIX}fixture")),
        pending_marker: None,
        font_size_pt: Some(11.0),
        reference_width_pt: Some(72.0),
        reference_height_pt: Some(18.0),
        formula_letter_font: None,
        word_math_font_name: None,
        word_document_math_font_name: None,
        power_point: None,
        document_import: None,
        client: None,
        agent: false,
    };

    assert!(validate_request(&request, &session_id)
        .unwrap_err()
        .contains("explicit source target"));

    request.source_object_id = Some("VT_E_31345678123442349234".to_string());
    validate_request(&request, &session_id)
        .expect("image formula edits with an explicit edit bookmark should validate");

    request.native_equation = true;
    request.source_object_id = Some(
        "VT_F_41345678-1234-4234-9234-123456789abc".to_string(),
    );
    validate_request(&request, &session_id)
        .expect("native formula edits with an explicit formula bookmark should validate");
}

#[test]
fn copied_formula_fork_validation_accepts_word_conversion_edits() {
    let session_id = "32345678-1234-4234-9234-123456789abc".to_string();
    let formula_id = "42345678-1234-4234-9234-123456789abc".to_string();
    let encoded_metadata = format!("{METADATA_PREFIX}fixture");
    let mut request = MacOfflineSessionRequest {
        protocol_version: OFFLINE_PROTOCOL_VERSION,
        session_id: session_id.clone(),
        host: "word".to_string(),
        mode: "edit".to_string(),
        operation: Some("imageToNative".to_string()),
        fork_copied_formula: true,
        formula_id: Some(formula_id),
        display_mode: "inline".to_string(),
        numbered: false,
        native_equation: true,
        native_edit_xml: None,
        source_document_id: Some("Document1".to_string()),
        source_object_id: Some("VT_E_32345678123442349234".to_string()),
        encoded_metadata: Some(encoded_metadata),
        pending_marker: None,
        font_size_pt: Some(14.0),
        reference_width_pt: Some(42.0),
        reference_height_pt: Some(18.0),
        formula_letter_font: None,
        word_math_font_name: None,
        word_document_math_font_name: None,
        power_point: None,
        document_import: None,
        client: None,
        agent: false,
    };

    validate_request(&request, &session_id)
        .expect("a copied Word image should be forkable while converting to OMML");

    request.operation = Some("nativeToImage".to_string());
    request.native_equation = false;
    validate_request(&request, &session_id)
        .expect("a copied Word OMML equation should be forkable while converting to image");

    request.native_equation = true;
    assert!(validate_request(&request, &session_id)
        .unwrap_err()
        .contains("matching Word edit output request"));
}

#[test]
fn windows_picture_keeps_its_font_size() {
    // E=mc^2 VisualTeX for Windows drew at 20 pt: 105.69 x 24.37 px, a 79.27 x 18.28 pt picture without the
    // Mac's scale tags (the add-in then reports its default 18 pt)
    let mut metadata = document_formula_metadata("latex", &["E=mc^2"], "E=mc^2", "inline", false);
    metadata.render_width_px = Some(105.694);
    metadata.render_height_px = Some(24.373);
    metadata.font_size_pt = Some(20.0);
    metadata.render_font_size_pt = Some(20.0);
    let mut powerpoint = MacOfflinePowerPointRequest {
        native_text_target: None,
        language_id: None,
        presentation_identity: "Deck".to_string(),
        slide_index: 1,
        slide_id: 2,
        shape_index: 3,
        shape_id: 3,
        shape_name: "VisualTeX_12345678-1234-4234-9234-123456789abc".to_string(),
        left: 60.0,
        top: 140.0,
        width: 79.27,
        height: 18.28,
        rotation: 0.0,
        z_order: 2,
        font_size_pt: Some(18.0),
        reference_width_pt: Some(61.65),
        reference_height_pt: Some(14.22),
    };
    assert_eq!(windows_picture_font_size(Some(&powerpoint), Some(&metadata)), Some(20.0));
    // twice as tall: resized to 40 pt
    powerpoint.height = 36.56;
    assert_eq!(windows_picture_font_size(Some(&powerpoint), Some(&metadata)), Some(40.0));
    // its scale tags were there (another size than the default): they stand
    powerpoint.font_size_pt = Some(24.0);
    assert_eq!(windows_picture_font_size(Some(&powerpoint), Some(&metadata)), None);
    // a Mac picture (no renderFontSizePt): the add-in's size stands
    powerpoint.font_size_pt = Some(18.0);
    metadata.render_font_size_pt = None;
    assert_eq!(windows_picture_font_size(Some(&powerpoint), Some(&metadata)), None);
}

#[test]
fn canonical_latex_of_a_windows_formula() {
    // VisualTeX for Windows stores codeFormat "latex": the rows as they are
    let metadata = document_formula_metadata("latex", &["a&=b+c", "d&=e-f"], "a&=b+c\nd&=e-f", "block", false);
    assert_eq!(canonical_document_formula_latex(&metadata).unwrap(), "a&=b+c\nd&=e-f");
}

#[test]
fn canonical_latex_of_the_mixed_copy_format() {
    let metadata = document_formula_metadata(
        "mixed-inline-display",
        &["a = b", "c"],
        "$$a = b$$\n$$c$$",
        "block",
        false,
    );
    assert_eq!(canonical_document_formula_latex(&metadata).unwrap(), "$$a = b$$\n$$c$$");
}

#[test]
fn document_formula_metadata_rebuilds_canonical_multiline_environments() {
    let cases = [
        (
            "align",
            vec!["a = b + c", "d = e"],
            r"\begin{align}
a &= b + c \\
d &= e
\end{align}",
        ),
        (
            "align-star",
            vec!["x = y", "y = z"],
            r"\begin{align*}
x &= y \\
y &= z
\end{align*}",
        ),
        (
            "aligned",
            vec!["p = q", "r = s"],
            r"\[
\begin{aligned}
p &= q \\
r &= s
\end{aligned}
\]",
        ),
        (
            "gather",
            vec!["a=b", "c=d"],
            r"\begin{gather}
a=b \\
c=d
\end{gather}",
        ),
        (
            "multline-star",
            vec!["a+b+c", "=d+e"],
            r"\begin{multline*}
a+b+c \\
=d+e
\end{multline*}",
        ),
        (
            "equation-split",
            vec!["a = b", "c = d"],
            r"\begin{equation}
\begin{split}
a &= b \\
c &= d
\end{split}
\end{equation}",
        ),
        (
            "equation-star-split",
            vec!["a = b", "c = d"],
            r"\begin{equation*}
\begin{split}
a &= b \\
c &= d
\end{split}
\end{equation*}",
        ),
    ];

    for (code_format, lines, expected) in cases {
        let metadata = document_formula_metadata(code_format, &lines, expected, "block", false);
        assert_eq!(
            canonical_document_formula_latex(&metadata).unwrap(),
            expected,
            "{code_format} canonical source"
        );
        assert_eq!(
            validate_document_formula_metadata_match(
                &metadata,
                &metadata.formula_id,
                expected,
                "block",
                false,
            )
            .unwrap(),
            expected,
            "{code_format} metadata match"
        );
    }
}

#[test]
fn document_formula_metadata_preserves_internal_equation_newlines() {
    let body = r"u(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}c_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y,\qquad
f(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}d_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y.";
    let canonical = r"\begin{equation*}
u(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}c_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y,\qquad
f(x,y)=\sum_{n=1}^{+\infty}\sum_{m=1}^{+\infty}d_{nm}\sin\frac{n\pi}{a}x\sin\frac{m\pi}{b}y.
\end{equation*}";
    let metadata =
        document_formula_metadata("equation-star", &[body], canonical, "block", false);

    assert_eq!(
        canonical_document_formula_latex(&metadata).unwrap(),
        canonical
    );
    assert_eq!(
        validate_document_formula_metadata_match(
            &metadata,
            &metadata.formula_id,
            canonical,
            "block",
            false,
        )
        .unwrap(),
        canonical,
    );

    let separate_rows = document_formula_metadata(
        "equation-star",
        &["a=b", "c=d"],
        r"\begin{equation*}
a=b
\end{equation*}

\begin{equation*}
c=d
\end{equation*}",
        "block",
        false,
    );
    assert_eq!(
        canonical_document_formula_latex(&separate_rows).unwrap(),
        separate_rows.latex,
        "separate logical rows must remain separate equation environments",
    );
}

#[test]
fn document_formula_metadata_accepts_frontend_environment_spacing() {
    let frontend_source = r"\begin {equation}
E=mc^2
\end {equation}";
    let metadata =
        document_formula_metadata("equation", &["E=mc^2"], frontend_source, "block", true);

    assert_ne!(
        canonical_document_formula_latex(&metadata).unwrap(),
        frontend_source,
        "the Rust formatter intentionally normalizes environment spacing",
    );
    assert_eq!(
        validate_document_formula_metadata_match(
            &metadata,
            &metadata.formula_id,
            frontend_source,
            "block",
            true,
        )
        .unwrap(),
        frontend_source,
        "a valid frontend serialization must not be rejected because Rust formats it differently",
    );
}

#[test]
fn document_formula_metadata_match_rejects_structural_drift() {
    let canonical = r"\begin{align}
a &= b \\
c &= d
\end{align}";
    let metadata =
        document_formula_metadata("align", &["a = b", "c = d"], canonical, "block", true);
    assert!(validate_document_formula_metadata_match(
        &metadata,
        &metadata.formula_id,
        canonical,
        "block",
        true,
    )
    .is_ok());
    assert!(validate_document_formula_metadata_match(
        &metadata,
        "22345678-1234-4234-9234-123456789abc",
        canonical,
        "block",
        true,
    )
    .is_err());
    assert!(validate_document_formula_metadata_match(
        &metadata,
        &metadata.formula_id,
        canonical,
        "inline",
        true,
    )
    .is_err());
    assert!(validate_document_formula_metadata_match(
        &metadata,
        &metadata.formula_id,
        canonical,
        "block",
        false,
    )
    .is_err());
    assert!(validate_document_formula_metadata_match(
        &metadata,
        &metadata.formula_id,
        r"\begin{align}
a &= b \\
c &= e
\end{align}",
        "block",
        true,
    )
    .is_err());

    let mut stale_metadata = metadata.clone();
    stale_metadata.latex = "a = b\nc = d".to_string();
    assert!(validate_document_formula_metadata_match(
        &stale_metadata,
        &stale_metadata.formula_id,
        canonical,
        "block",
        true,
    )
    .is_err());
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
fn office_geometry_preserves_visual_point_size_and_powerpoint_center() {
    let word_geometry = calculate_word_svg_geometry(100.0, 20.0, 15.0, 14.0)
        .expect("Word image geometry should apply its visual calibration");
    assert!((word_geometry.width - 82.5).abs() < 0.001);
    assert!((word_geometry.height - 16.5).abs() < 0.001);
    assert_eq!(word_geometry.baseline, -4);
    assert!((word_geometry.reference_width_pt - 82.5).abs() < 0.001);
    assert!((word_geometry.reference_height_pt - 16.5).abs() < 0.001);

    let request = MacOfflinePowerPointRequest {
        native_text_target: None,
        language_id: None,
        presentation_identity: "Deck".to_string(),
        slide_index: 1,
        slide_id: 2,
        shape_index: 3,
        shape_id: 3,
        shape_name: "VisualTeX_12345678-1234-4234-9234-123456789abc".to_string(),
        left: 100.0,
        top: 200.0,
        width: 120.0,
        height: 40.0,
        rotation: 0.0,
        z_order: 2,
        font_size_pt: None,
        reference_width_pt: None,
        reference_height_pt: None,
    };
    let session = OfficeFormulaSession {
        id: "32345678-1234-4234-9234-123456789abc".to_string(),
        mode: OfficeSessionMode::Edit,
        host: OfficeHost::Powerpoint,
        client: None,
        operation: None,
        native_equation: false,
        formula_id: "12345678-1234-4234-9234-123456789abc".to_string(),
        source_document_id: None,
        source_object_id: None,
        title: "Formula".to_string(),
        lines: vec![],
        active_line_id: None,
        code_format: "latex".to_string(),
        display_mode: "block".to_string(),
        inline_image_math_style: "text".to_string(),
        numbered: false,
        font_size_pt: None,
        formula_letter_font: "katex".to_string(),
        formula_chinese_font: "system".to_string(),
        export_width: 0.0,
        export_height: 0.0,
        export_result: Some(crate::office::sessions::OfficeExportResult {
            svg: "<svg/>".to_string(),
            svg_base64: String::new(),
            png_base64: None,
            omml_base64: None,
            omml_docx_base64: None,
            word_math_ml: None,
            width: 300.0,
            height: 50.0,
            baseline: None,
            ink_top_ratio: None,
            ink_bottom_ratio: None,
            ink_center_y_ratio: None,
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
            inline_image_math_style: Some("text".to_string()),
            numbered: false,
            render_width_px: Some(120.0),
            render_height_px: Some(40.0),
            render_font_size_pt: None,
            font_size_pt: None,
            formula_letter_font: None,
            formula_chinese_font: None,
            reference_width_pt: None,
            reference_height_pt: None,
            reference_baseline_pt: None,
            image_ink_center_y_ratio: None,
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
    let geometry =
        calculate_powerpoint_geometry(&request, &session).expect("geometry should scale");
    assert!((geometry.height - 50.0).abs() < 0.001);
    assert!((geometry.width - 300.0).abs() < 0.001);
    assert!((geometry.left + geometry.width / 2.0 - 160.0).abs() < 0.001);
    assert!((geometry.top + geometry.height / 2.0 - 220.0).abs() < 0.001);
    assert!((geometry.font_size_pt - 18.6666666667).abs() < 0.001);
    assert!((geometry.reference_width_pt - 225.0).abs() < 0.001);
    assert!((geometry.reference_height_pt - 37.5).abs() < 0.001);

    let mut create_request = request.clone();
    create_request.font_size_pt = Some(28.0);
    create_request.reference_width_pt = None;
    create_request.reference_height_pt = None;
    let mut create_session = session.clone();
    create_session.mode = OfficeSessionMode::Create;
    create_session.original_metadata = None;
    let create_geometry = calculate_powerpoint_geometry(&create_request, &create_session)
        .expect("declared PowerPoint point size should scale the SVG");
    assert!((create_geometry.font_size_pt - 28.0).abs() < 0.001);
    assert!((create_geometry.width - 450.0).abs() < 0.001);
    assert!((create_geometry.height - 75.0).abs() < 0.001);
    assert!((create_geometry.left + create_geometry.width / 2.0 - 160.0).abs() < 0.001);
    assert!((create_geometry.top + create_geometry.height / 2.0 - 220.0).abs() < 0.001);

    let mut edited_request = request.clone();
    edited_request.font_size_pt = Some(18.0);
    edited_request.reference_width_pt = Some(225.0);
    edited_request.reference_height_pt = Some(37.5);
    let mut edited_session = session.clone();
    edited_session.font_size_pt = Some(32.0);
    let edited_geometry = calculate_powerpoint_geometry(&edited_request, &edited_session)
        .expect("the editor-selected PowerPoint size should override the launch geometry");
    assert!((edited_geometry.font_size_pt - 32.0).abs() < 0.001);
    assert!(
        (edited_geometry.width / edited_geometry.reference_width_pt - 32.0 / 14.0).abs()
            < 0.001
    );
    assert!(
        (edited_geometry.height / edited_geometry.reference_height_pt - 32.0 / 14.0).abs()
            < 0.001
    );

    let word_request = MacOfflineSessionRequest {
        protocol_version: OFFLINE_PROTOCOL_VERSION,
        session_id: "42345678-1234-4234-9234-123456789abc".to_string(),
        host: "word".to_string(),
        mode: "edit".to_string(),
        operation: None,
        fork_copied_formula: false,
        formula_id: Some("12345678-1234-4234-9234-123456789abc".to_string()),
        display_mode: "inline".to_string(),
        numbered: false,
        native_equation: false,
        native_edit_xml: None,
        source_document_id: Some("Document".to_string()),
        source_object_id: Some("VT_F_12345678-1234-4234-9234-123456789abc".to_string()),
        encoded_metadata: None,
        pending_marker: None,
        font_size_pt: Some(10.5),
        reference_width_pt: Some(60.0),
        reference_height_pt: Some(15.0),
        formula_letter_font: None,
        word_math_font_name: None,
        word_document_math_font_name: None,
        power_point: None,
        document_import: None,
        client: None,
        agent: false,
    };
    let mut word_session = session.clone();
    word_session.host = OfficeHost::Word;
    word_session.font_size_pt = Some(24.0);
    if let Some(export) = word_session.export_result.as_mut() {
        export.baseline = Some(40.0);
    }
    if let Some(metadata) = word_session.original_metadata.as_mut() {
        metadata.font_size_pt = Some(10.5);
    }
    let edited_word_geometry = calculate_word_geometry(&word_request, &word_session)
        .expect("the editor-selected Word size should override the launch request");
    assert!((edited_word_geometry.font_size_pt - 24.0).abs() < 0.001);
    assert!(
        (edited_word_geometry.width / edited_word_geometry.reference_width_pt - 24.0 / 14.0)
            .abs()
            < 0.001
    );
    assert!(
        (edited_word_geometry.height / edited_word_geometry.reference_height_pt - 24.0 / 14.0)
            .abs()
            < 0.001
    );
}

#[test]
fn the_word_number_request_is_answered_through_the_inbox() {
    use base64::{engine::general_purpose::URL_SAFE_NO_PAD, Engine as _};
    let root = std::env::temp_dir().join(format!("visualtex-fast-number-{}", Uuid::new_v4()));
    let inbox = root.join("inbox");
    let sessions = root.join("sessions");
    fs::create_dir_all(&inbox).unwrap();
    let session_id = Uuid::new_v4().to_string();
    let session_directory = sessions.join(&session_id);
    fs::create_dir_all(&session_directory).unwrap();
    let encode = |value: &str| URL_SAFE_NO_PAD.encode(value.as_bytes());
    let rows = [
        "protocolVersion=1".to_string(),
        format!("sessionId={session_id}"),
        "sourceKind=nativeNumber".to_string(),
        format!("sourceDocumentId={}", encode("Macintosh HD:test.docx")),
        "bookmarkName=-".to_string(),
        "itemCount=1".to_string(),
        format!("mathFont={}", encode("Cambria Math")),
        format!("item1SequenceName={}", encode("Equation")),
        "item1HeadingLevel=0".to_string(),
        format!("item1Separator={}", encode(".")),
        "item1Initial=1".to_string(),
    ];
    fs::write(session_directory.join(FORMULA_RESTORE_SOURCE_FILE), rows.join("\n")).unwrap();
    // files that are not number requests stay where they are
    fs::write(inbox.join("resident-ready"), b"ready").unwrap();
    fs::write(inbox.join(format!("{session_id}.json")), b"{}").unwrap();
    fs::write(inbox.join("not-a-uuid.number"), b"nativeNumber").unwrap();
    fs::write(inbox.join(format!("{session_id}.number")), b"nativeNumber").unwrap();

    let directory_of = |id: &str| -> Result<PathBuf, String> { Ok(sessions.join(id)) };
    assert_eq!(fast_open::answer_fast_number_requests(&inbox, &directory_of).unwrap(), 1);
    assert!(!inbox.join(format!("{session_id}.number")).exists(), "the request is taken");
    assert!(inbox.join("resident-ready").exists() && inbox.join(format!("{session_id}.json")).exists());
    assert!(inbox.join("not-a-uuid.number").exists(), "only Session-named markers are answered");
    let answer = fs::read_to_string(inbox.join(format!("{session_id}.number-done"))).unwrap();
    assert!(answer.contains("operation=nativeNumber"), "{answer}");
    let docx = answer.lines().find_map(|row| row.strip_prefix("docxPath=")).expect("the answer names the package");
    assert!(Path::new(docx).is_file(), "{docx}");
    assert_eq!(fast_open::answer_fast_number_requests(&inbox, &directory_of).unwrap(), 0, "nothing is answered twice");

    // a Session with no manifest is answered with the error, not left waiting
    let missing = Uuid::new_v4().to_string();
    fs::write(inbox.join(format!("{missing}.number")), b"nativeNumber").unwrap();
    assert_eq!(fast_open::answer_fast_number_requests(&inbox, &directory_of).unwrap(), 1);
    let failure = fs::read_to_string(inbox.join(format!("{missing}.number-done"))).unwrap();
    assert!(failure.starts_with("error="), "{failure}");
    let _ = fs::remove_dir_all(&root);
}
