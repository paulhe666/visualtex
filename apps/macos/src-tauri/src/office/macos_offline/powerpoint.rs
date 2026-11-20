//! Committing a formula to PowerPoint (and WPS Presentation): the native equation being edited, and handing the
//! result over.

use super::*;

pub(super) struct PowerPointNativeEditSource { pub(super) omml: String, pub(super) clipboard: Vec<u8>, #[cfg(test)] pub(super) math_ml: String }

pub(super) fn capture_powerpoint_native_edit_source(
    app: Option<&AppHandle>, session_id: &str, recheck: bool,
) -> Result<Option<PowerPointNativeEditSource>, String> {
    let directory = session_directory(OfficeHost::Powerpoint, session_id)?;
    let status_path = directory.join("native-edit-copy-status.txt");
    let clipboard_path = directory.join("native-edit-current.bin");
    let mathml_path = directory.join("native-edit-current.mathml");
    atomic_write_runtime(&mathml_path,b"",0o600)?;
    atomic_write_runtime(&status_path, b"pending", 0o600)?;
    let quote = |path: &Path| serde_json::to_string(path.to_string_lossy().as_ref()).map_err(|e|e.to_string());
    let script = r#"use framework "AppKit"
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
set ownedCount to initialCount
try
    tell application "Microsoft PowerPoint" to run VB macro macro name "VisualTeX_CopyPowerPointNativeEditSource" list of parameters {}
    set ownedCount to (pb's changeCount()) as integer
    if ownedCount is not initialCount then
        set nativeData to pb's dataForType:"com.microsoft.Art--Text-ClipFormat"
        set mathData to pb's dataForType:"public.mathml"
        if nativeData is missing value then error "PowerPoint did not export the selected native equation as Office Math"
        if (nativeData's writeToFile:__NATIVE_PATH__ atomically:true) is false then error "Unable to preserve the original native equation"
        if mathData is not missing value then
            if (mathData's writeToFile:__MATHML_PATH__ atomically:true) is false then error "Unable to read the diagnostic MathML export"
        end if
    end if
on error errorText number errorNumber
    if ((pb's changeCount()) as integer) is ownedCount and ownedCount is not initialCount then
        pb's clearContents()
        pb's writeObjects:savedItems
    end if
    error errorText number errorNumber
end try
if ((pb's changeCount()) as integer) is ownedCount and ownedCount is not initialCount then
    pb's clearContents()
    pb's writeObjects:savedItems
end if"#.replace("__NATIVE_PATH__",&quote(&clipboard_path)?).replace("__MATHML_PATH__",&quote(&mathml_path)?);
    with_dispatch_pointer(OfficeHost::Powerpoint,session_id,|| match app {
        Some(app) => run_office_vba_script_on_main_thread(app,&script,"Read native PowerPoint equation"),
        None => run_office_vba_script(&script,"Read native PowerPoint equation"),
    })?;
    let status = fs::read_to_string(status_path).map_err(|e|e.to_string())?;
    if status.trim() == "completed" { return Ok(None); }
    if status.trim() != "ok" { return Err(format!("Cannot read the original PowerPoint equation: {}",status.trim())); }
    let bytes=fs::read(clipboard_path).map_err(|e|e.to_string())?;
    let omml=crate::office::powerpoint_omml::native_equation_from_clipboard(&bytes)?;
    let math_ml=fs::read_to_string(mathml_path).map_err(|e|e.to_string())?;
    if math_ml.len()>MAX_OMML_BYTES { return Err("The native MathML source is too large".into()); }
    let source_path=directory.join("native-edit-original.omml");
    if recheck {
        let original=fs::read_to_string(source_path).map_err(|e|format!("The original equation snapshot is unavailable: {e}"))?;
        if omml!=original { return Err("The native equation changed while the editor was open. Reopen the current equation; no content was replaced.".into()); }
    } else { atomic_write_runtime(&source_path,omml.as_bytes(),0o600)?; }
    Ok(Some(PowerPointNativeEditSource{omml,clipboard:bytes,#[cfg(test)] math_ml}))
}

pub(super) fn capture_powerpoint_native_paragraph(
    app: Option<&AppHandle>, session_id: &str,
) -> Result<Option<(Vec<u8>, usize)>, String> {
    let directory = session_directory(OfficeHost::Powerpoint, session_id)?;
    let context_path = directory.join("native-paragraph-context.bin");
    let status_path = directory.join("native-context-status.txt");
    atomic_write_runtime(&status_path, b"pending", 0o600)?;
    let quoted_path = serde_json::to_string(context_path.to_string_lossy().as_ref()).map_err(|e|e.to_string())?;
    let script = r#"use framework "AppKit"
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
set ownedCount to initialCount
try
    tell application "Microsoft PowerPoint" to run VB macro macro name "VisualTeX_CopyNativeParagraphContext" list of parameters {}
    set ownedCount to (pb's changeCount()) as integer
    if ownedCount is not initialCount then
        set payload to pb's dataForType:"com.microsoft.Art--GVML-ClipFormat"
        if payload is missing value then error "PowerPoint did not provide its native shape text format"
        set didWrite to payload's writeToFile:__CONTEXT_PATH__ atomically:true
        if didWrite is false then error "The native paragraph format could not be preserved"
    end if
on error errorText number errorNumber
    if ((pb's changeCount()) as integer) is ownedCount and ownedCount is not initialCount then
        pb's clearContents()
        pb's writeObjects:savedItems
    end if
    error errorText number errorNumber
end try
if ((pb's changeCount()) as integer) is ownedCount and ownedCount is not initialCount then
    pb's clearContents()
    pb's writeObjects:savedItems
end if"#.replace("__CONTEXT_PATH__", &quoted_path);
    with_dispatch_pointer(OfficeHost::Powerpoint,session_id,|| match app {
        Some(app) => run_office_vba_script_on_main_thread(app,&script,"PowerPoint native paragraph preservation"),
        None => run_office_vba_script(&script,"PowerPoint native paragraph preservation"),
    })?;
    let status = fs::read_to_string(&status_path).map_err(|e|e.to_string())?;
    if status.trim() == "completed" { return Ok(None); }
    if status.trim() != "ok" { return Err(format!("PowerPoint paragraph preservation failed: {}",status.trim())); }
    let bytes = fs::read(context_path).map_err(|e|format!("Cannot read the native paragraph format: {e}"))?;
    let index = fs::read_to_string(directory.join("native-context-paragraph.txt")).map_err(|e|e.to_string())?
        .trim().parse::<usize>().map_err(|_|"PowerPoint returned an invalid paragraph index")?;
    Ok(Some((bytes,index)))
}

pub(super) fn commit_powerpoint_native(
    app: Option<&AppHandle>, request: &MacOfflineSessionRequest, session: &OfficeFormulaSession,
) -> Result<(), String> {
    let wps = request.client.as_deref() == Some("wps");
    // WPS Presentation checks the equation's text itself when it writes.
    let edit_source = if request.mode == "edit" && !wps {
        let Some(source)=capture_powerpoint_native_edit_source(app,&session.id,true)? else {return Ok(());};
        Some(source)
    }else{None};
    let export = session.export_result.as_ref().ok_or("Native PowerPoint equation has no export")?;
    let table_cell = request.power_point.as_ref()
        .and_then(|ppt| ppt.native_text_target.as_ref()).is_some_and(|target| target.table_cell);
    // The equation comes from the shared writer, as Word's does: Windows'
    // MathML -> OMML (office::omml_write). The editor's own OMML remains the
    // fallback for an export that carries no MathML.
    let omml = match export.word_math_ml.as_deref() {
        Some(mathml) if table_cell => crate::office::omml_write::transform_mathml_to_omml(
            &crate::office::omml_write::prepare_mathml_for_table_cell(mathml)?)?,
        Some(mathml) => crate::office::omml_write::transform_mathml_to_omml(mathml)?,
        None => {
            let encoded = export.omml_base64.as_deref().ok_or("Native PowerPoint equation has no OMML export")?;
            let bytes = URL_SAFE_NO_PAD.decode(encoded).map_err(|_| "Invalid OMML encoding")?;
            if bytes.is_empty() || bytes.len() > MAX_OMML_BYTES { return Err("Invalid OMML payload size".into()); }
            String::from_utf8(bytes).map_err(|_| "OMML is not UTF-8")?
        }
    };
    let omml = omml.as_str();
    let ppt = request.power_point.as_ref().ok_or("Missing PowerPoint target")?;
    let in_text = request.source_object_id.as_deref().is_some_and(|id| id.starts_with("visualtex-ppt-native-text:"));
    let font_size = if in_text { ppt.font_size_pt } else { session.font_size_pt.or(ppt.font_size_pt) }.unwrap_or(DEFAULT_POWERPOINT_FONT_SIZE_PT);
    let display = session.display_mode == "block";
    if wps {
        return deliver_wps_powerpoint_native_commit(app, request, session, omml, font_size, display);
    }
    let leading = display && ppt.native_text_target.as_ref().is_some_and(|t| t.leading_paragraph);
    let trailing = display && ppt.native_text_target.as_ref().is_some_and(|t| t.trailing_paragraph);
    let generated = crate::office::powerpoint_omml::clipboard_entries(omml, font_size, display && request.mode != "edit", leading, trailing)?;
    // A table cell's equation is rebuilt from its linear format: the edited
    // equation's own context does not carry over (the source check above still ran).
    let entries = if let Some(source)=edit_source.filter(|_| !table_cell) {
        crate::office::powerpoint_omml::replace_native_edit_context(generated,&source.clipboard)?
    } else if leading {
        let Some((context,index)) = capture_powerpoint_native_paragraph(app,&session.id)? else { return Ok(()); };
        crate::office::powerpoint_omml::preserve_leading_paragraph(generated,&context,index)?
    } else {
        generated.into_iter().map(|(name,xml)|(name,xml.into_bytes())).collect()
    };
    let package = build_stored_zip(&entries)?;
    let path = session_directory(OfficeHost::Powerpoint, &session.id)?.join("native-math.bin");
    atomic_write_runtime(&path, &package, 0o600)?;
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()), ("sessionId",session.id.clone()),
        ("host","powerpoint".into()), ("action","commit".into()), ("mode","create".into()),
        ("nativeEquation","1".into()),
        ("leadingParagraph", if leading { "1" } else { "0" }.into()),
        ("trailingParagraph", if trailing { "1" } else { "0" }.into()),
        ("displayMode",session.display_mode.clone()), ("fontSizePt",font_size.to_string()),
    ])?;
    atomic_write_runtime(&dispatch_path(OfficeHost::Powerpoint,&session.id)?,dispatch.as_bytes(),0o600)?;
    // Use the independently replayable native TEXT flavor. The payload has no
    // process pointers and needs no hidden/visible intermediate Office document.
    let result_path = session_directory(OfficeHost::Powerpoint, &session.id)?.join("native-result.txt");
    atomic_write_runtime(&result_path, b"pending", 0o600)?;
    let quoted_path = serde_json::to_string(path.to_string_lossy().as_ref()).map_err(|e| e.to_string())?;
    let script = r#"use framework "AppKit"
use scripting additions
set payload to current application's NSData's dataWithContentsOfFile:__NATIVE_CLIPBOARD_PATH__
if payload is missing value then error "The native math clipboard payload is missing"
set pb to current application's NSPasteboard's generalPasteboard()
set savedItems to current application's NSMutableArray's array()
repeat with sourceItem in (pb's pasteboardItems())
    set savedItem to current application's NSPasteboardItem's alloc()'s init()
    repeat with t in sourceItem's types()
        savedItem's setData:(sourceItem's dataForType:t) forType:t
    end repeat
    savedItems's addObject:savedItem
end repeat
pb's clearContents()
set didSet to pb's setData:payload forType:"com.microsoft.Art--Text-ClipFormat"
set ownedCount to (pb's changeCount()) as integer
try
    if didSet is false then error "macOS could not prepare the native math clipboard"
    tell application "Microsoft PowerPoint" to run VB macro macro name "VisualTeX_ApplyPendingResult" list of parameters {}
on error errorText number errorNumber
    if ((pb's changeCount()) as integer) is ownedCount then
        pb's clearContents()
        pb's writeObjects:savedItems
    end if
    error errorText number errorNumber
end try
if ((pb's changeCount()) as integer) is ownedCount then
    pb's clearContents()
    pb's writeObjects:savedItems
end if"#.replace("__NATIVE_CLIPBOARD_PATH__", &quoted_path);
    with_dispatch_pointer(OfficeHost::Powerpoint,&session.id,|| {
        match app {
            Some(app) => run_office_vba_script_on_main_thread(app,&script,"PowerPoint native math insertion"),
            None => run_office_vba_script(&script,"PowerPoint native math insertion"),
        }
    })?;
    let result = fs::read_to_string(&result_path).map_err(|e| format!("Cannot read PowerPoint native insertion result: {e}"))?;
    if result.trim() == "ok" { return Ok(()); }
    if let Some(detail) = result.strip_prefix("error\n") {
        return Err(format!("PowerPoint native equation insertion failed: {}", detail.trim()));
    }
    Err("PowerPoint did not confirm the native equation insertion; the active add-in may need to be updated".into())
}

pub(super) fn commit_powerpoint(
    app: Option<&AppHandle>,
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
    metadata: &str,
    geometry: PowerPointGeometry,
) -> Result<(), String> {
    let commit_started = Instant::now();
    let powerpoint = request
        .power_point
        .as_ref()
        .ok_or_else(|| "PowerPoint request geometry is missing".to_string())?;
    let image_path = materialize_powerpoint_svg(session)?;
    let fallback_image_path = materialize_result_png(session).ok();
    let entries = [
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session.id.clone()),
        ("action", "commit".to_string()),
        ("host", "powerpoint".to_string()),
        ("mode", request.mode.clone()),
        ("formulaId", session.formula_id.clone()),
        ("displayMode", "block".to_string()),
        ("numbered", "0".to_string()),
        ("imagePath", image_path.to_string_lossy().to_string()),
        (
            "fallbackImagePath",
            fallback_image_path
                .as_ref()
                .map(|path| path.to_string_lossy().to_string())
                .unwrap_or_default(),
        ),
        ("metadata", metadata.to_string()),
        (
            "pendingMarker",
            request.pending_marker.clone().unwrap_or_default(),
        ),
        (
            "sourceMarker",
            request.encoded_metadata.clone().unwrap_or_default(),
        ),
        ("sourceShapeIndex", powerpoint.shape_index.to_string()),
        ("sourceShapeId", powerpoint.shape_id.to_string()),
        ("sourceShapeName", powerpoint.shape_name.clone()),
        ("shapeName", format!("VisualTeX_{}", session.formula_id)),
        ("targetLeft", format!("{:.6}", geometry.left)),
        ("targetTop", format!("{:.6}", geometry.top)),
        ("targetWidth", format!("{:.6}", geometry.width)),
        ("targetHeight", format!("{:.6}", geometry.height)),
        ("fontSizePt", format!("{:.6}", geometry.font_size_pt)),
        (
            "referenceWidthPt",
            format!("{:.6}", geometry.reference_width_pt),
        ),
        (
            "referenceHeightPt",
            format!("{:.6}", geometry.reference_height_pt),
        ),
        ("rotation", format!("{:.6}", powerpoint.rotation)),
        ("zOrder", powerpoint.z_order.to_string()),
        (
            "presentationIdentity",
            powerpoint.presentation_identity.clone(),
        ),
        ("slideIndex", powerpoint.slide_index.to_string()),
        ("slideId", powerpoint.slide_id.to_string()),
    ];
    if request.client.as_deref() == Some("wps") {
        return deliver_wps_powerpoint_commit(session, &entries, &image_path, fallback_image_path.as_deref());
    }
    let dispatch = dispatch_text(&entries)?;
    atomic_write_runtime(
        &dispatch_path(OfficeHost::Powerpoint, &session.id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    queue_editor_performance(
        OfficeHost::Powerpoint,
        &session.id,
        "apply-powerpoint-dispatch-ready",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    with_dispatch_pointer(OfficeHost::Powerpoint, &session.id, || {
        run_vba_callback_on_main_thread(app, OfficeHost::Powerpoint)
    })?;
    queue_editor_performance(
        OfficeHost::Powerpoint,
        &session.id,
        "apply-powerpoint-vba-complete",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    Ok(())
}

/// A native equation for the WPS Presentation add-in: a one-slide presentation
/// holding the equation's paragraph (powerpoint_omml::wps_equation_slide_entries)
/// for WPS to import and paste, and where it goes. An edit keeps its zone
/// (native_edit_paragraph). The add-in writes a display equation's paragraph
/// breaks itself, as WPS turns only the last CR of a string into a break.
/// The clipboard the add-in pastes through gets its contents back.
pub(super) fn deliver_wps_powerpoint_native_commit(
    app: Option<&AppHandle>,
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
    omml: &str,
    font_size: f64,
    display: bool,
) -> Result<(), String> {
    let ppt = request.power_point.as_ref().ok_or("Missing PowerPoint target")?;
    let language = ppt.language_id.and_then(crate::office::wps::presentation::language_tag);
    let paragraph = if request.mode == "edit" {
        let zone = request.native_edit_xml.as_deref().ok_or("The WPS equation edit has no equation")?;
        crate::office::powerpoint_omml::native_edit_paragraph(omml, font_size, zone, language)?
    } else {
        crate::office::powerpoint_omml::native_paragraphs(omml, font_size, display, false, false, language)?
    };
    let package = build_stored_zip(&crate::office::powerpoint_omml::wps_equation_slide_entries(&paragraph))?;
    let path = session_directory(OfficeHost::Powerpoint, &session.id)?.join("wps-equation.pptx");
    atomic_write_runtime(&path, &package, 0o600)?;
    let name = format!("{}-equation.pptx", session.id);
    let target = ppt.native_text_target.as_ref();
    let payload = json!({
        "sessionId": session.id,
        "mode": request.mode,
        "nativeEquation": true,
        "displayMode": session.display_mode,
        "fontSizePt": font_size,
        "sourceObjectId": request.source_object_id,
        "presentationIdentity": ppt.presentation_identity,
        "slideIndex": ppt.slide_index,
        "slideId": ppt.slide_id,
        "shapeId": ppt.shape_id,
        "standalone": target.is_none(),
        "rangeStart": target.map(|target| target.range_start),
        "rangeLength": target.map(|target| target.range_length),
        "files": { "slide": name },
    });
    #[cfg(target_os = "macos")]
    return crate::office::wps::pasteboard::kept(app, || crate::office::wps::bridge::deliver_commit(&session.id, payload, vec![(name, path)]));
    #[cfg(not(target_os = "macos"))]
    {
        let _ = app;
        crate::office::wps::bridge::deliver_commit(&session.id, payload, vec![(name, path)])
    }
}

/// A PowerPoint commit for the WPS Presentation add-in: the values the
/// PowerPoint macro reads from dispatch.txt (VTFinalizePowerPointDispatch),
/// the formula reference the macro would write as the picture's Title, and
/// the pictures to download in place of the local paths. Returns once the
/// add-in has written the slide.
pub(super) fn deliver_wps_powerpoint_commit(
    session: &OfficeFormulaSession,
    entries: &[(&str, String)],
    image_path: &Path,
    fallback_image_path: Option<&Path>,
) -> Result<(), String> {
    let mut result = serde_json::Map::new();
    for (key, value) in entries {
        if !matches!(*key, "imagePath" | "fallbackImagePath" | "protocolVersion" | "action" | "host") {
            result.insert((*key).to_string(), json!(value));
        }
    }
    result.insert(
        "formulaReference".to_string(),
        json!(crate::office::wps::payload::formula_reference(&session.formula_id, "block", false)),
    );
    let image = format!("{}-formula.svg", session.id);
    let mut files = vec![(image.clone(), image_path.to_path_buf())];
    let mut names = json!({ "image": image });
    if let Some(path) = fallback_image_path {
        let fallback = format!("{}-formula.png", session.id);
        files.push((fallback.clone(), path.to_path_buf()));
        names["fallback"] = json!(fallback);
    }
    result.insert("files".to_string(), names);
    crate::office::wps::bridge::deliver_commit(&session.id, Value::Object(result), files)
}
