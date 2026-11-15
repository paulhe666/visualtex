//! Reading and checking what an add-in asks for: the request file, formula metadata, and the LaTeX the editor gets.

use super::*;

pub(super) fn validate_uuid(value: &str, label: &str) -> Result<(), String> {
    if valid_uuid(value) {
        Ok(())
    } else {
        Err(format!("{label} must be a canonical UUID"))
    }
}

pub(super) fn validate_bounded_text(value: &str, maximum: usize, label: &str) -> Result<(), String> {
    if value.chars().count() > maximum || value.chars().any(char::is_control) {
        return Err(format!(
            "{label} contains unsupported characters or is too long"
        ));
    }
    Ok(())
}

pub(super) fn validate_finite_geometry(value: f64, label: &str) -> Result<(), String> {
    if !value.is_finite() || value.abs() > 10_000_000.0 {
        Err(format!("PowerPoint {label} is invalid"))
    } else {
        Ok(())
    }
}

pub(super) fn validate_latex_redraw_source_size(byte_len: u64) -> Result<(), String> {
    if byte_len == 0 || byte_len > MAX_LATEX_REDRAW_SOURCE_BYTES {
        Err("LaTeX redraw source must contain 1 byte to 5 MB".to_string())
    } else {
        Ok(())
    }
}

// Untagged Word equations use an ephemeral edit anchor and a session-local
// OMML snapshot, never a persistent formula metadata envelope.
pub(super) fn is_untagged_word_edit(request: &MacOfflineSessionRequest) -> bool {
    request.host == "word" && request.mode == "edit"
        && request.encoded_metadata.is_none()
        && !request.fork_copied_formula
        && request.source_object_id.as_deref() == Some(
            format!("VT_E_{}", request.session_id.chars().take(24).filter(|c| *c != '-').collect::<String>()).as_str())
        && ((request.operation.as_deref().unwrap_or("formula") == "formula" && request.native_equation)
            || request.operation.as_deref() == Some("nativeToImage"))
}

// A Word session converts between an image and a Word equation when its
// source and its result differ in kind. The source is a Word equation when it
// carries a VisualTeX OMML anchor (VT_F_) or is an untagged Word equation,
// opened through its session-only edit anchor.
pub(super) fn word_format_conversion_requested(request: &MacOfflineSessionRequest) -> bool {
    let source_is_word_equation = request
        .source_object_id
        .as_deref()
        .is_some_and(|value| value.starts_with("VT_F_"))
        || is_untagged_word_edit(request);
    request.host == "word" && source_is_word_equation != request.native_equation
}

// Ending a session without applying it must remove what the add-in put in the
// document for it: a new formula's placeholder, a Word edit's session-only
// anchor (VT_E_), or PowerPoint's native edit state.
pub(super) fn cancel_needs_host_callback(request: &MacOfflineSessionRequest) -> bool {
    request.mode == "create"
        || (request.host == "powerpoint" && request.native_equation)
        || (request.host == "word"
            && request.source_object_id.as_deref().is_some_and(|value| value.starts_with("VT_E_")))
}

pub(super) fn word_native_edit_latex(xml: &str, display: bool) -> Result<crate::office::powerpoint_omml::to_latex::NativeLatex, String> {
    let math_ml = word_omml_to_mathml(xml, display)?;
    Ok(crate::office::powerpoint_omml::to_latex::NativeLatex {
        latex: crate::office::omml_batch::mathml_to_latex_checked(&math_ml)?,
        letter_font: "cambria".to_string(),
    })
}

pub(super) fn word_native_edit_source(request: &MacOfflineSessionRequest) -> Result<crate::office::powerpoint_omml::to_latex::NativeLatex, String> {
    // Small live OMath snapshots travel in the same atomic inbox request as
    // the edit target. Older templates and oversized snapshots retain the
    // session-file transport, with the same parser and no metadata cache.
    let display = request.display_mode == "block";
    if let Some(xml) = request.native_edit_xml.as_deref() {
        return word_native_edit_latex(xml, display);
    }
    let path = session_directory(OfficeHost::Word, &request.session_id)?.join("native-edit-original.omml");
    let size = fs::metadata(&path).map_err(|e| format!("Native Word equation snapshot is missing: {e}"))?.len();
    if size == 0 || size > MAX_OMML_BYTES as u64 { return Err("Native Word equation snapshot has invalid size".into()); }
    let xml = fs::read_to_string(path).map_err(|e| e.to_string())?;
    word_native_edit_latex(&xml, display)
}

pub(super) fn validate_request(request: &MacOfflineSessionRequest, session_id: &str) -> Result<(), String> {
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
    // WPS Presentation sends the edited equation as WPS saved it: its
    // object model has no equations to read (wps::presentation).
    let wps_native_edit = request.client.as_deref() == Some("wps")
        && request.host == "powerpoint"
        && request.native_equation
        && request.mode == "edit";
    if let Some(xml) = request.native_edit_xml.as_deref() {
        if !(is_untagged_word_edit(request) || wps_native_edit) || xml.is_empty() || xml.len() > MAX_OMML_BYTES {
            return Err("An inline native Word snapshot requires an explicit native edit target and valid size".into());
        }
    } else if wps_native_edit {
        return Err("A WPS Presentation equation edit needs the equation it replaces".into());
    }
    let operation = request.operation.as_deref().unwrap_or("formula");
    if let Some(client) = request.client.as_deref() {
        if client != "wps" {
            return Err("Offline Office request client is unknown".to_string());
        }
        // A WPS document operation carries the same documentImport object as
        // Word's; the checks below apply to it unchanged.
        let document_operation = matches!(operation, "documentImport" | "latexRedraw" | "formulaRestore");
        let writer_request = request.host == "word"
            && (document_operation || matches!(operation, "formula" | "nativeToImage" | "imageToNative"))
            && !request.fork_copied_formula
            && request.power_point.is_none()
            && request.document_import.is_some() == document_operation;
        // WPS Presentation writes the PowerPoint add-in's picture formulas and
        // native equations, and redraws and restores formulas; the PowerPoint
        // checks below apply to its requests.
        let presentation_request = request.host == "powerpoint"
            && ((operation == "formula" && request.document_import.is_none())
                || (matches!(operation, "latexRedraw" | "formulaRestore") && request.document_import.is_some()));
        if !(writer_request || presentation_request) {
            return Err(
                "A WPS request must be a Word formula or document request, or a PowerPoint picture formula"
                    .to_string(),
            );
        }
    }
    if request.host == "powerpoint" && matches!(operation, "latexRedraw" | "formulaRestore") {
        return presentation_redraw::validate(request);
    }
    if matches!(operation, "documentImport" | "latexRedraw" | "formulaRestore") {
        if request.host != "word" || request.mode != "create" {
            return Err("Document import is supported only as a new Word operation".to_string());
        }
        let document_import = request
            .document_import
            .as_ref()
            .ok_or_else(|| "Document import request is missing its insertion anchor".to_string())?;
        let source_document_id = request.source_document_id.as_deref().ok_or_else(|| {
            "Document import request is missing the Word document identity".to_string()
        })?;
        validate_bounded_text(source_document_id, MAX_IDENTITY_CHARS, "sourceDocumentId")?;
        validate_bounded_text(
            &document_import.bookmark_name,
            40,
            "documentImport bookmarkName",
        )?;
        if !document_import.bookmark_name.starts_with("VT_D_") {
            return Err("Document import bookmark name is invalid".to_string());
        }
        if !document_import.default_font_size_pt.is_finite()
            || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT)
                .contains(&document_import.default_font_size_pt)
        {
            return Err(
                "Document import default font size is outside the supported range".to_string(),
            );
        }
        for size in [document_import.inline_font_size_pt, document_import.display_font_size_pt]
            .into_iter()
            .flatten()
        {
            if !size.is_finite() || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&size) {
                return Err(
                    "Document import conversion font size is outside the supported range"
                        .to_string(),
                );
            }
        }
        if operation == "latexRedraw" {
            if !matches!(
                document_import.redraw_scope.as_deref(),
                Some("selection" | "document")
            ) {
                return Err("LaTeX redraw scope must be selection or document".to_string());
            }
            if !matches!(
                document_import.output_kind.as_deref(),
                Some("omml" | "image")
            ) {
                return Err("LaTeX redraw output kind must be omml or image".to_string());
            }
            if document_import.source_kind.is_some() {
                return Err("LaTeX redraw request contains a formula restore source kind".to_string());
            }
            let source_path =
                session_directory(OfficeHost::Word, session_id)?.join(LATEX_REDRAW_SOURCE_FILE);
            let source_metadata = fs::metadata(&source_path)
                .map_err(|error| format!("Unable to read {}: {error}", source_path.display()))?;
            validate_latex_redraw_source_size(source_metadata.len())?;
        } else if operation == "formulaRestore" {
            if !matches!(
                document_import.redraw_scope.as_deref(),
                Some("selection" | "document")
            ) {
                return Err("Formula restore scope must be selection or document".to_string());
            }
            if !matches!(
                document_import.output_kind.as_deref(),
                Some("latex" | "image" | "omml")
            ) {
                return Err("Formula restore output kind must be latex, image or omml".to_string());
            }
            if !matches!(document_import.source_kind.as_deref(), Some("omml" | "image")) {
                return Err("Formula restore source kind must be omml or image".to_string());
            }
            if document_import.output_kind.as_deref() == Some("image")
                && document_import.source_kind.as_deref() != Some("omml")
            {
                return Err("Only native OMML can be restored to a VisualTeX image".to_string());
            }
            if document_import.output_kind.as_deref() == Some("omml")
                && document_import.source_kind.as_deref() != Some("image")
            {
                return Err("Only VisualTeX image formulas can be converted to Word OMML".to_string());
            }
            let source_metadata = fs::metadata(formula_restore_source_path(session_id)?)
                .map_err(|error| format!("Unable to inspect formula restore source: {error}"))?;
            if source_metadata.file_type().is_symlink()
                || !source_metadata.is_file()
                || source_metadata.len() == 0
                || source_metadata.len() as usize > MAX_DOCUMENT_IMPORT_MANIFEST_BYTES
            {
                return Err("Formula restore source has an invalid size".to_string());
            }
        } else if document_import.redraw_scope.is_some()
            || document_import.output_kind.is_some()
            || document_import.source_kind.is_some()
        {
            return Err("Document import request contains formula transform fields".to_string());
        }
        if request.formula_id.is_some()
            || request.encoded_metadata.is_some()
            || request.pending_marker.is_some()
            || request.power_point.is_some()
            || request.document_import.is_none()
            || request.numbered
            || request.native_equation
            || request.fork_copied_formula
        {
            return Err("Document import request contains formula-only fields".to_string());
        }
        return Ok(());
    }
    if !matches!(operation, "formula" | "nativeToImage" | "imageToNative")
        || request.document_import.is_some()
    {
        return Err("Unsupported offline Office operation".to_string());
    }
    if request.host == "word"
        && request.mode == "edit"
        && operation == "formula"
        && (request.formula_id.is_none()
            || request.source_document_id.is_none()
            || request.source_object_id.is_none()
            || (request.encoded_metadata.is_none() && !is_untagged_word_edit(request))
            || request.pending_marker.is_some()
            || request.power_point.is_some())
    {
        return Err("Word formula edit is missing its explicit source target".to_string());
    }
    if request.fork_copied_formula
        && (request.mode != "edit"
            || !matches!(operation, "formula" | "nativeToImage" | "imageToNative")
            || request.formula_id.is_none()
            || request.source_object_id.is_none()
            || (request.encoded_metadata.is_none() && !is_untagged_word_edit(request)))
    {
        return Err(
            "Copied-formula isolation requires an existing formula edit with an explicit target"
                .to_string(),
        );
    }
    if matches!(operation, "nativeToImage" | "imageToNative") {
        let output_matches_operation = if operation == "nativeToImage" {
            !request.native_equation
        } else {
            request.native_equation
        };
        if request.host != "word" || request.mode != "edit" || !output_matches_operation {
            return Err(format!(
                "{operation} conversion requires a matching Word edit output request"
            ));
        }
        if request.formula_id.is_none()
            || request.source_document_id.is_none()
            || request.source_object_id.is_none()
            || (request.encoded_metadata.is_none() && !is_untagged_word_edit(request))
            || request.pending_marker.is_some()
            || request.power_point.is_some()
        {
            return Err(format!(
                "{operation} conversion is missing its existing Word formula identity"
            ));
        }
    }
    if !matches!(request.display_mode.as_str(), "inline" | "block") {
        return Err("Offline Office displayMode must be inline or block".to_string());
    }
    if request.numbered && (request.host != "word" || request.display_mode != "block") {
        return Err("Only Word display formulas can be numbered".to_string());
    }
    if request.native_equation && request.host == "powerpoint" {
        if !matches!(request.mode.as_str(), "create" | "edit") || operation != "formula"
            || request.pending_marker.is_some() || request.encoded_metadata.is_some()
            || request.source_document_id.is_none()
            || !request.source_object_id.as_deref().is_some_and(|id|
                id.starts_with("visualtex-ppt-native-text:") || id.starts_with("visualtex-ppt-native-slide:"))
        {
            return Err("Native PowerPoint insertion requires an explicit, unmodified text or slide target".to_string());
        }
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
    for (value, label) in [
        (request.font_size_pt, "fontSizePt"),
        (request.reference_width_pt, "referenceWidthPt"),
        (request.reference_height_pt, "referenceHeightPt"),
    ] {
        if let Some(value) = value {
            if !value.is_finite() || value <= 0.0 {
                return Err(format!(
                    "Offline Office {label} must be a positive finite number"
                ));
            }
        }
    }
    if let Some(font_size) = request.font_size_pt {
        if !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&font_size) {
            return Err(
                "Offline Office Word fontSizePt is outside the supported range".to_string(),
            );
        }
    }
    if request.host != "word"
        && (request.font_size_pt.is_some()
            || request.reference_width_pt.is_some()
            || request.reference_height_pt.is_some())
    {
        return Err("PowerPoint requests must not contain Word font-size metadata".to_string());
    }
    if let Some(font) = request.formula_letter_font.as_deref() {
        if request.host != "word" || request.mode != "edit" {
            return Err(
                "Live Word formula font metadata is valid only for Word edit requests"
                    .to_string(),
            );
        }
        if !matches!(
            font,
            "katex" | "times" | "cambria" | "stix" | "palatino" | "helvetica"
        ) {
            return Err("Live Word formula font metadata is invalid".to_string());
        }
    }
    if let Some(font_name) = request.word_math_font_name.as_deref() {
        if request.host != "word" || request.mode != "edit" {
            return Err(
                "Live Word math font name is valid only for Word edit requests".to_string(),
            );
        }
        validate_bounded_text(font_name, 128, "wordMathFontName")?;
    }
    if let Some(font_name) = request.word_document_math_font_name.as_deref() {
        if request.host != "word" {
            return Err("The document math font is valid only for Word requests".to_string());
        }
        validate_bounded_text(font_name, 128, "wordDocumentMathFontName")?;
    }

    match (request.host.as_str(), request.power_point.as_ref()) {
        ("word", None) => {}
        ("word", Some(_)) => {
            return Err("Word request must not contain PowerPoint geometry".to_string())
        }
        ("powerpoint", None) => return Err("PowerPoint request requires geometry".to_string()),
        ("powerpoint", Some(powerpoint)) => {
            validate_bounded_text(
                &powerpoint.presentation_identity,
                MAX_IDENTITY_CHARS,
                "PowerPoint presentation identity",
            )?;
            let standalone_native = request.native_equation
                && request.source_object_id.as_deref().is_some_and(|id| id.starts_with("visualtex-ppt-native-slide:"));
            if request.native_equation && request.mode == "edit" && standalone_native {
                return Err("Native PowerPoint editing requires an existing mathematical text range".into());
            }
            match powerpoint.native_text_target.as_ref() {
                Some(text) if request.native_equation && !standalone_native => {
                    if request.mode == "edit" && (text.range_length == 0 || text.leading_paragraph || text.trailing_paragraph) {
                        return Err("Native PowerPoint edits must replace exactly the selected equation without paragraph insertion".into());
                    }
                    if text.range_start == 0 || text.range_start.checked_add(text.range_length).is_none() {
                        return Err("PowerPoint native text range is invalid".into());
                    }
                }
                None if !request.native_equation || standalone_native => {}
                _ => return Err("PowerPoint native text target does not match the requested insertion mode".into()),
            }
            if !standalone_native {
                validate_bounded_text(&powerpoint.shape_name, MAX_SHAPE_NAME_CHARS, "PowerPoint shape name")?;
            } else if powerpoint.shape_id != 0 || powerpoint.shape_index != 0 {
                return Err("A standalone native equation must not name an existing shape".to_string());
            }
            if powerpoint.slide_index == 0
                || powerpoint.slide_id == 0
                || (!standalone_native && (powerpoint.shape_index == 0
                    || powerpoint.shape_id == 0 || powerpoint.z_order == 0))
            {
                return Err(
                    "PowerPoint slide, shape-index, shape and z-order references must be positive".to_string(),
                );
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
            for (value, label) in [
                (powerpoint.font_size_pt, "fontSizePt"),
                (powerpoint.reference_width_pt, "referenceWidthPt"),
                (powerpoint.reference_height_pt, "referenceHeightPt"),
            ] {
                if let Some(value) = value {
                    if !value.is_finite() || value <= 0.0 {
                        return Err(format!(
                            "PowerPoint formula {label} must be a positive finite number"
                        ));
                    }
                }
            }
            if let Some(font_size) = powerpoint.font_size_pt {
                if !(MIN_POWERPOINT_FONT_SIZE_PT..=MAX_POWERPOINT_FONT_SIZE_PT).contains(&font_size)
                {
                    return Err(
                        "PowerPoint formula fontSizePt is outside the supported range".to_string(),
                    );
                }
            }
        }
        _ => return Err("Offline Office request host is inconsistent".to_string()),
    }
    Ok(())
}

pub(super) fn read_request(session_id: &str) -> Result<MacOfflineSessionRequest, String> {
    validate_uuid(session_id, "Session id")?;
    let mut candidates = Vec::new();
    for host in [OfficeHost::Word, OfficeHost::Powerpoint] {
        let path = request_path(host, session_id)?;
        match fs::symlink_metadata(&path) {
            Ok(metadata) => candidates.push((host, path, metadata)),
            Err(error) if error.kind() == std::io::ErrorKind::NotFound => {}
            Err(error) => {
                return Err(format!(
                    "Unable to inspect offline Office request metadata at {}: {error}",
                    path.display()
                ))
            }
        }
    }
    let (expected_host, path, metadata) = match candidates.len() {
        1 => candidates.remove(0),
        0 => {
            return Err(
                "Offline Office request was not found in either host runtime directory".to_string(),
            )
        }
        _ => {
            return Err(
                "The same Offline Office Session exists in both host runtime directories"
                    .to_string(),
            )
        }
    };
    if metadata.file_type().is_symlink()
        || !metadata.is_file()
        || metadata.len() == 0
        || metadata.len() > MAX_REQUEST_BYTES
    {
        return Err("Offline Office request has an invalid size".to_string());
    }
    let bytes = fs::read(&path)
        .map_err(|error| format!("Unable to read offline Office request: {error}"))?;
    let request: MacOfflineSessionRequest = serde_json::from_slice(&bytes)
        .map_err(|error| format!("Offline Office request contains invalid JSON: {error}"))?;
    validate_request(&request, session_id)?;
    if host_from_request_name(&request.host)? != expected_host {
        return Err(
            "Offline Office request host does not match its Application Scripts runtime directory"
                .to_string(),
        );
    }
    Ok(request)
}

pub(in crate::office) fn decode_metadata(encoded: &str) -> Result<VisualTeXFormulaMetadata, String> {
    let json = metadata_json(encoded)?;
    let metadata: VisualTeXFormulaMetadata = serde_json::from_slice(&json)
        .map_err(|error| format!("VisualTeX formula metadata JSON is invalid: {error}"))?;
    validate_metadata(&metadata)?;
    Ok(metadata)
}

/// The JSON text inside encoded formula metadata.
pub(super) fn metadata_json(encoded: &str) -> Result<Vec<u8>, String> {
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
    Ok(json)
}

pub(super) fn validate_metadata(metadata: &VisualTeXFormulaMetadata) -> Result<(), String> {
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
    for (value, label) in [
        (metadata.render_width_px, "renderWidthPx"),
        (metadata.render_height_px, "renderHeightPx"),
        (metadata.reference_width_pt, "referenceWidthPt"),
        (metadata.reference_height_pt, "referenceHeightPt"),
    ] {
        if let Some(value) = value {
            if !value.is_finite() || value <= 0.0 {
                return Err(format!(
                    "VisualTeX metadata {label} must be positive and finite"
                ));
            }
        }
    }
    if let Some(value) = metadata.font_size_pt {
        if !value.is_finite() || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&value) {
            return Err("VisualTeX metadata fontSizePt is outside the Office range".to_string());
        }
    }
    if let Some(value) = metadata.reference_baseline_pt {
        if !value.is_finite() || !(-256.0..=0.0).contains(&value) {
            return Err("VisualTeX metadata referenceBaselinePt is invalid".to_string());
        }
    }
    if let Some(value) = metadata.image_ink_center_y_ratio {
        if !value.is_finite() || !(0.0..=1.0).contains(&value) {
            return Err("VisualTeX metadata imageInkCenterYRatio is invalid".to_string());
        }
    }
    Ok(())
}

pub(super) fn replace_mathlive_latex_command(source: &str, command: &str, replacement: &str) -> String {
    let pattern = format!("\\{command}");
    let mut output = String::with_capacity(source.len());
    let mut cursor = 0;
    while let Some(relative) = source[cursor..].find(&pattern) {
        let start = cursor + relative;
        let end = start + pattern.len();
        output.push_str(&source[cursor..start]);
        let followed_by_command_letter = source[end..]
            .chars()
            .next()
            .is_some_and(|character| character.is_ascii_alphabetic());
        if followed_by_command_letter {
            output.push_str(&pattern);
        } else {
            output.push_str(replacement);
        }
        cursor = end;
    }
    output.push_str(&source[cursor..]);
    output
}

pub(super) fn normalize_mathlive_upright_commands(source: &str) -> String {
    let mut normalized = source.to_string();
    for (command, replacement) in [
        ("capitalDifferentialD", "\\mathrm{D}"),
        ("differentialD", "\\mathrm{d}"),
        ("exponentialE", "\\mathrm{e}"),
        ("imaginaryI", "\\mathrm{i}"),
        ("imaginaryJ", "\\mathrm{j}"),
    ] {
        normalized = replace_mathlive_latex_command(&normalized, command, replacement);
    }
    for prefix in ["\\mathrm{d", "\\textrm{d"] {
        let mut output = String::with_capacity(normalized.len());
        let mut cursor = 0;
        while let Some(relative) = normalized[cursor..].find(prefix) {
            let start = cursor + relative;
            let variable_start = start + prefix.len();
            output.push_str(&normalized[cursor..start]);
            let Some(variable) = normalized[variable_start..].chars().next() else {
                output.push_str(prefix);
                cursor = variable_start;
                continue;
            };
            let variable_end = variable_start + variable.len_utf8();
            if variable.is_ascii_alphabetic() && normalized[variable_end..].starts_with('}') {
                output.push_str("\\mathrm{d}");
                output.push(variable);
                cursor = variable_end + 1;
            } else {
                output.push_str(prefix);
                cursor = variable_start;
            }
        }
        output.push_str(&normalized[cursor..]);
        normalized = output;
    }
    normalized
}

pub(super) fn latex_character_is_escaped(source: &str, index: usize) -> bool {
    let bytes = source.as_bytes();
    let mut cursor = index;
    let mut slash_count = 0;
    while cursor > 0 && bytes[cursor - 1] == b'\\' {
        slash_count += 1;
        cursor -= 1;
    }
    slash_count % 2 == 1
}

pub(super) fn read_latex_environment_token(source: &str, index: usize) -> Option<(bool, String, usize)> {
    let rest = source.get(index..)?;
    let (is_begin, name_start) = if rest.starts_with("\\begin{") {
        (true, index + "\\begin{".len())
    } else if rest.starts_with("\\end{") {
        (false, index + "\\end{".len())
    } else {
        return None;
    };
    let name_end = source[name_start..].find('}')? + name_start;
    let name = &source[name_start..name_end];
    if name.is_empty()
        || !name.chars().enumerate().all(|(position, character)| {
            character.is_ascii_alphabetic()
                || (character == '*' && position == name.chars().count() - 1)
        })
    {
        return None;
    }
    Some((is_begin, name.to_string(), name_end + 1))
}

pub(super) fn update_latex_environment_stack(environments: &mut Vec<String>, is_begin: bool, name: String) {
    if is_begin {
        environments.push(name);
    } else if let Some(index) = environments.iter().rposition(|value| value == &name) {
        environments.remove(index);
    }
}

pub(super) fn has_top_level_alignment_marker(source: &str) -> bool {
    let mut brace_depth = 0_u32;
    let mut environments = Vec::new();
    let mut index = 0;
    while index < source.len() {
        if let Some((is_begin, name, end)) = read_latex_environment_token(source, index) {
            update_latex_environment_stack(&mut environments, is_begin, name);
            index = end;
            continue;
        }
        let Some(character) = source[index..].chars().next() else {
            break;
        };
        if character == '{' && !latex_character_is_escaped(source, index) {
            brace_depth += 1;
        } else if character == '}' && !latex_character_is_escaped(source, index) {
            brace_depth = brace_depth.saturating_sub(1);
        } else if character == '&'
            && !latex_character_is_escaped(source, index)
            && brace_depth == 0
            && environments.is_empty()
        {
            return true;
        }
        index += character.len_utf8();
    }
    false
}

pub(super) fn top_level_relation_index(source: &str) -> Option<usize> {
    const RELATION_COMMANDS: &[&str] = &[
        "\\Longleftrightarrow",
        "\\Longrightarrow",
        "\\Leftrightarrow",
        "\\Rightarrow",
        "\\leftrightarrow",
        "\\rightarrow",
        "\\leftarrow",
        "\\subseteq",
        "\\supseteq",
        "\\notin",
        "\\approx",
        "\\equiv",
        "\\simeq",
        "\\propto",
        "\\mapsto",
        "\\subset",
        "\\supset",
        "\\cong",
        "\\neq",
        "\\leq",
        "\\geq",
        "\\sim",
        "\\to",
        "\\ne",
        "\\le",
        "\\ge",
        "\\in",
    ];
    let mut brace_depth = 0_u32;
    let mut environments = Vec::new();
    let mut index = 0;
    while index < source.len() {
        if let Some((is_begin, name, end)) = read_latex_environment_token(source, index) {
            update_latex_environment_stack(&mut environments, is_begin, name);
            index = end;
            continue;
        }
        let character = source[index..].chars().next()?;
        if character == '{' && !latex_character_is_escaped(source, index) {
            brace_depth += 1;
            index += 1;
            continue;
        }
        if character == '}' && !latex_character_is_escaped(source, index) {
            brace_depth = brace_depth.saturating_sub(1);
            index += 1;
            continue;
        }
        if brace_depth == 0 && environments.is_empty() {
            if matches!(character, '=' | '<' | '>') {
                return Some(index);
            }
            if character == '\\' {
                for command in RELATION_COMMANDS {
                    if !source[index..].starts_with(command) {
                        continue;
                    }
                    let next = source[index + command.len()..].chars().next();
                    if next.is_some_and(|value| value.is_ascii_alphabetic()) {
                        continue;
                    }
                    return Some(index);
                }
            }
        }
        index += character.len_utf8();
    }
    None
}

pub(super) fn add_latex_alignment_marker(source: &str) -> String {
    if source.is_empty() || has_top_level_alignment_marker(source) {
        return source.to_string();
    }
    let Some(index) = top_level_relation_index(source) else {
        return source.to_string();
    };
    format!("{}&{}", &source[..index], &source[index..])
}

pub(super) fn wrap_latex_environment(name: &str, body: &str) -> String {
    format!("\\begin{{{name}}}\n{body}\n\\end{{{name}}}")
}

pub(super) fn format_document_formula_rows(lines: &[String], align_relations: bool) -> String {
    lines
        .iter()
        .enumerate()
        .map(|(index, line)| {
            let content = if align_relations {
                add_latex_alignment_marker(line)
            } else {
                line.clone()
            };
            if index + 1 < lines.len() {
                format!("{content} \\\\")
            } else {
                content
            }
        })
        .collect::<Vec<_>>()
        .join("\n")
}

pub(crate) fn canonical_document_formula_latex(metadata: &VisualTeXFormulaMetadata) -> Result<String, String> {
    // `metadata.lines` stores logical editor rows. A single logical formula may
    // itself contain source-formatting newlines, especially inside equation or
    // equation* environments. Keep those internal newlines inside the same row;
    // splitting them here would rebuild one equation as several equations and
    // make the Rust validator disagree with the TypeScript serializer.
    let mut lines = metadata
        .lines
        .iter()
        .map(|line| {
            normalize_mathlive_upright_commands(
                &line.latex.replace("\r\n", "\n").replace('\r', "\n"),
            )
            .trim()
            .to_string()
        })
        .filter(|line| !line.is_empty())
        .collect::<Vec<_>>();
    if lines.is_empty() {
        lines.push(String::new());
    }
    let joined = lines.join("\n");
    let formatted = match metadata.code_format.as_str() {
        // VisualTeX for Windows writes "latex" for its formulas' rows as they are
        "raw" | "latex" => joined,
        // the main app's default copy format (formatLatexLines: one $$row$$ per line)
        "mixed-inline-display" => lines
            .iter()
            .map(|line| format!("$${line}$$"))
            .collect::<Vec<_>>()
            .join("\n"),
        "inline-dollar" => lines
            .iter()
            .map(|line| format!("${line}$"))
            .collect::<Vec<_>>()
            .join("\n"),
        "inline-paren" => lines
            .iter()
            .map(|line| format!("\\({line}\\)"))
            .collect::<Vec<_>>()
            .join("\n"),
        "display-dollar" => lines
            .iter()
            .map(|line| format!("$$\n{line}\n$$"))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "display-bracket" => lines
            .iter()
            .map(|line| format!("\\[\n{line}\n\\]"))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "equation" => lines
            .iter()
            .map(|line| wrap_latex_environment("equation", line))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "equation-star" => lines
            .iter()
            .map(|line| wrap_latex_environment("equation*", line))
            .collect::<Vec<_>>()
            .join("\n\n"),
        "align" => wrap_latex_environment("align", &format_document_formula_rows(&lines, true)),
        "align-star" => {
            wrap_latex_environment("align*", &format_document_formula_rows(&lines, true))
        }
        "aligned" => format!(
            "\\[\n{}\n\\]",
            wrap_latex_environment("aligned", &format_document_formula_rows(&lines, true),)
        ),
        "gather" => wrap_latex_environment("gather", &format_document_formula_rows(&lines, false)),
        "gather-star" => {
            wrap_latex_environment("gather*", &format_document_formula_rows(&lines, false))
        }
        "multline" => {
            wrap_latex_environment("multline", &format_document_formula_rows(&lines, false))
        }
        "multline-star" => {
            wrap_latex_environment("multline*", &format_document_formula_rows(&lines, false))
        }
        "equation-split" => wrap_latex_environment(
            "equation",
            &wrap_latex_environment("split", &format_document_formula_rows(&lines, true)),
        ),
        "equation-star-split" => wrap_latex_environment(
            "equation*",
            &wrap_latex_environment("split", &format_document_formula_rows(&lines, true)),
        ),
        _ => {
            return Err(format!(
                "Document formula metadata codeFormat is unsupported: {}",
                metadata.code_format
            ))
        }
    };
    Ok(formatted)
}

pub(super) fn normalized_serialized_latex(source: &str) -> String {
    source
        .replace("\r\n", "\n")
        .replace('\r', "\n")
        .trim()
        .to_string()
}

pub(super) fn validate_document_formula_metadata_match(
    metadata: &VisualTeXFormulaMetadata,
    formula_id: &str,
    latex: &str,
    display_mode: &str,
    numbered: bool,
) -> Result<String, String> {
    validate_metadata(metadata)?;
    if metadata.formula_id != formula_id
        || metadata.display_mode != display_mode
        || metadata.numbered != numbered
    {
        return Err(
            "Document formula metadata identity does not match its formula block".to_string(),
        );
    }
    // The document-import frontend builds metadata, SVG, PNG and OMML from one
    // normalized editor document. Treat that submitted serialization as the
    // source of truth instead of rebuilding it again in Rust. Re-serializing
    // here creates false mismatches for harmless differences such as CRLF/LF,
    // environment whitespace, alignat arguments and internal equation
    // newlines, even though every rendered artifact belongs to the same
    // formula. Rust still validates the full metadata schema and identity above
    // and requires the formula block and metadata to carry the same non-empty
    // normalized source.
    let formula_latex = normalized_serialized_latex(latex);
    let metadata_latex = normalized_serialized_latex(&metadata.latex);
    if formula_latex.is_empty() || metadata_latex != formula_latex {
        return Err("Document formula metadata LaTeX does not match its formula block".to_string());
    }
    Ok(formula_latex)
}

pub(in crate::office) fn encode_metadata(metadata: &VisualTeXFormulaMetadata) -> Result<String, String> {
    validate_metadata(metadata)?;
    let json = serde_json::to_vec(metadata)
        .map_err(|error| format!("Unable to encode VisualTeX formula metadata: {error}"))?;
    encode_metadata_json(&json)
}

pub(super) fn encode_metadata_json(json: &[u8]) -> Result<String, String> {
    let mut encoder = DeflateEncoder::new(Vec::new(), Compression::best());
    encoder
        .write_all(json)
        .map_err(|error| format!("Unable to compress VisualTeX formula metadata: {error}"))?;
    let compressed = encoder
        .finish()
        .map_err(|error| format!("Unable to finish VisualTeX formula metadata: {error}"))?;
    Ok(format!(
        "{METADATA_PREFIX}{}",
        URL_SAFE_NO_PAD.encode(compressed)
    ))
}

/// The metadata of a VisualTeX image display formula after batch numbering
/// numbered it, so a later restore finds the numbering state that the image's
/// title and format record. Only the "numbered" member of the JSON text
/// changes: every other value keeps its exact text (a parse and serialize
/// round trip would rewrite floating-point values in their last digit).
pub(in crate::office) fn numbered_formula_metadata(encoded: &str) -> Result<String, String> {
    let metadata = decode_metadata(encoded)?;
    if metadata.display_mode != "block" {
        return Err("Only a display formula can be numbered".to_string());
    }
    if metadata.numbered {
        return Ok(encoded.to_string());
    }
    let json = String::from_utf8(metadata_json(encoded)?)
        .map_err(|_| "VisualTeX formula metadata JSON is not UTF-8".to_string())?;
    // A quote inside a JSON string is escaped, so this text is only ever the
    // member itself.
    let updated = match json.matches("\"numbered\":false").count() {
        1 => json.replacen("\"numbered\":false", "\"numbered\":true", 1),
        0 if !json.contains("\"numbered\"") && json.ends_with('}') => {
            format!("{},\"numbered\":true}}", &json[..json.len() - 1])
        }
        _ => return Err("VisualTeX formula metadata has an unexpected numbering member".to_string()),
    };
    let numbered = encode_metadata_json(updated.as_bytes())?;
    let check = decode_metadata(&numbered)?;
    if !check.numbered || check.formula_id != metadata.formula_id {
        return Err("VisualTeX formula metadata could not be marked numbered".to_string());
    }
    Ok(numbered)
}

pub(super) fn hex_encode(value: &str) -> String {
    value
        .as_bytes()
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

pub(super) fn fork_copied_metadata_identity(
    metadata: &mut VisualTeXFormulaMetadata,
    formula_id: &str,
) -> Result<(), String> {
    validate_uuid(formula_id, "Copied formula id")?;
    if metadata.formula_id == formula_id {
        return Err("Copied formula request must use a fresh formulaId".to_string());
    }
    metadata.formula_id = formula_id.to_string();
    for line in &mut metadata.lines {
        line.id = Uuid::new_v4().to_string();
    }
    Ok(())
}

pub(super) fn import_request(
    state: &OfficeCompanionState,
    request: MacOfflineSessionRequest,
) -> Result<OfficeFormulaSession, String> {
    match state.session_store.get(&request.session_id) {
        Ok(existing) => return Ok(existing),
        Err(SessionError::NotFound) => {}
        Err(error) => return Err(error.to_string()),
    }

    let mut original_metadata = request
        .encoded_metadata
        .as_deref()
        .map(decode_metadata)
        .transpose()?;
    if let (Some(metadata), Some(live_font)) = (
        original_metadata.as_mut(),
        request.formula_letter_font.as_ref(),
    ) {
        // A user may change a native Word equation's font after VisualTeX
        // created it. The live Word Range is authoritative for the next edit;
        // refresh the metadata baseline now so Apply does not restore the stale
        // font stored by an earlier VisualTeX session.
        metadata.formula_letter_font = Some(live_font.clone());
    }
    let metadata_formula_id = original_metadata
        .as_ref()
        .map(|value| value.formula_id.clone());
    let formula_id = if request.fork_copied_formula {
        let request_id = request
            .formula_id
            .clone()
            .ok_or_else(|| "Copied formula request is missing its new formulaId".to_string())?;
        let source_id = metadata_formula_id
            .as_deref()
            .ok_or_else(|| "Copied formula request is missing source metadata".to_string())?;
        if request_id == source_id {
            return Err("Copied formula request must use a fresh formulaId".to_string());
        }
        request_id
    } else {
        match (request.formula_id.clone(), metadata_formula_id) {
            (Some(request_id), Some(metadata_id)) if request_id != metadata_id => {
                return Err("Request formulaId does not match encoded metadata".to_string())
            }
            (Some(request_id), _) => request_id,
            (None, Some(metadata_id)) => metadata_id,
            (None, None) if request.mode == "create" || (request.host == "powerpoint" && request.native_equation) => Uuid::new_v4().to_string(),
            (None, None) => return Err("Edit request does not contain a formulaId".to_string()),
        }
    };
    validate_uuid(&formula_id, "Imported formula id")?;
    if request.fork_copied_formula {
        let metadata = original_metadata
            .as_mut()
            .ok_or_else(|| "Copied formula request is missing source metadata".to_string())?;
        fork_copied_metadata_identity(metadata, &formula_id)?;
    }

    let host = match request.host.as_str() {
        "word" => OfficeHost::Word,
        "powerpoint" => OfficeHost::Powerpoint,
        _ => return Err("Offline Office request host is invalid".to_string()),
    };
    let mode = match request.mode.as_str() {
        "create" => OfficeSessionMode::Create,
        "edit" => OfficeSessionMode::Edit,
        _ => return Err("Offline Office request mode is invalid".to_string()),
    };
    let native_edit_source = if host == OfficeHost::Powerpoint && request.native_equation && mode == OfficeSessionMode::Edit {
        let omml = if request.client.as_deref() == Some("wps") {
            // The equation as WPS saved it; kept as the snapshot an unchanged
            // edit is recognized by (native_powerpoint_edit_source_changed).
            let zone = request.native_edit_xml.as_deref().ok_or("The WPS equation edit has no equation")?;
            let omml = crate::office::wps::presentation::zone_equation(zone)?;
            atomic_write_runtime(
                &session_directory(OfficeHost::Powerpoint, &request.session_id)?.join("native-edit-original.omml"),
                omml.as_bytes(),
                0o600,
            )?;
            omml
        } else {
            capture_powerpoint_native_edit_source(state.app.as_ref(), &request.session_id, false)?
                .ok_or("The native equation edit session has already completed")?
                .omml
        };
        Some(crate::office::powerpoint_omml::to_latex::from_omml(&omml)?)
    } else if is_untagged_word_edit(&request) {
        Some(word_native_edit_source(&request)?)
    } else { None };
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
                latex: native_edit_source.as_ref().map(|s|s.latex.clone()).unwrap_or_default(),
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
            OfficeSessionMode::Edit => request
                .source_object_id
                .clone()
                .or_else(|| request.encoded_metadata.clone()),
        },
        OfficeHost::Powerpoint if request.native_equation => request.source_object_id.clone(),
        OfficeHost::Powerpoint => request.power_point.as_ref().map(|powerpoint| {
            format!(
                "visualtex-ppt-native-edit:{}:{}:{}",
                powerpoint.slide_id,
                powerpoint.shape_id,
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
        .unwrap_or_else(|| if native_edit_source.is_some() { "raw" } else { "latex" }.to_string());
    let font_size_pt = request
        .font_size_pt
        .or_else(|| windows_picture_font_size(request.power_point.as_ref(), original_metadata.as_ref()))
        .or_else(|| {
            request
                .power_point
                .as_ref()
                .and_then(|powerpoint| powerpoint.font_size_pt)
        })
        .or_else(|| {
            original_metadata
                .as_ref()
                .and_then(|metadata| metadata.font_size_pt)
        });

    let session_id = request.session_id.clone();
    let session_operation = request
        .operation
        .as_deref()
        .filter(|operation| *operation != "formula")
        .map(str::to_string)
        // the editor commits an agent's formula by itself, as it does a conversion
        .or_else(|| request.agent.then(|| "agent".to_string()));
    match state.session_store.create_external(
        session_id.clone(),
        CreateOfficeSessionInput {
            mode,
            host,
            client: request.client.clone(),
            operation: session_operation,
            native_equation: request.native_equation,
            formula_id: Some(formula_id),
            source_document_id,
            source_object_id,
            title: Some(title),
            lines: Some(lines),
            active_line_id: None,
            code_format: Some(code_format),
            display_mode: Some(request.display_mode),
            inline_image_math_style: None,
            numbered: Some(request.numbered),
            font_size_pt,
            formula_letter_font: request
                .formula_letter_font
                .clone()
                .or_else(|| native_edit_source.as_ref().map(|s| s.letter_font.clone()))
                .or_else(|| {
                    original_metadata
                        .as_ref()
                        .and_then(|metadata| metadata.formula_letter_font.clone())
                }),
            formula_chinese_font: original_metadata
                .as_ref()
                .and_then(|metadata| metadata.formula_chinese_font.clone()),
            export_width: None,
            export_height: None,
            original_metadata,
            auto_commit_on_close: Some(true),
        },
    ) {
        Ok(session) => Ok(session),
        Err(SessionError::Conflict(_)) => state
            .session_store
            .get(&session_id)
            .map_err(|error| error.to_string()),
        Err(error) => Err(error.to_string()),
    }
}
