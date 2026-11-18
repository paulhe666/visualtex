//! Importing a document into Word: the import request, its formulas as OMML, pictures or both, paragraph transfer,
//! LaTeX redraw font sizes, progress and cancelling.

use super::*;

pub(super) fn document_import_request_data(
    request: &MacOfflineSessionRequest,
) -> Result<MacOfflineDocumentImportPublicRequest, String> {
    let operation = request.operation.as_deref().unwrap_or("");
    if request.host == "powerpoint" && matches!(operation, "latexRedraw" | "formulaRestore") {
        return presentation_redraw::public_request(request);
    }
    if !matches!(operation, "documentImport" | "latexRedraw" | "formulaRestore")
        || request.host != "word"
    {
        return Err("Offline Office request is not a Word document operation".to_string());
    }
    let document_import = request
        .document_import
        .as_ref()
        .ok_or_else(|| "Document import request is missing insertion information".to_string())?;
    let (source, restore_targets) = if operation == "latexRedraw" {
        let source_path = session_directory(OfficeHost::Word, &request.session_id)?
            .join(LATEX_REDRAW_SOURCE_FILE);
        let bytes = fs::read(&source_path)
            .map_err(|error| format!("Unable to read {}: {error}", source_path.display()))?;
        if bytes.is_empty() || bytes.len() as u64 > MAX_LATEX_REDRAW_SOURCE_BYTES {
            return Err("LaTeX redraw source must contain 1 byte to 5 MB".to_string());
        }
        (
            Some(
                String::from_utf8(bytes)
                    .map_err(|_| "LaTeX redraw source is not valid UTF-8".to_string())?,
            ),
            None,
        )
    } else if operation == "formulaRestore" {
        let (source, targets) = parse_formula_restore_manifest(request)?;
        (Some(source), Some(targets))
    } else {
        (None, None)
    };
    Ok(MacOfflineDocumentImportPublicRequest {
        protocol_version: request.protocol_version,
        session_id: request.session_id.clone(),
        host: request.host.clone(),
        source_document_id: request.source_document_id.clone().ok_or_else(|| {
            "Document import request is missing the Word document identity".to_string()
        })?,
        bookmark_name: document_import.bookmark_name.clone(),
        default_font_size_pt: document_import.default_font_size_pt,
        inline_font_size_pt: document_import.inline_font_size_pt,
        display_font_size_pt: document_import.display_font_size_pt,
        operation: operation.to_string(),
        redraw_scope: document_import.redraw_scope.clone(),
        output_kind: document_import.output_kind.clone(),
        source_kind: document_import.source_kind.clone(),
        source,
        restore_targets,
        font_refresh: document_import.font_refresh,
        client: request.client.clone(),
        segments: None,
        agent: agent_import(request),
    })
}

pub(super) fn document_formula_file_path(
    session_id: &str,
    formula_id: &str,
    extension: &str,
) -> Result<PathBuf, String> {
    validate_uuid(formula_id, "Document formula id")?;
    if !extension.bytes().all(|byte| byte.is_ascii_alphanumeric()) {
        return Err("Document formula file extension is invalid".to_string());
    }
    Ok(session_directory(OfficeHost::Word, session_id)?
        .join(format!("document-formula-{formula_id}.{extension}")))
}

pub(super) fn validate_document_omml_payload(value: &str) -> Result<(), String> {
    let bytes = URL_SAFE_NO_PAD
        .decode(value)
        .map_err(|_| "Document formula OMML payload is not valid Base64URL".to_string())?;
    if bytes.is_empty() || bytes.len() > MAX_OMML_BYTES {
        return Err("Document formula OMML payload is empty or too large".to_string());
    }
    let omml = std::str::from_utf8(&bytes)
        .map_err(|_| "Document formula OMML payload is not UTF-8".to_string())?;
    if !omml.trim_start().starts_with("<m:oMath")
        || !omml.contains("http://schemas.openxmlformats.org/officeDocument/2006/math")
        || omml.contains("<!DOCTYPE")
        || omml.contains("<!ENTITY")
    {
        return Err("Document formula OMML payload is not a safe Office Math fragment".to_string());
    }
    Ok(())
}

pub(super) fn decode_document_native_docx(value: &str) -> Result<Vec<u8>, String> {
    let bytes = URL_SAFE_NO_PAD
        .decode(value)
        .map_err(|_| "Document formula native DOCX is not valid Base64URL".to_string())?;
    if bytes.len() < 128 || bytes.len() > MAX_OMML_BYTES * 8 || !bytes.starts_with(b"PK\x03\x04") {
        return Err("Document formula native DOCX payload is invalid or too large".to_string());
    }
    Ok(bytes)
}

pub(super) fn decode_document_image_fallback_png(value: Option<&str>) -> Result<Vec<u8>, String> {
    let value = value.ok_or_else(|| {
        "Image document formula is missing its PNG compatibility preview".to_string()
    })?;
    let bytes = decode_png(value)?;
    if bytes.len() < 24 {
        return Err("Image document formula PNG compatibility preview is truncated".to_string());
    }
    let width = u32::from_be_bytes(bytes[16..20].try_into().map_err(|_| {
        "Image document formula PNG compatibility preview has an invalid IHDR".to_string()
    })?);
    let height = u32::from_be_bytes(bytes[20..24].try_into().map_err(|_| {
        "Image document formula PNG compatibility preview has an invalid IHDR".to_string()
    })?);
    if width <= 1 || height <= 1 {
        return Err(
            "Image document formula PNG compatibility preview must contain rendered formula pixels"
                .to_string(),
        );
    }
    Ok(bytes)
}

pub(super) fn document_output_requires_image_artifacts(output_kind: &str) -> bool {
    output_kind == "image"
}

pub(super) fn calculate_document_image_geometry(
    width: f64,
    height: f64,
    baseline: f64,
    font_size_pt: f64,
    formula_letter_font: Option<&str>,
) -> Result<WordGeometry, String> {
    if !width.is_finite()
        || !height.is_finite()
        || !baseline.is_finite()
        || width <= 0.0
        || height <= 0.0
        || baseline < 0.0
        || baseline > height
    {
        return Err("Document formula SVG geometry is invalid".to_string());
    }
    calculate_word_svg_geometry_for_font(
        width,
        height,
        baseline,
        font_size_pt,
        formula_letter_font,
    )
}

pub(super) fn resolve_document_paragraph_transfer(
    paragraph_id: &Option<String>,
    paragraph_style: &Option<String>,
    paragraph_alignment: &Option<String>,
    list_kind: &Option<String>,
    list_level: Option<u32>,
    paragraph_start: bool,
    paragraph_end: bool,
) -> Result<Option<DocumentParagraphTransfer>, String> {
    let Some(id) = paragraph_id.as_deref() else {
        if paragraph_style.is_some()
            || paragraph_alignment.is_some()
            || list_kind.is_some()
            || list_level.is_some()
            || paragraph_start
            || paragraph_end
        {
            return Err("Document paragraph metadata is missing its paragraph id".to_string());
        }
        return Ok(None);
    };
    validate_uuid(id, "Document paragraph id")?;
    let style = paragraph_style.as_deref().unwrap_or("normal");
    if !matches!(
        style,
        "normal" | "heading1" | "heading2" | "heading3" | "heading4" | "quote" | "code"
    ) {
        return Err("Document paragraph style is invalid".to_string());
    }
    let alignment = paragraph_alignment.as_deref().unwrap_or("left");
    if !matches!(alignment, "left" | "center" | "right" | "justify") {
        return Err("Document paragraph alignment is invalid".to_string());
    }
    let resolved_list_kind = list_kind.as_deref().unwrap_or("none");
    if !matches!(resolved_list_kind, "none" | "bullet" | "number") {
        return Err("Document paragraph list kind is invalid".to_string());
    }
    let resolved_list_level = list_level.unwrap_or(0);
    if (resolved_list_kind == "none" && resolved_list_level != 0)
        || (resolved_list_kind != "none" && !(1..=9).contains(&resolved_list_level))
    {
        return Err("Document paragraph list level is invalid".to_string());
    }
    Ok(Some(DocumentParagraphTransfer {
        id: id.to_string(),
        style: style.to_string(),
        alignment: alignment.to_string(),
        list_kind: resolved_list_kind.to_string(),
        list_level: resolved_list_level,
        start: paragraph_start,
        end: paragraph_end,
    }))
}

pub(super) fn append_document_paragraph_entries(
    entries: &mut Vec<(String, String)>,
    prefix: &str,
    paragraph: Option<&DocumentParagraphTransfer>,
) {
    let value = |field: &str| format!("{prefix}{field}");
    if let Some(paragraph) = paragraph {
        entries.push((value("paragraphId"), paragraph.id.clone()));
        entries.push((value("paragraphStyle"), paragraph.style.clone()));
        entries.push((value("paragraphAlignment"), paragraph.alignment.clone()));
        entries.push((value("listKind"), paragraph.list_kind.clone()));
        entries.push((value("listLevel"), paragraph.list_level.to_string()));
        entries.push((
            value("paragraphStart"),
            if paragraph.start { "1" } else { "0" }.to_string(),
        ));
        entries.push((
            value("paragraphEnd"),
            if paragraph.end { "1" } else { "0" }.to_string(),
        ));
    }
}

pub(super) fn validate_latex_redraw_range(
    source_utf16: &[u16],
    previous_end: &mut usize,
    start: usize,
    end: usize,
    original: &str,
) -> Result<(), String> {
    let original_utf16 = original.encode_utf16().collect::<Vec<_>>();
    if start < *previous_end
        || end <= start
        || end > source_utf16.len()
        || source_utf16[start..end] != original_utf16
    {
        return Err(
            "LaTeX redraw formula range does not match the Word source snapshot".to_string(),
        );
    }
    *previous_end = end;
    Ok(())
}

pub(super) fn validate_formula_restore_range(
    previous_end: &mut usize,
    start: usize,
    end: usize,
    original: &str,
) -> Result<(), String> {
    let span = end.saturating_sub(start);
    if start < *previous_end
        || end <= start
        || end > 50_000_000
        || span > 5_000_000
        || original.is_empty()
        || original.len() > MAX_LATEX_REDRAW_SOURCE_BYTES as usize
    {
        return Err("Word formula restore range is invalid or overlapping".to_string());
    }
    *previous_end = end;
    Ok(())
}

pub(super) fn parse_latex_redraw_font_sizes(source: &str, expected_count: usize) -> Result<Vec<f64>, String> {
    let mut item_count = None;
    let mut values = vec![None; expected_count];
    for line in source.lines() {
        let Some((key, raw)) = line.split_once('=') else {
            continue;
        };
        if key == "itemCount" {
            item_count = raw.parse::<usize>().ok();
            continue;
        }
        let Some(index_text) = key
            .strip_prefix("item")
            .and_then(|value| value.strip_suffix("fontSizePt"))
        else {
            continue;
        };
        let index = index_text
            .parse::<usize>()
            .map_err(|_| "LaTeX redraw font-size result contains an invalid item index".to_string())?;
        if index >= expected_count || values[index].is_some() {
            return Err("LaTeX redraw font-size result contains an invalid or duplicate item".to_string());
        }
        let value = raw
            .parse::<f64>()
            .map_err(|_| "LaTeX redraw font-size result contains an invalid number".to_string())?;
        if !value.is_finite() || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&value) {
            return Err("LaTeX redraw font-size result is outside the supported range".to_string());
        }
        values[index] = Some(value);
    }
    if item_count != Some(expected_count) {
        return Err("LaTeX redraw font-size result count does not match the request".to_string());
    }
    values
        .into_iter()
        .map(|value| value.ok_or_else(|| "LaTeX redraw font-size result is incomplete".to_string()))
        .collect()
}

pub(super) fn resolve_latex_redraw_font_sizes_blocking(
    session_id: String,
    input: MacOfflineLatexRedrawFontQueryInput,
) -> Result<Vec<f64>, String> {
    let preflight_started = Instant::now();
    validate_uuid(&session_id, "Session id")?;
    queue_editor_performance(
        OfficeHost::Word,
        &session_id,
        "latex-redraw-font-backend-start",
        0.0,
        None,
        json!({ "itemCount": input.ranges.len() }),
    );
    if input.ranges.is_empty() || input.ranges.len() > 1000 {
        return Err("LaTeX redraw font preflight must contain 1 to 1000 ranges".to_string());
    }
    let request = read_request(&session_id)?;
    if request.host == "powerpoint" && request.operation.as_deref() == Some("latexRedraw") {
        return presentation_redraw::font_sizes(&request, &input);
    }
    let public_request = document_import_request_data(&request)?;
    if public_request.operation != "latexRedraw" {
        return Err("Only a Word LaTeX redraw session can resolve source font sizes".to_string());
    }
    let source = public_request
        .source
        .as_deref()
        .ok_or_else(|| "LaTeX redraw source snapshot is missing".to_string())?;
    let source_utf16 = source.encode_utf16().collect::<Vec<_>>();
    let mut previous_end = 0usize;
    for range in &input.ranges {
        validate_latex_redraw_range(
            &source_utf16,
            &mut previous_end,
            range.source_start,
            range.source_end,
            &range.source_text,
        )?;
        if !matches!(range.display_mode.as_str(), "inline" | "block") {
            return Err("LaTeX redraw font preflight contains an invalid display mode".to_string());
        }
    }

    if request.client.as_deref() == Some("wps") {
        // The WPS add-in reads each formula's size from the document as
        // VTInheritedLatexRedrawFontSize does.
        let ranges = input
            .ranges
            .iter()
            .map(|range| {
                json!({
                    "sourceStart": range.source_start,
                    "sourceEnd": range.source_end,
                    "sourceText": range.source_text,
                    "displayMode": range.display_mode,
                })
            })
            .collect::<Vec<_>>();
        let values = crate::office::wps::bridge::request_font_sizes(&session_id, Value::Array(ranges))?;
        if values.len() != input.ranges.len()
            || values
                .iter()
                .any(|value| !value.is_finite() || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(value))
        {
            return Err("WPS returned invalid LaTeX redraw font sizes".to_string());
        }
        queue_editor_performance(
            OfficeHost::Word,
            &session_id,
            "latex-redraw-font-backend-complete",
            preflight_started.elapsed().as_secs_f64() * 1000.0,
            None,
            json!({ "itemCount": values.len() }),
        );
        return Ok(values);
    }

    let manifest_path = latex_redraw_preflight_manifest_path(&session_id)?;
    let result_path = latex_redraw_font_sizes_path(&session_id)?;
    let mut entries = vec![
        ("protocolVersion".to_string(), OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId".to_string(), session_id.clone()),
        ("operation".to_string(), "latexRedraw".to_string()),
        (
            "outputKind".to_string(),
            public_request
                .output_kind
                .clone()
                .ok_or_else(|| "LaTeX redraw output kind is missing".to_string())?,
        ),
        (
            "sourceDocumentId".to_string(),
            public_request.source_document_id.clone(),
        ),
        ("bookmarkName".to_string(), public_request.bookmark_name.clone()),
        ("itemCount".to_string(), input.ranges.len().to_string()),
    ];
    for (index, range) in input.ranges.iter().enumerate() {
        let prefix = format!("item{index}");
        entries.push((format!("{prefix}sourceStart"), range.source_start.to_string()));
        entries.push((format!("{prefix}sourceEnd"), range.source_end.to_string()));
        entries.push((
            format!("{prefix}sourceTextBase64"),
            URL_SAFE_NO_PAD.encode(range.source_text.as_bytes()),
        ));
        entries.push((
            format!("{prefix}displayMode"),
            range.display_mode.clone(),
        ));
    }
    let manifest = dynamic_dispatch_text(&entries)?;
    if manifest.len() > MAX_DOCUMENT_IMPORT_MANIFEST_BYTES {
        return Err("LaTeX redraw font preflight exceeds the transfer limit".to_string());
    }
    atomic_write(&manifest_path, manifest.as_bytes(), 0o600)?;
    let _ = fs::remove_file(&result_path);

    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session_id.clone()),
        ("action", "latexRedrawPreflight".to_string()),
        ("host", "word".to_string()),
        (
            "sourceDocumentId",
            public_request.source_document_id.clone(),
        ),
        ("bookmarkName", public_request.bookmark_name.clone()),
        (
            "documentImportPath",
            manifest_path.to_string_lossy().to_string(),
        ),
        (
            "fontSizeResultPath",
            result_path.to_string_lossy().to_string(),
        ),
    ])?;
    atomic_write(
        &dispatch_path(OfficeHost::Word, &session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    with_dispatch_pointer(OfficeHost::Word, &session_id, || {
        run_vba_callback(OfficeHost::Word)
    })?;
    let result = fs::read_to_string(&result_path)
        .map_err(|error| format!("Unable to read Word LaTeX redraw font sizes: {error}"))?;
    let values = parse_latex_redraw_font_sizes(&result, input.ranges.len())?;
    queue_editor_performance(
        OfficeHost::Word,
        &session_id,
        "latex-redraw-font-backend-complete",
        preflight_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({ "itemCount": values.len() }),
    );
    let _ = fs::remove_file(manifest_path);
    let _ = fs::remove_file(result_path);
    Ok(values)
}

pub(super) fn commit_document_import_blocking(
    state: OfficeCompanionState,
    session_id: String,
    input: MacOfflineDocumentImportCommitInput,
) -> Result<(), String> {
    let commit_started = Instant::now();
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    if request.host == "powerpoint" {
        let committed = presentation_redraw::commit(&state, &request, input);
        presentation_redraw::remove_files(&session_id);
        let _ = cleanup_session_files(OfficeHost::Powerpoint, &session_id);
        return committed;
    }
    let public_request = document_import_request_data(&request)?;
    let is_redraw = public_request.operation == "latexRedraw";
    let is_formula_restore = public_request.operation == "formulaRestore";
    let is_range_replace = is_redraw || is_formula_restore;
    // A WPS operation hands the add-in each item with its own files instead of
    // Word's manifest and shared staging documents.
    let wps = request.client.as_deref() == Some("wps");
    let mut wps_items: Vec<Value> = Vec::new();
    let mut wps_files: Vec<(String, PathBuf)> = Vec::new();
    if is_redraw {
        queue_editor_performance(
            OfficeHost::Word,
            &session_id,
            "latex-redraw-backend-start",
            0.0,
            None,
            json!({ "itemCount": input.items.len(), "outputKind": input.output_kind }),
        );
    }
    let valid_output = if is_formula_restore {
        matches!(input.output_kind.as_str(), "latex" | "image" | "omml")
    } else {
        matches!(input.output_kind.as_str(), "omml" | "image")
    };
    if !valid_output {
        return Err("Document formula output kind is invalid for this operation".to_string());
    }
    if is_range_replace
        && public_request.output_kind.as_deref() != Some(input.output_kind.as_str())
    {
        return Err("Word range replacement output kind changed after the request".to_string());
    }
    let maximum_items = if is_range_replace { 1000 } else { 2048 };
    if input.items.is_empty() || input.items.len() > maximum_items {
        return Err(format!(
            "Document operation must contain 1 to {maximum_items} blocks"
        ));
    }
    let range_source_utf16 = public_request
        .source
        .as_deref()
        .map(|source| source.encode_utf16().collect::<Vec<_>>());
    if is_range_replace {
        let source = range_source_utf16
            .as_deref()
            .ok_or_else(|| "Word range replacement source snapshot is missing".to_string())?;
        let mut preflight_end = 0usize;
        for item in &input.items {
            let (source_start, source_end, source_text, has_paragraph_metadata, item_kind) =
                match item {
                    MacOfflineDocumentImportCommitItem::Formula {
                        source_start,
                        source_end,
                        source_text,
                        paragraph_id,
                        paragraph_style,
                        paragraph_alignment,
                        list_kind,
                        list_level,
                        paragraph_start,
                        paragraph_end,
                        ..
                    } => (
                        *source_start,
                        *source_end,
                        source_text.as_deref(),
                        paragraph_id.is_some()
                            || paragraph_style.is_some()
                            || paragraph_alignment.is_some()
                            || list_kind.is_some()
                            || list_level.is_some()
                            || *paragraph_start
                            || *paragraph_end,
                        "formula",
                    ),
                    MacOfflineDocumentImportCommitItem::Text {
                        source_start,
                        source_end,
                        source_text,
                        paragraph_id,
                        paragraph_style,
                        paragraph_alignment,
                        list_kind,
                        list_level,
                        paragraph_start,
                        paragraph_end,
                        ..
                    } => (
                        *source_start,
                        *source_end,
                        source_text.as_deref(),
                        paragraph_id.is_some()
                            || paragraph_style.is_some()
                            || paragraph_alignment.is_some()
                            || list_kind.is_some()
                            || list_level.is_some()
                            || *paragraph_start
                            || *paragraph_end,
                        "text",
                    ),
                };
            if has_paragraph_metadata {
                return Err("Word range replacement items cannot carry paragraph metadata".to_string());
            }
            if (is_redraw && item_kind != "formula")
                || (is_formula_restore
                    && ((input.output_kind == "latex" && item_kind != "text")
                        || (input.output_kind != "latex" && item_kind != "formula")))
            {
                return Err("Word range replacement contains an incompatible item kind".to_string());
            }
            let start = source_start
                .ok_or_else(|| "Word replacement item is missing sourceStart".to_string())?;
            let end = source_end
                .ok_or_else(|| "Word replacement item is missing sourceEnd".to_string())?;
            let original = source_text
                .ok_or_else(|| "Word replacement item is missing sourceText".to_string())?;
            if is_formula_restore {
                validate_formula_restore_range(
                    &mut preflight_end,
                    start,
                    end,
                    original,
                )?;
            } else {
                validate_latex_redraw_range(
                    source,
                    &mut preflight_end,
                    start,
                    end,
                    original,
                )?;
            }
        }
    }
    let mut previous_redraw_end = 0usize;

    let mut entries = vec![
        (
            "protocolVersion".to_string(),
            OFFLINE_PROTOCOL_VERSION.to_string(),
        ),
        ("sessionId".to_string(), session_id.clone()),
        ("operation".to_string(), public_request.operation.clone()),
        ("outputKind".to_string(), input.output_kind.clone()),
        (
            "sourceDocumentId".to_string(),
            public_request.source_document_id.clone(),
        ),
        (
            "bookmarkName".to_string(),
            public_request.bookmark_name.clone(),
        ),
        ("itemCount".to_string(), input.items.len().to_string()),
    ];
    if is_range_replace {
        if public_request.font_refresh {
            // the commit also brings the rest of the format up to date, in the same Undo step
            entries.push(("fontRefresh".to_string(), "1".to_string()));
        }
        entries.push((
            "redrawScope".to_string(),
            public_request.redraw_scope.clone().unwrap_or_default(),
        ));
        entries.push((
            "sourceKind".to_string(),
            public_request.source_kind.clone().unwrap_or_default(),
        ));
        let target_snapshot = public_request
            .source
            .as_deref()
            .ok_or_else(|| "Word range replacement source snapshot is missing".to_string())?;
        entries.push((
            "targetTextBase64".to_string(),
            URL_SAFE_NO_PAD.encode(target_snapshot.as_bytes()),
        ));
    }
    // Import and redraw use the same shared drawing carrier. Opening a separate
    // hidden Word document for every imported image dominates batch latency.
    let redraw_vector_batch_path = if input.output_kind == "image" && !wps {
        Some(latex_redraw_vector_batch_path(&session_id)?)
    } else {
        None
    };
    let native_batch_document_path = if input.output_kind == "omml" && !wps {
        Some(document_native_batch_path(&session_id)?)
    } else {
        None
    };
    if let Some(batch_path) = native_batch_document_path.as_ref() {
        entries.push((
            "nativeBatchDocumentPath".to_string(),
            batch_path.to_string_lossy().to_string(),
        ));
    }
    let mut redraw_vector_entries = Vec::new();
    let mut native_batch_entries = Vec::<String>::new();
    let mut metadata_to_cache = Vec::new();
    let mut formula_count = 0usize;
    let mut text_bytes = 0usize;
    let mut active_paragraph_id: Option<String> = None;

    for (index, item) in input.items.iter().enumerate() {
        let prefix = format!("item{index}");
        match item {
            MacOfflineDocumentImportCommitItem::Text {
                text,
                source_start,
                source_end,
                source_text,
                paragraph_id,
                paragraph_style,
                paragraph_alignment,
                list_kind,
                list_level,
                paragraph_start,
                paragraph_end,
            } => {
                if is_redraw {
                    return Err("LaTeX redraw accepts formula items only".to_string());
                }
                if !is_formula_restore
                    && (source_start.is_some() || source_end.is_some() || source_text.is_some())
                {
                    return Err("Document import text contains range replacement coordinates".to_string());
                }
                let paragraph = resolve_document_paragraph_transfer(
                    paragraph_id,
                    paragraph_style,
                    paragraph_alignment,
                    list_kind,
                    *list_level,
                    *paragraph_start,
                    *paragraph_end,
                )?;
                if let Some(paragraph) = paragraph.as_ref() {
                    if paragraph.start {
                        if active_paragraph_id.is_some() {
                            return Err(
                                "Document paragraphs overlap in the transfer stream".to_string()
                            );
                        }
                        active_paragraph_id = Some(paragraph.id.clone());
                    } else if active_paragraph_id.as_deref() != Some(paragraph.id.as_str()) {
                        return Err(
                            "Document paragraph continuation has no matching start".to_string()
                        );
                    }
                } else if active_paragraph_id.is_some() {
                    return Err(
                        "Document paragraph content is missing paragraph metadata".to_string()
                    );
                }
                text_bytes = text_bytes.saturating_add(text.len());
                if text_bytes > 4 * 1024 * 1024 {
                    return Err("Document import text exceeds the 4 MB limit".to_string());
                }
                entries.push((format!("{prefix}kind"), "text".to_string()));
                entries.push((
                    format!("{prefix}textBase64"),
                    URL_SAFE_NO_PAD.encode(text.as_bytes()),
                ));
                if is_formula_restore {
                    let (Some(source_start), Some(source_end), Some(source_text)) =
                        (*source_start, *source_end, source_text.as_deref())
                    else {
                        return Err(
                            "Formula restore text is missing validated source coordinates"
                                .to_string(),
                        );
                    };
                    entries.push((
                        format!("{prefix}sourceStart"),
                        source_start.to_string(),
                    ));
                    entries.push((
                        format!("{prefix}sourceEnd"),
                        source_end.to_string(),
                    ));
                    entries.push((
                        format!("{prefix}sourceTextBase64"),
                        URL_SAFE_NO_PAD.encode(source_text.as_bytes()),
                    ));
                }
                append_document_paragraph_entries(&mut entries, &prefix, paragraph.as_ref());
                if wps {
                    wps_items.push(json!({
                        "kind": "text",
                        "text": text,
                        "sourceStart": source_start,
                        "sourceEnd": source_end,
                        "sourceText": source_text,
                        "paragraph": paragraph.as_ref().map(wps_paragraph),
                    }));
                }
                if paragraph.as_ref().is_some_and(|value| value.end) {
                    active_paragraph_id = None;
                }
            }
            MacOfflineDocumentImportCommitItem::Formula {
                formula_id,
                latex,
                display_mode,
                numbered,
                font_size_pt,
                metadata,
                omml_base64,
                omml_docx_base64,
                word_math_ml,
                svg_base64,
                png_base64,
                width,
                height,
                baseline,
                ink_center_y_ratio,
                source_start,
                source_end,
                source_text,
                paragraph_id,
                paragraph_style,
                paragraph_alignment,
                list_kind,
                list_level,
                paragraph_start,
                paragraph_end,
            } => {
                let paragraph = resolve_document_paragraph_transfer(
                    paragraph_id,
                    paragraph_style,
                    paragraph_alignment,
                    list_kind,
                    *list_level,
                    *paragraph_start,
                    *paragraph_end,
                )?;
                if is_range_replace {
                    if paragraph.is_some() {
                        return Err(
                            "Word range replacement formulas cannot carry paragraph metadata".to_string()
                        );
                    }
                    let start = (*source_start)
                        .ok_or_else(|| "Word replacement formula is missing sourceStart".to_string())?;
                    let end = (*source_end)
                        .ok_or_else(|| "Word replacement formula is missing sourceEnd".to_string())?;
                    let original = source_text
                        .as_deref()
                        .ok_or_else(|| "Word replacement formula is missing sourceText".to_string())?;
                    let source = range_source_utf16
                        .as_ref()
                        .ok_or_else(|| "Word replacement source snapshot is missing".to_string())?;
                    if is_formula_restore {
                        validate_formula_restore_range(
                            &mut previous_redraw_end,
                            start,
                            end,
                            original,
                        )?;
                    } else {
                        validate_latex_redraw_range(
                            source,
                            &mut previous_redraw_end,
                            start,
                            end,
                            original,
                        )?;
                    }
                } else if source_start.is_some() || source_end.is_some() || source_text.is_some() {
                    return Err(
                        "Document import formula contains LaTeX redraw coordinates".to_string()
                    );
                }
                if display_mode == "block" && paragraph.is_some() {
                    return Err("Display formulas must own their Word paragraph".to_string());
                }
                if let Some(paragraph) = paragraph.as_ref() {
                    if paragraph.start {
                        if active_paragraph_id.is_some() {
                            return Err(
                                "Document paragraphs overlap in the transfer stream".to_string()
                            );
                        }
                        active_paragraph_id = Some(paragraph.id.clone());
                    } else if active_paragraph_id.as_deref() != Some(paragraph.id.as_str()) {
                        return Err(
                            "Document paragraph continuation has no matching start".to_string()
                        );
                    }
                } else if active_paragraph_id.is_some() {
                    return Err(
                        "Document paragraph content is missing paragraph metadata".to_string()
                    );
                }
                formula_count += 1;
                if formula_count > 1000 {
                    return Err("Document operations support at most 1000 formulas".to_string());
                }
                validate_uuid(formula_id, "Document formula id")?;
                if latex.trim().is_empty() || latex.len() > 1_000_000 || latex.contains('\0') {
                    return Err(
                        "A document formula contains invalid or excessive LaTeX".to_string()
                    );
                }
                if !matches!(display_mode.as_str(), "inline" | "block") {
                    return Err("Document formula display mode is invalid".to_string());
                }
                if *numbered && display_mode != "block" {
                    return Err("Only document display formulas can be numbered".to_string());
                }
                if !font_size_pt.is_finite()
                    || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(font_size_pt)
                {
                    return Err(
                        "Document formula font size is outside the supported range".to_string()
                    );
                }
                // Equation format v2: the shared writer makes the equation from
                // Windows' MathML (see commit_word).
                let shared_word_equation = word_math_ml
                    .as_deref()
                    .map(|mathml| {
                        crate::office::omml_write::word_equation_artifacts(
                            mathml,
                            word_equation_math_font(&request),
                        )
                    })
                    .transpose()?;
                let omml_base64: &String = match shared_word_equation.as_ref() {
                    Some(artifacts) => &artifacts.omml_base64,
                    None => omml_base64,
                };
                validate_document_omml_payload(omml_base64)?;
                let native_docx = match shared_word_equation.as_ref() {
                    Some(artifacts) => artifacts.docx.clone(),
                    None => decode_document_native_docx(omml_docx_base64)?,
                };
                let native_document_path = native_word_document_path(formula_id)?;
                atomic_write(&native_document_path, &native_docx, 0o600)?;
                let native_batch_document_index = if native_batch_document_path.is_some() {
                    let omml_fragment = match shared_word_equation.as_ref() {
                        Some(artifacts) => artifacts.insertion_omml.clone(),
                        None => {
                            let omml_bytes = URL_SAFE_NO_PAD
                                .decode(omml_base64)
                                .map_err(|_| "Document formula OMML payload is not valid Base64URL".to_string())?;
                            String::from_utf8(omml_bytes)
                                .map_err(|_| "Document formula OMML payload is not UTF-8".to_string())?
                        }
                    };
                    native_batch_entries.push(omml_fragment);
                    native_batch_entries.len()
                } else {
                    0
                };

                let mut resolved_metadata = metadata.clone();
                let canonical_latex = validate_document_formula_metadata_match(
                    &resolved_metadata,
                    formula_id,
                    latex,
                    display_mode,
                    *numbered,
                )?;
                resolved_metadata.latex = canonical_latex.clone();
                resolved_metadata.font_size_pt = Some(*font_size_pt);

                let mut image_path = String::new();
                let mut vector_document_path = String::new();
                let mut vector_document_index = 0usize;
                let mut fallback_image_path = String::new();
                let width = width
                    .ok_or_else(|| "Document formula width is missing".to_string())?;
                let height = height
                    .ok_or_else(|| "Document formula height is missing".to_string())?;
                let baseline = baseline.ok_or_else(|| {
                    "Document formula is missing its mathematical baseline".to_string()
                })?;
                let geometry = calculate_document_image_geometry(
                    width,
                    height,
                    baseline,
                    *font_size_pt,
                    resolved_metadata.formula_letter_font.as_deref(),
                )?;

                if document_output_requires_image_artifacts(&input.output_kind) {
                    let svg_value = svg_base64
                        .as_deref()
                        .ok_or_else(|| "Document image formula is missing SVG data".to_string())?;
                    let svg = decode_stabilized_word_svg(svg_value)?;
                    let png = decode_document_image_fallback_png(png_base64.as_deref())?;
                    let image_package =
                        build_word_svg_docx(&svg, &png, geometry.width, geometry.height)?;
                    let (cached_svg_path, cached_document_path, cached_png_path) =
                        word_image_cache_paths(formula_id)?;
                    atomic_write(&cached_svg_path, &svg, 0o600)?;
                    atomic_write(&cached_document_path, &image_package, 0o600)?;
                    atomic_write(&cached_png_path, &png, 0o600)?;
                    let svg_path = document_formula_file_path(&session_id, formula_id, "svg")?;
                    let png_path = document_formula_file_path(&session_id, formula_id, "png")?;
                    atomic_write(&svg_path, &svg, 0o600)?;
                    atomic_write(&png_path, &png, 0o600)?;
                    if let Some(batch_path) = redraw_vector_batch_path.as_ref() {
                        vector_document_index = redraw_vector_entries.len() + 1;
                        redraw_vector_entries.push(WordSvgBatchEntry {
                            svg: svg.clone(),
                            png: png.clone(),
                            width_points: geometry.width,
                            height_points: geometry.height,
                        });
                        vector_document_path = batch_path.to_string_lossy().to_string();
                    } else {
                        let vector_path =
                            document_formula_file_path(&session_id, formula_id, "docx")?;
                        atomic_write(&vector_path, &image_package, 0o600)?;
                        vector_document_path = vector_path.to_string_lossy().to_string();
                    }
                    image_path = svg_path.to_string_lossy().to_string();
                    fallback_image_path = png_path.to_string_lossy().to_string();
                }
                resolved_metadata.render_width_px = Some(width);
                resolved_metadata.render_height_px = Some(height);
                resolved_metadata.reference_width_pt = Some(geometry.reference_width_pt);
                resolved_metadata.reference_height_pt = Some(geometry.reference_height_pt);
                resolved_metadata.reference_baseline_pt = Some(geometry.reference_baseline_pt);
                resolved_metadata.image_ink_center_y_ratio = ink_center_y_ratio
                    .filter(|value| value.is_finite() && (0.0..=1.0).contains(value));
                let encoded_metadata = encode_metadata(&resolved_metadata)?;
                let item_letter_font = resolved_metadata.formula_letter_font.clone();
                metadata_to_cache.push(resolved_metadata);

                let preserve_formula_identity = is_formula_restore
                    && public_request
                        .restore_targets
                        .as_ref()
                        .and_then(|targets| targets.get(index))
                        .and_then(|target| target.formula_id.as_deref())
                        == Some(formula_id.as_str());
                if wps {
                    let latex_base64 = URL_SAFE_NO_PAD.encode(canonical_latex.as_bytes());
                    let ink_center = ink_center_y_ratio
                        .filter(|value| value.is_finite() && (0.0..=1.0).contains(value));
                    let mut files = json!({});
                    if input.output_kind == "omml" {
                        // The equation in WPS's package form (omml_write::wps_equation_docx).
                        let equation = match shared_word_equation.as_ref() {
                            Some(artifacts) => artifacts.insertion_omml.clone(),
                            None => String::from_utf8(
                                URL_SAFE_NO_PAD
                                    .decode(omml_base64)
                                    .map_err(|_| "Document formula OMML payload is not valid Base64URL".to_string())?,
                            )
                            .map_err(|_| "Document formula OMML payload is not UTF-8".to_string())?,
                        };
                        let docx = crate::office::omml_write::wps_equation_docx(&equation, word_equation_math_font(&request))?;
                        let path = session_directory(OfficeHost::Word, &session_id)?
                            .join(format!("document-formula-{formula_id}-wps.docx"));
                        atomic_write(&path, &docx, 0o600)?;
                        let native_name = format!("{session_id}-{index}.docx");
                        wps_files.push((native_name.clone(), path));
                        files["native"] = json!(native_name);
                    }
                    let variables = if document_output_requires_image_artifacts(&input.output_kind) {
                        let svg_name = format!("{session_id}-{index}.svg");
                        let png_name = format!("{session_id}-{index}.png");
                        wps_files.push((svg_name.clone(), PathBuf::from(&image_path)));
                        wps_files.push((png_name.clone(), PathBuf::from(&fallback_image_path)));
                        files["image"] = json!(svg_name);
                        files["fallback"] = json!(png_name);
                        Some(crate::office::wps::payload::word_image_variables(&crate::office::wps::payload::WordImageState {
                            formula_id,
                            display_mode,
                            numbered: *numbered,
                            latex_base64: &latex_base64,
                            omml_base64,
                            metadata: &encoded_metadata,
                            font_size_pt: *font_size_pt,
                            reference_width_pt: geometry.reference_width_pt,
                            reference_height_pt: geometry.reference_height_pt,
                            reference_baseline_pt: geometry.reference_baseline_pt,
                            ink_center_y_ratio: ink_center,
                            letter_font: item_letter_font.as_deref().unwrap_or(""),
                        })?)
                    } else {
                        None
                    };
                    // A numbered formula is written without its number and numbered
                    // with the others once all are in place (office::wps::numbering).
                    wps_items.push(json!({
                        "kind": "formula",
                        "formulaId": formula_id,
                        "preserveIdentity": preserve_formula_identity,
                        "displayMode": display_mode,
                        "numbered": numbered,
                        "nativeEquation": input.output_kind == "omml",
                        "fontSizePt": font_size_pt,
                        "metadata": encoded_metadata,
                        "formulaReference": crate::office::wps::payload::formula_reference(formula_id, display_mode, false),
                        "ownerBookmark": crate::office::wps::payload::owner_bookmark_name(formula_id),
                        "formulaLetterFont": item_letter_font,
                        "widthPoints": geometry.width,
                        "heightPoints": geometry.height,
                        "referenceWidthPt": geometry.reference_width_pt,
                        "referenceHeightPt": geometry.reference_height_pt,
                        "referenceBaselinePt": geometry.reference_baseline_pt,
                        "inkCenterYRatio": ink_center,
                        "sourceStart": source_start,
                        "sourceEnd": source_end,
                        "sourceText": source_text,
                        "paragraph": paragraph.as_ref().map(wps_paragraph),
                        "files": files,
                        "variables": variables,
                    }));
                }
                entries.push((format!("{prefix}kind"), "formula".to_string()));
                entries.push((format!("{prefix}formulaId"), formula_id.clone()));
                entries.push((
                    format!("{prefix}preserveIdentity"),
                    if preserve_formula_identity { "1" } else { "0" }.to_string(),
                ));
                entries.push((
                    format!("{prefix}latexBase64"),
                    URL_SAFE_NO_PAD.encode(canonical_latex.as_bytes()),
                ));
                entries.push((format!("{prefix}displayMode"), display_mode.clone()));
                entries.push((
                    format!("{prefix}numbered"),
                    if *numbered { "1" } else { "0" }.to_string(),
                ));
                entries.push((
                    format!("{prefix}fontSizePt"),
                    format!("{:.6}", font_size_pt),
                ));
                entries.push((format!("{prefix}metadata"), encoded_metadata));
                entries.push((format!("{prefix}ommlBase64"), omml_base64.clone()));
                entries.push((
                    format!("{prefix}nativeDocumentPath"),
                    native_document_path.to_string_lossy().to_string(),
                ));
                if native_batch_document_index > 0 {
                    entries.push((
                        format!("{prefix}nativeBatchDocumentIndex"),
                        native_batch_document_index.to_string(),
                    ));
                }
                entries.push((format!("{prefix}imagePath"), image_path));
                entries.push((format!("{prefix}vectorDocumentPath"), vector_document_path));
                if vector_document_index > 0 {
                    entries.push((
                        format!("{prefix}vectorDocumentIndex"),
                        vector_document_index.to_string(),
                    ));
                }
                entries.push((format!("{prefix}fallbackImagePath"), fallback_image_path));
                entries.push((
                    format!("{prefix}widthPoints"),
                    format!("{:.6}", geometry.width),
                ));
                entries.push((
                    format!("{prefix}heightPoints"),
                    format!("{:.6}", geometry.height),
                ));
                entries.push((format!("{prefix}baseline"), geometry.baseline.to_string()));
                entries.push((
                    format!("{prefix}referenceWidthPt"),
                    format!("{:.6}", geometry.reference_width_pt),
                ));
                entries.push((
                    format!("{prefix}referenceHeightPt"),
                    format!("{:.6}", geometry.reference_height_pt),
                ));
                entries.push((
                    format!("{prefix}referenceBaselinePt"),
                    format!("{:.6}", geometry.reference_baseline_pt),
                ));
                if let Some(value) = ink_center_y_ratio
                    .filter(|value| value.is_finite() && (0.0..=1.0).contains(value))
                {
                    entries.push((
                        format!("{prefix}inkCenterYRatio"),
                        format!("{value:.8}"),
                    ));
                }
                if let Some(font) = item_letter_font {
                    entries.push((format!("{prefix}formulaLetterFont"), font));
                }
                if is_range_replace {
                    let (Some(source_start), Some(source_end), Some(source_text)) =
                        (*source_start, *source_end, source_text.as_deref())
                    else {
                        return Err(
                            "Word replacement formula is missing validated source coordinates"
                                .to_string(),
                        );
                    };
                    entries.push((
                        format!("{prefix}sourceStart"),
                        source_start.to_string(),
                    ));
                    entries.push((
                        format!("{prefix}sourceEnd"),
                        source_end.to_string(),
                    ));
                    entries.push((
                        format!("{prefix}sourceTextBase64"),
                        URL_SAFE_NO_PAD.encode(source_text.as_bytes()),
                    ));
                }
                append_document_paragraph_entries(&mut entries, &prefix, paragraph.as_ref());
                if paragraph.as_ref().is_some_and(|value| value.end) {
                    active_paragraph_id = None;
                }
            }
        }
    }
    if active_paragraph_id.is_some() {
        return Err("Document paragraph transfer ended before its paragraph boundary".to_string());
    }
    if formula_count == 0 && text_bytes == 0 {
        return Err("Document import contains no visible content".to_string());
    }
    if wps {
        write_document_progress(&session_id, 0, input.items.len(), "inserting")?;
        let payload = json!({
            "operation": public_request.operation,
            "outputKind": input.output_kind,
            "sourceKind": public_request.source_kind,
            "redrawScope": public_request.redraw_scope,
            "fontRefresh": public_request.font_refresh,
            "bookmarkName": public_request.bookmark_name,
            "sourceDocumentId": public_request.source_document_id,
            "mathFont": crate::office::omml_write::normalize_math_font_name(word_equation_math_font(&request)),
            "items": wps_items,
        });
        crate::office::wps::bridge::deliver_document_commit(&session_id, payload, wps_files, input.items.len())?;
        for metadata in metadata_to_cache {
            let formula_id = metadata.formula_id.clone();
            let _ = state.formula_cache.put(&formula_id, metadata);
        }
        let _ = cleanup_session_files(OfficeHost::Word, &session_id);
        return Ok(());
    }
    if let Some(batch_path) = redraw_vector_batch_path.as_ref() {
        if !redraw_vector_entries.is_empty() {
            let package = build_word_svg_batch_docx(&redraw_vector_entries)?;
            atomic_write(batch_path, &package, 0o600)?;
            entries.push((
                "vectorBatchDocumentPath".to_string(),
                batch_path.to_string_lossy().to_string(),
            ));
            entries.push((
                "vectorBatchFormulaCount".to_string(),
                redraw_vector_entries.len().to_string(),
            ));
        }
    }
    if let Some(batch_path) = native_batch_document_path.as_ref() {
        let package = build_word_omml_batch_docx(&native_batch_entries)?;
        atomic_write(batch_path, &package, 0o600)?;
        entries.push((
            "nativeBatchFormulaCount".to_string(),
            native_batch_entries.len().to_string(),
        ));
    }

    let manifest = dynamic_dispatch_text(&entries)?;
    if manifest.len() > MAX_DOCUMENT_IMPORT_MANIFEST_BYTES {
        return Err(
            "Document import expands beyond the 16 MB Word transfer limit; split the source into smaller imports"
                .to_string(),
        );
    }
    let manifest_path = document_import_manifest_path(&session_id)?;
    atomic_write(&manifest_path, manifest.as_bytes(), 0o600)?;
    let progress_path =
        session_directory(OfficeHost::Word, &session_id)?.join(DOCUMENT_IMPORT_PROGRESS_FILE);
    atomic_write(
        &progress_path,
        format!("current=0\ntotal={}\nstage=preparing\n", input.items.len()).as_bytes(),
        0o600,
    )?;
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session_id.clone()),
        ("action", "documentCommit".to_string()),
        ("host", "word".to_string()),
        (
            "performanceTrace",
            if word_performance_trace_enabled() { "1" } else { "0" }.to_string(),
        ),
        (
            "sourceDocumentId",
            public_request.source_document_id.clone(),
        ),
        ("bookmarkName", public_request.bookmark_name.clone()),
        (
            "documentImportPath",
            manifest_path.to_string_lossy().to_string(),
        ),
    ])?;
    atomic_write(
        &dispatch_path(OfficeHost::Word, &session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    if is_redraw {
        queue_editor_performance(
            OfficeHost::Word,
            &session_id,
            "latex-redraw-backend-prepared",
            commit_started.elapsed().as_secs_f64() * 1000.0,
            None,
            json!({ "itemCount": input.items.len() }),
        );
        queue_editor_performance(
            OfficeHost::Word,
            &session_id,
            "latex-redraw-word-callback-start",
            commit_started.elapsed().as_secs_f64() * 1000.0,
            None,
            json!({ "itemCount": input.items.len() }),
        );
    }
    with_dispatch_pointer(OfficeHost::Word, &session_id, || {
        run_vba_callback(OfficeHost::Word)
    })?;
    if is_redraw {
        queue_editor_performance(
            OfficeHost::Word,
            &session_id,
            "latex-redraw-word-callback-complete",
            commit_started.elapsed().as_secs_f64() * 1000.0,
            None,
            json!({ "itemCount": input.items.len() }),
        );
    }

    for metadata in metadata_to_cache {
        let formula_id = metadata.formula_id.clone();
        let _ = state.formula_cache.put(&formula_id, metadata);
    }
    let _ = cleanup_session_files(OfficeHost::Word, &session_id);
    Ok(())
}

/// The import window's progress (get_macos_offline_document_import_progress),
/// written for an operation the WPS add-in carries out.
pub(crate) fn write_document_progress(
    session_id: &str,
    current: usize,
    total: usize,
    stage: &str,
) -> Result<(), String> {
    validate_uuid(session_id, "Session id")?;
    if current > total || total > 2048 || !matches!(stage, "preparing" | "inserting" | "complete" | "error") {
        return Err("Document operation progress is invalid".to_string());
    }
    let host = read_request(session_id)
        .ok()
        .and_then(|request| host_from_request_name(&request.host).ok())
        .unwrap_or(OfficeHost::Word);
    let path = session_directory(host, session_id)?.join(DOCUMENT_IMPORT_PROGRESS_FILE);
    atomic_write_runtime(
        &path,
        format!("current={current}\ntotal={total}\nstage={stage}\n").as_bytes(),
        0o600,
    )
}

pub(super) fn wps_paragraph(paragraph: &DocumentParagraphTransfer) -> Value {
    json!({
        "id": paragraph.id,
        "style": paragraph.style,
        "alignment": paragraph.alignment,
        "listKind": paragraph.list_kind,
        "listLevel": paragraph.list_level,
        "start": paragraph.start,
        "end": paragraph.end,
    })
}

pub(super) fn cancel_document_import_blocking(session_id: String) -> Result<(), String> {
    validate_uuid(&session_id, "Session id")?;
    let request = read_request(&session_id)?;
    if request.host == "powerpoint" {
        // A presentation operation changes nothing before its result arrives.
        if request.client.as_deref() == Some("wps") {
            crate::office::wps::bridge::deliver_cancel(&session_id);
        }
        presentation_redraw::remove_files(&session_id);
        let _ = cleanup_session_files(OfficeHost::Powerpoint, &session_id);
        return Ok(());
    }
    if request.client.as_deref() == Some("wps") {
        // The add-in removes its VT_D_ anchor when it sees the cancel.
        crate::office::wps::bridge::deliver_cancel(&session_id);
        let _ = cleanup_session_files(OfficeHost::Word, &session_id);
        return Ok(());
    }
    let public_request = document_import_request_data(&request)?;
    let dispatch = dispatch_text(&[
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session_id.clone()),
        ("action", "documentCancel".to_string()),
        ("host", "word".to_string()),
        ("sourceDocumentId", public_request.source_document_id),
        ("bookmarkName", public_request.bookmark_name),
    ])?;
    atomic_write(
        &dispatch_path(OfficeHost::Word, &session_id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    with_dispatch_pointer(OfficeHost::Word, &session_id, || {
        run_vba_callback(OfficeHost::Word)
    })?;
    let _ = cleanup_session_files(OfficeHost::Word, &session_id);
    Ok(())
}
