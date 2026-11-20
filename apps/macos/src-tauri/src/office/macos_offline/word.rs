//! Committing a formula to Word (and WPS Writer): the math font, numbering-only edits, and handing the result over.

use super::*;

pub(super) fn word_math_font_name_for_formula_letter_font(value: &str) -> &'static str {
    match value {
        "times" => "Times New Roman",
        "cambria" => "Cambria Math",
        "stix" => "STIX Two Math",
        "palatino" => "Palatino",
        "helvetica" => "Helvetica Neue",
        // KaTeX is a web font family and cannot be assigned to a native Word
        // OMath range. Keep Word's native default for that VisualTeX choice.
        _ => "Cambria Math",
    }
}

pub(super) fn resolved_word_math_font_name(
    captured_word_font: Option<&str>,
    initial_formula_font: &str,
    final_formula_font: &str,
) -> String {
    if final_formula_font == initial_formula_font {
        if let Some(font) = captured_word_font.map(str::trim).filter(|font| !font.is_empty()) {
            return font.to_string();
        }
    }
    word_math_font_name_for_formula_letter_font(final_formula_font).to_string()
}

/// The math font of a Word equation (equation format v2): the document's math
/// font, as Windows writes it; an add-in that predates the field gives the
/// edited equation's captured font; Cambria Math otherwise.
pub(super) fn word_equation_math_font(request: &MacOfflineSessionRequest) -> &str {
    request
        .word_document_math_font_name
        .as_deref()
        .or(request.word_math_font_name.as_deref())
        .unwrap_or_default()
}

pub(super) fn word_image_numbering_only_edit(
    request: &MacOfflineSessionRequest,
    original: Option<&VisualTeXFormulaMetadata>,
    current: &VisualTeXFormulaMetadata,
) -> bool {
    let Some(original) = original else {
        return false;
    };
    request.mode == "edit"
        && !request.native_equation
        && !request.fork_copied_formula
        && request.operation.as_deref().unwrap_or("formula") == "formula"
        && original.formula_id == current.formula_id
        && original.display_mode == "block"
        && current.display_mode == "block"
        && original.numbered != current.numbered
        && original.latex == current.latex
        && original.code_format == current.code_format
        && original.font_size_pt == current.font_size_pt
        && original.formula_letter_font == current.formula_letter_font
        && original.formula_chinese_font == current.formula_chinese_font
}

pub(super) fn commit_word(
    app: Option<&AppHandle>,
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
    metadata: &str,
    canonical_latex: &str,
    geometry: WordGeometry,
) -> Result<(), String> {
    let commit_started = Instant::now();
    let numbering_only = word_image_numbering_only_edit(
        request,
        session.original_metadata.as_ref(),
        &decode_metadata(metadata)?,
    );
    let export = session
        .export_result
        .as_ref()
        .ok_or_else(|| "Word Session has no formula export".to_string())?;
    // Equation format v2: the editor sends Windows' MathML and the shared
    // writer (the port of Windows' C# writer) makes the equation, so both
    // platforms insert identical OMML. A number is added later by the add-in
    // through the same writer, once it knows the document's numbering.
    let shared_word_equation = export
        .word_math_ml
        .as_deref()
        .map(|mathml| {
            crate::office::omml_write::word_equation_artifacts(mathml, word_equation_math_font(request))
        })
        .transpose()?;
    let omml_base64: String = match shared_word_equation.as_ref() {
        Some(artifacts) => artifacts.omml_base64.clone(),
        None => export
            .omml_base64
            .as_deref()
            .ok_or_else(|| "Word formula export has no OMML payload".to_string())?
            .to_string(),
    };
    let omml_base64 = omml_base64.as_str();
    let omml_bytes = URL_SAFE_NO_PAD
        .decode(omml_base64)
        .map_err(|_| "Word formula OMML payload is not valid Base64URL".to_string())?;
    if omml_bytes.is_empty() || omml_bytes.len() > MAX_OMML_BYTES {
        return Err("Word formula OMML payload is empty or too large".to_string());
    }
    let omml = std::str::from_utf8(&omml_bytes)
        .map_err(|_| "Word formula OMML payload is not UTF-8".to_string())?;
    if !omml.trim_start().starts_with("<m:oMath")
        || !omml.contains("http://schemas.openxmlformats.org/officeDocument/2006/math")
        || omml.contains("<!DOCTYPE")
        || omml.contains("<!ENTITY")
    {
        return Err("Word formula OMML payload is not a safe Office Math fragment".to_string());
    }
    // Keep both Word representations warm for every formula. Image formulas
    // already carry OMML in their editor export, so dropping its DOCX package
    // forced the first image-to-OMML conversion back through the renderer.
    let omml_docx = match shared_word_equation {
        Some(artifacts) => artifacts.docx,
        None => {
            let omml_docx_base64 = export
                .omml_docx_base64
                .as_deref()
                .ok_or_else(|| "Word formula export has no native DOCX payload".to_string())?;
            URL_SAFE_NO_PAD
                .decode(omml_docx_base64)
                .map_err(|_| "Word formula native DOCX payload is not valid Base64URL".to_string())?
        }
    };
    if omml_docx.len() < 128
        || omml_docx.len() > MAX_OMML_BYTES * 8
        || !omml_docx.starts_with(b"PK\x03\x04")
    {
        return Err("Word formula native DOCX payload is invalid or too large".to_string());
    }
    let native_document_path = native_word_document_path(&session.formula_id)?;
    atomic_write(&native_document_path, &omml_docx, 0o600)?;

    // Every successful Word export refreshes the durable image cache. Native
    // OMML commits do not pass the Session-scoped image paths to VBA, but the
    // formula-scoped cache makes the later OMML-to-image command a sub-second
    // in-Word transaction with no renderer wake-up.
    let prepared_image_artifacts = materialize_word_svg_package(session, geometry)?;
    let image_artifacts = if request.native_equation {
        None
    } else {
        Some(prepared_image_artifacts)
    };
    queue_editor_performance(
        OfficeHost::Word,
        &session.id,
        "apply-word-artifacts-ready",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({ "nativeEquation": request.native_equation }),
    );
    let source_marker = request
        .source_object_id
        .clone()
        .or_else(|| request.encoded_metadata.clone())
        .unwrap_or_default();
    let pending_marker = request.pending_marker.clone().unwrap_or_default();
    let latex = canonical_latex.trim();
    if latex.is_empty() {
        return Err("Word native-equation conversion requires non-empty LaTeX".to_string());
    }
    let latex_base64 = URL_SAFE_NO_PAD.encode(latex.as_bytes());
    let initial_formula_letter_font = request
        .formula_letter_font
        .as_deref()
        .or_else(|| {
            session
                .original_metadata
                .as_ref()
                .and_then(|metadata| metadata.formula_letter_font.as_deref())
        })
        .unwrap_or("cambria");
    let final_word_math_font_name = resolved_word_math_font_name(
        request.word_math_font_name.as_deref(),
        initial_formula_letter_font,
        session.formula_letter_font.as_str(),
    );
    let entries = vec![
        ("protocolVersion", OFFLINE_PROTOCOL_VERSION.to_string()),
        ("sessionId", session.id.clone()),
        ("action", "commit".to_string()),
        ("host", "word".to_string()),
        (
            "operation",
            request
                .operation
                .clone()
                .unwrap_or_else(|| "formula".to_string()),
        ),
        (
            "performanceTrace",
            if word_performance_trace_enabled() { "1" } else { "0" }.to_string(),
        ),
        ("mode", request.mode.clone()),
        ("formulaId", session.formula_id.clone()),
        ("displayMode", session.display_mode.clone()),
        (
            "numberingOnly",
            if numbering_only { "1" } else { "0" }.to_string(),
        ),
        (
            "numbered",
            if session.numbered { "1" } else { "0" }.to_string(),
        ),
        (
            "nativeEquation",
            if request.native_equation { "1" } else { "0" }.to_string(),
        ),
        (
            "imagePath",
            image_artifacts
                .as_ref()
                .map(|artifacts| artifacts.0.to_string_lossy().to_string())
                .unwrap_or_default(),
        ),
        (
            "vectorDocumentPath",
            image_artifacts
                .as_ref()
                .map(|artifacts| artifacts.1.to_string_lossy().to_string())
                .unwrap_or_default(),
        ),
        (
            "fallbackImagePath",
            image_artifacts
                .as_ref()
                .map(|artifacts| artifacts.2.to_string_lossy().to_string())
                .unwrap_or_default(),
        ),
        ("metadata", metadata.to_string()),
        ("latexBase64", latex_base64),
        ("ommlBase64", omml_base64.to_string()),
        (
            "nativeDocumentPath",
            native_document_path.to_string_lossy().to_string(),
        ),
        ("pendingMarker", pending_marker),
        ("sourceMarker", source_marker),
        (
            "forkCopiedFormula",
            if request.fork_copied_formula { "1" } else { "0" }.to_string(),
        ),
        (
            "sourceDocumentId",
            request.source_document_id.clone().unwrap_or_default(),
        ),
        ("widthPoints", format!("{:.6}", geometry.width)),
        ("heightPoints", format!("{:.6}", geometry.height)),
        ("baseline", geometry.baseline.to_string()),
        ("fontSizePt", format!("{:.6}", geometry.font_size_pt)),
        (
            "referenceWidthPt",
            format!("{:.6}", geometry.reference_width_pt),
        ),
        (
            "referenceHeightPt",
            format!("{:.6}", geometry.reference_height_pt),
        ),
        (
            "referenceBaselinePt",
            format!("{:.6}", geometry.reference_baseline_pt),
        ),
        ("wordMathFontName", final_word_math_font_name),
        // An image formula's Equation number takes its letter font.
        ("formulaLetterFont", session.formula_letter_font.clone()),
        (
            "inkCenterYRatio",
            session
                .export_result
                .as_ref()
                .and_then(|value| value.ink_center_y_ratio)
                .filter(|value| value.is_finite() && (0.0..=1.0).contains(value))
                .map(|value| format!("{value:.8}"))
                .unwrap_or_default(),
        ),
    ];
    if request.client.as_deref() == Some("wps") {
        return deliver_wps_word_commit(
            request,
            session,
            &entries,
            geometry,
            image_artifacts.as_ref(),
            &native_document_path,
        );
    }
    let dispatch = dispatch_text(&entries)?;
    atomic_write_runtime(
        &dispatch_path(OfficeHost::Word, &session.id)?,
        dispatch.as_bytes(),
        0o600,
    )?;
    queue_editor_performance(
        OfficeHost::Word,
        &session.id,
        "apply-word-dispatch-ready",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    with_dispatch_pointer(OfficeHost::Word, &session.id, || {
        run_vba_callback_on_main_thread(app, OfficeHost::Word)
    })?;
    queue_editor_performance(
        OfficeHost::Word,
        &session.id,
        "apply-word-vba-complete",
        commit_started.elapsed().as_secs_f64() * 1000.0,
        None,
        json!({}),
    );
    Ok(())
}

/// A Word commit for the WPS add-in: the values the Word macro would read
/// from dispatch.txt, plus the document state it would write (office::wps).
/// Returns once the add-in has written the document.
pub(super) fn deliver_wps_word_commit(
    request: &MacOfflineSessionRequest,
    session: &OfficeFormulaSession,
    entries: &[(&str, String)],
    geometry: WordGeometry,
    image_artifacts: Option<&(PathBuf, PathBuf, PathBuf)>,
    native_document_path: &Path,
) -> Result<(), String> {
    use crate::office::wps::payload;
    let value = |key: &str| {
        entries
            .iter()
            .find(|(name, _)| *name == key)
            .map(|(_, value)| value.as_str())
            .unwrap_or("")
    };
    // An edited numbered formula keeps the number it has in the document; a
    // new one is numbered by the add-in through the bridge (office::wps::numbering).
    let display_mode = session.display_mode.as_str();
    let numbered = session.numbered && display_mode == "block";
    let file_name = |suffix: &str| format!("{}-{suffix}", session.id);
    let mut files = Vec::new();
    let mut file_names = json!({});
    // A picture converted to an equation has no equation number to keep: the
    // add-in numbers the new equation through the bridge.
    let keeps_number = numbered && request.mode == "edit" && request.operation.as_deref().unwrap_or("formula") == "formula";
    let variables = if request.native_equation {
        // The equation in WPS's package form (omml_write::wps_equation_docx).
        let semantic = URL_SAFE_NO_PAD
            .decode(value("ommlBase64"))
            .ok()
            .and_then(|bytes| String::from_utf8(bytes).ok())
            .ok_or_else(|| "The WPS equation has no OMML".to_string())?;
        let font = word_equation_math_font(request);
        let insertion = crate::office::omml_write::insertion_omml(&semantic, font)?;
        let equation = if keeps_number {
            // The edited formula goes back in front of the equation's own
            // "#" and (number), fields and bookmarks as they were.
            let original = request
                .native_edit_xml
                .as_deref()
                .ok_or_else(|| "The numbered WPS equation has no original to keep its number from".to_string())?;
            crate::office::omml_write::keep_equation_number(original, &insertion)?
        } else {
            insertion
        };
        let docx = crate::office::omml_write::wps_equation_docx(&equation, font)?;
        let native_path = native_document_path.with_extension("wps.docx");
        atomic_write(&native_path, &docx, 0o600)?;
        files.push((file_name("equation.docx"), native_path));
        file_names["native"] = json!(file_name("equation.docx"));
        None
    } else {
        let (svg, _, png) = image_artifacts
            .ok_or_else(|| "The WPS image formula has no image artifacts".to_string())?;
        files.push((file_name("formula.svg"), svg.clone()));
        files.push((file_name("formula.png"), png.clone()));
        file_names["image"] = json!(file_name("formula.svg"));
        file_names["fallback"] = json!(file_name("formula.png"));
        Some(payload::word_image_variables(&payload::WordImageState {
            formula_id: &session.formula_id,
            display_mode,
            numbered,
            latex_base64: value("latexBase64"),
            omml_base64: value("ommlBase64"),
            metadata: value("metadata"),
            font_size_pt: geometry.font_size_pt,
            reference_width_pt: geometry.reference_width_pt,
            reference_height_pt: geometry.reference_height_pt,
            reference_baseline_pt: geometry.reference_baseline_pt,
            ink_center_y_ratio: value("inkCenterYRatio").parse::<f64>().ok(),
            letter_font: value("formulaLetterFont"),
        })?)
    };
    let result = json!({
        "sessionId": session.id,
        "mode": request.mode,
        "formulaId": session.formula_id,
        "displayMode": display_mode,
        "nativeEquation": request.native_equation,
        "numbered": numbered,
        // What a new number needs (the bridge's numbering call): the equation
        // body, the document's math font and the letter font of the number.
        "ommlBase64": value("ommlBase64"),
        "mathFont": crate::office::omml_write::normalize_math_font_name(word_equation_math_font(request)),
        "formulaLetterFont": value("formulaLetterFont"),
        "metadata": value("metadata"),
        "formulaReference": payload::formula_reference(&session.formula_id, display_mode, numbered),
        "ownerBookmark": payload::owner_bookmark_name(&session.formula_id),
        "widthPoints": geometry.width,
        "heightPoints": geometry.height,
        "fontSizePt": geometry.font_size_pt,
        "referenceWidthPt": geometry.reference_width_pt,
        "referenceHeightPt": geometry.reference_height_pt,
        "referenceBaselinePt": geometry.reference_baseline_pt,
        // Where the formula's ink is centred, from the top (0-1): a display
        // picture is placed by it (VTExpectedImageFormulaPosition).
        "inkCenterYRatio": value("inkCenterYRatio").parse::<f64>().ok(),
        "variables": variables,
        "files": file_names,
    });
    crate::office::wps::bridge::deliver_commit(&session.id, result, files)
}
