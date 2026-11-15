//! LaTeX redraw and formula restore in PowerPoint and WPS Presentation.
//!
//! The PowerPoint add-in (VTPowerPointRedraw.bas) and the WPS Presentation
//! add-in (office/wps/module-wpp.js) send the text of the selected text boxes,
//! slides or presentation (a redraw), or the formulas to turn back into LaTeX
//! (a restore). Word's redraw window renders the formulas
//! (office/redraw/WordLatexRedrawApp.tsx) and the result goes back to the
//! add-in, which writes it: an equation in place of its LaTeX, a VisualTeX
//! picture in place of a text box that holds only one formula, or LaTeX text
//! in place of an equation or a picture formula.

use super::*;

/// What the add-in sends for a redraw (a JSON file in the Session folder).
pub(super) const REDRAW_SOURCE_FILE: &str = "presentation-redraw-source.json";
/// What the add-in sends for a restore.
pub(super) const RESTORE_SOURCE_FILE: &str = "presentation-restore-source.json";
/// Why PowerPoint could not copy the equations of a restore.
const COPY_ERROR_FILE: &str = "presentation-copy-error.txt";
const MAX_SOURCE_BYTES: u64 = 8 * 1024 * 1024;
const MAX_SEGMENTS: usize = 4096;
const MAX_ITEMS: usize = 1000;
/// Between two text boxes in the redraw window's source text; each box is
/// searched for LaTeX on its own (the window gets the boxes' bounds).
const SEGMENT_SEPARATOR: &str = "\n\n";
/// The redraw window accepts formula sizes from 5 to 200 pt.
const MIN_TEXT_SIZE_PT: f64 = 5.0;
const MAX_TEXT_SIZE_PT: f64 = 200.0;

/// PowerPoint's copy of an equation to restore (copy_powerpoint_equations).
fn copied_equation_name(index: usize) -> String {
    format!("restore-equation-{index}.bin")
}

/// The text of one text box, or the selected part of it.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct TextSegment {
    slide_id: u32,
    shape_id: u32,
    /// Where `text` starts in the shape's text, 1-based (UTF-16 units).
    start: usize,
    text: String,
    /// `text` is the box's whole text (a picture formula may replace the box).
    #[serde(default)]
    whole: bool,
    /// Font sizes over `text`: [start, length, size] runs, 0-based UTF-16.
    #[serde(default)]
    sizes: Vec<(usize, usize, f64)>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct RedrawSource {
    segments: Vec<TextSegment>,
}

/// A formula to turn back into LaTeX.
#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct RestoreTarget {
    slide_id: u32,
    shape_id: u32,
    /// "omml" (a native equation in the shape's text) or "image" (a VisualTeX
    /// picture formula, the whole shape).
    kind: String,
    /// A native equation's range in the shape's text, 1-based (UTF-16 units).
    #[serde(default)]
    start: Option<usize>,
    #[serde(default)]
    length: Option<usize>,
    display_mode: String,
    font_size_pt: f64,
    /// A picture formula's metadata.
    #[serde(default)]
    metadata: Option<String>,
    /// A native equation's a14:m element as WPS saved it (wps::presentation);
    /// PowerPoint copies its equations instead (copy_powerpoint_equations).
    #[serde(default)]
    zone: Option<String>,
}

#[derive(Debug, Clone, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct RestoreSource {
    targets: Vec<RestoreTarget>,
}

/// Where a text box's text is in the redraw window's source text.
#[derive(Debug, Clone, Serialize, PartialEq)]
#[serde(rename_all = "camelCase")]
pub struct PresentationSegment {
    start: usize,
    end: usize,
    whole: bool,
}

fn read_source_file(request: &MacOfflineSessionRequest, name: &str) -> Result<String, String> {
    let path = session_directory(OfficeHost::Powerpoint, &request.session_id)?.join(name);
    let metadata = fs::symlink_metadata(&path).map_err(|error| format!("Unable to inspect {}: {error}", path.display()))?;
    if !metadata.is_file() || metadata.len() == 0 || metadata.len() > MAX_SOURCE_BYTES {
        return Err("The presentation's formula source has an invalid size".to_string());
    }
    fs::read_to_string(&path).map_err(|error| format!("Unable to read {}: {error}", path.display()))
}

fn redraw_source(request: &MacOfflineSessionRequest) -> Result<RedrawSource, String> {
    let source: RedrawSource = serde_json::from_str(&read_source_file(request, REDRAW_SOURCE_FILE)?)
        .map_err(|error| format!("The presentation's redraw source is invalid: {error}"))?;
    if source.segments.is_empty() || source.segments.len() > MAX_SEGMENTS {
        return Err("A presentation redraw needs 1 to 4096 text boxes".to_string());
    }
    for segment in &source.segments {
        let length = segment.text.encode_utf16().count();
        if segment.slide_id == 0 || segment.shape_id == 0 || segment.start == 0 || length == 0 {
            return Err("The presentation's redraw source names an invalid text box".to_string());
        }
        for &(start, run, size) in &segment.sizes {
            if start.checked_add(run).is_none_or(|end| end > length) || !size.is_finite() || size <= 0.0 {
                return Err("The presentation's redraw source has invalid font sizes".to_string());
            }
        }
    }
    Ok(source)
}

/// The boxes' texts one after another, and where each one is.
fn joined(source: &RedrawSource) -> (String, Vec<PresentationSegment>) {
    let mut text = String::new();
    let mut units = 0usize;
    let mut segments = Vec::with_capacity(source.segments.len());
    for (index, segment) in source.segments.iter().enumerate() {
        if index > 0 {
            text.push_str(SEGMENT_SEPARATOR);
            units += SEGMENT_SEPARATOR.encode_utf16().count();
        }
        let length = segment.text.encode_utf16().count();
        segments.push(PresentationSegment { start: units, end: units + length, whole: segment.whole });
        text.push_str(&segment.text);
        units += length;
    }
    (text, segments)
}

fn decoded_latex(metadata: &str) -> Result<VisualTeXFormulaMetadata, String> {
    decode_metadata(metadata)
}

fn clamp_size(size: f64) -> f64 {
    if size.is_finite() { size.clamp(MIN_TEXT_SIZE_PT, MAX_TEXT_SIZE_PT) } else { 18.0 }
}

/// The OMML of an equation PowerPoint copied for this restore.
fn copied_equation(directory: &Path, index: usize) -> Result<String, String> {
    if let Ok(reason) = fs::read_to_string(directory.join(COPY_ERROR_FILE)) {
        return Err(reason);
    }
    let bytes = fs::read(directory.join(copied_equation_name(index)))
        .map_err(|_| format!("PowerPoint did not copy equation {} for VisualTeX", index + 1))?;
    super::super::powerpoint_omml::native_equation_from_clipboard(&bytes)
}

fn restore_source(request: &MacOfflineSessionRequest) -> Result<RestoreSource, String> {
    let source: RestoreSource = serde_json::from_str(&read_source_file(request, RESTORE_SOURCE_FILE)?)
        .map_err(|error| format!("The presentation's restore source is invalid: {error}"))?;
    if source.targets.is_empty() || source.targets.len() > MAX_ITEMS {
        return Err("A presentation restore needs 1 to 1000 formulas".to_string());
    }
    let wps = request.client.as_deref() == Some("wps");
    for target in &source.targets {
        let valid = target.slide_id > 0
            && target.shape_id > 0
            && matches!(target.display_mode.as_str(), "inline" | "block")
            && target.font_size_pt.is_finite()
            && target.font_size_pt > 0.0
            && match target.kind.as_str() {
                "omml" => target.start.is_some_and(|start| start > 0) && target.length.is_some_and(|length| length > 0)
                    && target.zone.is_some() == wps,
                "image" => target.metadata.is_some() && target.zone.is_none(),
                _ => false,
            };
        if !valid {
            return Err("The presentation's restore source names an invalid formula".to_string());
        }
    }
    Ok(source)
}

/// The LaTeX a restore target becomes, with what the redraw window shows.
fn restore_targets(request: &MacOfflineSessionRequest, source: &RestoreSource) -> Result<Vec<MacOfflineFormulaRestoreTarget>, String> {
    let directory = session_directory(OfficeHost::Powerpoint, &request.session_id)?;
    let mut targets = Vec::with_capacity(source.targets.len());
    for (index, target) in source.targets.iter().enumerate() {
        let (latex, metadata, formula_id) = if target.kind == "image" {
            let metadata = decoded_latex(target.metadata.as_deref().unwrap_or_default())?;
            (metadata.latex.clone(), Some(metadata.clone()), Some(metadata.formula_id.clone()))
        } else {
            let equation = match &target.zone {
                Some(zone) => super::super::wps::presentation::zone_equation(zone)?,
                None => copied_equation(&directory, index)?,
            };
            (super::super::powerpoint_omml::to_latex::from_omml(&equation)?.latex, None, None)
        };
        targets.push(MacOfflineFormulaRestoreTarget {
            source_start: index * 2,
            source_end: index * 2 + 1,
            source_text: "\u{FFFC}".to_string(),
            display_mode: target.display_mode.clone(),
            font_size_pt: clamp_size(target.font_size_pt),
            source_kind: target.kind.clone(),
            formula_id,
            numbered: false,
            metadata,
            math_ml: None,
            latex: Some(latex),
        });
    }
    Ok(targets)
}

/// The redraw window's view of a presentation operation: the source text and
/// its boxes, or the formulas to restore.
pub(super) fn public_request(request: &MacOfflineSessionRequest) -> Result<MacOfflineDocumentImportPublicRequest, String> {
    let operation = request.operation.as_deref().unwrap_or_default();
    let document_import = request
        .document_import
        .as_ref()
        .ok_or_else(|| "The presentation operation is missing its options".to_string())?;
    let (source, segments, restore_targets) = match operation {
        "latexRedraw" => {
            let (text, segments) = joined(&redraw_source(request)?);
            (Some(text), Some(segments), None)
        }
        "formulaRestore" => (None, None, Some(restore_targets(request, &restore_source(request)?)?)),
        _ => return Err("A presentation supports LaTeX redraw and formula restore only".to_string()),
    };
    Ok(MacOfflineDocumentImportPublicRequest {
        protocol_version: request.protocol_version,
        session_id: request.session_id.clone(),
        host: request.host.clone(),
        source_document_id: request.source_document_id.clone().ok_or("The presentation operation has no presentation")?,
        bookmark_name: document_import.bookmark_name.clone(),
        default_font_size_pt: clamp_size(document_import.default_font_size_pt),
        inline_font_size_pt: None,
        display_font_size_pt: None,
        operation: operation.to_string(),
        redraw_scope: document_import.redraw_scope.clone(),
        output_kind: document_import.output_kind.clone(),
        source_kind: document_import.source_kind.clone(),
        source,
        restore_targets,
        client: request.client.clone(),
        segments,
        agent: None,
    })
}

/// A redraw's font sizes come with its source: each formula takes the size
/// of the text where its LaTeX starts.
pub(super) fn font_sizes(request: &MacOfflineSessionRequest, input: &MacOfflineLatexRedrawFontQueryInput) -> Result<Vec<f64>, String> {
    let source = redraw_source(request)?;
    let (text, segments) = joined(&source);
    let units: Vec<u16> = text.encode_utf16().collect();
    let default = request.document_import.as_ref().map(|options| options.default_font_size_pt).unwrap_or(18.0);
    input
        .ranges
        .iter()
        .map(|range| {
            let index = segments
                .iter()
                .position(|segment| range.source_start >= segment.start && range.source_end <= segment.end && range.source_start < range.source_end)
                .ok_or("A formula lies outside the presentation's text boxes")?;
            if String::from_utf16_lossy(&units[range.source_start..range.source_end]) != range.source_text {
                return Err("A formula's text differs from the presentation's".to_string());
            }
            let offset = range.source_start - segments[index].start;
            let size = source.segments[index]
                .sizes
                .iter()
                .find(|(start, length, _)| offset >= *start && offset < start + length)
                .map(|(_, _, size)| *size)
                .unwrap_or(default);
            Ok(clamp_size(size))
        })
        .collect()
}

/// Where an item goes in a shape's text.
struct Placement {
    slide_id: u32,
    shape_id: u32,
    start: usize,
    length: usize,
    source_text: String,
    whole: bool,
    /// The box's text as VisualTeX read it, from the start of what the add-in
    /// sent up to the formula's end: checked before the formula is written
    /// (the text after it may already hold the later formulas).
    check_start: usize,
    check_text: String,
}

fn placement(source: &RedrawSource, segments: &[PresentationSegment], units: &[u16], start: usize, end: usize, text: &str) -> Result<Placement, String> {
    let index = segments
        .iter()
        .position(|segment| start >= segment.start && end <= segment.end && start < end)
        .ok_or("A formula lies outside the presentation's text boxes")?;
    if String::from_utf16_lossy(&units[start..end]) != text {
        return Err("A formula's text differs from the presentation's".to_string());
    }
    let segment = &source.segments[index];
    Ok(Placement {
        slide_id: segment.slide_id,
        shape_id: segment.shape_id,
        start: segment.start + (start - segments[index].start),
        length: end - start,
        source_text: text.to_string(),
        whole: segment.whole && segment.text.trim() == text,
        check_start: segment.start,
        check_text: String::from_utf16_lossy(&units[segments[index].start..end]),
    })
}

/// PowerPoint pictures are sized from their 14 pt rendering (calculate_powerpoint_geometry).
fn picture_geometry(width_px: f64, height_px: f64, font_size_pt: f64) -> Result<(f64, f64, f64, f64), String> {
    if !(width_px.is_finite() && height_px.is_finite() && width_px > 0.0 && height_px > 0.0) {
        return Err("A picture formula has no size".to_string());
    }
    let size = font_size_pt.clamp(MIN_POWERPOINT_FONT_SIZE_PT, MAX_POWERPOINT_FONT_SIZE_PT);
    let reference_width = width_px * 0.75;
    let reference_height = height_px * 0.75;
    Ok((
        reference_width * size / POWERPOINT_REFERENCE_FONT_SIZE_PT,
        reference_height * size / POWERPOINT_REFERENCE_FONT_SIZE_PT,
        reference_width,
        reference_height,
    ))
}

/// Writes the operation's result into the presentation through the add-in that
/// started it, last item first, so the earlier items' positions stay valid.
pub(super) fn commit(state: &OfficeCompanionState, request: &MacOfflineSessionRequest, input: MacOfflineDocumentImportCommitInput) -> Result<(), String> {
    let session_id = request.session_id.clone();
    let operation = request.operation.as_deref().unwrap_or_default();
    let options = request.document_import.as_ref().ok_or("The presentation operation is missing its options")?;
    if options.output_kind.as_deref() != Some(input.output_kind.as_str()) {
        return Err("The output kind changed after the request".to_string());
    }
    if input.items.is_empty() || input.items.len() > MAX_ITEMS {
        return Err("A presentation operation writes 1 to 1000 formulas".to_string());
    }
    let wps = request.client.as_deref() == Some("wps");
    let directory = session_directory(OfficeHost::Powerpoint, &session_id)?;
    let mut items = Vec::with_capacity(input.items.len());
    let mut files: Vec<(String, PathBuf)> = Vec::new();
    // WPS: one presentation holding every equation, one text box each.
    let mut wps_boxes: Vec<String> = Vec::new();
    match operation {
        "latexRedraw" => {
            let source = redraw_source(request)?;
            let (text, segments) = joined(&source);
            let units: Vec<u16> = text.encode_utf16().collect();
            let mut previous_end = 0usize;
            for (index, item) in input.items.iter().enumerate() {
                let MacOfflineDocumentImportCommitItem::Formula {
                    formula_id, display_mode, font_size_pt, metadata, word_math_ml, svg_base64, png_base64, width, height,
                    source_start, source_end, source_text, ..
                } = item else {
                    return Err("A LaTeX redraw writes formulas only".to_string());
                };
                let (start, end) = (source_start.ok_or("A formula has no position")?, source_end.ok_or("A formula has no position")?);
                if start < previous_end {
                    return Err("The formulas overlap or are out of order".to_string());
                }
                previous_end = end;
                let place = placement(&source, &segments, &units, start, end, source_text.as_deref().unwrap_or_default())?;
                let display = display_mode == "block";
                let size = font_size_pt.clamp(MIN_POWERPOINT_FONT_SIZE_PT, MAX_POWERPOINT_FONT_SIZE_PT);
                let mut entry = json!({
                    "index": index,
                    "slideId": place.slide_id,
                    "shapeId": place.shape_id,
                    "start": place.start,
                    "length": place.length,
                    "sourceText": place.source_text,
                    "checkStart": place.check_start,
                    "checkText": place.check_text,
                    "displayMode": display_mode,
                    "fontSizePt": size,
                });
                if input.output_kind == "omml" {
                    let math_ml = word_math_ml.as_deref().ok_or("A redrawn equation has no MathML")?;
                    let omml = super::super::omml_write::transform_mathml_to_omml(math_ml)?;
                    entry["kind"] = json!("equation");
                    if wps {
                        wps_boxes.push(super::super::powerpoint_omml::native_paragraphs(&omml, size, display, false, false, None)?);
                        entry["box"] = json!(wps_boxes.len());
                    } else {
                        let package = build_stored_zip(&super::super::powerpoint_omml::clipboard_entries(&omml, size, display, false, false)?)?;
                        let path = directory.join(format!("redraw-equation-{index}.bin"));
                        atomic_write_runtime(&path, &package, 0o600)?;
                        entry["package"] = json!(path.to_string_lossy());
                    }
                } else {
                    if !place.whole {
                        return Err("A picture formula can only replace a text box that holds just that formula".to_string());
                    }
                    let (svg, png) = (svg_base64.as_deref().ok_or("A picture formula has no SVG")?, png_base64.as_deref());
                    let (width_pt, height_pt, reference_width, reference_height) =
                        picture_geometry(width.unwrap_or_default(), height.unwrap_or_default(), size)?;
                    let mut metadata = metadata.clone();
                    metadata.display_mode = "block".to_string();
                    metadata.numbered = false;
                    metadata.font_size_pt = Some(size);
                    metadata.reference_width_pt = Some(reference_width);
                    metadata.reference_height_pt = Some(reference_height);
                    metadata.reference_baseline_pt = None;
                    if metadata.formula_id != *formula_id {
                        return Err("A picture formula's metadata names another formula".to_string());
                    }
                    let svg_name = format!("{session_id}-redraw-{index}.svg");
                    let svg_path = directory.join(&svg_name);
                    atomic_write_runtime(&svg_path, &decode_svg(svg)?, 0o600)?;
                    files.push((svg_name.clone(), svg_path.clone()));
                    let mut picture = json!({ "image": svg_name, "imagePath": svg_path.to_string_lossy() });
                    if let Some(png) = png {
                        let png_name = format!("{session_id}-redraw-{index}.png");
                        let png_path = directory.join(&png_name);
                        atomic_write_runtime(&png_path, &decode_png(png)?, 0o600)?;
                        files.push((png_name.clone(), png_path.clone()));
                        picture["fallback"] = json!(png_name);
                        picture["fallbackPath"] = json!(png_path.to_string_lossy());
                    }
                    entry["kind"] = json!("picture");
                    entry["formulaId"] = json!(formula_id);
                    entry["metadata"] = json!(encode_metadata(&metadata)?);
                    entry["formulaReference"] = json!(super::super::wps::payload::formula_reference(formula_id, "block", false));
                    entry["widthPt"] = json!(width_pt);
                    entry["heightPt"] = json!(height_pt);
                    entry["referenceWidthPt"] = json!(reference_width);
                    entry["referenceHeightPt"] = json!(reference_height);
                    entry["files"] = picture;
                    let _ = state.formula_cache.put(formula_id, metadata);
                }
                items.push(entry);
            }
        }
        "formulaRestore" => {
            let source = restore_source(request)?;
            for (index, item) in input.items.iter().enumerate() {
                let MacOfflineDocumentImportCommitItem::Text { text, source_start, .. } = item else {
                    return Err("A formula restore writes LaTeX text only".to_string());
                };
                let target_index = source_start.ok_or("A LaTeX text has no formula")? / 2;
                let target = source.targets.get(target_index).ok_or("A LaTeX text names no formula")?;
                if text.is_empty() || text.len() > 1_000_000 {
                    return Err("A formula's LaTeX is empty or too long".to_string());
                }
                let mut entry = json!({
                    "index": index,
                    "kind": "text",
                    "replace": if target.kind == "image" { "picture" } else { "equation" },
                    "slideId": target.slide_id,
                    "shapeId": target.shape_id,
                    "displayMode": target.display_mode,
                    "fontSizePt": clamp_size(target.font_size_pt),
                    "text": text,
                });
                if target.kind == "image" {
                    // The picture must still be this formula.
                    entry["metadata"] = json!(target.metadata);
                } else {
                    entry["start"] = json!(target.start);
                    entry["length"] = json!(target.length);
                }
                items.push(entry);
            }
        }
        _ => return Err("A presentation supports LaTeX redraw and formula restore only".to_string()),
    }
    if wps {
        if !wps_boxes.is_empty() {
            let name = format!("{session_id}-equations.pptx");
            let path = directory.join(&name);
            atomic_write_runtime(&path, &build_stored_zip(&super::super::powerpoint_omml::wps_equation_slide_boxes(&wps_boxes))?, 0o600)?;
            files.push((name.clone(), path));
            let payload = json!({ "operation": operation, "outputKind": input.output_kind, "items": items, "equations": name });
            // The add-in pastes WPS's own copies of the equations.
            #[cfg(target_os = "macos")]
            return super::super::wps::pasteboard::kept(state.app.as_ref(), || {
                super::super::wps::bridge::deliver_document_commit(&session_id, payload, files, input.items.len())
            });
            #[cfg(not(target_os = "macos"))]
            return super::super::wps::bridge::deliver_document_commit(&session_id, payload, files, input.items.len());
        }
        let payload = json!({ "operation": operation, "outputKind": input.output_kind, "items": items });
        return super::super::wps::bridge::deliver_document_commit(&session_id, payload, files, input.items.len());
    }
    apply_in_powerpoint(state, request, &items)
}

/// Rows the PowerPoint add-in reads (VTReadPresentationItem): key=value, one
/// per line; texts are base64url (UTF-8) so that none spans lines.
fn item_rows(session_id: &str, fields: &[(String, String)]) -> Result<String, String> {
    let mut rows = format!("protocolVersion={OFFLINE_PROTOCOL_VERSION}\nsessionId={session_id}\n");
    for (key, value) in fields {
        if value.contains(['\r', '\n', '\0']) {
            return Err(format!("A formula's {key} cannot be passed to PowerPoint"));
        }
        rows.push_str(&format!("{key}={value}\n"));
    }
    Ok(rows)
}

/// An item as the PowerPoint add-in reads it (the WPS add-in gets the JSON).
fn powerpoint_item(session_id: &str, identity: &str, item: &Value) -> Result<String, String> {
    let mut fields = vec![("presentationIdentity".to_string(), identity.to_string())];
    let Some(object) = item.as_object() else {
        return Err("A formula item is not an object".to_string());
    };
    for (key, value) in object {
        match (key.as_str(), value) {
            ("package" | "box", _) | (_, Value::Null) => {}
            ("files", files) => {
                for (from, to) in [("imagePath", "imagePath"), ("fallbackPath", "fallbackImagePath")] {
                    if let Some(path) = files[from].as_str() {
                        fields.push((to.to_string(), path.to_string()));
                    }
                }
            }
            ("sourceText" | "checkText" | "text", Value::String(text)) => {
                fields.push((key.clone(), URL_SAFE_NO_PAD.encode(text.as_bytes())));
            }
            (_, Value::String(text)) => fields.push((key.clone(), text.clone())),
            (_, Value::Number(number)) => fields.push((key.clone(), number.to_string())),
            (_, Value::Bool(flag)) => fields.push((key.clone(), if *flag { "1" } else { "0" }.to_string())),
            _ => return Err(format!("A formula item's {key} cannot be passed to PowerPoint")),
        }
    }
    item_rows(session_id, &fields)
}

/// The PowerPoint add-in writes one item per macro call
/// (VisualTeX_ApplyPresentationItem), last first; an equation is pasted from
/// the clipboard, so its package goes on the clipboard first, and the user's
/// clipboard comes back at the end. A formula that cannot be written is
/// skipped; the message names it.
fn apply_in_powerpoint(state: &OfficeCompanionState, request: &MacOfflineSessionRequest, items: &[Value]) -> Result<(), String> {
    let session_id = request.session_id.as_str();
    let directory = session_directory(OfficeHost::Powerpoint, session_id)?;
    let identity = request.source_document_id.clone().ok_or("The presentation operation has no presentation")?;
    let app = state.app.as_ref();
    let total = items.len();
    let run = || -> Result<(), String> {
        let mut written = 0usize;
        let mut failures: Vec<String> = Vec::new();
        for (done, item) in items.iter().rev().enumerate() {
            let index = item["index"].as_u64().ok_or("An item has no index")? as usize;
            let result_path = directory.join(format!("presentation-item-{index}.txt"));
            let _ = fs::remove_file(&result_path);
            atomic_write_runtime(
                &directory.join(format!("presentation-item-{index}.in")),
                powerpoint_item(session_id, &identity, item)?.as_bytes(),
                0o600,
            )?;
            if let Some(package) = item["package"].as_str() {
                let bytes = fs::read(package).map_err(|error| format!("Unable to read an equation package: {error}"))?;
                place_native_math_clipboard(app, bytes)?;
            }
            // The script stays the same for every item (NSAppleScript compiles
            // it once): the add-in reads the Session from the dispatch pointer
            // and the item from presentation-next.txt.
            atomic_write_runtime(&directory.join(NEXT_ITEM_FILE), index.to_string().as_bytes(), 0o600)?;
            // A failed macro call means PowerPoint or its add-in cannot take
            // the others either.
            with_dispatch_pointer(OfficeHost::Powerpoint, session_id, || match app {
                Some(app) => run_office_vba_script_on_main_thread(app, APPLY_ITEM_SCRIPT, "PowerPoint formula operation"),
                None => run_office_vba_script(APPLY_ITEM_SCRIPT, "PowerPoint formula operation"),
            })
            .map_err(|error| summary(written, total, &failures, &error))?;
            let result = fs::read_to_string(&result_path).unwrap_or_default();
            if result.trim() == "ok" {
                written += 1;
            } else if let Some(detail) = result.strip_prefix("error\n") {
                failures.push(format!("formula {}: {}", index + 1, detail.trim()));
            } else {
                return Err(summary(
                    written,
                    total,
                    &failures,
                    "PowerPoint did not confirm the formula; the VisualTeX add-in in PowerPoint may need to be updated",
                ));
            }
            let _ = write_document_progress(session_id, done + 1, total, "inserting");
        }
        let _ = write_document_progress(session_id, total, total, if failures.is_empty() { "complete" } else { "error" });
        if failures.is_empty() {
            Ok(())
        } else {
            Err(summary(written, total, &failures, ""))
        }
    };
    #[cfg(target_os = "macos")]
    return super::super::wps::pasteboard::kept(app, run);
    #[cfg(not(target_os = "macos"))]
    run()
}

/// Which item VisualTeX_ApplyPresentationItem writes next.
const NEXT_ITEM_FILE: &str = "presentation-next.txt";
const APPLY_ITEM_SCRIPT: &str = "with timeout of 300 seconds
tell application \"Microsoft PowerPoint\" to run VB macro macro name \"VisualTeX_ApplyPresentationItem\" list of parameters {}
end timeout";

/// What was written and what was not.
fn summary(written: usize, total: usize, failures: &[String], stop: &str) -> String {
    let mut message = if written == 0 {
        "PowerPoint wrote none of the formulas.".to_string()
    } else {
        format!("PowerPoint wrote {written} of the {total} formulas.")
    };
    for failure in failures.iter().take(3) {
        message.push_str(&format!(" {failure}."));
    }
    if failures.len() > 3 {
        message.push_str(&format!(" {} more could not be written.", failures.len() - 3));
    }
    if !stop.is_empty() {
        message.push_str(&format!(" {stop}"));
    }
    message
}

/// PowerPoint's object model has no OMML: before the redraw window opens, the
/// add-in copies each equation of a restore (VisualTeX_CopyPresentationEquation)
/// and VisualTeX keeps its clipboard package, as an edit reads its equation
/// (capture_powerpoint_native_edit_source). The clipboard is put back.
pub(super) fn copy_powerpoint_equations(session_id: &str) -> Result<(), String> {
    let directory = session_directory(OfficeHost::Powerpoint, session_id)?;
    let copied = (|| {
        let request = read_request(session_id)?;
        let source = restore_source(&request)?;
        let identity = request.source_document_id.clone().ok_or("The restore has no presentation")?;
        let mut indices = Vec::new();
        for (index, target) in source.targets.iter().enumerate().filter(|(_, target)| target.kind == "omml") {
            let fields = [
                ("presentationIdentity", identity.clone()),
                ("slideId", target.slide_id.to_string()),
                ("shapeId", target.shape_id.to_string()),
                ("start", target.start.unwrap_or_default().to_string()),
                ("length", target.length.unwrap_or_default().to_string()),
            ]
            .map(|(key, value)| (key.to_string(), value));
            atomic_write_runtime(&directory.join(format!("presentation-copy-{index}.in")), item_rows(session_id, &fields)?.as_bytes(), 0o600)?;
            let _ = fs::remove_file(directory.join(format!("presentation-copy-{index}.txt")));
            let _ = fs::remove_file(directory.join(copied_equation_name(index)));
            indices.push(index);
        }
        if indices.is_empty() {
            return Ok(());
        }
        let quoted = serde_json::to_string(directory.to_string_lossy().as_ref()).map_err(|error| error.to_string())?;
        let list = indices.iter().map(usize::to_string).collect::<Vec<_>>().join(", ");
        let script = COPY_SCRIPT
            .replace("__SESSION__", session_id)
            .replace("__INDICES__", &list)
            .replace("__DIRECTORY__", &quoted);
        with_dispatch_pointer(OfficeHost::Powerpoint, session_id, || run_office_vba_script(&script, "PowerPoint equation copy"))
            .map_err(|error| {
                // The add-in's own reason, when it gave one.
                indices
                    .iter()
                    .find_map(|index| fs::read_to_string(directory.join(format!("presentation-copy-{index}.txt"))).ok())
                    .map(|reason| reason.trim_start_matches("error\n").trim().to_string())
                    .filter(|reason| !reason.is_empty())
                    .unwrap_or(error)
            })
    })();
    if let Err(error) = &copied {
        let _ = atomic_write_runtime(&directory.join(COPY_ERROR_FILE), error.as_bytes(), 0o600);
    }
    copied
}

/// A restore started by the PowerPoint add-in reads its equations from
/// PowerPoint first (copy_powerpoint_equations).
pub(super) fn needs_copied_equations(request: &MacOfflineSessionRequest) -> bool {
    request.host == "powerpoint" && request.client.is_none() && request.operation.as_deref() == Some("formulaRestore")
}

/// Copies the equations one after another with the user's clipboard saved
/// around them; each package is written beside the restore source.
const COPY_SCRIPT: &str = r#"use framework "AppKit"
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
set indexList to {__INDICES__}
set failureText to ""
try
    repeat with k from 1 to (count of indexList)
        set itemIndex to item k of indexList
        set countBefore to (pb's changeCount()) as integer
        with timeout of 120 seconds
            tell application "Microsoft PowerPoint" to run VB macro macro name "VisualTeX_CopyPresentationEquation" list of parameters {"__SESSION__", (itemIndex as text)}
        end timeout
        if ((pb's changeCount()) as integer) is countBefore then error "PowerPoint did not copy equation " & (itemIndex + 1)
        set nativeData to pb's dataForType:"com.microsoft.Art--Text-ClipFormat"
        if nativeData is missing value then error "PowerPoint did not copy equation " & (itemIndex + 1) & " as Office Math"
        set targetPath to __DIRECTORY__ & "/restore-equation-" & (itemIndex as text) & ".bin"
        if (nativeData's writeToFile:targetPath atomically:true) is false then error "Unable to keep equation " & (itemIndex + 1)
    end repeat
on error errorText
    set failureText to errorText
end try
if ((pb's changeCount()) as integer) is not initialCount then
    pb's clearContents()
    pb's writeObjects:savedItems
end if
if failureText is not "" then error failureText"#;

/// The operation's own files (cleanup_session_files removes the shared ones).
pub(super) fn remove_files(session_id: &str) {
    let Ok(directory) = session_directory(OfficeHost::Powerpoint, session_id) else {
        return;
    };
    let Ok(entries) = fs::read_dir(&directory) else {
        return;
    };
    for entry in entries.flatten() {
        let name = entry.file_name();
        let name = name.to_string_lossy();
        if name.starts_with("presentation-")
            || name.starts_with("redraw-equation-")
            || name.starts_with("restore-equation-")
            || name.starts_with(session_id)
            || name == DOCUMENT_IMPORT_PROGRESS_FILE
        {
            let _ = fs::remove_file(entry.path());
        }
    }
}

/// PowerPoint pastes an equation from its own clipboard format.
#[cfg(target_os = "macos")]
fn place_native_math_clipboard(app: Option<&AppHandle>, bytes: Vec<u8>) -> Result<(), String> {
    use objc2_app_kit::NSPasteboard;
    use objc2_foundation::{NSData, NSString};
    let write = move || {
        let pasteboard = NSPasteboard::generalPasteboard();
        pasteboard.clearContents();
        pasteboard.setData_forType(Some(&NSData::with_bytes(&bytes)), &NSString::from_str("com.microsoft.Art--Text-ClipFormat"))
    };
    let written = match app {
        Some(app) => {
            let (sender, receiver) = mpsc::sync_channel(1);
            app.run_on_main_thread(move || {
                let _ = sender.send(write());
            })
            .map_err(|error| format!("Unable to reach the clipboard: {error}"))?;
            receiver.recv_timeout(Duration::from_secs(5)).map_err(|_| "The clipboard did not answer".to_string())?
        }
        None => write(),
    };
    if written { Ok(()) } else { Err("macOS could not prepare the equation on the clipboard".to_string()) }
}

#[cfg(not(target_os = "macos"))]
fn place_native_math_clipboard(_app: Option<&AppHandle>, _bytes: Vec<u8>) -> Result<(), String> {
    Err("Native PowerPoint equations need macOS".to_string())
}

/// The request rules of a presentation operation (validate_request).
pub(super) fn validate(request: &MacOfflineSessionRequest) -> Result<(), String> {
    let operation = request.operation.as_deref().unwrap_or_default();
    let options = request.document_import.as_ref().ok_or("The presentation operation is missing its options")?;
    validate_bounded_text(&options.bookmark_name, 40, "documentImport bookmarkName")?;
    if request.mode != "create"
        || !options.bookmark_name.starts_with("VT_D_")
        || !options.default_font_size_pt.is_finite()
        || !(MIN_POWERPOINT_FONT_SIZE_PT..=MAX_POWERPOINT_FONT_SIZE_PT).contains(&options.default_font_size_pt)
        || !matches!(options.redraw_scope.as_deref(), Some("selection" | "document"))
        || options.inline_font_size_pt.is_some()
        || options.display_font_size_pt.is_some()
        || request.formula_id.is_some()
        || request.encoded_metadata.is_some()
        || request.pending_marker.is_some()
        || request.power_point.is_some()
        || request.numbered
        || request.native_equation
        || request.fork_copied_formula
        || request.source_document_id.is_none()
    {
        return Err("The presentation operation's options are invalid".to_string());
    }
    let (file, outputs): (&str, &[&str]) = match operation {
        "latexRedraw" if options.source_kind.is_none() => (REDRAW_SOURCE_FILE, &["omml", "image"]),
        // A restore of equations and pictures together names no source kind.
        "formulaRestore" if matches!(options.source_kind.as_deref(), None | Some("omml" | "image")) => (RESTORE_SOURCE_FILE, &["latex"]),
        _ => return Err("A presentation supports LaTeX redraw and formula restore only".to_string()),
    };
    if !options.output_kind.as_deref().is_some_and(|kind| outputs.contains(&kind)) {
        return Err("The presentation operation's output kind is invalid".to_string());
    }
    let path = session_directory(OfficeHost::Powerpoint, &request.session_id)?.join(file);
    let metadata = fs::symlink_metadata(&path).map_err(|error| format!("Unable to read {}: {error}", path.display()))?;
    if !metadata.is_file() || metadata.len() == 0 || metadata.len() > MAX_SOURCE_BYTES {
        return Err("The presentation's formula source has an invalid size".to_string());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn segment(shape_id: u32, start: usize, text: &str, whole: bool) -> TextSegment {
        TextSegment { slide_id: 256, shape_id, start, text: text.to_string(), whole, sizes: vec![(0, text.encode_utf16().count(), 24.0)] }
    }

    #[test]
    fn boxes_are_joined_with_their_bounds() {
        let source = RedrawSource { segments: vec![segment(5, 1, "Let $x$ be", false), segment(6, 3, "$$y^2$$", true)] };
        let (text, segments) = joined(&source);
        assert_eq!(text, "Let $x$ be\n\n$$y^2$$");
        assert_eq!(segments, vec![
            PresentationSegment { start: 0, end: 10, whole: false },
            PresentationSegment { start: 12, end: 19, whole: true },
        ]);
        let units: Vec<u16> = text.encode_utf16().collect();
        let first = placement(&source, &segments, &units, 4, 7, "$x$").unwrap();
        assert_eq!((first.shape_id, first.start, first.length, first.whole), (5, 5, 3, false));
        assert_eq!((first.check_start, first.check_text.as_str()), (1, "Let $x$"));
        let second = placement(&source, &segments, &units, 12, 19, "$$y^2$$").unwrap();
        assert_eq!((second.shape_id, second.start, second.length, second.whole), (6, 3, 7, true));
        assert_eq!((second.check_start, second.check_text.as_str()), (3, "$$y^2$$"));
        assert!(placement(&source, &segments, &units, 8, 14, "be\n\n$$").is_err(), "across two boxes");
        assert!(placement(&source, &segments, &units, 4, 7, "$y$").is_err(), "other text");
    }

    #[test]
    fn utf16_positions_count_surrogate_pairs() {
        let source = RedrawSource { segments: vec![segment(5, 1, "\u{1D465} $z$", false)] };
        let (text, segments) = joined(&source);
        let units: Vec<u16> = text.encode_utf16().collect();
        let place = placement(&source, &segments, &units, 3, 6, "$z$").unwrap();
        assert_eq!((place.start, place.length), (4, 3));
    }

    #[test]
    fn powerpoint_items_are_one_line_rows() {
        let item = json!({
            "index": 2, "kind": "picture", "slideId": 256, "shapeId": 7, "start": 1, "length": 9,
            "sourceText": "$$a\rb$$", "checkStart": 1, "checkText": "$$a\rb$$", "displayMode": "block",
            "fontSizePt": 24.5, "package": "/tmp/x.bin", "box": 3, "metadata": null,
            "files": { "image": "f.svg", "imagePath": "/r/f.svg", "fallback": "f.png", "fallbackPath": "/r/f.png" },
        });
        let rows = powerpoint_item("5028ce4b-71c7-4d06-a165-61151fd64183", "Deck.pptx", &item).unwrap();
        let lines: Vec<&str> = rows.lines().collect();
        assert_eq!(&lines[..3], ["protocolVersion=1", "sessionId=5028ce4b-71c7-4d06-a165-61151fd64183", "presentationIdentity=Deck.pptx"]);
        let encoded = URL_SAFE_NO_PAD.encode("$$a\rb$$".as_bytes());
        for row in ["index=2", "kind=picture", "slideId=256", "start=1", "length=9", "fontSizePt=24.5", "displayMode=block",
            "imagePath=/r/f.svg", "fallbackImagePath=/r/f.png"] {
            assert!(lines.contains(&row), "{row} in {rows}");
        }
        assert!(lines.contains(&format!("sourceText={encoded}").as_str()) && lines.contains(&format!("checkText={encoded}").as_str()));
        assert!(!rows.contains("package") && !rows.contains("box=") && !rows.contains("metadata"), "{rows}");
        assert!(powerpoint_item("5028ce4b-71c7-4d06-a165-61151fd64183", "Deck\n2", &item).is_err(), "a row cannot break");
    }

    #[test]
    fn a_partial_write_names_what_was_not_written() {
        let failures = vec!["formula 2: The text box changed".to_string()];
        assert_eq!(summary(3, 4, &failures, ""), "PowerPoint wrote 3 of the 4 formulas. formula 2: The text box changed.");
        assert_eq!(summary(0, 1, &[], "PowerPoint is not open"), "PowerPoint wrote none of the formulas. PowerPoint is not open");
    }

    #[test]
    fn pictures_take_the_powerpoint_reference_size() {
        let (width, height, reference_width, reference_height) = picture_geometry(100.0, 40.0, 28.0).unwrap();
        assert_eq!((reference_width, reference_height), (75.0, 30.0));
        assert_eq!((width, height), (150.0, 60.0));
        assert!(picture_geometry(0.0, 40.0, 28.0).is_err());
    }
}
