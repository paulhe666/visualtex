//! Restoring formulas from Word OMML: the restore manifest, the formula identity, and the OMath fragment as MathML.

use super::*;

pub(super) fn parse_formula_restore_manifest(
    request: &MacOfflineSessionRequest,
) -> Result<(String, Vec<MacOfflineFormulaRestoreTarget>), String> {
    let path = formula_restore_source_path(&request.session_id)?;
    let metadata = fs::metadata(&path)
        .map_err(|error| format!("Unable to inspect {}: {error}", path.display()))?;
    if metadata.file_type().is_symlink()
        || !metadata.is_file()
        || metadata.len() == 0
        || metadata.len() as usize > MAX_DOCUMENT_IMPORT_MANIFEST_BYTES
    {
        return Err("Word formula restore source has an invalid size".to_string());
    }
    let text = fs::read_to_string(&path)
        .map_err(|error| format!("Unable to read {}: {error}", path.display()))?;
    let mut values = std::collections::HashMap::<String, String>::new();
    for line in text.lines() {
        let (key, value) = line
            .split_once('=')
            .ok_or_else(|| "Word formula restore source contains an invalid line".to_string())?;
        if key.is_empty()
            || !key.bytes().all(|byte| byte.is_ascii_alphanumeric())
            || values.insert(key.to_string(), value.to_string()).is_some()
        {
            return Err("Word formula restore source contains an invalid key".to_string());
        }
    }
    let required = |key: &str| {
        values
            .get(key)
            .map(String::as_str)
            .ok_or_else(|| format!("Word formula restore source is missing {key}"))
    };
    if required("protocolVersion")? != OFFLINE_PROTOCOL_VERSION.to_string()
        || required("sessionId")? != request.session_id
    {
        return Err("Word formula restore source identity is invalid".to_string());
    }
    let item_count = required("itemCount")?
        .parse::<usize>()
        .map_err(|_| "Word formula restore item count is invalid".to_string())?;
    if item_count == 0 || item_count > 1000 {
        return Err("Word formula restore supports 1 to 1000 formulas".to_string());
    }
    let target_text = String::from_utf8(
        URL_SAFE_NO_PAD
            .decode(required("targetTextBase64")?)
            .map_err(|error| format!("Unable to decode Word restore target text: {error}"))?,
    )
    .map_err(|_| "Word formula restore target text is not UTF-8".to_string())?;
    if target_text.len() > MAX_LATEX_REDRAW_SOURCE_BYTES as usize {
        return Err("Word formula restore target exceeds 5 MB".to_string());
    }
    let source_kind = required("sourceKind")?;
    if !matches!(source_kind, "omml" | "image") {
        return Err("Word formula restore source kind is invalid".to_string());
    }

    let mut targets = Vec::with_capacity(item_count);
    let mut previous_end = 0usize;
    for index in 0..item_count {
        let prefix = format!("item{index}");
        let start = required(&format!("{prefix}sourceStart"))?
            .parse::<usize>()
            .map_err(|_| "Word formula restore sourceStart is invalid".to_string())?;
        let end = required(&format!("{prefix}sourceEnd"))?
            .parse::<usize>()
            .map_err(|_| "Word formula restore sourceEnd is invalid".to_string())?;
        let source_text = String::from_utf8(
            URL_SAFE_NO_PAD
                .decode(required(&format!("{prefix}sourceTextBase64"))?)
                .map_err(|error| format!("Unable to decode Word formula source text: {error}"))?,
        )
        .map_err(|_| "Word formula source text is not UTF-8".to_string())?;
        validate_formula_restore_range(
            &mut previous_end,
            start,
            end,
            &source_text,
        )?;
        let display_mode = required(&format!("{prefix}displayMode"))?.to_string();
        if !matches!(display_mode.as_str(), "inline" | "block") {
            return Err("Word formula restore display mode is invalid".to_string());
        }
        let font_size_pt = required(&format!("{prefix}fontSizePt"))?
            .parse::<f64>()
            .map_err(|_| "Word formula restore font size is invalid".to_string())?;
        if !font_size_pt.is_finite()
            || !(MIN_WORD_FONT_SIZE_PT..=MAX_WORD_FONT_SIZE_PT).contains(&font_size_pt)
        {
            return Err("Word formula restore font size is outside the supported range".to_string());
        }
        let formula_id = values
            .get(&format!("{prefix}formulaId"))
            .map(String::as_str)
            .filter(|value| !value.is_empty())
            .map(str::to_string);
        let numbered = match values.get(&format!("{prefix}numbered")).map(String::as_str) {
            Some("1") => true,
            Some("0") | None => false,
            _ => return Err("Word formula restore numbered flag is invalid".to_string()),
        };
        let encoded_identity_metadata = values
            .get(&format!("{prefix}metadata"))
            .map(String::as_str)
            .filter(|value| !value.is_empty());
        let payload = String::from_utf8(
            URL_SAFE_NO_PAD
                .decode(required(&format!("{prefix}payloadBase64"))?)
                .map_err(|error| format!("Unable to decode Word formula restore payload: {error}"))?,
        )
        .map_err(|_| "Word formula restore payload is not UTF-8".to_string())?;
        if payload.is_empty() || payload.len() > MAX_OMML_BYTES.max(MAX_METADATA_BYTES) {
            return Err("Word formula restore payload is invalid or excessive".to_string());
        }
        let (resolved_formula_id, resolved_numbered, resolved_metadata, math_ml, latex) =
            if source_kind == "omml" {
                let metadata = omml_restore_identity(
                    formula_id.as_deref(),
                    encoded_identity_metadata,
                    &display_mode,
                    numbered,
                )?;
                // The restore reads the equation with the same Word reader as
                // Windows; the editor no longer re-derives LaTeX from MathML.
                let math_ml = word_omml_to_mathml(&payload, display_mode == "block")?;
                let latex = crate::office::omml_batch::mathml_to_latex(&math_ml)?;
                (formula_id, numbered, metadata, Some(math_ml), Some(latex))
            } else {
                let metadata = decode_metadata(&payload)?;
                let latex = canonical_document_formula_latex(&metadata)?;
                if metadata.display_mode != display_mode {
                    return Err("Word image formula display metadata changed".to_string());
                }
                if let Some(manifest_formula_id) = formula_id.as_deref() {
                    if metadata.formula_id != manifest_formula_id {
                        return Err("Word image formula id changed during restore".to_string());
                    }
                }
                if metadata.numbered != numbered {
                    return Err("Word image formula numbering metadata changed".to_string());
                }
                if let Some(encoded) = encoded_identity_metadata {
                    let manifest_metadata = decode_metadata(encoded)?;
                    if manifest_metadata.formula_id != metadata.formula_id
                        || manifest_metadata.numbered != metadata.numbered
                    {
                        return Err("Word image formula identity metadata changed".to_string());
                    }
                }
                (
                    Some(metadata.formula_id.clone()),
                    metadata.numbered,
                    Some(metadata),
                    None,
                    Some(latex),
                )
            };
        targets.push(MacOfflineFormulaRestoreTarget {
            source_start: start,
            source_end: end,
            source_text,
            display_mode,
            font_size_pt,
            source_kind: source_kind.to_string(),
            formula_id: resolved_formula_id,
            numbered: resolved_numbered,
            metadata: resolved_metadata,
            math_ml,
            latex,
        });
    }
    Ok((target_text, targets))
}

/// The VisualTeX identity of one OMML restore item. An equation without a
/// VisualTeX id can still carry Word's own number (a batch import or a number
/// made in Word); it is restored as a new numbered formula. Metadata without
/// an id is a broken identity.
pub(super) fn omml_restore_identity(
    formula_id: Option<&str>,
    encoded_metadata: Option<&str>,
    display_mode: &str,
    numbered: bool,
) -> Result<Option<VisualTeXFormulaMetadata>, String> {
    let Some(formula_id) = formula_id else {
        if encoded_metadata.is_some() {
            return Err("Unmanaged Word OMML contains partial VisualTeX identity".to_string());
        }
        return Ok(None);
    };
    validate_uuid(formula_id, "Word managed OMML formula id")?;
    let encoded = encoded_metadata
        .ok_or_else(|| "Managed Word OMML formula is missing VisualTeX metadata".to_string())?;
    let metadata = decode_metadata(encoded)?;
    if metadata.formula_id != formula_id
        || metadata.display_mode != display_mode
        || metadata.numbered != numbered
    {
        return Err("Managed Word OMML identity does not match its VisualTeX metadata".to_string());
    }
    Ok(Some(metadata))
}

pub(super) fn extract_omath_fragment(word_open_xml: &str) -> Result<String, String> {
    if word_open_xml.is_empty()
        || word_open_xml.len() > MAX_OMML_BYTES
        || word_open_xml.contains("<!DOCTYPE")
        || word_open_xml.contains("<!ENTITY")
    {
        return Err("Word OMML XML is invalid or excessive".to_string());
    }
    let starts = ["<m:oMath>", "<m:oMath "];
    let start = starts
        .iter()
        .filter_map(|marker| word_open_xml.find(marker))
        .min()
        .ok_or_else(|| "Word range XML does not contain one native OMath".to_string())?;
    let relative_end = word_open_xml[start..]
        .find("</m:oMath>")
        .ok_or_else(|| "Word native OMath XML is incomplete".to_string())?;
    let end = start + relative_end + "</m:oMath>".len();
    if word_open_xml[end..].contains("<m:oMath>")
        || word_open_xml[end..].contains("<m:oMath ")
    {
        return Err("Word formula restore item contains multiple OMath objects".to_string());
    }
    let mut fragment = word_open_xml[start..end].to_string();
    let tag_end = fragment
        .find('>')
        .ok_or_else(|| "Word native OMath start tag is incomplete".to_string())?;
    let mut declarations = String::new();
    let mut cursor = 0usize;
    while let Some(offset) = word_open_xml[cursor..].find("xmlns") {
        let attribute_start = cursor + offset;
        let Some(equal_offset) = word_open_xml[attribute_start..].find('=') else {
            break;
        };
        let equal = attribute_start + equal_offset;
        let name = word_open_xml[attribute_start..equal].trim();
        if !name.bytes().all(|byte| {
            byte.is_ascii_alphanumeric() || matches!(byte, b':' | b'-' | b'_')
        }) {
            cursor = equal + 1;
            continue;
        }
        let Some(quote) = word_open_xml.as_bytes().get(equal + 1).copied() else {
            break;
        };
        if quote != b'\'' && quote != b'"' {
            cursor = equal + 1;
            continue;
        }
        let value_start = equal + 2;
        let Some(value_end_offset) = word_open_xml[value_start..].find(quote as char) else {
            break;
        };
        let value_end = value_start + value_end_offset;
        if !fragment[..tag_end].contains(&format!("{name}="))
            && !declarations.contains(&format!(" {name}="))
        {
            declarations.push(' ');
            declarations.push_str(name);
            declarations.push('=');
            declarations.push(quote as char);
            declarations.push_str(&word_open_xml[value_start..value_end]);
            declarations.push(quote as char);
        }
        cursor = value_end + 1;
    }
    if !fragment[..tag_end].contains("xmlns:m=") && !declarations.contains(" xmlns:m=") {
        declarations.push_str(
            " xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"",
        );
    }
    if !fragment[..tag_end].contains("xmlns:w=") && !declarations.contains(" xmlns:w=") {
        declarations.push_str(
            " xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"",
        );
    }
    fragment.insert_str(tag_end, &declarations);
    Ok(fragment)
}

pub(super) fn strip_visualtex_numbered_equation_array(fragment: &str) -> Option<String> {
    // Recognize Word's native one-row #() number host, including merged math
    // runs and nested arrays in the formula body. Preserve the original XML
    // rather than rebuilding the mathematics through a text conversion.
    #[derive(Debug)]
    struct Element {
        name: Vec<u8>,
        start: usize,
        content: usize,
        close: usize,
        end: usize,
        parent: Option<usize>,
    }
    use quick_xml::{events::Event, Reader};
    let mut reader = Reader::from_str(fragment);
    let mut elements: Vec<Element> = Vec::new();
    let mut stack: Vec<usize> = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        match reader.read_event().ok()? {
            Event::Start(tag) => {
                let index = elements.len();
                elements.push(Element {
                    name: tag.name().as_ref().to_vec(),
                    start,
                    content: reader.buffer_position() as usize,
                    close: 0,
                    end: 0,
                    parent: stack.last().copied(),
                });
                stack.push(index);
            }
            Event::End(_) => {
                let index = stack.pop()?;
                elements[index].close = start;
                elements[index].end = reader.buffer_position() as usize;
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() {
        return None;
    }
    let root = elements.iter().position(|e| e.name == b"m:oMath")?;
    let arrays: Vec<_> = elements
        .iter()
        .enumerate()
        .filter(|(_, e)| e.parent == Some(root) && e.name == b"m:eqArr")
        .map(|(i, _)| i)
        .collect();
    if arrays.len() != 1 {
        return None;
    }
    let array = arrays[0];
    let entries: Vec<_> = elements
        .iter()
        .enumerate()
        .filter(|(_, e)| e.parent == Some(array) && e.name == b"m:e")
        .map(|(i, _)| i)
        .collect();
    if entries.len() != 1 {
        return None;
    }
    let entry = entries[0];
    let children: Vec<_> = elements
        .iter()
        .enumerate()
        .filter(|(_, e)| e.parent == Some(entry))
        .map(|(i, _)| i)
        .collect();
    if children.len() < 2 {
        return None;
    }
    let delimiter = &elements[*children.last()?];
    let marker_index = children[children.len() - 2];
    let marker = &elements[marker_index];
    if delimiter.name != b"m:d" || marker.name != b"m:r" {
        return None;
    }
    let text = elements
        .iter()
        .rev()
        .find(|e| e.parent == Some(marker_index) && e.name == b"m:t")?;
    let marker_text = &fragment[text.content..text.close];
    if !marker_text.ends_with('#') {
        return None;
    }
    let tail = &fragment[delimiter.start..delimiter.end];
    // Word for Mac stores field instructions as m:t inside math runs; Windows
    // may use w:instrText. Joining text handles a localized identifier split
    // into a separate run. Require an actual field, not a typed "SEQ" number.
    let mut tail_reader = Reader::from_str(tail);
    let mut field = false;
    let mut instruction = String::new();
    loop {
        match tail_reader.read_event().ok()? {
            Event::Start(e) | Event::Empty(e) if e.name().as_ref() == b"w:fldChar" => field = true,
            Event::Text(e) => instruction.push_str(&e.decode().ok()?),
            Event::Eof => break,
            _ => {}
        }
    }
    let instruction = instruction.to_ascii_uppercase();
    if !field || !(instruction.contains("SEQ ") || instruction.contains("REF VT_N_")) {
        return None;
    }
    if instruction.contains("REF ") && !instruction.contains("REF VT_N_") {
        return None;
    }
    let mut body = fragment[elements[entry].content..marker.start].to_string();
    if marker_text.len() > 1 {
        body.push_str(&fragment[marker.start..text.close - 1]);
        body.push_str(&fragment[text.close..marker.end]);
    }
    if body.trim().is_empty() {
        return None;
    }
    Some(format!(
        "{}{}{}",
        &fragment[..elements[array].start],
        body,
        &fragment[elements[array].end..]
    ))
}

/// One Word equation as MathML. VisualTeX's own OMML reader (shared with the
/// Windows add-in) replaces Word's omml2mathml.xsl, whose output changes with
/// Office updates.
pub(crate) fn word_omml_to_mathml(word_open_xml: &str, display: bool) -> Result<String, String> {
    let fragment = extract_omath_fragment(word_open_xml)?;
    let fragment = strip_visualtex_numbered_equation_array(&fragment).unwrap_or(fragment);
    let math_ml = crate::office::omml_latex::omml_to_mathml(&fragment, display)?;
    if math_ml.trim().is_empty() || math_ml.len() > MAX_OMML_BYTES {
        return Err("The Word equation produced empty or excessive MathML".to_string());
    }
    Ok(math_ml)
}
