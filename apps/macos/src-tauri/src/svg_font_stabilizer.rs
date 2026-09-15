use crate::system_math_glyphs::{extract_stable_macos_svg_glyph, StableSvgGlyphOutline};
use std::collections::HashMap;

const LETTER_MARKER: &str = "data-visualtex-output-letter-font";
const TEXT_MARKER: &str = "data-visualtex-output-text-font";
const MAX_STABILIZED_SVG_BYTES: usize = 16 * 1024 * 1024;

fn xml_unescape(value: &str) -> String {
    let mut output = value
        .replace("&quot;", "\"")
        .replace("&apos;", "'")
        .replace("&lt;", "<")
        .replace("&gt;", ">");
    output = output.replace("&amp;", "&");
    output
}

fn xml_escape_attribute(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('"', "&quot;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

fn find_tag_end(source: &str, start: usize) -> Option<usize> {
    let bytes = source.as_bytes();
    let mut quote = 0_u8;
    let mut index = start;
    while index < bytes.len() {
        let byte = bytes[index];
        if quote == 0 {
            if byte == b'\'' || byte == b'"' {
                quote = byte;
            } else if byte == b'>' {
                return Some(index);
            }
        } else if byte == quote {
            quote = 0;
        }
        index += 1;
    }
    None
}

fn parse_attributes(opening: &str) -> Result<HashMap<String, String>, String> {
    let bytes = opening.as_bytes();
    let mut attributes = HashMap::new();
    let mut index = opening
        .find(char::is_whitespace)
        .unwrap_or(opening.len());
    while index < bytes.len() {
        while index < bytes.len() && bytes[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= bytes.len() || bytes[index] == b'>' || bytes[index] == b'/' {
            break;
        }
        let name_start = index;
        while index < bytes.len()
            && !bytes[index].is_ascii_whitespace()
            && bytes[index] != b'='
            && bytes[index] != b'>'
        {
            index += 1;
        }
        if index == name_start {
            return Err("VisualTeX stable SVG text attribute is malformed.".to_string());
        }
        let name = opening[name_start..index].to_ascii_lowercase();
        while index < bytes.len() && bytes[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= bytes.len() || bytes[index] != b'=' {
            return Err("VisualTeX stable SVG text attribute has no value.".to_string());
        }
        index += 1;
        while index < bytes.len() && bytes[index].is_ascii_whitespace() {
            index += 1;
        }
        if index >= bytes.len() || (bytes[index] != b'\'' && bytes[index] != b'"') {
            return Err("VisualTeX stable SVG text attribute is not quoted.".to_string());
        }
        let quote = bytes[index];
        index += 1;
        let value_start = index;
        while index < bytes.len() && bytes[index] != quote {
            index += 1;
        }
        if index >= bytes.len() {
            return Err("VisualTeX stable SVG text attribute is unterminated.".to_string());
        }
        attributes.insert(name, xml_unescape(&opening[value_start..index]));
        index += 1;
    }
    Ok(attributes)
}

fn split_font_families(value: &str) -> Vec<String> {
    let mut result = Vec::new();
    let mut current = String::new();
    let mut quote = None;
    for character in value.chars() {
        match quote {
            Some(active) if character == active => quote = None,
            Some(_) => current.push(character),
            None if character == '\'' || character == '"' => quote = Some(character),
            None if character == ',' => {
                let candidate = current.trim();
                if !candidate.is_empty() {
                    result.push(candidate.to_string());
                }
                current.clear();
            }
            None => current.push(character),
        }
    }
    let candidate = current.trim();
    if !candidate.is_empty() {
        result.push(candidate.to_string());
    }
    result.retain(|family| {
        !matches!(
            family.to_ascii_lowercase().as_str(),
            "serif"
                | "sans-serif"
                | "monospace"
                | "system-ui"
                | "-apple-system"
                | "blinkmacsystemfont"
        )
    });
    result
}

fn parse_svg_number(value: Option<&String>, fallback: f64) -> f64 {
    let Some(value) = value else {
        return fallback;
    };
    let trimmed = value.trim();
    let numeric = trimmed
        .strip_suffix("px")
        .or_else(|| trimmed.strip_suffix("pt"))
        .unwrap_or(trimmed)
        .trim()
        .parse::<f64>();
    numeric
        .ok()
        .filter(|number| number.is_finite())
        .unwrap_or(fallback)
}

fn is_bold_font_weight(value: Option<&String>) -> bool {
    let Some(value) = value else {
        return false;
    };
    let normalized = value.trim().to_ascii_lowercase();
    if normalized == "bold" || normalized == "bolder" {
        return true;
    }
    normalized
        .parse::<u16>()
        .ok()
        .is_some_and(|weight| weight >= 600)
}

fn strip_text_orientation_flip(value: Option<&String>) -> String {
    let Some(value) = value else {
        return String::new();
    };
    value
        .replace("scale(1,-1)", "")
        .replace("scale(1, -1)", "")
        .split_whitespace()
        .collect::<Vec<_>>()
        .join(" ")
}

fn format_number(value: f64) -> String {
    let normalized = if value.abs() < 0.000_000_5 { 0.0 } else { value };
    let mut output = format!("{normalized:.6}");
    while output.contains('.') && output.ends_with('0') {
        output.pop();
    }
    if output.ends_with('.') {
        output.pop();
    }
    output
}

fn outline_cache_key(
    families: &[String],
    character: char,
    italic: bool,
    bold: bool,
) -> String {
    format!(
        "{}\u{0}{}\u{0}{}\u{0}{}",
        families.join("\u{1f}"), character, italic, bold
    )
}

fn replacement_for_text(
    opening: &str,
    body: &str,
    cache: &mut HashMap<String, StableSvgGlyphOutline>,
) -> Result<String, String> {
    let attributes = parse_attributes(opening)?;
    let family_value = attributes
        .get("font-family")
        .ok_or_else(|| "VisualTeX stable SVG text is missing font-family.".to_string())?;
    let families = split_font_families(family_value);
    if families.is_empty() {
        return Err("VisualTeX stable SVG text has no concrete font family.".to_string());
    }
    if body.contains('<') || body.contains('>') {
        return Err("VisualTeX stable SVG text contains nested markup.".to_string());
    }
    let text = xml_unescape(body);
    if text.is_empty() {
        return Ok(String::new());
    }
    let italic = attributes
        .get("font-style")
        .is_some_and(|value| value.trim().eq_ignore_ascii_case("italic"));
    let bold = is_bold_font_weight(attributes.get("font-weight"));
    let font_size = parse_svg_number(attributes.get("font-size"), 1000.0);
    if font_size <= 0.0 || font_size > 100_000.0 {
        return Err("VisualTeX stable SVG text has an invalid font-size.".to_string());
    }
    let scale = font_size / 1000.0;
    let x = parse_svg_number(attributes.get("x"), 0.0);
    let y = parse_svg_number(attributes.get("y"), 0.0);
    let outer_transform = strip_text_orientation_flip(attributes.get("transform"));

    let mut cursor = 0.0_f64;
    let mut paths = String::new();
    let mut resolved_families = Vec::new();
    for character in text.chars() {
        if character.is_control() {
            return Err("VisualTeX stable SVG text contains a control character.".to_string());
        }
        let key = outline_cache_key(&families, character, italic, bold);
        let outline = if let Some(outline) = cache.get(&key) {
            outline.clone()
        } else {
            let outline = extract_stable_macos_svg_glyph(
                &families,
                &character.to_string(),
                italic,
                bold,
            )?;
            cache.insert(key, outline.clone());
            outline
        };
        if !resolved_families
            .iter()
            .any(|family: &String| family == &outline.resolved_family)
        {
            resolved_families.push(outline.resolved_family.clone());
        }
        if !outline.path.is_empty() {
            if cursor.abs() < 0.000_000_5 {
                paths.push_str(&format!("<path d=\"{}\"></path>", outline.path));
            } else {
                paths.push_str(&format!(
                    "<path d=\"{}\" transform=\"translate({},0)\"></path>",
                    outline.path,
                    format_number(cursor)
                ));
            }
        }
        cursor += outline.advance_units;
    }

    let mut local_transform = String::new();
    if x.abs() >= 0.000_000_5 || y.abs() >= 0.000_000_5 {
        local_transform.push_str(&format!(
            "translate({},{})",
            format_number(x),
            format_number(y)
        ));
    }
    if (scale - 1.0).abs() >= 0.000_000_5 {
        if !local_transform.is_empty() {
            local_transform.push(' ');
        }
        local_transform.push_str(&format!("scale({})", format_number(scale)));
    }

    let resolved = xml_escape_attribute(&resolved_families.join(", "));
    let local = if local_transform.is_empty() {
        paths
    } else {
        format!("<g transform=\"{}\">{}</g>", local_transform, paths)
    };
    if outer_transform.is_empty() {
        Ok(format!(
            "<g data-visualtex-stable-font-outline=\"true\" data-visualtex-resolved-font=\"{}\">{}</g>",
            resolved, local
        ))
    } else {
        Ok(format!(
            "<g data-visualtex-stable-font-outline=\"true\" data-visualtex-resolved-font=\"{}\" transform=\"{}\">{}</g>",
            resolved,
            xml_escape_attribute(&outer_transform),
            local
        ))
    }
}

pub(crate) fn stabilize_word_svg_font_outlines(svg: &[u8]) -> Result<Vec<u8>, String> {
    let source = std::str::from_utf8(svg)
        .map_err(|_| "VisualTeX stable Word SVG is not UTF-8.".to_string())?;
    if !source.contains(LETTER_MARKER) && !source.contains(TEXT_MARKER) {
        return Ok(svg.to_vec());
    }

    let mut output = String::with_capacity(source.len().saturating_mul(2));
    let mut cursor = 0_usize;
    let mut cache = HashMap::new();
    while let Some(relative_start) = source[cursor..].find("<text") {
        let start = cursor + relative_start;
        output.push_str(&source[cursor..start]);
        let end = find_tag_end(source, start)
            .ok_or_else(|| "VisualTeX stable SVG contains an unterminated text tag.".to_string())?;
        let opening = &source[start..=end];
        let closing_relative = source[end + 1..]
            .find("</text>")
            .ok_or_else(|| "VisualTeX stable SVG contains an unterminated text element.".to_string())?;
        let closing_start = end + 1 + closing_relative;
        let closing_end = closing_start + "</text>".len();
        let body = &source[end + 1..closing_start];
        if opening.contains(LETTER_MARKER) || opening.contains(TEXT_MARKER) {
            output.push_str(&replacement_for_text(opening, body, &mut cache)?);
        } else {
            output.push_str(&source[start..closing_end]);
        }
        cursor = closing_end;
        if output.len() > MAX_STABILIZED_SVG_BYTES {
            return Err("VisualTeX stable Word SVG became too large.".to_string());
        }
    }
    output.push_str(&source[cursor..]);
    if output.len() > MAX_STABILIZED_SVG_BYTES {
        return Err("VisualTeX stable Word SVG became too large.".to_string());
    }
    for marker in [LETTER_MARKER, TEXT_MARKER] {
        if output
            .match_indices("<text")
            .any(|(index, _)| output[index..].find('>').is_some_and(|end| output[index..index + end].contains(marker)))
        {
            return Err("VisualTeX stable Word SVG still contains a host-font text glyph.".to_string());
        }
    }
    Ok(output.into_bytes())
}

#[cfg(all(test, target_os = "macos"))]
mod tests {
    use super::*;

    #[test]
    fn vectorizes_visualtex_letter_text_without_a_host_font_dependency() {
        let svg = br#"<svg xmlns="http://www.w3.org/2000/svg"><g transform="scale(1,-1)"><text data-c="78" data-visualtex-output-letter-font="helvetica" transform="translate(120,0) scale(1,-1)" font-size="1000px" font-family="Helvetica, Arial, sans-serif" font-style="italic">x</text></g></svg>"#;
        let stabilized = stabilize_word_svg_font_outlines(svg).expect("Helvetica should vectorize");
        let text = String::from_utf8(stabilized).expect("stabilized SVG should remain UTF-8");
        assert!(!text.contains("<text"));
        assert!(!text.contains("font-family"));
        assert!(text.contains("<path d=\"M"));
        assert!(text.contains("transform=\"translate(120,0)\""));
        assert!(text.contains("data-visualtex-stable-font-outline=\"true\""));
    }

    #[test]
    fn vectorizes_cjk_formula_text_into_the_requested_system_font_outline() {
        let svg = r#"<svg xmlns="http://www.w3.org/2000/svg"><g transform="scale(1,-1)"><text data-variant="normal" transform="scale(1,-1)" font-size="1000px" font-family="&quot;PingFang SC&quot;, sans-serif" data-visualtex-output-text-font="pingfang">速</text></g></svg>"#;
        let stabilized = stabilize_word_svg_font_outlines(svg.as_bytes()).expect("PingFang CJK should vectorize");
        let text = String::from_utf8(stabilized).expect("stabilized SVG should remain UTF-8");
        assert!(!text.contains("<text"));
        assert!(!text.contains("font-family"));
        assert!(text.contains("<path d=\"M"));
        assert!(text.contains("data-visualtex-resolved-font=\"PingFang SC\""));
    }

    #[test]
    fn leaves_mathjax_path_only_svg_byte_identical() {
        let svg = br#"<svg xmlns="http://www.w3.org/2000/svg"><path d="M0 0L10 10"></path></svg>"#;
        assert_eq!(stabilize_word_svg_font_outlines(svg).unwrap(), svg);
    }
}
