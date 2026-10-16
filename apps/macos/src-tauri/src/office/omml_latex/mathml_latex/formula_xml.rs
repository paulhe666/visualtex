//! FormulaXml.Parse(xml, kind, LoadOptions.PreserveWhitespace): XmlReader with
//! DTD prohibited, 4,000,000 characters, depth < 256, at most 250,000 elements.
//!
//! dom.rs does the tree building. The XmlReader behaviors it does not cover
//! are applied here: end-of-line normalization, attribute-value whitespace
//! normalization, rejecting content at the root level, duplicate attributes
//! and undeclared attribute prefixes.

use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::dom::{Doc, ParseOptions};
use super::netstr::len16;

pub const MAXIMUM_CHARACTERS: usize = 4_000_000;
pub const MAXIMUM_DEPTH: usize = 256;
pub const MAXIMUM_ELEMENTS: usize = 250_000;

pub fn parse(xml: &str, kind: &str) -> Result<Doc, String> {
    if len16(xml) > MAXIMUM_CHARACTERS {
        return Err(format!("{kind} exceeds the formula limit of {MAXIMUM_CHARACTERS} characters."));
    }
    let normalized = normalize_input(xml);
    check_reader_rules(&normalized)?;
    let doc = Doc::parse(&normalized, ParseOptions { keep_whitespace: true, max_bytes: usize::MAX })?;
    check_tree(&doc, kind)?;
    Ok(doc)
}

/// XML 1.0 section 2.11 (CR LF / CR -> LF) and 3.3.3 (literal white space in
/// a CDATA-typed attribute value becomes a space; no DTD means all CDATA).
fn normalize_input(xml: &str) -> String {
    let text = if xml.contains('\r') { xml.replace("\r\n", "\n").replace('\r', "\n") } else { xml.to_string() };
    let mut bytes = text.into_bytes();
    let length = bytes.len();
    let find = |bytes: &[u8], from: usize, needle: &[u8]| -> usize {
        bytes[from.min(bytes.len())..]
            .windows(needle.len())
            .position(|window| window == needle)
            .map(|offset| from + offset + needle.len())
            .unwrap_or(bytes.len())
    };
    let mut index = 0;
    while index < length {
        if bytes[index] != b'<' {
            index += 1;
            continue;
        }
        let rest = &bytes[index..];
        if rest.starts_with(b"<!--") {
            index = find(&bytes, index + 4, b"-->");
        } else if rest.starts_with(b"<![CDATA[") {
            index = find(&bytes, index + 9, b"]]>");
        } else if rest.starts_with(b"<?") {
            index = find(&bytes, index + 2, b"?>");
        } else if rest.starts_with(b"<!") {
            index = find(&bytes, index + 2, b">");
        } else {
            index += 1;
            let mut quote: Option<u8> = None;
            while index < length {
                let byte = bytes[index];
                match quote {
                    Some(open) => {
                        if byte == open {
                            quote = None;
                        } else if byte == b'\t' || byte == b'\n' {
                            bytes[index] = b' ';
                        }
                    }
                    None => {
                        if byte == b'"' || byte == b'\'' {
                            quote = Some(byte);
                        } else if byte == b'>' {
                            index += 1;
                            break;
                        }
                    }
                }
                index += 1;
            }
        }
    }
    String::from_utf8(bytes).expect("ASCII-only replacements keep UTF-8 valid")
}

fn is_xml_space(text: &str) -> bool {
    text.chars().all(|c| matches!(c, ' ' | '\t' | '\r' | '\n'))
}

fn check_reader_rules(xml: &str) -> Result<(), String> {
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut depth = 0usize;
    loop {
        let event = reader.read_event().map_err(|error| format!("Unable to parse XML: {error}"))?;
        let opens = matches!(event, Event::Start(_));
        let closes = matches!(event, Event::End(_));
        match event {
            Event::Start(start) | Event::Empty(start) => {
                for attribute in start.attributes().with_checks(false) {
                    let attribute = attribute.map_err(|error| error.to_string())?;
                    let raw = attribute.key.as_ref();
                    if raw == b"xmlns" || raw.starts_with(b"xmlns:") || raw.starts_with(b"xml:") {
                        continue;
                    }
                    if attribute.key.prefix().is_some() {
                        if let (ResolveResult::Unknown(prefix), _) = reader.resolver().resolve_attribute(attribute.key) {
                            return Err(format!(
                                "'{}' is an undeclared prefix.",
                                String::from_utf8_lossy(&prefix)
                            ));
                        }
                    }
                }
            }
            Event::End(_) => {}
            Event::Text(text) if depth == 0 => {
                let decoded = text.decode().map_err(|error| error.to_string())?;
                if !is_xml_space(&decoded) {
                    return Err("Data at the root level is invalid.".to_string());
                }
            }
            Event::CData(_) | Event::GeneralRef(_) if depth == 0 => {
                return Err("Data at the root level is invalid.".to_string());
            }
            Event::DocType(_) => return Err("DTD is prohibited in this XML document.".to_string()),
            Event::Eof => break,
            _ => {}
        }
        if opens {
            depth += 1;
        }
        if closes {
            depth = depth.saturating_sub(1);
        }
    }
    Ok(())
}

fn check_tree(doc: &Doc, kind: &str) -> Result<(), String> {
    let root = doc.root.ok_or_else(|| format!("{kind} has no root element."))?;
    let mut elements = 0usize;
    let mut stack = vec![(root, 0usize)];
    while let Some((node, depth)) = stack.pop() {
        if depth >= MAXIMUM_DEPTH {
            return Err(format!("{kind} exceeds the formula nesting limit of {MAXIMUM_DEPTH} levels."));
        }
        elements += 1;
        if elements > MAXIMUM_ELEMENTS {
            return Err(format!("{kind} exceeds the formula limit of {MAXIMUM_ELEMENTS} elements."));
        }
        let attrs = doc.attrs(node);
        for (index, (name, _)) in attrs.iter().enumerate() {
            if attrs[..index].iter().any(|(other, _)| other == name) {
                return Err(format!("'{}' is a duplicate attribute name.", name.local));
            }
        }
        for child in doc.elements(node) {
            stack.push((child, depth + 1));
        }
    }
    Ok(())
}
