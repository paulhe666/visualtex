//! XmlReader + XDocument.Load / XDocument.Parse / XElement.Parse equivalents on
//! top of the arena DOM, as System.Xml.Linq builds its trees:
//!
//! * namespace declarations stay attributes (`xmlns` in no namespace,
//!   `xmlns:p` as `{http://www.w3.org/2000/xmlns/}p`), in source order, so the
//!   XLinq serializer (`xml_writer`) sees the same declarations;
//! * an element written `<a></a>` (or whose only content was skipped) gets
//!   empty-string content (XContainer.ReadContentFrom sets content = "" at the
//!   end tag), which XLinq writes back as `<a></a>`;
//! * text is decided per text node while reading (XmlTextReaderImpl.ParseText):
//!   end-of-line normalization, whitespace-only nodes dropped only when
//!   IgnoreWhitespace is on and xml:space="preserve" is not in scope;
//! * attribute values get attribute-value normalization (literal CR/LF/TAB
//!   become spaces, character references keep their characters).
//!
//! Comments and processing instructions are not represented in the DOM; the
//! formula payloads this port reads never contain them (see the crate docs).
//! CDATA sections become ordinary text.
//!
//! Based on the reference port's omml_mathml/xml.rs.

use crate::office::omml_write::proj::dom::{Data, Doc, Name, NodeId, XMLNS_NS, XML_NS};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

/// FormulaXml limits and the XmlReaderSettings.MaxCharactersInDocument value.
pub const MAXIMUM_CHARACTERS: usize = 4_000_000;
pub const MAXIMUM_DEPTH: usize = 256;
pub const MAXIMUM_ELEMENTS: usize = 250_000;

#[derive(Clone, Copy, Debug)]
pub struct ReadOptions {
    /// XmlReaderSettings.IgnoreWhitespace: insignificant whitespace text nodes
    /// (outside xml:space="preserve") are not reported.
    pub ignore_whitespace: bool,
    /// FormulaXml.BoundedReader depth and element limits (and its kind label).
    pub formula_kind: Option<&'static str>,
    /// XmlReaderSettings.MaxCharactersInDocument (None = unlimited).
    pub max_characters: Option<usize>,
}

/// XDocument.Parse(xml, LoadOptions.PreserveWhitespace) / XElement.Parse(xml,
/// LoadOptions.PreserveWhitespace): every text node is kept.
pub const PRESERVE_WHITESPACE: ReadOptions =
    ReadOptions { ignore_whitespace: false, formula_kind: None, max_characters: None };

/// OmmlMathMlProjection.ReaderSettings(): IgnoreWhitespace = true.
pub const PROJECTION_READER: ReadOptions =
    ReadOptions { ignore_whitespace: true, formula_kind: None, max_characters: Some(MAXIMUM_CHARACTERS) };

/// The `new XmlReaderSettings { DtdProcessing = Prohibit, IgnoreComments = true,
/// IgnoreWhitespace = false, MaxCharactersInDocument = 4_000_000 }` readers of
/// WordOmmlConverter (ReplaceExtendedIntegralsWithOfficePlaceholders, ...).
pub const CONVERTER_READER: ReadOptions =
    ReadOptions { ignore_whitespace: false, formula_kind: None, max_characters: Some(MAXIMUM_CHARACTERS) };

/// IsBlockMathMl's reader: IgnoreWhitespace = true.
pub const IGNORE_WHITESPACE_READER: ReadOptions =
    ReadOptions { ignore_whitespace: true, formula_kind: None, max_characters: Some(MAXIMUM_CHARACTERS) };

/// FormulaXml.Parse(xml, kind, ...): IgnoreWhitespace = false plus the
/// BoundedReader limits.
pub const fn formula_reader(kind: &'static str) -> ReadOptions {
    ReadOptions { ignore_whitespace: false, formula_kind: Some(kind), max_characters: Some(MAXIMUM_CHARACTERS) }
}

pub fn utf16_len(value: &str) -> usize {
    value.chars().map(char::len_utf16).sum()
}

/// XmlCharType.IsWhiteSpace: the four XML whitespace characters.
pub fn is_xml_whitespace(c: char) -> bool {
    matches!(c, ' ' | '\t' | '\r' | '\n')
}

fn normalize_line_ends(value: &str) -> String {
    if value.contains('\r') {
        value.replace("\r\n", "\n").replace('\r', "\n")
    } else {
        value.to_string()
    }
}

/// XmlTextReaderImpl xml:space handling: XmlConvert.TrimString, then only
/// "preserve" and "default" are accepted.
fn xml_space_value(value: &str) -> Result<bool, String> {
    match value.trim_matches(is_xml_whitespace) {
        "preserve" => Ok(true),
        "default" => Ok(false),
        other => Err(format!("XmlException: '{other}' is an invalid xml:space value.")),
    }
}

/// Reads a document. The returned Doc's `root` is the document element.
pub fn parse(xml: &str, options: ReadOptions) -> Result<Doc, String> {
    parse_document(xml, options).map(|document| document.doc)
}

/// XDocument.Load: the tree plus the whitespace text nodes the XDocument
/// keeps before and after its root element (reported only when the reader
/// does not ignore whitespace). XDocument.ToString writes them around the root.
pub struct Document {
    pub doc: Doc,
    pub before_root: String,
    pub after_root: String,
}

pub fn parse_document(xml: &str, options: ReadOptions) -> Result<Document, String> {
    if let Some(kind) = options.formula_kind {
        if utf16_len(xml) > MAXIMUM_CHARACTERS {
            return Err(format!(
                "InvalidDataException: {kind} exceeds the formula limit of {MAXIMUM_CHARACTERS} characters."
            ));
        }
    }
    if let Some(limit) = options.max_characters {
        if utf16_len(xml) > limit {
            return Err(format!("XmlException: The input document has exceeded a limit set by MaxCharactersInDocument."));
        }
    }
    let mut doc = Doc::new();
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().expand_empty_elements = false;
    // (element, xml:space="preserve" in scope)
    let mut stack: Vec<(NodeId, bool)> = Vec::new();
    let mut pending = String::new();
    let mut elements = 0usize;
    let mut outside = (String::new(), String::new());
    loop {
        let (resolve, event) =
            reader.read_resolved_event().map_err(|error| format!("XmlException: Unable to parse XML: {error}"))?;
        let element_ns = match resolve {
            ResolveResult::Bound(namespace) => Ok(String::from_utf8_lossy(namespace.as_ref()).into_owned()),
            ResolveResult::Unbound => Ok(String::new()),
            ResolveResult::Unknown(prefix) => {
                Err(format!("XmlException: '{}' is an undeclared prefix.", String::from_utf8_lossy(&prefix)))
            }
        };
        let event = event.into_owned();
        match &event {
            Event::Text(text) => {
                let decoded = text.decode().map_err(|error| format!("XmlException: {error}"))?;
                pending.push_str(&normalize_line_ends(&decoded));
                continue;
            }
            Event::GeneralRef(reference) => {
                let raw = format!("&{};", reference.decode().map_err(|error| format!("XmlException: {error}"))?);
                let value = quick_xml::escape::unescape(&raw).map_err(|error| format!("XmlException: {error}"))?;
                pending.push_str(&value);
                continue;
            }
            _ => {}
        }
        flush_text(&mut doc, &stack, &mut pending, options, &mut outside)?;
        match event {
            Event::Start(start) => {
                let id = open_element(&mut doc, &reader, element_ns?, &start, &stack, options, &mut elements)?;
                let inherited = stack.last().map(|&(_, preserve)| preserve).unwrap_or(false);
                let own = match doc.attr_ns(id, XML_NS, "space") {
                    Some(value) => Some(xml_space_value(value)?),
                    None => None,
                };
                stack.push((id, own.unwrap_or(inherited)));
            }
            Event::Empty(start) => {
                let id = open_element(&mut doc, &reader, element_ns?, &start, &stack, options, &mut elements)?;
                if let Some(value) = doc.attr_ns(id, XML_NS, "space") {
                    xml_space_value(value)?;
                }
            }
            Event::End(_) => {
                let (id, _) = stack.pop().ok_or_else(|| "XmlException: unexpected end tag".to_string())?;
                // XContainer.ReadContentFrom: an end tag with no content read
                // leaves content = string.Empty.
                if doc.child_nodes(id).is_empty() {
                    doc.empty_string_content.insert(id);
                }
            }
            Event::CData(text) => {
                let decoded = text.decode().map_err(|error| format!("XmlException: {error}"))?.into_owned();
                match stack.last() {
                    Some(&(parent, _)) => doc.append_text(parent, &normalize_line_ends(&decoded)),
                    None => return Err("XmlException: Data at the root level is invalid.".to_string()),
                }
            }
            Event::DocType(_) => {
                return Err("XmlException: For security reasons DTD is prohibited in this XML document.".to_string());
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() {
        return Err("XmlException: Unexpected end of file has occurred.".to_string());
    }
    flush_text(&mut doc, &stack, &mut pending, options, &mut outside)?;
    if doc.root.is_none() {
        return Err("XmlException: Root element is missing.".to_string());
    }
    Ok(Document { doc, before_root: outside.0, after_root: outside.1 })
}

fn flush_text(
    doc: &mut Doc,
    stack: &[(NodeId, bool)],
    pending: &mut String,
    options: ReadOptions,
    outside: &mut (String, String),
) -> Result<(), String> {
    if pending.is_empty() {
        return Ok(());
    }
    let text = std::mem::take(pending);
    let whitespace = text.chars().all(is_xml_whitespace);
    match stack.last() {
        None => {
            if whitespace {
                if !options.ignore_whitespace {
                    if doc.root.is_none() {
                        outside.0.push_str(&text);
                    } else {
                        outside.1.push_str(&text);
                    }
                }
                Ok(())
            } else {
                Err("XmlException: Data at the root level is invalid.".to_string())
            }
        }
        Some(&(parent, preserve)) => {
            if !(options.ignore_whitespace && whitespace && !preserve) {
                doc.append_text(parent, &text);
            }
            Ok(())
        }
    }
}

fn open_element(
    doc: &mut Doc,
    reader: &NsReader<&[u8]>,
    ns: String,
    start: &BytesStart,
    stack: &[(NodeId, bool)],
    options: ReadOptions,
    elements: &mut usize,
) -> Result<NodeId, String> {
    if let Some(kind) = options.formula_kind {
        if stack.len() >= MAXIMUM_DEPTH {
            return Err(format!("InvalidDataException: {kind} exceeds the formula nesting limit of {MAXIMUM_DEPTH} levels."));
        }
        *elements += 1;
        if *elements > MAXIMUM_ELEMENTS {
            return Err(format!("InvalidDataException: {kind} exceeds the formula limit of {MAXIMUM_ELEMENTS} elements."));
        }
    }
    let local = String::from_utf8_lossy(start.local_name().as_ref()).into_owned();
    let id = doc.new_element(Name { ns, local });
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| format!("XmlException: {error}"))?;
        let key = attribute.key;
        let raw_key = String::from_utf8_lossy(key.as_ref()).into_owned();
        let name = if raw_key == "xmlns" {
            Name::new("", "xmlns")
        } else if let Some(prefix) = raw_key.strip_prefix("xmlns:") {
            Name::new(XMLNS_NS, prefix)
        } else {
            let (resolved, local_name) = reader.resolver().resolve_attribute(key);
            let attribute_ns = if raw_key.starts_with("xml:") {
                XML_NS.to_string()
            } else {
                match resolved {
                    ResolveResult::Bound(namespace) => String::from_utf8_lossy(namespace.as_ref()).into_owned(),
                    ResolveResult::Unbound => String::new(),
                    ResolveResult::Unknown(prefix) => {
                        return Err(format!(
                            "XmlException: '{}' is an undeclared prefix.",
                            String::from_utf8_lossy(&prefix)
                        ))
                    }
                }
            };
            Name { ns: attribute_ns, local: String::from_utf8_lossy(local_name.as_ref()).into_owned() }
        };
        // Attribute-value normalization: literal line ends and tabs become
        // spaces; character references keep their characters.
        let raw = std::str::from_utf8(&attribute.value).map_err(|error| format!("XmlException: {error}"))?;
        let normalized = raw.replace("\r\n", " ").replace(['\r', '\n', '\t'], " ");
        let value =
            quick_xml::escape::unescape(&normalized).map_err(|error| format!("XmlException: {error}"))?.into_owned();
        if doc.attrs(id).iter().any(|(existing, _)| *existing == name) {
            return Err(format!("XmlException: '{}' is a duplicate attribute name.", raw_key));
        }
        if let Data::Element { attrs, .. } = &mut doc.nodes[id].data {
            attrs.push((name, value));
        }
    }
    if let Some(&(parent, _)) = stack.last() {
        doc.append(parent, id);
    } else if doc.root.replace(id).is_some() {
        return Err("XmlException: There are multiple root elements.".to_string());
    }
    Ok(id)
}
