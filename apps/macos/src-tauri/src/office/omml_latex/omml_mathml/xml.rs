//! XmlReader + XDocument.Load equivalents on top of the shared arena DOM.
//!
//! dom::Doc::parse decides whitespace per element after the fact; the C#
//! readers decide it per text node while reading (XmlTextReaderImpl.ParseText),
//! so this module reads with quick-xml directly and builds the same tree.

use super::dom::{Data, Doc, Name, NodeId, XML_NS};
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
    /// FormulaXml.BoundedReader depth and element limits.
    pub formula_limits: bool,
}

/// OmmlMathMlProjection.ReaderSettings().
pub const PROJECTION_READER: ReadOptions = ReadOptions { ignore_whitespace: true, formula_limits: false };
/// FormulaXml.Parse(xml, kind, LoadOptions.None): IgnoreWhitespace = false.
pub const FORMULA_READER: ReadOptions = ReadOptions { ignore_whitespace: false, formula_limits: true };
/// XDocument.Parse(xml, LoadOptions.PreserveWhitespace).
pub const PRESERVE_READER: ReadOptions = ReadOptions { ignore_whitespace: false, formula_limits: false };

fn utf16_len(value: &str) -> usize {
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
        other => Err(format!("'{other}' is an invalid xml:space value.")),
    }
}

pub fn parse(xml: &str, options: ReadOptions) -> Result<Doc, String> {
    if utf16_len(xml) > MAXIMUM_CHARACTERS {
        return Err(format!("The XML input exceeds the limit of {MAXIMUM_CHARACTERS} characters."));
    }
    let mut doc = Doc::new();
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().expand_empty_elements = false;
    // (element, xml:space="preserve" in scope)
    let mut stack: Vec<(NodeId, bool)> = Vec::new();
    let mut pending = String::new();
    let mut elements = 0usize;
    loop {
        let (resolve, event) =
            reader.read_resolved_event().map_err(|error| format!("Unable to parse XML: {error}"))?;
        let element_ns = match resolve {
            ResolveResult::Bound(namespace) => Ok(String::from_utf8_lossy(namespace.as_ref()).into_owned()),
            ResolveResult::Unbound => Ok(String::new()),
            ResolveResult::Unknown(prefix) => {
                Err(format!("Unknown XML namespace prefix {}", String::from_utf8_lossy(&prefix)))
            }
        };
        let event = event.into_owned();
        match &event {
            Event::Text(text) => {
                let decoded = text.decode().map_err(|error| error.to_string())?;
                pending.push_str(&normalize_line_ends(&decoded));
                continue;
            }
            Event::GeneralRef(reference) => {
                let raw = format!("&{};", reference.decode().map_err(|error| error.to_string())?);
                let value = quick_xml::escape::unescape(&raw).map_err(|error| error.to_string())?;
                pending.push_str(&value);
                continue;
            }
            _ => {}
        }
        flush_text(&mut doc, &stack, &mut pending, options)?;
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
                stack.pop().ok_or_else(|| "XML contains an unexpected closing tag".to_string())?;
            }
            Event::CData(text) => {
                let decoded = text.decode().map_err(|error| error.to_string())?.into_owned();
                // A CDATA section is its own node and never insignificant whitespace.
                match stack.last() {
                    Some(&(parent, _)) => doc.append_text(parent, &normalize_line_ends(&decoded)),
                    None => return Err("XML contains data outside the root element".to_string()),
                }
            }
            Event::DocType(_) => {
                return Err("XML input must not contain a document type declaration".to_string());
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() {
        return Err("XML input is incomplete".to_string());
    }
    if doc.root.is_none() {
        return Err("XML input has no root element".to_string());
    }
    Ok(doc)
}

fn flush_text(doc: &mut Doc, stack: &[(NodeId, bool)], pending: &mut String, options: ReadOptions) -> Result<(), String> {
    if pending.is_empty() {
        return Ok(());
    }
    let text = std::mem::take(pending);
    let whitespace = text.chars().all(is_xml_whitespace);
    match stack.last() {
        None => {
            if whitespace {
                Ok(())
            } else {
                Err("XML contains text outside the root element".to_string())
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
    if options.formula_limits {
        if stack.len() >= MAXIMUM_DEPTH {
            return Err(format!("XML exceeds the formula nesting limit of {MAXIMUM_DEPTH} levels."));
        }
        *elements += 1;
        if *elements > MAXIMUM_ELEMENTS {
            return Err(format!("XML exceeds the formula limit of {MAXIMUM_ELEMENTS} elements."));
        }
    }
    let local = String::from_utf8_lossy(start.local_name().as_ref()).into_owned();
    let id = doc.new_element(Name { ns, local });
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| error.to_string())?;
        let key = attribute.key;
        let raw_key = String::from_utf8_lossy(key.as_ref()).into_owned();
        if raw_key == "xmlns" || raw_key.starts_with("xmlns:") {
            continue;
        }
        let (resolved, local_name) = reader.resolver().resolve_attribute(key);
        let attribute_ns = if raw_key.starts_with("xml:") {
            XML_NS.to_string()
        } else {
            match resolved {
                ResolveResult::Bound(namespace) => String::from_utf8_lossy(namespace.as_ref()).into_owned(),
                ResolveResult::Unbound => String::new(),
                ResolveResult::Unknown(prefix) => {
                    return Err(format!("Unknown XML namespace prefix {}", String::from_utf8_lossy(&prefix)))
                }
            }
        };
        // Attribute-value normalization: literal line ends and tabs become
        // spaces; character references keep their characters.
        let raw = std::str::from_utf8(&attribute.value).map_err(|error| error.to_string())?;
        let normalized = raw.replace("\r\n", " ").replace(['\r', '\n', '\t'], " ");
        let value = quick_xml::escape::unescape(&normalized).map_err(|error| error.to_string())?.into_owned();
        let name = Name { ns: attribute_ns, local: String::from_utf8_lossy(local_name.as_ref()).into_owned() };
        doc.set_attr_ns(id, &name.ns, &name.local, Some(&value));
    }
    if let Some(&(parent, _)) = stack.last() {
        doc.append(parent, id);
    } else if doc.root.replace(id).is_some() {
        return Err("XML contains multiple root elements".to_string());
    }
    Ok(id)
}

/// The observable effect of XElement.ToString followed by a parse with the
/// same tree (XmlWriter NewLineHandling.Replace, then XML end-of-line
/// normalization): CR and CRLF in text become LF. Attributes are entitized
/// by the writer and survive unchanged.
pub fn normalize_text_line_ends(doc: &mut Doc, root: NodeId) {
    for node in doc.descendant_nodes(root) {
        if let Data::Text(text) = &doc.nodes[node].data {
            if text.contains('\r') {
                let normalized = normalize_line_ends(text);
                doc.set_text(node, &normalized);
            }
        }
    }
}

/// Re-reading a serialized tree with IgnoreWhitespace = true: whitespace-only
/// text outside xml:space="preserve" disappears.
pub fn drop_insignificant_whitespace(doc: &mut Doc, root: NodeId) -> Result<(), String> {
    drop_whitespace_in(doc, root, false)
}

fn drop_whitespace_in(doc: &mut Doc, node: NodeId, inherited: bool) -> Result<(), String> {
    let preserve = match doc.attr_ns(node, XML_NS, "space") {
        Some(value) => xml_space_value(value)?,
        None => inherited,
    };
    for child in doc.child_nodes(node).to_vec() {
        match &doc.nodes[child].data {
            Data::Text(text) => {
                if !preserve && text.chars().all(is_xml_whitespace) {
                    doc.detach(child);
                }
            }
            Data::Element { .. } => drop_whitespace_in(doc, child, preserve)?,
        }
    }
    Ok(())
}

/// XElement.ToString(SaveOptions.DisableFormatting) for a tree whose elements
/// are all in one default namespace (the MathML output).
pub fn to_default_namespace_xml(doc: &Doc, root: NodeId, namespace: &str) -> String {
    let mut out = String::new();
    write_node(doc, root, Some(namespace), &mut out);
    out
}

fn write_node(doc: &Doc, node: NodeId, declare: Option<&str>, out: &mut String) {
    match &doc.nodes[node].data {
        Data::Text(text) => out.push_str(&super::dom::escape_text(text)),
        Data::Element { name, attrs } => {
            out.push('<');
            out.push_str(&name.local);
            for (attribute, value) in attrs {
                out.push(' ');
                if attribute.ns == XML_NS {
                    out.push_str("xml:");
                } else if !attribute.ns.is_empty() {
                    out.push_str("p1:");
                }
                out.push_str(&attribute.local);
                out.push_str("=\"");
                out.push_str(&super::dom::escape_attr(value));
                out.push('"');
            }
            if let Some(namespace) = declare {
                out.push_str(" xmlns=\"");
                out.push_str(&super::dom::escape_attr(namespace));
                out.push('"');
            }
            let children = doc.child_nodes(node);
            if children.is_empty() {
                out.push_str(" />");
                return;
            }
            out.push('>');
            for &child in children {
                write_node(doc, child, None, out);
            }
            out.push_str("</");
            out.push_str(&name.local);
            out.push('>');
        }
    }
}
