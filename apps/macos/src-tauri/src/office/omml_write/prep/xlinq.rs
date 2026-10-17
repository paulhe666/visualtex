//! System.Xml.Linq as the Windows pre-passes use it, on top of the arena DOM:
//!
//! * [`parse`]: `XmlReader.Create(StringReader, settings)` + `XDocument.Load`
//!   (or `XDocument.Parse`), with the observable reader rules: end-of-line
//!   normalization in text, attribute-value normalization, character
//!   references, XML character validity, whitespace text nodes kept or
//!   dropped, `<x></x>` kept distinct from `<x />` (XLinq's empty-string
//!   content), namespace declarations kept as attributes in source order, and
//!   top-level whitespace kept by the XDocument.
//! * [`XDocument::to_string`] / [`element_to_string`]:
//!   `XDocument.ToString(SaveOptions.DisableFormatting)` and
//!   `XElement.ToString(SaveOptions.DisableFormatting)` on Windows: XLinq's
//!   ElementWriter (namespace resolver) feeding an XmlWellFormedWriter
//!   (implicit/explicit namespace declarations, generated `p1` prefixes) and
//!   an XmlEncodedRawTextWriter with NewLineHandling.Replace and
//!   NewLineChars = "\r\n" (text newlines become CRLF; attribute `\t \n \r`
//!   become `&#x9; &#xA; &#xD;`).
//! * Tree edits with XLinq semantics: adding a node that already has a parent
//!   adds a deep copy of it (XContainer.AddNode / Inserter.AddNode), the
//!   document root can be replaced, and so on.
//!
//! Not modelled (none of them occurs in MathJax or VisualTeX MathML; every
//! reader of the chain after s04 drops comments anyway):
//! * comments and processing instructions are dropped while reading (XLinq
//!   keeps XComment/XProcessingInstruction nodes where IgnoreComments is off);
//! * CDATA sections are read as ordinary text (XLinq keeps an XCData node and
//!   writes it back as a CDATA section);
//! * a document type declaration is rejected (XDocument.Parse would parse it;
//!   FormulaXml, and so ValidateMathMlForOmml, rejects it as well).

use crate::office::omml_write::prep::dom::{Data, Doc, Name, NodeId, XMLNS_NS, XML_NS};
use crate::office::omml_write::prep::netstr::len16;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

/// FormulaXml limits (also XmlReaderSettings.MaxCharactersInDocument of the
/// pre-pass readers).
pub const MAXIMUM_CHARACTERS: usize = 4_000_000;
pub const MAXIMUM_DEPTH: usize = 256;
pub const MAXIMUM_ELEMENTS: usize = 250_000;

/// The XmlReaderSettings/LoadOptions combination a pass reads with.
#[derive(Clone, Copy, Debug)]
pub struct ReaderSettings {
    /// XmlReaderSettings.IgnoreComments.
    pub ignore_comments: bool,
    /// XmlReaderSettings.IgnoreWhitespace (LoadOptions without PreserveWhitespace).
    pub ignore_whitespace: bool,
    /// XmlReaderSettings.MaxCharactersInDocument (0 = unlimited in .NET).
    pub max_characters: Option<usize>,
    /// FormulaXml.BoundedReader: the `kind` used in its messages.
    pub formula_kind: Option<&'static str>,
}

/// `XDocument.Parse(xml, LoadOptions.PreserveWhitespace)`.
pub const XDOCUMENT_PARSE_PRESERVE: ReaderSettings =
    ReaderSettings { ignore_comments: false, ignore_whitespace: false, max_characters: None, formula_kind: None };

/// The explicit reader of NormalizeNestedEmptyBaseScripts, NormalizeMathMlAccents,
/// NormalizeNaryArguments and ReplaceExtendedIntegralsWithOfficePlaceholders:
/// DtdProcessing.Prohibit, IgnoreComments, IgnoreWhitespace = false,
/// MaxCharactersInDocument = 4,000,000, loaded with PreserveWhitespace.
pub const PREPASS_READER: ReaderSettings = ReaderSettings {
    ignore_comments: true,
    ignore_whitespace: false,
    max_characters: Some(MAXIMUM_CHARACTERS),
    formula_kind: None,
};

/// The reader of IsBlockMathMl: as [`PREPASS_READER`] but IgnoreWhitespace.
pub const BLOCK_READER: ReaderSettings = ReaderSettings {
    ignore_comments: true,
    ignore_whitespace: true,
    max_characters: Some(MAXIMUM_CHARACTERS),
    formula_kind: None,
};

/// `FormulaXml.Parse(xml, kind)` (LoadOptions.PreserveWhitespace).
pub const fn formula_reader(kind: &'static str) -> ReaderSettings {
    ReaderSettings {
        ignore_comments: false,
        ignore_whitespace: false,
        max_characters: Some(MAXIMUM_CHARACTERS),
        formula_kind: Some(kind),
    }
}

/// An XDocument: the element tree plus the whitespace text nodes it keeps
/// before and after the root element.
#[derive(Clone, Debug)]
pub struct XDocument {
    pub doc: Doc,
    pub before_root: String,
    pub after_root: String,
}

impl XDocument {
    /// XDocument.Root.
    pub fn root(&self) -> Option<NodeId> {
        self.doc.root
    }

    /// XDocument.Descendants(): every element in document order.
    pub fn descendants(&self) -> Vec<NodeId> {
        match self.doc.root {
            Some(root) => self.doc.descendants_and_self(root),
            None => Vec::new(),
        }
    }

    /// XDocument.Descendants(name).
    pub fn descendants_named(&self, ns: &str, local: &str) -> Vec<NodeId> {
        self.descendants().into_iter().filter(|&id| self.doc.is(id, ns, local)).collect()
    }

    /// `XDocument.ToString(SaveOptions.DisableFormatting)`.
    pub fn to_string(&self) -> Result<String, String> {
        let mut out = String::new();
        escape_text_into(&self.before_root, &mut out);
        if let Some(root) = self.doc.root {
            out.push_str(&element_to_string(&self.doc, root)?);
        }
        escape_text_into(&self.after_root, &mut out);
        Ok(out)
    }
}

// ---- reading ---------------------------------------------------------------

/// XmlCharType.IsWhiteSpace: the four XML white-space characters.
pub fn is_xml_whitespace(c: char) -> bool {
    matches!(c, ' ' | '\t' | '\r' | '\n')
}

/// XML 1.0 Char production (XmlReaderSettings.CheckCharacters).
fn is_xml_char(c: char) -> bool {
    matches!(c, '\t' | '\n' | '\r' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
}

fn check_chars(value: &str) -> Result<(), String> {
    match value.chars().find(|&c| !is_xml_char(c)) {
        Some(c) => Err(format!("'{}', hexadecimal value 0x{:02X}, is an invalid character.", c.escape_unicode(), c as u32)),
        None => Ok(()),
    }
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
fn xml_space_preserve(value: &str) -> Result<bool, String> {
    match value.trim_matches(is_xml_whitespace) {
        "preserve" => Ok(true),
        "default" => Ok(false),
        other => Err(format!("'{other}' is an invalid xml:space value.")),
    }
}

fn resolve_reference(name: &str) -> Result<String, String> {
    let resolved = quick_xml::escape::unescape(&format!("&{name};"))
        .map_err(|_| format!("Reference to undeclared entity '{name}'."))?
        .into_owned();
    check_chars(&resolved)?;
    Ok(resolved)
}

struct ReadState {
    doc: Doc,
    before_root: String,
    after_root: String,
    /// Open elements and whether xml:space="preserve" is in scope.
    stack: Vec<(NodeId, bool)>,
    pending: String,
    elements: usize,
}

/// Reads `xml` the way `XDocument.Load(XmlReader.Create(new StringReader(xml),
/// settings), options)` does (see the module documentation for what is not
/// modelled). Errors carry XmlException-like messages.
pub fn parse(xml: &str, settings: &ReaderSettings) -> Result<XDocument, String> {
    if let Some(limit) = settings.max_characters {
        if len16(xml) > limit {
            return Err(match settings.formula_kind {
                Some(kind) => format!("{kind} exceeds the formula limit of {limit} characters."),
                None => format!(
                    "The input document has exceeded a limit set by MaxCharactersInDocument ({limit})."
                ),
            });
        }
    }
    // A StringReader hands XmlReader the characters as they are: a leading
    // U+FEFF is content at the root level, not an encoding signature.
    if xml.starts_with('\u{FEFF}') {
        return Err("Data at the root level is invalid. Line 1, position 1.".to_string());
    }
    let mut state = ReadState {
        doc: Doc::new(),
        before_root: String::new(),
        after_root: String::new(),
        stack: Vec::new(),
        pending: String::new(),
        elements: 0,
    };
    let mut reader = NsReader::from_str(xml);
    reader.config_mut().trim_text(false);
    reader.config_mut().expand_empty_elements = false;
    reader.config_mut().check_comments = true;
    let mut first_event = true;
    loop {
        let (resolve, event) =
            reader.read_resolved_event().map_err(|error| format!("Unable to parse XML: {error}"))?;
        let element_ns = match resolve {
            ResolveResult::Bound(namespace) => Ok(String::from_utf8_lossy(namespace.as_ref()).into_owned()),
            ResolveResult::Unbound => Ok(String::new()),
            ResolveResult::Unknown(prefix) => {
                Err(format!("'{}' is an undeclared prefix.", String::from_utf8_lossy(&prefix)))
            }
        };
        let is_first = std::mem::replace(&mut first_event, false);
        match &event {
            Event::Text(text) => {
                let decoded = text.decode().map_err(|error| error.to_string())?;
                if decoded.contains("]]>") {
                    return Err("']]>' is not allowed in character data.".to_string());
                }
                check_chars(&decoded)?;
                state.pending.push_str(&normalize_line_ends(&decoded));
                continue;
            }
            Event::GeneralRef(reference) => {
                let name = reference.decode().map_err(|error| error.to_string())?;
                let value = resolve_reference(&name)?;
                if state.stack.is_empty() {
                    return Err("Data at the root level is invalid.".to_string());
                }
                state.pending.push_str(&value);
                continue;
            }
            _ => {}
        }
        flush_text(&mut state, settings)?;
        match event {
            Event::Start(start) => {
                let id = open_element(&mut state, &reader, element_ns?, &start, settings)?;
                let inherited = state.stack.last().map(|&(_, preserve)| preserve).unwrap_or(false);
                let own = match state.doc.attr_ns(id, XML_NS, "space") {
                    Some(value) => Some(xml_space_preserve(value)?),
                    None => None,
                };
                state.stack.push((id, own.unwrap_or(inherited)));
            }
            Event::Empty(start) => {
                let id = open_element(&mut state, &reader, element_ns?, &start, settings)?;
                if let Some(value) = state.doc.attr_ns(id, XML_NS, "space") {
                    xml_space_preserve(value)?;
                }
            }
            Event::End(_) => {
                let (id, _) =
                    state.stack.pop().ok_or_else(|| "XML contains an unexpected end tag.".to_string())?;
                // XContainer.ContentReader: an element read as <x></x> (or whose
                // only content was skipped) gets content == "".
                if state.doc.nodes[id].children.is_empty() {
                    state.doc.nodes[id].empty_string_content = true;
                }
            }
            Event::CData(text) => {
                let decoded = text.decode().map_err(|error| error.to_string())?.into_owned();
                check_chars(&decoded)?;
                match state.stack.last() {
                    Some(&(parent, _)) => state.doc.append_text(parent, &normalize_line_ends(&decoded)),
                    None => return Err("Data at the root level is invalid.".to_string()),
                }
            }
            Event::Comment(_) | Event::PI(_) => {}
            Event::Decl(_) => {
                if !is_first {
                    return Err("Unexpected XML declaration. The XML declaration must be the first node in the document, and no whitespace characters are allowed to appear before it.".to_string());
                }
            }
            Event::DocType(_) => {
                return Err("For security reasons DTD is prohibited in this XML document.".to_string());
            }
            Event::Eof => break,
            Event::Text(_) | Event::GeneralRef(_) => unreachable!(),
        }
    }
    if !state.stack.is_empty() {
        return Err("Unexpected end of file has occurred. The following elements are not closed.".to_string());
    }
    if state.doc.root.is_none() {
        return Err("Root element is missing.".to_string());
    }
    Ok(XDocument { doc: state.doc, before_root: state.before_root, after_root: state.after_root })
}

fn flush_text(state: &mut ReadState, settings: &ReaderSettings) -> Result<(), String> {
    if state.pending.is_empty() {
        return Ok(());
    }
    let text = std::mem::take(&mut state.pending);
    let whitespace = text.chars().all(is_xml_whitespace);
    match state.stack.last() {
        None => {
            if !whitespace {
                return Err("Data at the root level is invalid.".to_string());
            }
            if !settings.ignore_whitespace {
                if state.doc.root.is_none() {
                    state.before_root.push_str(&text);
                } else {
                    state.after_root.push_str(&text);
                }
            }
            Ok(())
        }
        Some(&(parent, preserve)) => {
            if !(settings.ignore_whitespace && whitespace && !preserve) {
                state.doc.append_text(parent, &text);
            }
            Ok(())
        }
    }
}

fn open_element(
    state: &mut ReadState,
    reader: &NsReader<&[u8]>,
    ns: String,
    start: &BytesStart,
    settings: &ReaderSettings,
) -> Result<NodeId, String> {
    if let Some(kind) = settings.formula_kind {
        // FormulaXml.BoundedReader: XmlReader.Depth of the element.
        if state.stack.len() >= MAXIMUM_DEPTH {
            return Err(format!("{kind} exceeds the formula nesting limit of {MAXIMUM_DEPTH} levels."));
        }
        state.elements += 1;
        if state.elements > MAXIMUM_ELEMENTS {
            return Err(format!("{kind} exceeds the formula limit of {MAXIMUM_ELEMENTS} elements."));
        }
    }
    check_qualified_name(&String::from_utf8_lossy(start.name().as_ref()))?;
    let local = String::from_utf8_lossy(start.local_name().as_ref()).into_owned();
    let id = state.doc.new_element(Name { ns, local });
    for attribute in start.attributes() {
        let attribute = attribute.map_err(|error| format!("Unable to parse XML: {error}"))?;
        let key = attribute.key;
        let raw_key = String::from_utf8_lossy(key.as_ref()).into_owned();
        check_qualified_name(&raw_key)?;
        let name = if raw_key == "xmlns" {
            Name::new("", "xmlns")
        } else if let Some(prefix) = raw_key.strip_prefix("xmlns:") {
            Name::new(XMLNS_NS, prefix)
        } else if let Some(local) = raw_key.strip_prefix("xml:") {
            Name::new(XML_NS, local)
        } else {
            let (resolved, local_name) = reader.resolver().resolve_attribute(key);
            let attribute_ns = match resolved {
                ResolveResult::Bound(namespace) => String::from_utf8_lossy(namespace.as_ref()).into_owned(),
                ResolveResult::Unbound => String::new(),
                ResolveResult::Unknown(prefix) => {
                    return Err(format!("'{}' is an undeclared prefix.", String::from_utf8_lossy(&prefix)))
                }
            };
            Name { ns: attribute_ns, local: String::from_utf8_lossy(local_name.as_ref()).into_owned() }
        };
        let raw = std::str::from_utf8(&attribute.value).map_err(|error| error.to_string())?;
        if raw.contains('<') {
            return Err("'<', hexadecimal value 0x3C, is an invalid attribute character.".to_string());
        }
        check_chars(raw)?;
        // Attribute-value normalization (no DTD: every attribute is CDATA):
        // literal line ends and tabs become spaces; references keep their
        // characters.
        let normalized = raw.replace("\r\n", " ").replace(['\r', '\n', '\t'], " ");
        let value = unescape_attribute(&normalized)?;
        if state.doc.attrs(id).iter().any(|(existing, _)| *existing == name) {
            return Err(format!("'{}' is a duplicate attribute name.", name.local));
        }
        if name.ns == XMLNS_NS && value.is_empty() {
            return Err("Invalid namespace declaration.".to_string());
        }
        if let Data::Element { attrs, .. } = &mut state.doc.nodes[id].data {
            attrs.push((name, value));
        }
    }
    match state.stack.last() {
        Some(&(parent, _)) => state.doc.append(parent, id),
        None => {
            if state.doc.root.replace(id).is_some() {
                return Err("There are multiple root elements.".to_string());
            }
        }
    }
    Ok(id)
}

/// XML 1.0 (fifth edition) NameStartChar without ':'.
fn is_name_start_char(c: char) -> bool {
    matches!(c,
        'A'..='Z' | '_' | 'a'..='z' | '\u{C0}'..='\u{D6}' | '\u{D8}'..='\u{F6}' | '\u{F8}'..='\u{2FF}'
        | '\u{370}'..='\u{37D}' | '\u{37F}'..='\u{1FFF}' | '\u{200C}'..='\u{200D}' | '\u{2070}'..='\u{218F}'
        | '\u{2C00}'..='\u{2FEF}' | '\u{3001}'..='\u{D7FF}' | '\u{F900}'..='\u{FDCF}' | '\u{FDF0}'..='\u{FFFD}'
        | '\u{10000}'..='\u{EFFFF}')
}

/// XML 1.0 (fifth edition) NameChar without ':'.
fn is_name_char(c: char) -> bool {
    is_name_start_char(c)
        || matches!(c, '-' | '.' | '0'..='9' | '\u{B7}' | '\u{300}'..='\u{36F}' | '\u{203F}'..='\u{2040}')
}

/// XmlTextReaderImpl.ParseQName with namespaces: NCName or NCName:NCName.
fn check_qualified_name(name: &str) -> Result<(), String> {
    let valid_ncname = |part: &str| {
        let mut chars = part.chars();
        chars.next().is_some_and(is_name_start_char) && chars.all(is_name_char)
    };
    let valid = match name.split_once(':') {
        Some((prefix, local)) => valid_ncname(prefix) && valid_ncname(local),
        None => valid_ncname(name),
    };
    if valid {
        Ok(())
    } else {
        Err(format!("Name '{name}' is not a valid XML name."))
    }
}

fn unescape_attribute(value: &str) -> Result<String, String> {
    let mut out = String::with_capacity(value.len());
    let mut rest = value;
    while let Some(index) = rest.find('&') {
        out.push_str(&rest[..index]);
        let after = &rest[index + 1..];
        let end = after.find(';').ok_or_else(|| "An error occurred while parsing EntityName.".to_string())?;
        out.push_str(&resolve_reference(&after[..end])?);
        rest = &after[end + 1..];
    }
    out.push_str(rest);
    Ok(out)
}

// ---- writing ---------------------------------------------------------------

/// XmlEncodedRawTextWriter.WriteElementTextBlock with NewLineHandling.Replace
/// and NewLineChars = "\r\n".
pub fn escape_text_into(value: &str, out: &mut String) {
    let mut chars = value.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '\r' => {
                if chars.peek() == Some(&'\n') {
                    chars.next();
                }
                out.push_str("\r\n");
            }
            '\n' => out.push_str("\r\n"),
            _ => out.push(c),
        }
    }
}

/// XmlEncodedRawTextWriter.WriteAttributeTextBlock (NewLineHandling.Replace).
pub fn escape_attribute_into(value: &str, out: &mut String) {
    for c in value.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\t' => out.push_str("&#x9;"),
            '\n' => out.push_str("&#xA;"),
            '\r' => out.push_str("&#xD;"),
            _ => out.push(c),
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum NamespaceKind {
    Written,
    NeedToWrite,
    Implied,
    Special,
}

struct WriterNamespace {
    prefix: String,
    uri: String,
    kind: NamespaceKind,
}

enum SpecialAttribute {
    DefaultXmlns,
    PrefixedXmlns(String),
    XmlSpace,
    XmlLang,
}

/// The namespace bookkeeping of XmlWellFormedWriter (no predefined
/// namespaces, OmitDuplicateNamespaces off) over the raw text output.
struct WellFormedWriter {
    out: String,
    namespaces: Vec<WriterNamespace>,
    /// Per open element: (prevNSTop, prefix, local name).
    elements: Vec<(usize, String, String)>,
    start_tag_open: bool,
    /// Attributes of the open start tag: (prefix, local, namespace).
    attributes: Vec<(String, String, String)>,
}

impl WellFormedWriter {
    fn new() -> Self {
        let namespace = |prefix: &str, uri: &str, kind| WriterNamespace {
            prefix: prefix.to_string(),
            uri: uri.to_string(),
            kind,
        };
        Self {
            out: String::new(),
            namespaces: vec![
                namespace("xmlns", XMLNS_NS, NamespaceKind::Special),
                namespace("xml", XML_NS, NamespaceKind::Special),
                namespace("", "", NamespaceKind::Implied),
            ],
            elements: Vec::new(),
            start_tag_open: false,
            attributes: Vec::new(),
        }
    }

    fn top(&self) -> usize {
        self.namespaces.len() - 1
    }

    fn prev_ns_top(&self) -> usize {
        self.elements.last().map(|&(prev, _, _)| prev).unwrap_or(0)
    }

    fn lookup_namespace_index(&self, prefix: &str) -> Option<usize> {
        self.namespaces.iter().rposition(|ns| ns.prefix == prefix)
    }

    fn lookup_namespace(&self, prefix: &str) -> Option<&str> {
        self.lookup_namespace_index(prefix).map(|index| self.namespaces[index].uri.as_str())
    }

    /// LookupLocalNamespace: declarations of the current start tag only.
    fn lookup_local_namespace(&self, prefix: &str) -> Option<String> {
        let prev = self.prev_ns_top();
        (prev + 1..self.namespaces.len())
            .rev()
            .find(|&index| self.namespaces[index].prefix == prefix)
            .map(|index| self.namespaces[index].uri.clone())
    }

    /// LookupPrefix: the most recent declaration of `ns`, unless its prefix
    /// was redeclared later.
    fn lookup_prefix(&self, ns: &str) -> Option<String> {
        let index = self.namespaces.iter().rposition(|declaration| declaration.uri == ns)?;
        let prefix = &self.namespaces[index].prefix;
        if self.namespaces[index + 1..].iter().any(|declaration| declaration.prefix == *prefix) {
            return None;
        }
        Some(prefix.clone())
    }

    fn generate_prefix(&self) -> String {
        let generated = format!("p{}", self.top() - 2);
        if self.lookup_namespace(&generated).is_none() {
            return generated;
        }
        let mut counter = 0;
        loop {
            let candidate = format!("{generated}{counter}");
            counter += 1;
            if self.lookup_namespace(&candidate).is_none() {
                return candidate;
            }
        }
    }

    fn push_namespace_implicit(&mut self, prefix: &str, ns: &str) -> Result<(), String> {
        let kind = match self.lookup_namespace_index(prefix) {
            Some(index) if index > self.prev_ns_top() => {
                if self.namespaces[index].uri != ns {
                    return Err(format!(
                        "The prefix '{prefix}' cannot be redefined from '{}' to '{ns}' within the same start element tag.",
                        self.namespaces[index].uri
                    ));
                }
                return Ok(());
            }
            Some(index) => {
                if self.namespaces[index].kind == NamespaceKind::Special {
                    if prefix == "xml" && ns == XML_NS {
                        NamespaceKind::Implied
                    } else {
                        return Err(format!("Prefix \"{prefix}\" is reserved for use by XML."));
                    }
                } else if self.namespaces[index].uri == ns {
                    NamespaceKind::Implied
                } else {
                    NamespaceKind::NeedToWrite
                }
            }
            None => {
                if (ns == XML_NS && prefix != "xml") || (ns == XMLNS_NS && prefix != "xmlns") {
                    return Err(format!("Prefix '{prefix}' cannot be mapped to namespace name reserved for \"xml\" or \"xmlns\"."));
                }
                NamespaceKind::NeedToWrite
            }
        };
        self.namespaces.push(WriterNamespace { prefix: prefix.to_string(), uri: ns.to_string(), kind });
        Ok(())
    }

    /// PushNamespaceExplicit; returns whether the declaration is written.
    fn push_namespace_explicit(&mut self, prefix: &str, ns: &str) -> Result<bool, String> {
        match self.lookup_namespace_index(prefix) {
            Some(index) if index > self.prev_ns_top() => {
                if self.namespaces[index].uri != ns {
                    return Err(format!(
                        "The prefix '{prefix}' cannot be redefined from '{}' to '{ns}' within the same start element tag.",
                        self.namespaces[index].uri
                    ));
                }
                if self.namespaces[index].kind == NamespaceKind::Written {
                    let name = if prefix.is_empty() { "xmlns".to_string() } else { format!("xmlns:{prefix}") };
                    return Err(format!("'{name}' is a duplicate attribute name."));
                }
                self.namespaces[index].kind = NamespaceKind::Written;
                return Ok(true);
            }
            Some(_) => {}
            None => {
                if (ns == XML_NS && prefix != "xml") || (ns == XMLNS_NS && prefix != "xmlns") {
                    return Err(format!("Prefix '{prefix}' cannot be mapped to namespace name reserved for \"xml\" or \"xmlns\"."));
                }
                if prefix == "xml" && ns != XML_NS {
                    return Err("Prefix \"xml\" is reserved for use by XML and can be mapped only to namespace name \"http://www.w3.org/XML/1998/namespace\".".to_string());
                }
                if prefix == "xmlns" {
                    return Err("Prefix \"xmlns\" is reserved for use by XML.".to_string());
                }
            }
        }
        self.namespaces.push(WriterNamespace {
            prefix: prefix.to_string(),
            uri: ns.to_string(),
            kind: NamespaceKind::Written,
        });
        Ok(true)
    }

    fn write_declaration(&mut self, prefix: &str, uri: &str) {
        if prefix.is_empty() {
            self.out.push_str(" xmlns=\"");
        } else {
            self.out.push_str(" xmlns:");
            self.out.push_str(prefix);
            self.out.push_str("=\"");
        }
        escape_attribute_into(uri, &mut self.out);
        self.out.push('"');
    }

    /// StartElementContent without the closing '>' of the start tag.
    fn write_pending_declarations(&mut self) {
        let prev = self.prev_ns_top();
        for index in (prev + 1..self.namespaces.len()).rev() {
            if self.namespaces[index].kind == NamespaceKind::NeedToWrite {
                let (prefix, uri) = (self.namespaces[index].prefix.clone(), self.namespaces[index].uri.clone());
                self.write_declaration(&prefix, &uri);
            }
        }
    }

    fn start_element_content(&mut self) {
        if self.start_tag_open {
            self.write_pending_declarations();
            self.out.push('>');
            self.start_tag_open = false;
        }
    }

    fn write_start_element(&mut self, prefix: Option<String>, local: &str, ns: &str) -> Result<(), String> {
        self.start_element_content();
        let prefix = prefix.or_else(|| self.lookup_prefix(ns)).unwrap_or_default();
        if !prefix.is_empty() && ns.is_empty() {
            return Err("Cannot use a prefix with an empty namespace.".to_string());
        }
        self.out.push('<');
        if !prefix.is_empty() {
            self.out.push_str(&prefix);
            self.out.push(':');
        }
        self.out.push_str(local);
        self.elements.push((self.top(), prefix.clone(), local.to_string()));
        self.push_namespace_implicit(&prefix, ns)?;
        self.start_tag_open = true;
        self.attributes.clear();
        Ok(())
    }

    /// WriteAttributeString(prefix, localName, ns, value).
    fn write_attribute_string(&mut self, prefix: Option<String>, local: &str, ns: &str, value: &str) -> Result<(), String> {
        let mut prefix = match prefix {
            Some(prefix) => prefix,
            None => {
                if local == "xmlns" && ns == XMLNS_NS {
                    String::new()
                } else {
                    self.lookup_prefix(ns).unwrap_or_default()
                }
            }
        };
        let mut special = None;
        if prefix.is_empty() {
            if local == "xmlns" {
                if !ns.is_empty() && ns != XMLNS_NS {
                    return Err("Prefix \"xmlns\" is reserved for use by XML.".to_string());
                }
                special = Some(SpecialAttribute::DefaultXmlns);
            } else if !ns.is_empty() {
                prefix = match self.lookup_prefix(ns) {
                    Some(found) if !found.is_empty() => found,
                    _ => self.generate_prefix(),
                };
            }
        } else {
            if prefix == "xmlns" {
                if !ns.is_empty() && ns != XMLNS_NS {
                    return Err("Prefix \"xmlns\" is reserved for use by XML.".to_string());
                }
                special = Some(SpecialAttribute::PrefixedXmlns(local.to_string()));
            } else if prefix == "xml" {
                if !ns.is_empty() && ns != XML_NS {
                    return Err("Prefix \"xml\" is reserved for use by XML.".to_string());
                }
                special = match local {
                    "space" => Some(SpecialAttribute::XmlSpace),
                    "lang" => Some(SpecialAttribute::XmlLang),
                    _ => None,
                };
            }
            if special.is_none() {
                if ns.is_empty() {
                    prefix = String::new();
                } else if let Some(defined) = self.lookup_local_namespace(&prefix) {
                    if defined != ns {
                        prefix = self.generate_prefix();
                    }
                }
            }
        }
        if special.is_none() && !prefix.is_empty() {
            self.push_namespace_implicit(&prefix, ns)?;
        }
        // AddAttribute: duplicate detection.
        if self
            .attributes
            .iter()
            .any(|(other_prefix, other_local, other_ns)| other_local == local && (*other_prefix == prefix || other_ns == ns))
        {
            return Err(format!("'{local}' is a duplicate attribute name."));
        }
        self.attributes.push((prefix.clone(), local.to_string(), ns.to_string()));
        match special {
            None => {
                self.out.push(' ');
                if !prefix.is_empty() {
                    self.out.push_str(&prefix);
                    self.out.push(':');
                }
                self.out.push_str(local);
                self.out.push_str("=\"");
                escape_attribute_into(value, &mut self.out);
                self.out.push('"');
            }
            Some(SpecialAttribute::DefaultXmlns) => {
                if self.push_namespace_explicit("", value)? {
                    self.write_declaration("", value);
                }
            }
            Some(SpecialAttribute::PrefixedXmlns(declared)) => {
                if value.is_empty() {
                    return Err("Cannot use a prefix with an empty namespace.".to_string());
                }
                if value == XMLNS_NS || (value == XML_NS && declared != "xml") {
                    return Err("Cannot bind to the reserved namespace.".to_string());
                }
                if self.push_namespace_explicit(&declared, value)? {
                    self.write_declaration(&declared, value);
                }
            }
            Some(SpecialAttribute::XmlSpace) => {
                let trimmed = value.trim_matches(is_xml_whitespace);
                if trimmed != "default" && trimmed != "preserve" {
                    return Err(format!("'{trimmed}' is an invalid xml:space value."));
                }
                self.out.push_str(" xml:space=\"");
                escape_attribute_into(trimmed, &mut self.out);
                self.out.push('"');
            }
            Some(SpecialAttribute::XmlLang) => {
                self.out.push_str(" xml:lang=\"");
                escape_attribute_into(value, &mut self.out);
                self.out.push('"');
            }
        }
        Ok(())
    }

    fn write_string(&mut self, text: &str) {
        self.start_element_content();
        escape_text_into(text, &mut self.out);
    }

    /// WriteEndElement: the shortcut " />" when nothing followed the start tag.
    fn write_end_element(&mut self) {
        let shortcut = self.start_tag_open;
        if shortcut {
            // StartElementContent wrote the pending declarations and '>', which
            // the raw writer overwrites with " />".
            self.write_pending_declarations();
            self.out.push_str(" />");
            self.start_tag_open = false;
        }
        let (prev, prefix, local) = self.elements.pop().expect("open element");
        if !shortcut {
            self.write_end_tag(&prefix, &local);
        }
        self.namespaces.truncate(prev + 1);
    }

    fn write_full_end_element(&mut self) {
        self.start_element_content();
        let (prev, prefix, local) = self.elements.pop().expect("open element");
        self.write_end_tag(&prefix, &local);
        self.namespaces.truncate(prev + 1);
    }

    fn write_end_tag(&mut self, prefix: &str, local: &str) {
        self.out.push_str("</");
        if !prefix.is_empty() {
            self.out.push_str(prefix);
            self.out.push(':');
        }
        self.out.push_str(local);
        self.out.push('>');
    }
}

/// XLinq's ElementWriter.NamespaceResolver: in-scope namespace declarations
/// taken from the xmlns attributes of the tree, newest last.
#[derive(Default)]
struct NamespaceResolver {
    declarations: Vec<(String, String, usize)>,
    scope: usize,
}

impl NamespaceResolver {
    fn push_scope(&mut self) {
        self.scope += 1;
    }

    fn pop_scope(&mut self) {
        while self.declarations.last().is_some_and(|&(_, _, scope)| scope == self.scope) {
            self.declarations.pop();
        }
        self.scope -= 1;
    }

    fn add(&mut self, prefix: String, ns: String) {
        self.declarations.push((prefix, ns, self.scope));
    }

    fn add_first(&mut self, prefix: String, ns: String) {
        self.declarations.insert(0, (prefix, ns, self.scope));
    }

    /// GetPrefixOfNamespace: the newest declaration of `ns` whose prefix is
    /// not redeclared by a newer one (default namespace only for elements).
    fn get_prefix_of_namespace(&self, ns: &str, allow_default: bool) -> Option<String> {
        for index in (0..self.declarations.len()).rev() {
            let (prefix, uri, _) = &self.declarations[index];
            if uri != ns {
                continue;
            }
            let shadowed = self.declarations[index + 1..].iter().any(|(other, _, _)| other == prefix);
            if shadowed {
                continue;
            }
            if allow_default || !prefix.is_empty() {
                return Some(prefix.clone());
            }
        }
        None
    }
}

/// XAttribute.IsNamespaceDeclaration.
pub fn is_namespace_declaration(name: &Name) -> bool {
    if name.ns.is_empty() {
        name.local == "xmlns"
    } else {
        name.ns == XMLNS_NS
    }
}

fn declared_prefix(name: &Name) -> String {
    if name.ns.is_empty() { String::new() } else { name.local.clone() }
}

struct ElementWriter<'a> {
    doc: &'a Doc,
    writer: WellFormedWriter,
    resolver: NamespaceResolver,
}

impl ElementWriter<'_> {
    fn prefix_of_namespace(&self, ns: &str, allow_default: bool) -> Option<String> {
        if ns.is_empty() {
            return Some(String::new());
        }
        if let Some(prefix) = self.resolver.get_prefix_of_namespace(ns, allow_default) {
            return Some(prefix);
        }
        match ns {
            XML_NS => Some("xml".to_string()),
            XMLNS_NS => Some("xmlns".to_string()),
            _ => None,
        }
    }

    fn push_ancestors(&mut self, node: NodeId) {
        let mut current = self.doc.parent(node);
        while let Some(ancestor) = current {
            for (name, value) in self.doc.attrs(ancestor) {
                if is_namespace_declaration(name) {
                    self.resolver.add_first(declared_prefix(name), value.clone());
                }
            }
            current = self.doc.parent(ancestor);
        }
    }

    fn write_element(&mut self, node: NodeId) -> Result<(), String> {
        let Data::Element { name, attrs } = &self.doc.nodes[node].data else {
            unreachable!("write_element on a text node")
        };
        self.resolver.push_scope();
        for (attribute, value) in attrs {
            if is_namespace_declaration(attribute) {
                self.resolver.add(declared_prefix(attribute), value.clone());
            }
        }
        let prefix = self.prefix_of_namespace(&name.ns, true);
        self.writer.write_start_element(prefix, &name.local, &name.ns)?;
        for (attribute, value) in attrs {
            let prefix = self.prefix_of_namespace(&attribute.ns, false);
            let ns = if attribute.ns.is_empty() && attribute.local == "xmlns" { XMLNS_NS } else { &attribute.ns };
            self.writer.write_attribute_string(prefix, &attribute.local, ns, value)?;
        }
        let children = &self.doc.nodes[node].children;
        if children.is_empty() {
            if self.doc.nodes[node].empty_string_content {
                self.writer.write_string("");
                self.writer.write_full_end_element();
            } else {
                self.writer.write_end_element();
            }
        } else {
            for &child in children {
                match &self.doc.nodes[child].data {
                    Data::Text(text) => self.writer.write_string(text),
                    Data::Element { .. } => self.write_element(child)?,
                }
            }
            self.writer.write_full_end_element();
        }
        self.resolver.pop_scope();
        Ok(())
    }
}

/// `XElement.ToString(SaveOptions.DisableFormatting)` on Windows.
pub fn element_to_string(doc: &Doc, node: NodeId) -> Result<String, String> {
    let mut writer = ElementWriter { doc, writer: WellFormedWriter::new(), resolver: NamespaceResolver::default() };
    writer.push_ancestors(node);
    writer.write_element(node)?;
    Ok(writer.writer.out)
}

// ---- tree edits with XLinq semantics ----------------------------------------

/// XElement.Parent: the parent element (None for the root and detached nodes).
pub fn parent_element(doc: &Doc, node: NodeId) -> Option<NodeId> {
    doc.parent(node)
}

/// Whether XLinq would copy `node` when it is added somewhere: it has a
/// parent, is the document root, or is the top of `target`'s own tree.
fn needs_copy(doc: &Doc, node: NodeId, target: NodeId) -> bool {
    if doc.parent(node).is_some() || doc.root == Some(node) {
        return true;
    }
    let mut top = target;
    while let Some(parent) = doc.parent(top) {
        top = parent;
    }
    top == node
}

/// `new XElement(other)`: a detached deep copy.
pub fn copy_element(doc: &mut Doc, node: NodeId) -> NodeId {
    doc.deep_clone(node)
}

/// XContainer.Add(XNode) / `new XElement(name, node)`: appends `node`, or a
/// deep copy of it when it already belongs to a tree. Returns the node added.
pub fn add_node(doc: &mut Doc, parent: NodeId, node: NodeId) -> NodeId {
    let node = if needs_copy(doc, node, parent) { doc.deep_clone(node) } else { node };
    doc.append(parent, node);
    node
}

/// `new XElement(ns + local, nodes...)`.
pub fn new_element_with(doc: &mut Doc, ns: &str, local: &str, nodes: &[NodeId]) -> NodeId {
    let element = doc.new_element(Name::new(ns, local));
    for &node in nodes {
        add_node(doc, element, node);
    }
    element
}

/// XNode.ReplaceWith(XNode). The document root can be replaced.
pub fn replace_with(doc: &mut Doc, old: NodeId, new: NodeId) -> Result<(), String> {
    let new = if doc.parent(new).is_some() || doc.root == Some(new) { doc.deep_clone(new) } else { new };
    if doc.root == Some(old) {
        doc.detach(old);
        doc.root = Some(new);
        return Ok(());
    }
    if doc.parent(old).is_none() {
        return Err("The parent is missing.".to_string());
    }
    doc.insert_before(old, new);
    doc.detach(old);
    Ok(())
}

/// XNode.AddAfterSelf(XNode).
pub fn add_after_self(doc: &mut Doc, reference: NodeId, node: NodeId) -> Result<(), String> {
    if doc.root == Some(reference) {
        return Err("This operation would create an incorrectly structured document.".to_string());
    }
    if doc.parent(reference).is_none() {
        return Err("The parent is missing.".to_string());
    }
    let node = if doc.parent(node).is_some() || doc.root == Some(node) { doc.deep_clone(node) } else { node };
    doc.insert_after(reference, node);
    Ok(())
}

/// XNode.Remove().
pub fn remove(doc: &mut Doc, node: NodeId) -> Result<(), String> {
    if doc.parent(node).is_none() && doc.root != Some(node) {
        return Err("The parent is missing.".to_string());
    }
    doc.detach(node);
    Ok(())
}

/// XContainer.ReplaceNodes(nodes): the content snapshot is taken, every
/// child node is removed, then the nodes are added (former children are
/// parentless by then and move; nodes of other trees are copied).
pub fn replace_nodes(doc: &mut Doc, container: NodeId, nodes: &[NodeId]) {
    doc.clear_children(container);
    for &node in nodes {
        add_node(doc, container, node);
    }
}

/// XNode.ElementsAfterSelf().
pub fn elements_after_self(doc: &Doc, node: NodeId) -> Vec<NodeId> {
    let Some(parent) = doc.parent(node) else { return Vec::new() };
    let siblings = doc.child_nodes(parent);
    let index = siblings.iter().position(|&id| id == node).expect("child of its parent");
    siblings[index + 1..].iter().copied().filter(|&id| doc.is_element(id)).collect()
}

/// XContainer.Nodes() as XLinq enumerates it.
pub fn nodes(doc: &Doc, node: NodeId) -> Vec<NodeId> {
    doc.child_nodes(node).to_vec()
}

/// XContainer.HasElements.
pub fn has_elements(doc: &Doc, node: NodeId) -> bool {
    doc.child_nodes(node).iter().any(|&id| doc.is_element(id))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn roundtrip(xml: &str) -> String {
        parse(xml, &XDOCUMENT_PARSE_PRESERVE).unwrap().to_string().unwrap()
    }

    #[test]
    fn writes_like_xlinq_disable_formatting() {
        assert_eq!(
            roundtrip("<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">\n  <mi>x</mi>\n  <mrow></mrow><mrow/>\n</math>"),
            "<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"block\">\r\n  <mi>x</mi>\r\n  <mrow></mrow><mrow />\r\n</math>"
        );
        assert_eq!(
            roundtrip("<math display=\"block\" xmlns=\"http://www.w3.org/1998/Math/MathML\"><mo>&#x2212;&amp;&lt;&gt;&quot;'</mo><mi a=\"&#xA;&#9;&quot;&lt;'\">\r\r\n</mi></math>"),
            "<math display=\"block\" xmlns=\"http://www.w3.org/1998/Math/MathML\"><mo>\u{2212}&amp;&lt;&gt;\"'</mo><mi a=\"&#xA;&#x9;&quot;&lt;'\">\r\n\r\n</mi></math>"
        );
        assert_eq!(roundtrip("<m:math xmlns:m=\"urn:m\"><m:mi/><mi/></m:math>"), "<m:math xmlns:m=\"urn:m\"><m:mi /><mi /></m:math>");
        assert_eq!(roundtrip("<a xmlns=\"urn:a\"><b xmlns=\"\"/></a>"), "<a xmlns=\"urn:a\"><b xmlns=\"\" /></a>");
        assert_eq!(roundtrip("<a x='1' y=\"a\tb\nc\"/>"), "<a x=\"1\" y=\"a b c\" />");
    }

    #[test]
    fn writes_generated_namespace_declarations() {
        let mut doc = parse("<a xmlns=\"urn:a\"><b/></a>", &XDOCUMENT_PARSE_PRESERVE).unwrap();
        let root = doc.root().unwrap();
        let child = doc.doc.elements(root)[0];
        doc.doc.set_attr_ns(child, "urn:x", "y", Some("1"));
        let foreign = doc.doc.new_element(Name::new("urn:f", "f"));
        doc.doc.append(child, foreign);
        let plain = doc.doc.new_element(Name::new("", "p"));
        doc.doc.append(child, plain);
        assert_eq!(
            doc.to_string().unwrap(),
            "<a xmlns=\"urn:a\"><b p2:y=\"1\" xmlns:p2=\"urn:x\"><f xmlns=\"urn:f\" /><p xmlns=\"\" /></b></a>"
        );
        assert_eq!(element_to_string(&doc.doc, child).unwrap(), "<b p1:y=\"1\" xmlns:p1=\"urn:x\" xmlns=\"urn:a\"><f xmlns=\"urn:f\" /><p xmlns=\"\" /></b>");
    }

    #[test]
    fn rejects_what_xml_reader_rejects() {
        assert!(parse("<a>&foo;</a>", &XDOCUMENT_PARSE_PRESERVE).is_err());
        assert!(parse("<a>&#x1;</a>", &XDOCUMENT_PARSE_PRESERVE).is_err());
        assert!(parse("<a/><b/>", &XDOCUMENT_PARSE_PRESERVE).is_err());
        assert!(parse("x<a/>", &XDOCUMENT_PARSE_PRESERVE).is_err());
        assert!(parse("<a p:x=\"1\"/>", &XDOCUMENT_PARSE_PRESERVE).is_err());
        assert!(parse("<!DOCTYPE a><a/>", &PREPASS_READER).is_err());
        assert!(parse("<a x=\"1\" x=\"2\"/>", &XDOCUMENT_PARSE_PRESERVE).is_err());
        assert!(parse("\u{FEFF}<a/>", &XDOCUMENT_PARSE_PRESERVE).is_err());
        assert!(parse("<1a/>", &XDOCUMENT_PARSE_PRESERVE).is_err());
        assert!(parse("<a b:c:d=\"1\"/>", &XDOCUMENT_PARSE_PRESERVE).is_err());
        assert!(parse("<a>]]></a>", &XDOCUMENT_PARSE_PRESERVE).is_err());
        assert!(parse(" <?xml version=\"1.0\"?><a/>", &XDOCUMENT_PARSE_PRESERVE).is_err());
    }
}
