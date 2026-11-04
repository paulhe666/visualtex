//! XElement serialization exactly as .NET 8 produces it on Windows:
//! `XElement.ToString(SaveOptions.DisableFormatting)` and
//! `XElement.WriteTo(XmlWriter.Create(StringBuilder, settings))` without
//! indentation.
//!
//! Three .NET layers are mirrored, because their interplay decides every
//! prefix and namespace declaration in the output:
//!
//! * System.Xml.Linq.ElementWriter + NamespaceResolver: the prefixes XLinq
//!   knows from `xmlns` attributes of the element and its ancestors. A
//!   namespace without a known prefix is passed to the XmlWriter as `null`.
//! * System.Xml.XmlWellFormedWriter: its own namespace stack. Every element
//!   start pushes one entry for its own prefix (Implied when already in scope,
//!   NeedToWrite otherwise). An element namespace without a prefix becomes the
//!   default namespace (`xmlns="..."`); an attribute namespace without a
//!   prefix gets a generated `p{N}` prefix where N = _nsTop - 2, i.e. the
//!   number of namespace-stack entries above the three built-in ones. Pending
//!   declarations are written when the start tag is closed, newest first.
//! * System.Xml.XmlEncodedRawTextWriter: `<a />` for empty elements, text and
//!   attribute escaping, NewLineHandling.Replace with Environment.NewLine
//!   ("\r\n" on Windows, where the reference output was recorded).

use crate::office::omml_write::proj::dom::{Data, Doc, Name, NodeId, XMLNS_NS, XML_NS};

/// XmlWriterSettings.NewLineChars on Windows (Environment.NewLine).
pub const NEW_LINE: &str = "\r\n";

/// XElement.ToString(SaveOptions.DisableFormatting).
pub fn to_string_disable_formatting(doc: &Doc, element: NodeId) -> Result<String, String> {
    let mut writer = ElementWriter { resolver: NamespaceResolver::default(), xml: WellFormedWriter::new() };
    writer.write_element(doc, element)?;
    Ok(writer.xml.out)
}

/// XDocument.ToString(SaveOptions.DisableFormatting) for a document whose
/// nodes are whitespace text, the root element and whitespace text
/// (XText children of an XDocument are written with WriteWhitespace).
pub fn document_to_string_disable_formatting(
    doc: &Doc,
    root: NodeId,
    before_root: &str,
    after_root: &str,
) -> Result<String, String> {
    let mut out = String::new();
    escape_text(&mut out, before_root)?;
    out.push_str(&to_string_disable_formatting(doc, root)?);
    escape_text(&mut out, after_root)?;
    Ok(out)
}

/// XElement.WriteTo(XmlWriter) for `XmlWriter.Create(StringBuilder, new
/// XmlWriterSettings { OmitXmlDeclaration = true, ... })` (no indentation):
/// the same bytes as ToString(SaveOptions.DisableFormatting).
pub fn write_to_string(doc: &Doc, element: NodeId) -> Result<String, String> {
    to_string_disable_formatting(doc, element)
}

// ---- System.Xml.Linq.NamespaceResolver -------------------------------------

struct Declaration {
    prefix: String,
    ns: String,
    scope: usize,
}

/// NamespaceResolver: a circular list in .NET; here oldest first. `Add`
/// appends (newest), `AddFirst` inserts as the oldest, PopScope removes the
/// newest declarations of the current scope. The rover cache never changes
/// a lookup result, so it is not modelled.
#[derive(Default)]
struct NamespaceResolver {
    declarations: Vec<Declaration>,
    scope: usize,
}

impl NamespaceResolver {
    fn push_scope(&mut self) {
        self.scope += 1;
    }

    fn pop_scope(&mut self) {
        while self.declarations.last().is_some_and(|d| d.scope == self.scope) {
            self.declarations.pop();
        }
        self.scope -= 1;
    }

    fn add(&mut self, prefix: &str, ns: &str) {
        self.declarations.push(Declaration { prefix: prefix.to_string(), ns: ns.to_string(), scope: self.scope });
    }

    fn add_first(&mut self, prefix: &str, ns: &str) {
        self.declarations.insert(0, Declaration { prefix: prefix.to_string(), ns: ns.to_string(), scope: self.scope });
    }

    /// GetPrefixOfNamespace: the newest declaration of `ns` whose prefix is
    /// not redeclared by a newer one; attributes skip default declarations.
    fn get_prefix_of_namespace(&self, ns: &str, allow_default_namespace: bool) -> Option<String> {
        for index in (0..self.declarations.len()).rev() {
            let declaration = &self.declarations[index];
            if declaration.ns != ns {
                continue;
            }
            let shadowed = self.declarations[index + 1..].iter().any(|newer| newer.prefix == declaration.prefix);
            if shadowed {
                continue;
            }
            if allow_default_namespace || !declaration.prefix.is_empty() {
                return Some(declaration.prefix.clone());
            }
        }
        None
    }
}

// ---- System.Xml.Linq.ElementWriter -------------------------------------------

struct ElementWriter {
    resolver: NamespaceResolver,
    xml: WellFormedWriter,
}

fn declaration_prefix(name: &Name) -> &str {
    if name.ns.is_empty() {
        ""
    } else {
        &name.local
    }
}

impl ElementWriter {
    /// WriteElement: iterative, like XLinq's, so deep trees need no stack.
    fn write_element(&mut self, doc: &Doc, element: NodeId) -> Result<(), String> {
        self.push_ancestors(doc, element);
        if doc.is_text(element) {
            return self.xml.write_string(doc.text(element).unwrap_or(""));
        }
        // (element, index of the next child node to write)
        let mut stack: Vec<(NodeId, usize)> = Vec::new();
        if self.write_start(doc, element)? {
            stack.push((element, 0));
        }
        while let Some(top) = stack.last_mut() {
            let (node, index) = *top;
            let children = doc.child_nodes(node);
            if index < children.len() {
                top.1 += 1;
                let child = children[index];
                match &doc.nodes[child].data {
                    Data::Text(text) => self.xml.write_string(text)?,
                    Data::Element { .. } => {
                        if self.write_start(doc, child)? {
                            stack.push((child, 0));
                        }
                    }
                }
            } else {
                stack.pop();
                // WriteFullEndElement
                self.xml.write_end_element(true)?;
                self.resolver.pop_scope();
            }
        }
        Ok(())
    }

    /// WriteStartElement (PushElement, start tag, attributes). An element
    /// without child nodes is also ended here (WriteEndElement for null
    /// content, WriteString("") + WriteFullEndElement for empty-string
    /// content); returns whether child nodes follow.
    fn write_start(&mut self, doc: &Doc, element: NodeId) -> Result<bool, String> {
        let Data::Element { name, attrs } = &doc.nodes[element].data else { unreachable!() };
        // PushElement
        self.resolver.push_scope();
        for (attribute, value) in attrs {
            if Doc::is_namespace_declaration(attribute) {
                self.resolver.add(declaration_prefix(attribute), value);
            }
        }
        let prefix = self.get_prefix_of_namespace(&name.ns, true);
        self.xml.write_start_element(prefix.as_deref(), &name.local, &name.ns)?;
        for (attribute, value) in attrs {
            let prefix = self.get_prefix_of_namespace(&attribute.ns, false);
            let ns = if attribute.ns.is_empty() && attribute.local == "xmlns" { XMLNS_NS } else { &attribute.ns };
            self.xml.write_attribute_string(prefix.as_deref(), &attribute.local, ns, value)?;
        }
        if !doc.child_nodes(element).is_empty() {
            return Ok(true);
        }
        if doc.has_empty_string_content(element) {
            self.xml.write_string("")?;
            self.xml.write_end_element(true)?;
        } else {
            self.xml.write_end_element(false)?;
        }
        self.resolver.pop_scope();
        Ok(false)
    }

    fn push_ancestors(&mut self, doc: &Doc, element: NodeId) {
        for ancestor in doc.ancestors(element) {
            for (name, value) in doc.attrs(ancestor) {
                if Doc::is_namespace_declaration(name) {
                    self.resolver.add_first(declaration_prefix(name), value);
                }
            }
        }
    }

    fn get_prefix_of_namespace(&self, ns: &str, allow_default_namespace: bool) -> Option<String> {
        if ns.is_empty() {
            return Some(String::new());
        }
        if let Some(prefix) = self.resolver.get_prefix_of_namespace(ns, allow_default_namespace) {
            return Some(prefix);
        }
        if ns == XML_NS {
            return Some("xml".to_string());
        }
        if ns == XMLNS_NS {
            return Some("xmlns".to_string());
        }
        None
    }

}

// ---- System.Xml.XmlWellFormedWriter + XmlEncodedRawTextWriter --------------

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum NamespaceKind {
    Written,
    NeedToWrite,
    Implied,
    Special,
}

struct Namespace {
    prefix: String,
    uri: String,
    kind: NamespaceKind,
}

struct ElementScope {
    prefix: String,
    local: String,
    prev_ns_top: usize,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum SpecialAttribute {
    No,
    DefaultXmlns,
    PrefixedXmlns,
    XmlSpace,
    XmlLang,
}

struct WellFormedWriter {
    out: String,
    /// _nsStack[0..=_nsTop].
    ns_stack: Vec<Namespace>,
    /// _elemScopeStack[0..=_elemTop]; [0] is the document level.
    elem_stack: Vec<ElementScope>,
    /// The current start tag is still open (attributes may follow).
    start_tag_open: bool,
}

impl WellFormedWriter {
    fn new() -> Self {
        let ns_stack = vec![
            Namespace { prefix: "xmlns".into(), uri: XMLNS_NS.into(), kind: NamespaceKind::Special },
            Namespace { prefix: "xml".into(), uri: XML_NS.into(), kind: NamespaceKind::Special },
            Namespace { prefix: String::new(), uri: String::new(), kind: NamespaceKind::Implied },
        ];
        let elem_stack = vec![ElementScope { prefix: String::new(), local: String::new(), prev_ns_top: 2 }];
        Self { out: String::new(), ns_stack, elem_stack, start_tag_open: false }
    }

    fn ns_top(&self) -> usize {
        self.ns_stack.len() - 1
    }

    fn current_prev_ns_top(&self) -> usize {
        self.elem_stack.last().unwrap().prev_ns_top
    }

    /// LookupPrefix: the prefix of the newest entry for `uri`, or None when
    /// that prefix was redeclared later (no further search).
    fn lookup_prefix(&self, uri: &str) -> Option<String> {
        for index in (0..self.ns_stack.len()).rev() {
            if self.ns_stack[index].uri == uri {
                let prefix = &self.ns_stack[index].prefix;
                if self.ns_stack[index + 1..].iter().any(|later| &later.prefix == prefix) {
                    return None;
                }
                return Some(prefix.clone());
            }
        }
        None
    }

    fn lookup_namespace_index(&self, prefix: &str) -> Option<usize> {
        (0..self.ns_stack.len()).rev().find(|&index| self.ns_stack[index].prefix == prefix)
    }

    fn lookup_namespace(&self, prefix: &str) -> Option<&str> {
        self.lookup_namespace_index(prefix).map(|index| self.ns_stack[index].uri.as_str())
    }

    fn lookup_local_namespace(&self, prefix: &str) -> Option<&str> {
        let floor = self.current_prev_ns_top();
        (floor + 1..self.ns_stack.len())
            .rev()
            .find(|&index| self.ns_stack[index].prefix == prefix)
            .map(|index| self.ns_stack[index].uri.as_str())
    }

    fn generate_prefix(&self) -> String {
        let generated = format!("p{}", self.ns_top() - 2);
        if self.lookup_namespace(&generated).is_none() {
            return generated;
        }
        let mut index = 0usize;
        loop {
            let candidate = format!("{generated}{index}");
            index += 1;
            if self.lookup_namespace(&candidate).is_none() {
                return candidate;
            }
        }
    }

    fn push_namespace_implicit(&mut self, prefix: &str, uri: &str) -> Result<(), String> {
        let kind;
        match self.lookup_namespace_index(prefix) {
            Some(existing) => {
                if existing > self.current_prev_ns_top() {
                    if self.ns_stack[existing].uri != uri {
                        return Err(format!(
                            "XmlException: The prefix '{prefix}' cannot be redefined from '{}' to '{uri}' within the same start element tag.",
                            self.ns_stack[existing].uri
                        ));
                    }
                    return Ok(());
                }
                if self.ns_stack[existing].kind == NamespaceKind::Special {
                    if prefix == "xml" {
                        if uri != self.ns_stack[existing].uri {
                            return Err("ArgumentException: Prefix \"xml\" is reserved for use by XML and can be mapped only to namespace name \"http://www.w3.org/XML/1998/namespace\".".to_string());
                        }
                        kind = NamespaceKind::Implied;
                    } else {
                        return Err("ArgumentException: Prefix \"xmlns\" is reserved for use by XML.".to_string());
                    }
                } else if self.ns_stack[existing].uri == uri {
                    kind = NamespaceKind::Implied;
                } else {
                    kind = NamespaceKind::NeedToWrite;
                }
            }
            None => {
                if (uri == XML_NS && prefix != "xml") || (uri == XMLNS_NS && prefix != "xmlns") {
                    return Err(format!("ArgumentException: Cannot use a prefix with an empty namespace or bind '{prefix}' to a reserved namespace."));
                }
                kind = NamespaceKind::NeedToWrite;
            }
        }
        self.ns_stack.push(Namespace { prefix: prefix.to_string(), uri: uri.to_string(), kind });
        Ok(())
    }

    /// Returns whether the declaration is written (always, without
    /// NamespaceHandling.OmitDuplicates).
    fn push_namespace_explicit(&mut self, prefix: &str, uri: &str) -> Result<bool, String> {
        if let Some(existing) = self.lookup_namespace_index(prefix) {
            if existing > self.current_prev_ns_top() {
                if self.ns_stack[existing].uri != uri {
                    return Err(format!(
                        "XmlException: The prefix '{prefix}' cannot be redefined from '{}' to '{uri}' within the same start element tag.",
                        self.ns_stack[existing].uri
                    ));
                }
                if self.ns_stack[existing].kind == NamespaceKind::Written {
                    return Err("XmlException: Duplicate namespace declaration.".to_string());
                }
                self.ns_stack[existing].kind = NamespaceKind::Written;
                return Ok(true);
            }
        }
        if (uri == XML_NS && prefix != "xml") || (uri == XMLNS_NS && prefix != "xmlns") {
            return Err(format!("ArgumentException: Cannot bind '{prefix}' to a reserved namespace."));
        }
        if prefix == "xml" && uri != XML_NS {
            return Err("ArgumentException: Prefix \"xml\" is reserved for use by XML.".to_string());
        }
        if prefix == "xmlns" {
            return Err("ArgumentException: Prefix \"xmlns\" is reserved for use by XML.".to_string());
        }
        self.ns_stack.push(Namespace { prefix: prefix.to_string(), uri: uri.to_string(), kind: NamespaceKind::Written });
        Ok(true)
    }

    /// StartElementContent: pending declarations (newest first), then '>'.
    fn start_element_content(&mut self) -> Result<(), String> {
        let floor = self.current_prev_ns_top();
        for index in (floor + 1..self.ns_stack.len()).rev() {
            if self.ns_stack[index].kind == NamespaceKind::NeedToWrite {
                let prefix = self.ns_stack[index].prefix.clone();
                let uri = self.ns_stack[index].uri.clone();
                self.write_namespace_declaration(&prefix, &uri)?;
            }
        }
        self.out.push('>');
        self.start_tag_open = false;
        Ok(())
    }

    fn write_namespace_declaration(&mut self, prefix: &str, uri: &str) -> Result<(), String> {
        if prefix.is_empty() {
            self.out.push_str(" xmlns=\"");
        } else {
            self.out.push_str(" xmlns:");
            self.out.push_str(prefix);
            self.out.push_str("=\"");
        }
        escape_attribute(&mut self.out, uri)?;
        self.out.push('"');
        Ok(())
    }

    fn close_start_tag(&mut self) -> Result<(), String> {
        if self.start_tag_open {
            self.start_element_content()?;
        }
        Ok(())
    }

    fn write_start_element(&mut self, prefix: Option<&str>, local: &str, ns: &str) -> Result<(), String> {
        self.close_start_tag()?;
        let prefix = match prefix {
            None => self.lookup_prefix(ns).unwrap_or_default(),
            Some(prefix) => {
                if !prefix.is_empty() && ns.is_empty() {
                    return Err("ArgumentException: Cannot use a prefix with an empty namespace.".to_string());
                }
                prefix.to_string()
            }
        };
        self.out.push('<');
        if !prefix.is_empty() {
            self.out.push_str(&prefix);
            self.out.push(':');
        }
        self.out.push_str(local);
        let prev_ns_top = self.ns_top();
        self.elem_stack.push(ElementScope { prefix: prefix.clone(), local: local.to_string(), prev_ns_top });
        self.push_namespace_implicit(&prefix, ns)?;
        self.start_tag_open = true;
        Ok(())
    }

    /// WriteAttributeString(prefix, localName, ns, value) =
    /// WriteStartAttribute + WriteString + WriteEndAttribute.
    fn write_attribute_string(&mut self, prefix: Option<&str>, local: &str, ns: &str, value: &str) -> Result<(), String> {
        if !self.start_tag_open {
            return Err("InvalidOperationException: Token StartAttribute in state Content would result in an invalid XML document.".to_string());
        }
        let mut prefix: Option<String> = prefix.map(str::to_string);
        let mut local = local.to_string();
        if local.is_empty() {
            if prefix.as_deref() == Some("xmlns") {
                local = "xmlns".to_string();
                prefix = Some(String::new());
            } else {
                return Err("ArgumentException: The empty string '' is not a valid local name.".to_string());
            }
        }
        let namespace_name = ns.to_string();
        let mut prefix = match prefix {
            Some(prefix) => prefix,
            None => {
                let mut found = None;
                if !(local == "xmlns" && namespace_name == XMLNS_NS) {
                    found = self.lookup_prefix(&namespace_name);
                }
                found.unwrap_or_default()
            }
        };
        let mut special = SpecialAttribute::No;
        let mut declared_prefix = String::new();
        if prefix.is_empty() {
            if local == "xmlns" {
                if !namespace_name.is_empty() && namespace_name != XMLNS_NS {
                    return Err("ArgumentException: Prefix \"xmlns\" is reserved for use by XML.".to_string());
                }
                special = SpecialAttribute::DefaultXmlns;
            } else if !namespace_name.is_empty() {
                prefix = match self.lookup_prefix(&namespace_name) {
                    Some(found) if !found.is_empty() => found,
                    _ => self.generate_prefix(),
                };
            }
        } else {
            if prefix == "xmlns" {
                if !namespace_name.is_empty() && namespace_name != XMLNS_NS {
                    return Err("ArgumentException: Prefix \"xmlns\" is reserved for use by XML.".to_string());
                }
                declared_prefix = local.clone();
                special = SpecialAttribute::PrefixedXmlns;
            } else if prefix == "xml" {
                if !namespace_name.is_empty() && namespace_name != XML_NS {
                    return Err("ArgumentException: Prefix \"xml\" is reserved for use by XML.".to_string());
                }
                match local.as_str() {
                    "space" => special = SpecialAttribute::XmlSpace,
                    "lang" => special = SpecialAttribute::XmlLang,
                    _ => {}
                }
            }
            if special == SpecialAttribute::No {
                if namespace_name.is_empty() {
                    // attributes cannot have default namespace
                    prefix = String::new();
                } else if let Some(defined) = self.lookup_local_namespace(&prefix) {
                    if defined != namespace_name {
                        prefix = self.generate_prefix();
                    }
                }
            }
        }
        if special == SpecialAttribute::No && !prefix.is_empty() {
            self.push_namespace_implicit(&prefix, &namespace_name)?;
        }
        match special {
            SpecialAttribute::No => {
                self.out.push(' ');
                if !prefix.is_empty() {
                    self.out.push_str(&prefix);
                    self.out.push(':');
                }
                self.out.push_str(&local);
                self.out.push_str("=\"");
                escape_attribute(&mut self.out, value)?;
                self.out.push('"');
            }
            SpecialAttribute::DefaultXmlns => {
                if self.push_namespace_explicit("", value)? {
                    self.write_namespace_declaration("", value)?;
                }
            }
            SpecialAttribute::PrefixedXmlns => {
                if value.is_empty() {
                    return Err("ArgumentException: Cannot use a prefix with an empty namespace.".to_string());
                }
                if value == XMLNS_NS || (value == XML_NS && declared_prefix != "xml") {
                    return Err("ArgumentException: Cannot bind to the reserved namespace.".to_string());
                }
                if self.push_namespace_explicit(&declared_prefix, value)? {
                    self.write_namespace_declaration(&declared_prefix, value)?;
                }
            }
            SpecialAttribute::XmlSpace => {
                let trimmed = value.trim_matches(|c| matches!(c, ' ' | '\t' | '\r' | '\n'));
                if trimmed != "default" && trimmed != "preserve" {
                    return Err(format!("ArgumentException: '{trimmed}' is an invalid xml:space value."));
                }
                self.out.push_str(" xml:space=\"");
                escape_attribute(&mut self.out, trimmed)?;
                self.out.push('"');
            }
            SpecialAttribute::XmlLang => {
                self.out.push_str(" xml:lang=\"");
                escape_attribute(&mut self.out, value)?;
                self.out.push('"');
            }
        }
        Ok(())
    }

    fn write_string(&mut self, text: &str) -> Result<(), String> {
        self.close_start_tag()?;
        escape_text(&mut self.out, text)
    }

    /// WriteEndElement (`full` = false: `<a />` when nothing was written
    /// since the start tag) or WriteFullEndElement (`full` = true).
    fn write_end_element(&mut self, full: bool) -> Result<(), String> {
        let scope = self.elem_stack.pop().ok_or_else(|| "XmlException: There was no XML start tag open.".to_string())?;
        // Keep the scope visible to start_element_content while closing.
        let prev_ns_top = scope.prev_ns_top;
        if self.start_tag_open {
            self.elem_stack.push(scope);
            self.start_element_content()?;
            let scope = self.elem_stack.pop().unwrap();
            if full {
                self.write_end_tag(&scope);
            } else {
                // overwrite the '>' with the shortcut syntax " />"
                self.out.pop();
                self.out.push_str(" />");
            }
        } else {
            self.write_end_tag(&scope);
        }
        self.ns_stack.truncate(prev_ns_top + 1);
        Ok(())
    }

    fn write_end_tag(&mut self, scope: &ElementScope) {
        self.out.push_str("</");
        if !scope.prefix.is_empty() {
            self.out.push_str(&scope.prefix);
            self.out.push(':');
        }
        self.out.push_str(&scope.local);
        self.out.push('>');
    }
}

fn invalid_character(c: char) -> String {
    format!("ArgumentException: '{}', hexadecimal value 0x{:02X}, is an invalid character.", c.escape_unicode(), c as u32)
}

/// XmlEncodedRawTextWriter.WriteElementTextBlock with NewLineHandling.Replace
/// and CheckCharacters = true.
pub fn escape_text(out: &mut String, text: &str) -> Result<(), String> {
    let mut chars = text.chars().peekable();
    while let Some(c) = chars.next() {
        match c {
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '&' => out.push_str("&amp;"),
            '\n' => out.push_str(NEW_LINE),
            '\r' => {
                if chars.peek() == Some(&'\n') {
                    chars.next();
                }
                out.push_str(NEW_LINE);
            }
            '\t' => out.push('\t'),
            c if (c as u32) < 0x20 || c == '\u{FFFE}' || c == '\u{FFFF}' => return Err(invalid_character(c)),
            c => out.push(c),
        }
    }
    Ok(())
}

/// XmlEncodedRawTextWriter.WriteAttributeTextBlock (NewLineHandling.Replace).
pub fn escape_attribute(out: &mut String, value: &str) -> Result<(), String> {
    for c in value.chars() {
        match c {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\t' => out.push_str("&#x9;"),
            '\r' => out.push_str("&#xD;"),
            '\n' => out.push_str("&#xA;"),
            c if (c as u32) < 0x20 || c == '\u{FFFE}' || c == '\u{FFFF}' => return Err(invalid_character(c)),
            c => out.push(c),
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::office::omml_write::proj::dom::{Doc, Name, OMML_NS};
    use crate::office::omml_write::proj::xml_reader::{parse, PRESERVE_WHITESPACE};

    fn round_trip(xml: &str) -> String {
        let doc = parse(xml, PRESERVE_WHITESPACE).unwrap();
        to_string_disable_formatting(&doc, doc.root.unwrap()).unwrap()
    }

    #[test]
    fn escaping_follows_xml_encoded_raw_text_writer() {
        let mut doc = Doc::new();
        let root = doc.new_element(Name::local("a"));
        doc.set_attr(root, "v", "<>&\"'\t\n\r");
        doc.append_text(root, "<>&\"'\t|\r\n|\r|\n");
        assert_eq!(
            to_string_disable_formatting(&doc, root).unwrap(),
            "<a v=\"&lt;&gt;&amp;&quot;'&#x9;&#xA;&#xD;\">&lt;&gt;&amp;\"'\t|\r\n|\r\n|\r\n</a>"
        );
    }

    #[test]
    fn empty_string_content_survives_a_round_trip() {
        assert_eq!(round_trip("<a><b></b><c/><d> </d></a>"), "<a><b></b><c /><d> </d></a>");
    }

    #[test]
    fn unprefixed_namespaces_become_default_and_generated_prefixes() {
        // Elements built without xmlns attributes: default namespace on the
        // root, attributes get p{_nsTop - 2}.
        let mut doc = Doc::new();
        let root = doc.new_element(Name::new(OMML_NS, "oMath"));
        let child = doc.new_element(Name::new(OMML_NS, "maxDist"));
        doc.set_attr_ns(child, OMML_NS, "val", Some("1"));
        doc.append(root, child);
        assert_eq!(
            to_string_disable_formatting(&doc, root).unwrap(),
            format!("<oMath xmlns=\"{OMML_NS}\"><maxDist p2:val=\"1\" xmlns:p2=\"{OMML_NS}\" /></oMath>")
        );
        // The root's xmlns:p3 raises _nsTop, so maxDist's attribute asks for
        // p3, finds it bound and takes p30. XmlWellFormedWriter.LookupPrefix
        // then hands that generated p30 (the newest entry for the namespace)
        // to the child element itself: generated prefixes leak downwards.
        doc.set_attr_ns(root, XMLNS_NS, "p3", Some("urn:x"));
        let inner = doc.new_element(Name::new(OMML_NS, "x"));
        doc.set_attr_ns(inner, OMML_NS, "val", Some("2"));
        doc.append(child, inner);
        assert_eq!(
            to_string_disable_formatting(&doc, root).unwrap(),
            format!(
                "<oMath xmlns:p3=\"urn:x\" xmlns=\"{OMML_NS}\"><maxDist p30:val=\"1\" xmlns:p30=\"{OMML_NS}\">\
                 <p30:x p30:val=\"2\" /></maxDist></oMath>"
            )
        );
    }

    #[test]
    fn declared_prefixes_are_reused_and_nested_elements_redeclare() {
        let xml = format!(
            "<m:oMath xmlns:m=\"{OMML_NS}\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
             <m:r><w:rPr><w:color w:val=\"FF0000\" /></w:rPr><m:t xml:space=\"preserve\"> a </m:t></m:r></m:oMath>"
        );
        assert_eq!(round_trip(&xml), xml);
        // Serializing a nested element declares what its ancestors declared.
        let doc = parse(&xml, PRESERVE_WHITESPACE).unwrap();
        let run = doc.elements(doc.root.unwrap())[0];
        assert_eq!(
            to_string_disable_formatting(&doc, run).unwrap(),
            format!(
                "<m:r xmlns:m=\"{OMML_NS}\"><w:rPr xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">\
                 <w:color w:val=\"FF0000\" /></w:rPr><m:t xml:space=\"preserve\"> a </m:t></m:r>"
            )
        );
    }
}
