//! Minimal mutable XML tree for the OMML -> MathML -> LaTeX port.
//!
//! It mirrors the subset of System.Xml.Linq that the Windows converters use:
//! namespaced element/attribute names, text nodes (adjacent text merged, as
//! XDocument does), parent links, pre-order descendants and in-place edits.
//! Nodes live in an arena; detached nodes simply stay unreachable.

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

pub const MATHML_NS: &str = "http://www.w3.org/1998/Math/MathML";
pub const OMML_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
pub const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";
/// [writer port] XNamespace.Xmlns: the namespace of `xmlns:p` declaration attributes.
pub const XMLNS_NS: &str = "http://www.w3.org/2000/xmlns/";

pub type NodeId = usize;

#[derive(Clone, Debug, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Name {
    pub ns: String,
    pub local: String,
}

impl Name {
    pub fn new(ns: &str, local: &str) -> Self {
        Self { ns: ns.to_string(), local: local.to_string() }
    }
    pub fn local(local: &str) -> Self {
        Self::new("", local)
    }
    /// System.Xml.Linq's XName.ToString(): "{ns}local", or "local" without a namespace.
    pub fn expanded(&self) -> String {
        if self.ns.is_empty() { self.local.clone() } else { format!("{{{}}}{}", self.ns, self.local) }
    }
}

#[derive(Clone, Debug)]
pub enum Data {
    Element { name: Name, attrs: Vec<(Name, String)> },
    Text(String),
}

#[derive(Clone, Debug)]
pub struct Node {
    pub data: Data,
    pub parent: Option<NodeId>,
    pub children: Vec<NodeId>,
}

#[derive(Clone, Debug, Default)]
pub struct Doc {
    pub nodes: Vec<Node>,
    pub root: Option<NodeId>,
    /// [writer port] Elements whose XLinq content is the empty string rather
    /// than null: a parsed `<a></a>`, `new XElement(name, "")` or `Value = ""`.
    /// XLinq writes them as `<a></a>` instead of `<a />`. The mark is cleared
    /// when a node is added and kept by deep copies (XContainer(XContainer)).
    pub empty_string_content: std::collections::HashSet<NodeId>,
}

#[derive(Clone, Copy, Debug)]
pub struct ParseOptions {
    /// Keep whitespace-only text between elements. When false, whitespace-only
    /// text is dropped unless it is inside xml:space="preserve" (XmlReader with
    /// IgnoreWhitespace = true keeps such SignificantWhitespace).
    pub keep_whitespace: bool,
    pub max_bytes: usize,
}

impl Default for ParseOptions {
    fn default() -> Self {
        Self { keep_whitespace: true, max_bytes: 16 * 1024 * 1024 }
    }
}

impl Doc {
    pub fn new() -> Self {
        Self::default()
    }

    // ---- construction -----------------------------------------------------

    pub fn new_element(&mut self, name: Name) -> NodeId {
        self.nodes.push(Node { data: Data::Element { name, attrs: Vec::new() }, parent: None, children: Vec::new() });
        self.nodes.len() - 1
    }

    pub fn new_text(&mut self, text: &str) -> NodeId {
        self.nodes.push(Node { data: Data::Text(text.to_string()), parent: None, children: Vec::new() });
        self.nodes.len() - 1
    }

    /// Detaches `child` from its current parent and appends it to `parent`.
    /// Appending text next to a text node merges them, as XContainer.Add does.
    pub fn append(&mut self, parent: NodeId, child: NodeId) {
        self.detach(child);
        self.empty_string_content.remove(&parent);
        if let Data::Text(added) = self.nodes[child].data.clone() {
            if let Some(&last) = self.nodes[parent].children.last() {
                if let Data::Text(existing) = &mut self.nodes[last].data {
                    existing.push_str(&added);
                    return;
                }
            }
        }
        self.nodes[child].parent = Some(parent);
        self.nodes[parent].children.push(child);
    }

    pub fn append_text(&mut self, parent: NodeId, text: &str) {
        if text.is_empty() {
            return;
        }
        let id = self.new_text(text);
        self.append(parent, id);
    }

    /// Inserts `node` before `reference` (XNode.AddBeforeSelf).
    pub fn insert_before(&mut self, reference: NodeId, node: NodeId) {
        self.detach(node);
        let parent = self.nodes[reference].parent.expect("reference has no parent");
        let index = self.index_in_parent(reference).unwrap();
        self.empty_string_content.remove(&parent);
        self.nodes[node].parent = Some(parent);
        self.nodes[parent].children.insert(index, node);
    }

    /// Inserts `node` after `reference` (XNode.AddAfterSelf).
    pub fn insert_after(&mut self, reference: NodeId, node: NodeId) {
        self.detach(node);
        let parent = self.nodes[reference].parent.expect("reference has no parent");
        let index = self.index_in_parent(reference).unwrap();
        self.empty_string_content.remove(&parent);
        self.nodes[node].parent = Some(parent);
        self.nodes[parent].children.insert(index + 1, node);
    }

    /// Replaces `old` in its parent with `new` (XNode.ReplaceWith).
    pub fn replace(&mut self, old: NodeId, new: NodeId) {
        self.insert_before(old, new);
        self.detach(old);
    }

    /// Removes the node from its parent (XNode.Remove).
    pub fn detach(&mut self, node: NodeId) {
        if let Some(parent) = self.nodes[node].parent.take() {
            self.nodes[parent].children.retain(|&id| id != node);
        }
        if self.root == Some(node) {
            self.root = None;
        }
    }

    /// Removes all child nodes (XContainer.RemoveNodes).
    pub fn clear_children(&mut self, node: NodeId) {
        self.empty_string_content.remove(&node);
        for child in std::mem::take(&mut self.nodes[node].children) {
            self.nodes[child].parent = None;
        }
    }

    /// Deep copy inside this document; the copy is detached.
    pub fn deep_clone(&mut self, node: NodeId) -> NodeId {
        let data = self.nodes[node].data.clone();
        let copy = match data {
            Data::Element { name, attrs } => {
                let id = self.new_element(name);
                if let Data::Element { attrs: target, .. } = &mut self.nodes[id].data {
                    *target = attrs;
                }
                id
            }
            Data::Text(text) => self.new_text(&text),
        };
        if self.empty_string_content.contains(&node) {
            self.empty_string_content.insert(copy);
        }
        for child in self.nodes[node].children.clone() {
            let child_copy = self.deep_clone(child);
            self.nodes[child_copy].parent = Some(copy);
            self.nodes[copy].children.push(child_copy);
        }
        copy
    }

    /// Deep copy of a node of another document into this one; the copy is detached.
    pub fn import(&mut self, other: &Doc, node: NodeId) -> NodeId {
        let copy = match &other.nodes[node].data {
            Data::Element { name, attrs } => {
                let id = self.new_element(name.clone());
                if let Data::Element { attrs: target, .. } = &mut self.nodes[id].data {
                    *target = attrs.clone();
                }
                id
            }
            Data::Text(text) => self.new_text(text),
        };
        if other.empty_string_content.contains(&node) {
            self.empty_string_content.insert(copy);
        }
        for &child in &other.nodes[node].children {
            let child_copy = self.import(other, child);
            self.nodes[child_copy].parent = Some(copy);
            self.nodes[copy].children.push(child_copy);
        }
        copy
    }

    // ---- inspection -------------------------------------------------------

    pub fn is_element(&self, node: NodeId) -> bool {
        matches!(self.nodes[node].data, Data::Element { .. })
    }

    pub fn is_text(&self, node: NodeId) -> bool {
        matches!(self.nodes[node].data, Data::Text(_))
    }

    pub fn name(&self, node: NodeId) -> Option<&Name> {
        match &self.nodes[node].data {
            Data::Element { name, .. } => Some(name),
            Data::Text(_) => None,
        }
    }

    /// Local name of an element ("" for text).
    pub fn local(&self, node: NodeId) -> &str {
        self.name(node).map(|name| name.local.as_str()).unwrap_or("")
    }

    pub fn is(&self, node: NodeId, ns: &str, local: &str) -> bool {
        self.name(node).is_some_and(|name| name.ns == ns && name.local == local)
    }

    pub fn set_name(&mut self, node: NodeId, new_name: Name) {
        if let Data::Element { name, .. } = &mut self.nodes[node].data {
            *name = new_name;
        }
    }

    pub fn text(&self, node: NodeId) -> Option<&str> {
        match &self.nodes[node].data {
            Data::Text(text) => Some(text),
            Data::Element { .. } => None,
        }
    }

    pub fn set_text(&mut self, node: NodeId, value: &str) {
        if let Data::Text(text) = &mut self.nodes[node].data {
            *text = value.to_string();
        }
    }

    /// XElement.Value: the concatenated text of all descendants.
    pub fn value(&self, node: NodeId) -> String {
        let mut result = String::new();
        self.collect_text(node, &mut result);
        result
    }

    fn collect_text(&self, node: NodeId, out: &mut String) {
        match &self.nodes[node].data {
            Data::Text(text) => out.push_str(text),
            Data::Element { .. } => {
                for &child in &self.nodes[node].children {
                    self.collect_text(child, out);
                }
            }
        }
    }

    /// XElement.Value = text: replaces all children with one text node.
    /// [writer port] `Value = ""` leaves empty-string content (`<a></a>`).
    pub fn set_value(&mut self, node: NodeId, value: &str) {
        self.clear_children(node);
        self.add_string(node, value);
    }

    pub fn attrs(&self, node: NodeId) -> &[(Name, String)] {
        match &self.nodes[node].data {
            Data::Element { attrs, .. } => attrs,
            Data::Text(_) => &[],
        }
    }

    /// Attribute without a namespace ((string?)element.Attribute("x")).
    pub fn attr(&self, node: NodeId, local: &str) -> Option<&str> {
        self.attr_ns(node, "", local)
    }

    pub fn attr_ns(&self, node: NodeId, ns: &str, local: &str) -> Option<&str> {
        self.attrs(node)
            .iter()
            .find(|(name, _)| name.ns == ns && name.local == local)
            .map(|(_, value)| value.as_str())
    }

    /// SetAttributeValue: replaces in place, appends when new, removes on None.
    pub fn set_attr_ns(&mut self, node: NodeId, ns: &str, local: &str, value: Option<&str>) {
        if let Data::Element { attrs, .. } = &mut self.nodes[node].data {
            let position = attrs.iter().position(|(name, _)| name.ns == ns && name.local == local);
            match (position, value) {
                (Some(index), Some(value)) => attrs[index].1 = value.to_string(),
                (Some(index), None) => {
                    attrs.remove(index);
                }
                (None, Some(value)) => attrs.push((Name::new(ns, local), value.to_string())),
                (None, None) => {}
            }
        }
    }

    pub fn set_attr(&mut self, node: NodeId, local: &str, value: &str) {
        self.set_attr_ns(node, "", local, Some(value));
    }

    pub fn remove_attr(&mut self, node: NodeId, local: &str) {
        self.set_attr_ns(node, "", local, None);
    }

    pub fn parent(&self, node: NodeId) -> Option<NodeId> {
        self.nodes[node].parent
    }

    /// All child nodes (elements and text), XContainer.Nodes().
    pub fn child_nodes(&self, node: NodeId) -> &[NodeId] {
        &self.nodes[node].children
    }

    /// Child elements, XContainer.Elements().
    pub fn elements(&self, node: NodeId) -> Vec<NodeId> {
        self.nodes[node].children.iter().copied().filter(|&id| self.is_element(id)).collect()
    }

    /// Child elements with the given namespace and local name, XContainer.Elements(name).
    pub fn elements_named(&self, node: NodeId, ns: &str, local: &str) -> Vec<NodeId> {
        self.nodes[node].children.iter().copied().filter(|&id| self.is(id, ns, local)).collect()
    }

    /// First child element with the name, XContainer.Element(name).
    pub fn element(&self, node: NodeId, ns: &str, local: &str) -> Option<NodeId> {
        self.nodes[node].children.iter().copied().find(|&id| self.is(id, ns, local))
    }

    /// Descendant elements in document (pre-)order, excluding `node`.
    pub fn descendants(&self, node: NodeId) -> Vec<NodeId> {
        let mut out = Vec::new();
        for &child in &self.nodes[node].children {
            self.collect_descendants(child, &mut out);
        }
        out
    }

    /// Descendant elements including `node` itself (DescendantsAndSelf()).
    pub fn descendants_and_self(&self, node: NodeId) -> Vec<NodeId> {
        let mut out = Vec::new();
        self.collect_descendants(node, &mut out);
        out
    }

    fn collect_descendants(&self, node: NodeId, out: &mut Vec<NodeId>) {
        if !self.is_element(node) {
            return;
        }
        out.push(node);
        for &child in &self.nodes[node].children {
            self.collect_descendants(child, out);
        }
    }

    /// All descendant nodes, text included, in document order (DescendantNodes()).
    pub fn descendant_nodes(&self, node: NodeId) -> Vec<NodeId> {
        let mut out = Vec::new();
        for &child in &self.nodes[node].children {
            out.push(child);
            out.extend(self.descendant_nodes(child));
        }
        out
    }

    pub fn ancestors(&self, node: NodeId) -> Vec<NodeId> {
        let mut out = Vec::new();
        let mut current = self.nodes[node].parent;
        while let Some(id) = current {
            out.push(id);
            current = self.nodes[id].parent;
        }
        out
    }

    pub fn index_in_parent(&self, node: NodeId) -> Option<usize> {
        let parent = self.nodes[node].parent?;
        self.nodes[parent].children.iter().position(|&id| id == node)
    }

    /// Previous sibling element (ElementsBeforeSelf().LastOrDefault()).
    pub fn previous_element(&self, node: NodeId) -> Option<NodeId> {
        let parent = self.nodes[node].parent?;
        let siblings = &self.nodes[parent].children;
        let index = siblings.iter().position(|&id| id == node)?;
        siblings[..index].iter().rev().copied().find(|&id| self.is_element(id))
    }

    /// Next sibling element (ElementsAfterSelf().FirstOrDefault()).
    pub fn next_element(&self, node: NodeId) -> Option<NodeId> {
        let parent = self.nodes[node].parent?;
        let siblings = &self.nodes[parent].children;
        let index = siblings.iter().position(|&id| id == node)?;
        siblings[index + 1..].iter().copied().find(|&id| self.is_element(id))
    }

    pub fn previous_node(&self, node: NodeId) -> Option<NodeId> {
        let parent = self.nodes[node].parent?;
        let index = self.index_in_parent(node)?;
        index.checked_sub(1).map(|i| self.nodes[parent].children[i])
    }

    pub fn next_node(&self, node: NodeId) -> Option<NodeId> {
        let parent = self.nodes[node].parent?;
        let index = self.index_in_parent(node)?;
        self.nodes[parent].children.get(index + 1).copied()
    }

    // ---- [writer port] additional XLinq operations ------------------------

    /// XContainer.Add(string): appends text (merging with a trailing text
    /// node); an empty string turns null content into empty-string content.
    pub fn add_string(&mut self, parent: NodeId, text: &str) {
        if text.is_empty() {
            if self.nodes[parent].children.is_empty() && self.is_element(parent) {
                self.empty_string_content.insert(parent);
            }
            return;
        }
        self.append_text(parent, text);
    }

    /// Whether the element's content is the empty string (`<a></a>`).
    pub fn has_empty_string_content(&self, node: NodeId) -> bool {
        self.empty_string_content.contains(&node)
    }

    /// XContainer.AddFirst(node): detaches `child` and inserts it as the first child.
    pub fn add_first(&mut self, parent: NodeId, child: NodeId) {
        self.detach(child);
        self.empty_string_content.remove(&parent);
        self.nodes[child].parent = Some(parent);
        self.nodes[parent].children.insert(0, child);
    }

    /// XElement.HasElements.
    pub fn has_elements(&self, node: NodeId) -> bool {
        self.nodes[node].children.iter().any(|&id| self.is_element(id))
    }

    /// XNode.ElementsBeforeSelf(): preceding sibling elements in document order.
    pub fn elements_before_self(&self, node: NodeId) -> Vec<NodeId> {
        let Some(parent) = self.nodes[node].parent else { return Vec::new() };
        let siblings = &self.nodes[parent].children;
        let index = siblings.iter().position(|&id| id == node).unwrap_or(0);
        siblings[..index].iter().copied().filter(|&id| self.is_element(id)).collect()
    }

    /// XNode.ElementsAfterSelf(): following sibling elements in document order.
    pub fn elements_after_self(&self, node: NodeId) -> Vec<NodeId> {
        let Some(parent) = self.nodes[node].parent else { return Vec::new() };
        let siblings = &self.nodes[parent].children;
        let Some(index) = siblings.iter().position(|&id| id == node) else { return Vec::new() };
        siblings[index + 1..].iter().copied().filter(|&id| self.is_element(id)).collect()
    }

    /// Descendants(name): descendant elements with the name, document order.
    pub fn descendants_named(&self, node: NodeId, ns: &str, local: &str) -> Vec<NodeId> {
        self.descendants(node).into_iter().filter(|&id| self.is(id, ns, local)).collect()
    }

    /// XAttribute.IsNamespaceDeclaration for an attribute name.
    pub fn is_namespace_declaration(name: &Name) -> bool {
        if name.ns.is_empty() {
            name.local == "xmlns"
        } else {
            name.ns == XMLNS_NS
        }
    }

    // ---- parsing ----------------------------------------------------------

    pub fn parse(xml: &str, options: ParseOptions) -> Result<Doc, String> {
        if xml.len() > options.max_bytes {
            return Err("XML input is too large".to_string());
        }
        if xml.contains("<!DOCTYPE") || xml.contains("<!ENTITY") {
            return Err("XML input must not contain a document type declaration".to_string());
        }
        let mut doc = Doc::new();
        let mut reader = NsReader::from_str(xml);
        reader.config_mut().trim_text(false);
        reader.config_mut().expand_empty_elements = false;
        let mut stack: Vec<NodeId> = Vec::new();
        // Per open element: whether xml:space="preserve" is in effect.
        let mut preserve: Vec<bool> = Vec::new();
        loop {
            let (resolve, event) = reader
                .read_resolved_event()
                .map_err(|error| format!("Unable to parse XML: {error}"))?;
            let element_ns = match resolve {
                ResolveResult::Bound(namespace) => Ok(String::from_utf8_lossy(namespace.as_ref()).into_owned()),
                ResolveResult::Unbound => Ok(String::new()),
                ResolveResult::Unknown(prefix) => Err(format!(
                    "Unknown XML namespace prefix {}",
                    String::from_utf8_lossy(&prefix)
                )),
            };
            let event = event.into_owned();
            match event {
                Event::Start(start) => {
                    let id = open_element(&mut doc, &reader, element_ns?, &start, &stack)?;
                    let inherited = preserve.last().copied().unwrap_or(false);
                    let own = doc.attr_ns(id, XML_NS, "space").map(|value| value == "preserve");
                    stack.push(id);
                    preserve.push(own.unwrap_or(inherited));
                }
                Event::Empty(start) => {
                    open_element(&mut doc, &reader, element_ns?, &start, &stack)?;
                }
                Event::Text(text) => {
                    let decoded = text.decode().map_err(|error| error.to_string())?.into_owned();
                    push_text(&mut doc, &stack, &decoded);
                }
                Event::GeneralRef(reference) => {
                    let raw = format!("&{};", reference.decode().map_err(|error| error.to_string())?);
                    let value = quick_xml::escape::unescape(&raw).map_err(|error| error.to_string())?;
                    push_text(&mut doc, &stack, &value);
                }
                Event::CData(text) => {
                    let decoded = text.decode().map_err(|error| error.to_string())?.into_owned();
                    push_text(&mut doc, &stack, &decoded);
                }
                Event::End(_) => {
                    let id = stack.pop().ok_or_else(|| "XML contains an unexpected closing tag".to_string())?;
                    let keep = preserve.pop().unwrap_or(false);
                    if !options.keep_whitespace && !keep {
                        drop_whitespace_text(&mut doc, id);
                    }
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

    // ---- serialization ----------------------------------------------------

    /// Serializes like XElement.ToString(SaveOptions.DisableFormatting) for the
    /// documents this port builds: a default namespace on the root (and where
    /// it changes), prefixes m/w/xml for namespaced attributes and elements of
    /// a different namespace, empty elements as "<name />".
    pub fn to_xml(&self, node: NodeId) -> String {
        let mut out = String::new();
        self.write_xml(node, None, true, &mut out);
        out
    }

    fn write_xml(&self, node: NodeId, default_ns: Option<&str>, is_root: bool, out: &mut String) {
        match &self.nodes[node].data {
            Data::Text(text) => out.push_str(&escape_text(text)),
            Data::Element { name, attrs } => {
                let (tag, declare_default) = match prefix_for(&name.ns) {
                    Some(prefix) if name.ns != default_ns.unwrap_or("") => (format!("{prefix}:{}", name.local), false),
                    _ => (name.local.clone(), default_ns != Some(name.ns.as_str())),
                };
                out.push('<');
                out.push_str(&tag);
                if declare_default && !name.ns.is_empty() {
                    out.push_str(&format!(" xmlns=\"{}\"", escape_attr(&name.ns)));
                }
                if is_root {
                    let root_prefixed = !declare_default && prefix_for(&name.ns).is_some();
                    let mut prefixes: Vec<(&str, &str)> = Vec::new();
                    for id in self.descendants_and_self(node) {
                        let element_name = self.name(id).unwrap();
                        let mut namespaces = vec![element_name.ns.as_str()];
                        namespaces.extend(self.attrs(id).iter().map(|(attribute, _)| attribute.ns.as_str()));
                        for ns in namespaces {
                            if ns.is_empty() || ns == XML_NS || (ns == name.ns && !root_prefixed) {
                                continue;
                            }
                            if let Some(prefix) = prefix_for(ns) {
                                if !prefixes.iter().any(|(p, _)| *p == prefix) {
                                    prefixes.push((prefix, ns));
                                }
                            }
                        }
                    }
                    for (prefix, ns) in prefixes {
                        out.push_str(&format!(" xmlns:{prefix}=\"{}\"", escape_attr(ns)));
                    }
                }
                for (attribute, value) in attrs {
                    out.push(' ');
                    if attribute.ns.is_empty() {
                        out.push_str(&attribute.local);
                    } else {
                        out.push_str(prefix_for(&attribute.ns).unwrap_or("ns"));
                        out.push(':');
                        out.push_str(&attribute.local);
                    }
                    out.push_str("=\"");
                    out.push_str(&escape_attr(value));
                    out.push('"');
                }
                let children = &self.nodes[node].children;
                if children.is_empty() {
                    out.push_str(" />");
                    return;
                }
                out.push('>');
                let child_default = if prefix_for(&name.ns).is_some() && name.ns != default_ns.unwrap_or("") {
                    default_ns
                } else {
                    Some(name.ns.as_str())
                };
                for &child in children {
                    self.write_xml(child, child_default, false, out);
                }
                out.push_str("</");
                out.push_str(&tag);
                out.push('>');
            }
        }
    }

    /// OmmlMathMlProjectionOracleTests.Canonical: element names (MathML
    /// namespace as local names, others as {ns}local), attributes sorted by
    /// expanded name, text in brackets, non-ASCII as \uXXXX UTF-16 units.
    pub fn canonical(&self, node: NodeId) -> String {
        let mut out = String::new();
        self.write_canonical(node, &mut out);
        out
    }

    fn write_canonical(&self, node: NodeId, out: &mut String) {
        match &self.nodes[node].data {
            Data::Text(text) => {
                out.push('[');
                out.push_str(&visible(text));
                out.push(']');
            }
            Data::Element { name, attrs } => {
                let shown = if name.ns == MATHML_NS { name.local.clone() } else { name.expanded() };
                out.push('<');
                out.push_str(&shown);
                let mut sorted: Vec<&(Name, String)> = attrs.iter().collect();
                sorted.sort_by(|a, b| a.0.expanded().cmp(&b.0.expanded()));
                for (attribute, value) in sorted {
                    out.push(' ');
                    out.push_str(&attribute.expanded());
                    out.push_str("=\"");
                    out.push_str(&visible(value));
                    out.push('"');
                }
                out.push('>');
                for &child in &self.nodes[node].children {
                    self.write_canonical(child, out);
                }
                out.push_str("</");
                out.push_str(&name.local);
                out.push('>');
            }
        }
    }
}

fn open_element(
    doc: &mut Doc,
    reader: &NsReader<&[u8]>,
    ns: String,
    start: &BytesStart,
    stack: &[NodeId],
) -> Result<NodeId, String> {
    let id = element_from_start(doc, reader, ns, start)?;
    if let Some(&parent) = stack.last() {
        doc.nodes[id].parent = Some(parent);
        doc.nodes[parent].children.push(id);
    } else if doc.root.replace(id).is_some() {
        return Err("XML contains multiple root elements".to_string());
    }
    Ok(id)
}

fn element_from_start(
    doc: &mut Doc,
    reader: &NsReader<&[u8]>,
    ns: String,
    start: &BytesStart,
) -> Result<NodeId, String> {
    let local = String::from_utf8_lossy(start.local_name().as_ref()).into_owned();
    let id = doc.new_element(Name { ns, local });
    for attribute in start.attributes().with_checks(false) {
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
                _ => String::new(),
            }
        };
        let value = attribute
            .decode_and_unescape_value(reader.decoder())
            .map_err(|error| error.to_string())?
            .into_owned();
        let name = Name { ns: attribute_ns, local: String::from_utf8_lossy(local_name.as_ref()).into_owned() };
        if let Data::Element { attrs, .. } = &mut doc.nodes[id].data {
            attrs.push((name, value));
        }
    }
    Ok(id)
}

fn push_text(doc: &mut Doc, stack: &[NodeId], text: &str) {
    if let Some(&parent) = stack.last() {
        doc.append_text(parent, text);
    }
}

fn drop_whitespace_text(doc: &mut Doc, element: NodeId) {
    let children = doc.nodes[element].children.clone();
    for child in children {
        if let Data::Text(text) = &doc.nodes[child].data {
            // XmlReader with IgnoreWhitespace drops whitespace-only text,
            // including an element's sole content, outside xml:space="preserve".
            if text.chars().all(|c| matches!(c, ' ' | '\t' | '\r' | '\n')) {
                doc.detach(child);
            }
        }
    }
}

fn prefix_for(ns: &str) -> Option<&'static str> {
    match ns {
        OMML_NS => Some("m"),
        WORD_NS => Some("w"),
        XML_NS => Some("xml"),
        _ => None,
    }
}

pub fn escape_text(value: &str) -> String {
    value.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;")
}

pub fn escape_attr(value: &str) -> String {
    value
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\n', "&#xA;")
        .replace('\r', "&#xD;")
        .replace('\t', "&#x9;")
}

/// OmmlMathMlProjectionOracleTests.Visible: printable ASCII as is, every other
/// UTF-16 code unit as \uXXXX (uppercase hex).
pub fn visible(value: &str) -> String {
    let mut out = String::new();
    for unit in value.encode_utf16() {
        if (0x20..=0x7E).contains(&unit) {
            out.push(unit as u8 as char);
        } else {
            out.push_str(&format!("\\u{unit:04X}"));
        }
    }
    out
}
