//! Native equations in WPS Presentation's text (VisualTeX for Windows'
//! WpsMathText and WpsNativeEquations). WPS has no equation objects in its API
//! (MathZones is a placeholder); its text spells every equation inline: the
//! math characters as WPS stores them (an italic x is U+1D465), each structure
//! (fraction, script, radical, delimiter, matrix, ...) opened by U+001A and
//! each of its arguments closed by U+001B, in the order the arguments appear
//! in the OMML. Spelling WPS's own saved copy of the shape the same way gives
//! each equation's exact range in the live text, and its OMML for the editor.

use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, XmlVersion};
use serde::Serialize;
use std::collections::BTreeMap;
use std::fs::File;
use std::io::Read;
use std::path::{Path, PathBuf};

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const M: &str = "http://schemas.openxmlformats.org/officeDocument/2006/math";
const MC: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const A14: &str = "http://schemas.microsoft.com/office/drawing/2010/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(crate) const STRUCTURE_START: char = '\u{1A}';
pub(crate) const ARGUMENT_END: char = '\u{1B}';
const STRUCTURES: [&str; 19] = [
    "acc", "bar", "borderBox", "box", "d", "eqArr", "f", "func", "groupChr", "limLow", "limUpp", "m", "nary",
    "phant", "rad", "sPre", "sSub", "sSubSup", "sSup",
];
const MAX_PART_BYTES: u64 = 64 * 1024 * 1024;
const MAX_DEPTH: usize = 512;
/// The saved copy the add-in makes for one look at a shape (module-wpp.js).
const SAVED_COPY_PREFIX: &str = "visualtex-wpp-";
const WPS_BUNDLE_IDS: [&str; 2] = ["com.kingsoft.wpsoffice.mac", "com.kingsoft.wpsoffice.mac.global"];

#[derive(Debug, Clone)]
struct Element {
    ns: String,
    local: String,
    /// The prefix the element was written with, for a namespace VisualTeX has no prefix for.
    prefix: String,
    attrs: Vec<Attribute>,
    children: Vec<Content>,
}

#[derive(Debug, Clone)]
struct Attribute {
    ns: String,
    local: String,
    prefix: String,
    value: String,
}

#[derive(Debug, Clone)]
enum Content {
    Element(Element),
    Text(String),
}

impl Element {
    fn is(&self, ns: &str, local: &str) -> bool {
        self.ns == ns && self.local == local
    }

    fn elements(&self) -> impl Iterator<Item = &Element> {
        self.children.iter().filter_map(|child| match child {
            Content::Element(element) => Some(element),
            Content::Text(_) => None,
        })
    }

    fn child(&self, ns: &str, local: &str) -> Option<&Element> {
        self.elements().find(|element| element.is(ns, local))
    }

    fn attr(&self, ns: &str, local: &str) -> Option<&str> {
        self.attrs
            .iter()
            .find(|attr| attr.ns == ns && attr.local == local)
            .map(|attr| attr.value.as_str())
    }

    fn text(&self) -> String {
        self.children
            .iter()
            .map(|child| match child {
                Content::Text(text) => text.clone(),
                Content::Element(element) => element.text(),
            })
            .collect()
    }
}

fn namespace(result: &ResolveResult) -> Result<String, String> {
    match result {
        ResolveResult::Bound(ns) => Ok(String::from_utf8_lossy(ns.as_ref()).into_owned()),
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(format!(
            "WPS saved XML with the undeclared prefix {}",
            String::from_utf8_lossy(prefix)
        )),
    }
}

fn prefix_of(qualified: &[u8]) -> String {
    let name = String::from_utf8_lossy(qualified);
    name.split_once(':').map(|(prefix, _)| prefix.to_string()).unwrap_or_default()
}

fn start_element(reader: &NsReader<&[u8]>, ns: String, start: &BytesStart) -> Result<Element, String> {
    let mut attrs = Vec::new();
    for attr in start.attributes() {
        let attr = attr.map_err(|error| format!("WPS saved invalid XML: {error}"))?;
        let key = attr.key.as_ref();
        // Declarations are written again for what an output uses.
        if key == b"xmlns" || key.starts_with(b"xmlns:") {
            continue;
        }
        let (resolved, local) = reader.resolver().resolve_attribute(attr.key);
        let ns = namespace(&resolved)?;
        let value = attr
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
            .map_err(|error| format!("WPS saved invalid XML: {error}"))?
            .into_owned();
        attrs.push(Attribute {
            ns,
            local: String::from_utf8_lossy(local.as_ref()).into_owned(),
            prefix: prefix_of(key),
            value,
        });
    }
    Ok(Element {
        ns,
        local: String::from_utf8_lossy(start.local_name().as_ref()).into_owned(),
        prefix: prefix_of(start.name().as_ref()),
        attrs,
        children: Vec::new(),
    })
}

fn parse(xml: &str) -> Result<Element, String> {
    let mut reader = NsReader::from_str(xml);
    let mut stack: Vec<Element> = Vec::new();
    let invalid = |error: quick_xml::Error| format!("WPS saved invalid XML: {error}");
    loop {
        let (resolved, event) = reader.read_resolved_event().map_err(invalid)?;
        let ns = namespace(&resolved)?;
        match event {
            Event::Start(start) => {
                let element = start_element(&reader, ns, &start)?;
                stack.push(element);
                if stack.len() > MAX_DEPTH {
                    return Err("WPS saved XML that is nested too deeply".to_string());
                }
            }
            Event::Empty(start) => {
                let element = start_element(&reader, ns, &start)?;
                match stack.last_mut() {
                    Some(parent) => parent.children.push(Content::Element(element)),
                    None => return Ok(element),
                }
            }
            Event::End(_) => {
                let element = stack.pop().ok_or("WPS saved unbalanced XML")?;
                match stack.last_mut() {
                    Some(parent) => parent.children.push(Content::Element(element)),
                    None => return Ok(element),
                }
            }
            Event::Text(text) => {
                if let Some(parent) = stack.last_mut() {
                    let text = text.decode().map_err(|error| format!("WPS saved invalid text: {error}"))?;
                    parent.children.push(Content::Text(text.into_owned()));
                }
            }
            Event::CData(data) => {
                if let Some(parent) = stack.last_mut() {
                    let text = data.decode().map_err(|error| format!("WPS saved invalid text: {error}"))?;
                    parent.children.push(Content::Text(text.into_owned()));
                }
            }
            Event::GeneralRef(reference) => {
                if let Some(parent) = stack.last_mut() {
                    let name = reference.decode().map_err(|error| format!("WPS saved invalid text: {error}"))?;
                    let escaped = format!("&{name};");
                    let text = quick_xml::escape::unescape(&escaped)
                        .map_err(|error| format!("WPS saved an unknown character reference: {error}"))?
                        .into_owned();
                    parent.children.push(Content::Text(text));
                }
            }
            Event::DocType(_) => return Err("WPS saved XML with a document type".to_string()),
            Event::Eof => return Err("WPS saved XML without a root element".to_string()),
            _ => {}
        }
    }
}

/// Child elements, an mc:AlternateContent replaced by its first choice (the
/// branch WPS and PowerPoint show).
fn live_children(element: &Element) -> Vec<&Element> {
    let mut children = Vec::new();
    for child in element.elements() {
        if !child.is(MC, "AlternateContent") {
            children.push(child);
            continue;
        }
        let live = child
            .elements()
            .find(|branch| branch.is(MC, "Choice"))
            .or_else(|| child.child(MC, "Fallback"));
        if let Some(live) = live {
            children.extend(live.elements());
        }
    }
    children
}

/// Text with its length in UTF-16 code units, the positions WPS's text ranges use.
#[derive(Default)]
struct Spelled {
    text: String,
    units: usize,
}

impl Spelled {
    fn push(&mut self, character: char) {
        self.text.push(character);
        self.units += character.len_utf16();
    }

    fn push_str(&mut self, text: &str) {
        self.text.push_str(text);
        self.units += text.encode_utf16().count();
    }
}

fn spell(element: &Element, text: &mut Spelled) {
    if element.ns != M || element.local.ends_with("Pr") {
        return;
    }
    if element.local == "r" {
        for run in element.elements().filter(|child| child.is(M, "t")) {
            text.push_str(&run.text());
        }
        return;
    }
    if !STRUCTURES.contains(&element.local.as_str()) {
        spell_children(element, text);
        return;
    }
    text.push(STRUCTURE_START);
    for argument in element.elements() {
        if argument.ns != M || argument.local.ends_with("Pr") {
            continue;
        }
        // A matrix row adds no mark of its own: each cell is an argument.
        let cells: Vec<&Element> = if argument.is(M, "mr") {
            argument.elements().filter(|cell| cell.is(M, "e")).collect()
        } else {
            vec![argument]
        };
        for cell in cells {
            spell_children(cell, text);
            text.push(ARGUMENT_END);
        }
    }
}

fn spell_children(element: &Element, text: &mut Spelled) {
    for child in live_children(element) {
        spell(child, text);
    }
}

/// One native equation of a shape's text.
#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub(crate) struct Zone {
    /// 1-based position and length in the live text, in UTF-16 code units.
    pub start: usize,
    pub length: usize,
    /// The equation is alone on its line (a display equation).
    pub display: bool,
    /// Its a14:m element as WPS saved it, with the namespaces it uses.
    pub zone: String,
}

/// The native equations of a saved text shape, in document order; None when
/// the saved shape does not spell `live_text` exactly, so no position could
/// be trusted.
fn locate(shape: &Element, live_text: &str) -> Option<Vec<Zone>> {
    let mut text = Spelled::default();
    let mut found: Vec<(usize, usize, &Element)> = Vec::new();
    let paragraphs: Vec<&Element> = shape
        .child(P, "txBody")
        .map(|body| body.elements().filter(|element| element.is(A, "p")).collect())
        .unwrap_or_default();
    for (index, paragraph) in paragraphs.iter().enumerate() {
        if index > 0 {
            text.push('\r');
        }
        for child in live_children(paragraph) {
            if child.is(A, "r") || child.is(A, "fld") {
                // WPS saves the paragraph break before a display equation as a
                // CR inside the preceding run (Before text.&#13;).
                let run = child.child(A, "t").map(Element::text).unwrap_or_default();
                text.push_str(&run.replace("\r\n", "\r").replace('\n', "\r"));
            } else if child.is(A, "br") {
                text.push('\u{b}');
            } else if child.is(A14, "m") {
                let start = text.units;
                spell_children(child, &mut text);
                found.push((start + 1, text.units - start, child));
            }
        }
    }
    if text.text != live_text {
        return None;
    }
    let units: Vec<u16> = text.text.encode_utf16().collect();
    Some(
        found
            .into_iter()
            .map(|(start, length, zone)| Zone {
                start,
                length,
                display: alone_on_its_line(&units, start - 1, start - 1 + length),
                zone: serialize(zone),
            })
            .collect(),
    )
}

/// Whether nothing but spaces shares the equation's paragraph (VTNativeEditRequest).
fn alone_on_its_line(units: &[u16], start: usize, end: usize) -> bool {
    const CR: u16 = 0x0D;
    let blank = |unit: &u16| matches!(*unit, 0x20 | 0x09 | 0xA0);
    let before = units[..start].iter().rev().take_while(|unit| **unit != CR);
    let after = units[end..].iter().take_while(|unit| **unit != CR);
    before.chain(after).all(blank)
}

/// The one equation that holds the caret, or that a selection overlaps
/// (positions 1-based, as WPS's text ranges count them).
pub(crate) fn zone_at(zones: &[(usize, usize)], selected_start: usize, selected_length: usize) -> Option<usize> {
    let mut found = None;
    for (index, &(start, length)) in zones.iter().enumerate() {
        let hit = if selected_length == 0 {
            selected_start >= start && selected_start < start + length
        } else {
            selected_start < start + length && selected_start + selected_length > start
        };
        if !hit {
            continue;
        }
        if found.is_some() {
            return None;
        }
        found = Some(index);
    }
    found
}

/// False when the text certainly holds no equation (structures or math
/// letters); an equation of digits and operators alone passes unnoticed.
pub(crate) fn may_contain_math(text: &str) -> bool {
    text.chars()
        .any(|c| c == STRUCTURE_START || c == '\u{210E}' || ('\u{1D400}'..='\u{1D7FF}').contains(&c))
}

fn prefix_for(ns: &str, written: &str) -> String {
    match ns {
        M => "m".to_string(),
        A => "a".to_string(),
        A14 => "a14".to_string(),
        W => "w".to_string(),
        MC => "mc".to_string(),
        P => "p".to_string(),
        R => "r".to_string(),
        _ if !written.is_empty() => written.to_string(),
        _ => "ns".to_string(),
    }
}

fn escape_text(text: &str, attribute: bool) -> String {
    let mut out = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' if attribute => out.push_str("&quot;"),
            '\r' => out.push_str("&#13;"),
            '\n' if attribute => out.push_str("&#10;"),
            '\t' if attribute => out.push_str("&#9;"),
            _ => out.push(character),
        }
    }
    out
}

fn collect_namespaces(element: &Element, used: &mut BTreeMap<String, String>) {
    used.entry(prefix_for(&element.ns, &element.prefix)).or_insert_with(|| element.ns.clone());
    for attr in element.attrs.iter().filter(|attr| !attr.ns.is_empty()) {
        used.entry(prefix_for(&attr.ns, &attr.prefix)).or_insert_with(|| attr.ns.clone());
    }
    for child in element.elements() {
        collect_namespaces(child, used);
    }
}

fn write(element: &Element, out: &mut String, declarations: Option<&BTreeMap<String, String>>) {
    let name = |ns: &str, prefix: &str, local: &str| {
        if ns.is_empty() {
            local.to_string()
        } else {
            format!("{}:{local}", prefix_for(ns, prefix))
        }
    };
    let element_name = name(&element.ns, &element.prefix, &element.local);
    out.push('<');
    out.push_str(&element_name);
    if let Some(declarations) = declarations {
        for (prefix, uri) in declarations {
            out.push_str(&format!(" xmlns:{prefix}=\"{}\"", escape_text(uri, true)));
        }
    }
    for attr in &element.attrs {
        out.push_str(&format!(" {}=\"{}\"", name(&attr.ns, &attr.prefix, &attr.local), escape_text(&attr.value, true)));
    }
    if element.children.is_empty() {
        out.push_str("/>");
        return;
    }
    out.push('>');
    for child in &element.children {
        match child {
            Content::Element(child) => write(child, out, None),
            Content::Text(text) => out.push_str(&escape_text(text, false)),
        }
    }
    out.push_str(&format!("</{element_name}>"));
}

/// The element as XML that declares every namespace it uses.
fn serialize(element: &Element) -> String {
    let mut used = BTreeMap::new();
    collect_namespaces(element, &mut used);
    let mut out = String::new();
    write(element, &mut out, Some(&used));
    out
}

/// The equation of a zone (an a14:m element) as one m:oMath, as the editor
/// reads PowerPoint's (powerpoint_omml::native_equation_from_clipboard).
pub(crate) fn zone_equation(zone: &str) -> Result<String, String> {
    let zone = parse(zone)?;
    if !zone.is(A14, "m") {
        return Err("The equation to edit is not a native math zone".to_string());
    }
    fn collect(element: &Element, children: &mut Vec<Content>, count: &mut usize) {
        if element.is(M, "oMath") {
            *count += 1;
            children.extend(element.children.iter().cloned());
            return;
        }
        for child in live_children(element) {
            collect(child, children, count);
        }
    }
    let mut children = Vec::new();
    let mut count = 0;
    collect(&zone, &mut children, &mut count);
    if count == 0 {
        return Err("The native math zone holds no equation".to_string());
    }
    let equation = Element { ns: M.to_string(), local: "oMath".to_string(), prefix: "m".to_string(), attrs: Vec::new(), children };
    let mut used = BTreeMap::new();
    collect_namespaces(&equation, &mut used);
    // The editor's reader expects the math, DrawingML and Word prefixes declared.
    for (prefix, uri) in [("m", M), ("a", A), ("w", W)] {
        used.entry(prefix.to_string()).or_insert_with(|| uri.to_string());
    }
    let mut out = String::new();
    write(&equation, &mut out, Some(&used));
    Ok(out)
}

fn read_part<R: Read + std::io::Seek>(archive: &mut zip::ZipArchive<R>, name: &str) -> Result<String, String> {
    let mut part = archive
        .by_name(name)
        .map_err(|_| format!("WPS's saved copy has no {name}"))?;
    if part.size() > MAX_PART_BYTES {
        return Err(format!("{name} in WPS's saved copy is too large"));
    }
    let mut xml = String::new();
    part.by_ref()
        .take(MAX_PART_BYTES + 1)
        .read_to_string(&mut xml)
        .map_err(|error| format!("Unable to read {name} from WPS's saved copy: {error}"))?;
    Ok(xml)
}

/// A shape anywhere on the slide (in groups, and in the first branch of a
/// markup-compatibility choice) by its id.
fn find_shape<'a>(element: &'a Element, shape_id: &str) -> Option<&'a Element> {
    for child in live_children(element) {
        if child.is(P, "sp")
            && child
                .child(P, "nvSpPr")
                .and_then(|properties| properties.child(P, "cNvPr"))
                .and_then(|properties| properties.attr("", "id"))
                == Some(shape_id)
        {
            return Some(child);
        }
        if let Some(found) = find_shape(child, shape_id) {
            return Some(found);
        }
    }
    None
}

fn slide_part(archive: &mut zip::ZipArchive<File>, slide_id: u32) -> Result<String, String> {
    let presentation = parse(&read_part(archive, "ppt/presentation.xml")?)?;
    let slide_id = slide_id.to_string();
    let relationship = presentation
        .child(P, "sldIdLst")
        .and_then(|list| {
            list.elements()
                .find(|slide| slide.is(P, "sldId") && slide.attr("", "id") == Some(slide_id.as_str()))
        })
        .and_then(|slide| slide.attr(R, "id"))
        .ok_or("The slide is not in WPS's saved copy")?
        .to_string();
    let relationships = parse(&read_part(archive, "ppt/_rels/presentation.xml.rels")?)?;
    let target = relationships
        .elements()
        .find(|item| item.local == "Relationship" && item.attr("", "Id") == Some(relationship.as_str()))
        .and_then(|item| item.attr("", "Target"))
        .ok_or("The slide's part is not in WPS's saved copy")?;
    let part = match target.strip_prefix('/') {
        Some(absolute) => absolute.to_string(),
        None => format!("ppt/{target}"),
    };
    if part.contains("..") {
        return Err("WPS's saved copy names its slide outside the presentation".to_string());
    }
    read_part(archive, &part)
}

/// The native equations of one text shape in a presentation WPS saved
/// (Presentation.SaveCopyAs), checked against the shape's live text.
pub(crate) fn saved_shape_equations(path: &Path, slide_id: u32, shape_id: u32, live_text: &str) -> Result<Vec<Zone>, String> {
    let file = File::open(path).map_err(|error| format!("Unable to open WPS's saved copy: {error}"))?;
    let mut archive =
        zip::ZipArchive::new(file).map_err(|error| format!("WPS's saved copy is not a presentation: {error}"))?;
    let slide = parse(&slide_part(&mut archive, slide_id)?)?;
    let shape = find_shape(&slide, &shape_id.to_string()).ok_or("The text box is not in WPS's saved copy")?;
    locate(shape, live_text).ok_or_else(|| {
        "无法确定这个文本框中原生公式的位置（WPS 的文本与其保存内容不一致）。".to_string()
    })
}

/// The native equations of several text shapes, from one saved copy: each
/// shape's zones, or why they could not be located.
pub(crate) fn saved_shapes_equations(path: &Path, shapes: &[(u32, u32, String)]) -> Result<Vec<Result<Vec<Zone>, String>>, String> {
    let file = File::open(path).map_err(|error| format!("Unable to open WPS's saved copy: {error}"))?;
    let mut archive =
        zip::ZipArchive::new(file).map_err(|error| format!("WPS's saved copy is not a presentation: {error}"))?;
    let mut slides: std::collections::HashMap<u32, Result<Element, String>> = std::collections::HashMap::new();
    let mut results = Vec::with_capacity(shapes.len());
    for (slide_id, shape_id, text) in shapes {
        if !slides.contains_key(slide_id) {
            let slide = slide_part(&mut archive, *slide_id).and_then(|xml| parse(&xml));
            slides.insert(*slide_id, slide);
        }
        let result = match &slides[slide_id] {
            Ok(slide) => find_shape(slide, &shape_id.to_string())
                .ok_or_else(|| "The text box is not in WPS's saved copy".to_string())
                .and_then(|shape| {
                    locate(shape, text).ok_or_else(|| {
                        "无法确定这个文本框中原生公式的位置（WPS 的文本与其保存内容不一致）。".to_string()
                    })
                }),
            Err(error) => Err(error.clone()),
        };
        results.push(result);
    }
    Ok(results)
}

/// The saved copy the add-in names: a visualtex-wpp-<uuid>.pptx directly in a
/// WPS container's data folder tree, so no other file can be read through it.
pub(crate) fn saved_copy_path(path: &str, home: &Path) -> Result<PathBuf, String> {
    let path = Path::new(path);
    let name = path.file_name().and_then(|name| name.to_str()).unwrap_or_default();
    let id = name
        .strip_prefix(SAVED_COPY_PREFIX)
        .and_then(|rest| rest.strip_suffix(".pptx"))
        .unwrap_or_default();
    if uuid::Uuid::parse_str(id).map(|uuid| uuid.to_string() != id).unwrap_or(true) {
        return Err("The saved copy's name is not one the WPS add-in makes".to_string());
    }
    let canonical = path
        .canonicalize()
        .map_err(|error| format!("Unable to find WPS's saved copy: {error}"))?;
    let inside = WPS_BUNDLE_IDS.iter().any(|bundle_id| {
        home.join("Library/Containers")
            .join(bundle_id)
            .join("Data")
            .canonicalize()
            .is_ok_and(|data| canonical.starts_with(data))
    });
    let metadata = std::fs::symlink_metadata(&canonical).map_err(|error| error.to_string())?;
    if !inside || !metadata.is_file() {
        return Err("The saved copy is not in WPS's own folder".to_string());
    }
    Ok(canonical)
}

/// The BCP 47 tag of an Office language id (msoLanguageID), for the runs of
/// an equation. Right-to-left languages are left to WPS.
pub(crate) fn language_tag(id: u32) -> Option<&'static str> {
    Some(match id {
        1033 => "en-US",
        2057 => "en-GB",
        3081 => "en-AU",
        4105 => "en-CA",
        2052 => "zh-CN",
        1028 => "zh-TW",
        3076 => "zh-HK",
        4100 => "zh-SG",
        5124 => "zh-MO",
        1041 => "ja-JP",
        1042 => "ko-KR",
        1031 => "de-DE",
        1036 => "fr-FR",
        1034 | 3082 => "es-ES",
        2058 => "es-MX",
        1040 => "it-IT",
        1046 => "pt-BR",
        2070 => "pt-PT",
        1049 => "ru-RU",
        1045 => "pl-PL",
        1043 => "nl-NL",
        1053 => "sv-SE",
        1030 => "da-DK",
        1044 => "nb-NO",
        1035 => "fi-FI",
        1055 => "tr-TR",
        1029 => "cs-CZ",
        1038 => "hu-HU",
        1032 => "el-GR",
        1058 => "uk-UA",
        1066 => "vi-VN",
        1054 => "th-TH",
        1057 => "id-ID",
        1086 => "ms-MY",
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    const S: &str = "\u{1A}";
    const E: &str = "\u{1B}";
    const X: &str = "\u{1D465}"; // math italic x
    const Y: &str = "\u{1D466}"; // math italic y
    const RPR: &str = "<a:rPr sz=\"2400\"><a:latin typeface=\"Cambria Math\"/></a:rPr>";

    fn r(text: &str) -> String {
        format!("<m:r>{RPR}<m:t>{text}</m:t></m:r>")
    }

    fn ctrl(name: &str) -> String {
        format!("<m:{name}Pr><m:ctrlPr>{RPR}</m:ctrlPr></m:{name}Pr>")
    }

    fn spelled(math: &str) -> String {
        let zone = parse(&format!("<a14:m xmlns:a14=\"{A14}\" xmlns:a=\"{A}\"><m:oMath xmlns:m=\"{M}\">{math}</m:oMath></a14:m>")).unwrap();
        let mut text = Spelled::default();
        spell_children(&zone, &mut text);
        text.text
    }

    fn shape(paragraphs: &str) -> Element {
        parse(&format!(
            "<p:sp xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:a14=\"{A14}\" xmlns:m=\"{M}\"><p:txBody><a:bodyPr/>{paragraphs}</p:txBody></p:sp>"
        ))
        .unwrap()
    }

    fn math(math: &str) -> String {
        format!("<a14:m><m:oMath>{math}</m:oMath></a14:m>")
    }

    fn positions(zones: &[Zone]) -> Vec<(usize, usize)> {
        zones.iter().map(|zone| (zone.start, zone.length)).collect()
    }

    // How WPS Presentation spells native equations (measured in WPS 12.1 on Windows,
    // VisualTeX.WindowsOffice.Tests WpsMathTextTests).
    #[test]
    fn spells_runs_as_saved() {
        assert_eq!(spelled(&(r(X) + &r("=") + &r("1"))), format!("{X}=1"));
    }

    #[test]
    fn spells_every_structure_with_its_arguments_in_xml_order() {
        assert_eq!(spelled(&format!("<m:sSup>{}<m:e>{}</m:e><m:sup>{}</m:sup></m:sSup>", ctrl("sSup"), r(X), r("2"))), format!("{S}{X}{E}2{E}"));
        // a hidden radical degree still is an (empty) argument
        assert_eq!(spelled(&format!("<m:rad><m:radPr><m:degHide m:val=\"on\"/></m:radPr><m:deg/><m:e>{}</m:e></m:rad>", r(X))), format!("{S}{E}{X}{E}"));
        assert_eq!(
            spelled(&format!("<m:nary><m:naryPr><m:chr m:val=\"S\"/></m:naryPr><m:sub>{}</m:sub><m:sup>{}</m:sup><m:e>{}</m:e></m:nary>", r("i=1"), r("n"), r(X))),
            format!("{S}i=1{E}n{E}{X}{E}")
        );
        // a 2x2 matrix: one argument per cell, rows add nothing
        assert_eq!(
            spelled(&format!("<m:m><m:mr><m:e>{}</m:e><m:e>{}</m:e></m:mr><m:mr><m:e>{}</m:e><m:e>{}</m:e></m:mr></m:m>", r("a"), r("b"), r("c"), r("d"))),
            format!("{S}a{E}b{E}c{E}d{E}")
        );
        // delimiters and accents: their characters live in properties, not in the text
        assert_eq!(spelled(&format!("<m:d><m:dPr><m:begChr m:val=\"[\"/></m:dPr><m:e>{}</m:e><m:e>{}</m:e></m:d>", r(X), r(Y))), format!("{S}{X}{E}{Y}{E}"));
        assert_eq!(spelled(&format!("<m:acc><m:accPr><m:chr m:val=\"^\"/></m:accPr><m:e>{}</m:e></m:acc>", r(X))), format!("{S}{X}{E}"));
        assert_eq!(spelled(&format!("<m:func><m:fName>{}</m:fName><m:e>{}</m:e></m:func>", r("sin"), r(X))), format!("{S}sin{E}{X}{E}"));
        assert_eq!(
            spelled(&format!("<m:sPre><m:sub>{}</m:sub><m:sup>{}</m:sup><m:e><m:f><m:num>{}</m:num><m:den>{}</m:den></m:f></m:e></m:sPre>", r("1"), r("2"), r(X), r(Y))),
            format!("{S}1{E}2{E}{S}{X}{E}{Y}{E}{E}")
        );
    }

    #[test]
    fn locates_equations_between_text_runs_breaks_and_paragraphs() {
        let shape = shape(&format!(
            "<a:p><a:r><a:t>Let </a:t></a:r>{}<a:r><a:t> and </a:t></a:r>{}</a:p><a:p><a:r><a:t>line1</a:t></a:r><a:br/><a:r><a:t>line2</a:t></a:r></a:p><a:p><a:pPr algn=\"ctr\"/><a14:m><m:oMathPara><m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr><m:oMath>{}</m:oMath></m:oMathPara></a14:m></a:p>",
            math(&r(X)),
            math(&format!("<m:sSup><m:e>{}</m:e><m:sup>{}</m:sup></m:sSup>", r(Y), r("2"))),
            r("1")
        ));
        let live = format!("Let {X} and {S}{Y}{E}2{E}\rline1\u{b}line2\r1");
        let zones = locate(&shape, &live).unwrap();
        assert_eq!(positions(&zones), vec![(5, 2), (12, 6), (31, 1)]);
        assert_eq!(zones.iter().map(|zone| zone.display).collect::<Vec<_>>(), vec![false, false, true]);
        assert!(zones[2].zone.starts_with("<a14:m xmlns:a=") && zones[2].zone.contains("<m:jc m:val=\"centerGroup\"/>"));
    }

    #[test]
    fn separates_adjacent_equations() {
        let shape = shape(&format!("<a:p>{}{}</a:p>", math(&r(X)), math(&r(Y))));
        assert_eq!(positions(&locate(&shape, &format!("{X}{Y}")).unwrap()), vec![(1, 2), (3, 2)]);
    }

    // WPS saves the break before a display equation as a CR inside the preceding run.
    #[test]
    fn reads_the_cr_wps_keeps_in_a_run_before_a_display_equation() {
        let shape = shape(&format!("<a:p><a:r><a:t>Before.&#13;</a:t></a:r>{}</a:p>", math(&r(X))));
        let zones = locate(&shape, &format!("Before.\r{X}")).unwrap();
        assert_eq!(positions(&zones), vec![(9, 2)]);
        assert!(zones[0].display);
    }

    #[test]
    fn trusts_no_position_when_the_text_differs() {
        let shape = shape(&format!("<a:p><a:r><a:t>Let </a:t></a:r>{}</a:p>", math(&r(X))));
        assert!(locate(&shape, "Let x").is_none());
        assert!(locate(&shape, &format!("Let {X} more")).is_none());
    }

    #[test]
    fn finds_the_one_equation_at_a_caret_or_selection() {
        let zones = [(5, 2), (13, 6)];
        assert_eq!(zone_at(&zones, 5, 0), Some(0));
        assert_eq!(zone_at(&zones, 6, 0), Some(0));
        assert_eq!(zone_at(&zones, 7, 0), None);
        assert_eq!(zone_at(&zones, 13, 6), Some(1));
        assert_eq!(zone_at(&zones, 4, 12), None);
    }

    #[test]
    fn filters_text_without_math() {
        assert!(!may_contain_math("plain text 2+3"));
        assert!(may_contain_math(&format!("Let {X}")));
        assert!(may_contain_math(&format!("{S}a{E}")));
    }

    #[test]
    fn a_zone_gives_its_equation_to_the_editor() {
        let shape = shape(&format!("<a:p>{}</a:p>", math(&format!("<m:f><m:num>{}</m:num><m:den>{}</m:den></m:f>", r("x+1"), r("2")))));
        let zones = locate(&shape, &format!("{S}x+1{E}2{E}")).unwrap();
        let equation = zone_equation(&zones[0].zone).unwrap();
        assert!(equation.starts_with("<m:oMath xmlns:a=") && equation.contains("xmlns:w="));
        assert_eq!(crate::office::powerpoint_omml::to_latex::from_omml(&equation).unwrap().latex, r"\frac{x+1}{2}");
        // The zone is the edit's context: powerpoint_omml::native_edit_paragraph takes it.
        let edited = crate::office::powerpoint_omml::native_edit_paragraph(
            &format!("<m:oMath xmlns:m=\"{M}\"><m:r><m:t>y</m:t></m:r></m:oMath>"), 20.0, &zones[0].zone, Some("zh-CN"),
        )
        .unwrap();
        assert!(edited.contains("<m:t>y</m:t>") && !edited.contains("x+1") && edited.contains("lang=\"zh-CN\""));
        assert!(zone_equation("<p:sp/>").is_err());
    }

    #[test]
    fn saved_copies_are_read_by_slide_and_shape_id() {
        use std::io::Write;
        let temp = tempfile::TempDir::new().unwrap();
        let path = temp.path().join("deck.pptx");
        let slide = format!(
            "<p:sld xmlns:p=\"{P}\" xmlns:a=\"{A}\" xmlns:a14=\"{A14}\" xmlns:m=\"{M}\" xmlns:mc=\"{MC}\"><p:cSld><p:spTree><p:grpSp><mc:AlternateContent><mc:Choice Requires=\"a14\"><p:sp><p:nvSpPr><p:cNvPr id=\"7\" name=\"Box\"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Let </a:t></a:r>{}</a:p></p:txBody></p:sp></mc:Choice><mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id=\"7\" name=\"Box\"/></p:nvSpPr><p:txBody><a:p><a:r><a:t>Let x</a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent></p:grpSp></p:spTree></p:cSld></p:sld>",
            math(&r(X))
        );
        let file = std::fs::File::create(&path).unwrap();
        let mut zip = zip::ZipWriter::new(file);
        let options = zip::write::SimpleFileOptions::default();
        for (name, xml) in [
            ("ppt/presentation.xml".to_string(), format!("<p:presentation xmlns:p=\"{P}\" xmlns:r=\"{R}\"><p:sldIdLst><p:sldId id=\"256\" r:id=\"rId2\"/><p:sldId id=\"300\" r:id=\"rId9\"/></p:sldIdLst></p:presentation>")),
            ("ppt/_rels/presentation.xml.rels".to_string(), "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId9\" Type=\"x\" Target=\"slides/slide4.xml\"/></Relationships>".to_string()),
            ("ppt/slides/slide4.xml".to_string(), slide),
        ] {
            zip.start_file(name, options).unwrap();
            zip.write_all(xml.as_bytes()).unwrap();
        }
        zip.finish().unwrap();
        let zones = saved_shape_equations(&path, 300, 7, &format!("Let {X}")).unwrap();
        assert_eq!(positions(&zones), vec![(5, 2)]);
        assert!(!zones[0].display);
        assert!(saved_shape_equations(&path, 256, 7, "x").is_err(), "slide without a part");
        assert!(saved_shape_equations(&path, 300, 8, "x").is_err(), "no such shape");
        assert!(saved_shape_equations(&path, 300, 7, "Let x").is_err(), "the fallback is not read");
    }

    #[test]
    fn only_the_add_ins_saved_copies_in_wps_folders_are_read() {
        let home = tempfile::TempDir::new().unwrap();
        let data = home.path().join("Library/Containers/com.kingsoft.wpsoffice.mac/Data/tmp");
        std::fs::create_dir_all(&data).unwrap();
        let id = "0f8fad5b-d9cb-469f-a165-70867728950e";
        let good = data.join(format!("visualtex-wpp-{id}.pptx"));
        std::fs::write(&good, b"x").unwrap();
        assert!(saved_copy_path(good.to_str().unwrap(), home.path()).is_ok());
        let other = data.join("notes.pptx");
        std::fs::write(&other, b"x").unwrap();
        assert!(saved_copy_path(other.to_str().unwrap(), home.path()).is_err());
        let outside = home.path().join(format!("visualtex-wpp-{id}.pptx"));
        std::fs::write(&outside, b"x").unwrap();
        assert!(saved_copy_path(outside.to_str().unwrap(), home.path()).is_err());
        let upper = data.join(format!("visualtex-wpp-{}.pptx", id.to_uppercase()));
        std::fs::write(&upper, b"x").unwrap();
        assert!(saved_copy_path(upper.to_str().unwrap(), home.path()).is_err());
        assert!(saved_copy_path(data.join(format!("visualtex-wpp-{id}.pptx/../x")).to_str().unwrap(), home.path()).is_err());
    }

    #[test]
    fn languages_map_to_tags_and_right_to_left_ones_are_left_out() {
        assert_eq!(language_tag(2052), Some("zh-CN"));
        assert_eq!(language_tag(1033), Some("en-US"));
        assert_eq!(language_tag(1025), None);
        assert_eq!(language_tag(0), None);
    }
}
