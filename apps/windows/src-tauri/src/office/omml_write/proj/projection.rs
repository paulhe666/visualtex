//! Port of MathMlOmmlProjection.cs: VisualTeX's own presentation MathML ->
//! Office Math (OMML) projection (the replacement for Office's MML2OMML.XSL).
//!
//! Every C# helper keeps its name (snake_case) so the two files can be read
//! side by side. Source elements live in `src` (read with
//! OmmlMathMlProjection.ReaderSettings(), i.e. IgnoreWhitespace = true);
//! projected OMML nodes are created in `out`.

use crate::office::omml_write::proj::dom::{Doc, Name, NodeId, MATHML_NS, OMML_NS, WORD_NS, XMLNS_NS, XML_NS};
use crate::office::omml_write::proj::netstr::len16;
use crate::office::omml_write::proj::xml_reader::{self, PROJECTION_READER};
use crate::office::omml_write::proj::xml_writer;

const MML: &str = MATHML_NS;
const M: &str = OMML_NS;
const W: &str = WORD_NS;
const RUN_ROLE_ATTRIBUTE: &str = "data-visualtex-omml-run-role";
const SPACE_TEXT_ATTRIBUTE: &str = "data-visualtex-omml-space";

const NARY_OPERATORS: &[&str] = &[
    "\u{222B}", "\u{222C}", "\u{222D}", "\u{222E}", "\u{222F}", "\u{2230}", "\u{2232}", "\u{2233}", "\u{2231}",
    "\u{2229}", "\u{222A}", "\u{220F}", "\u{2210}", "\u{2211}", "\u{22C0}", "\u{22C1}", "\u{22C2}", "\u{22C3}",
];

const GROWING_NARY_OPERATORS: &[&str] = &[
    "\u{222B}", "\u{222E}", "\u{222F}", "\u{2232}", "\u{2233}", "\u{2229}", "\u{222A}", "\u{220F}", "\u{2211}",
    "\u{22C0}", "\u{22C1}", "\u{22C2}", "\u{22C3}",
];

/// MathMlOmmlProjection.Project: projects one MathML document; the result is
/// one m:oMath element carrying the stylesheet's namespace declarations.
pub fn project(mathml: &str) -> Result<String, String> {
    let source = xml_reader::parse(mathml, PROJECTION_READER)?;
    let mut projector = Projector { src: &source, out: Doc::new() };
    let equation = projector.out.new_element(Name::new(M, "oMath"));
    // The stylesheet's literal m:oMath carried every namespace in its scope.
    projector.out.set_attr_ns(equation, XMLNS_NS, "m", Some(M));
    projector.out.set_attr_ns(equation, XMLNS_NS, "mml", Some(MML));
    projector.out.set_attr_ns(equation, XMLNS_NS, "w", Some(W));
    projector.out.root = Some(equation);
    if let Some(root) = source.root {
        let nodes = projector.apply(root)?;
        projector.add_all(equation, nodes);
    }
    xml_writer::write_to_string(&projector.out, equation)
}

/// One item of `params object?[] content` for Om(...).
enum Part {
    Node(NodeId),
    Nodes(Vec<NodeId>),
    Attr(&'static str, &'static str, String),
    Text(String),
    Nothing,
}

impl From<NodeId> for Part {
    fn from(node: NodeId) -> Self {
        Part::Node(node)
    }
}

impl From<Option<NodeId>> for Part {
    fn from(node: Option<NodeId>) -> Self {
        match node {
            Some(node) => Part::Node(node),
            None => Part::Nothing,
        }
    }
}

impl From<Vec<NodeId>> for Part {
    fn from(nodes: Vec<NodeId>) -> Self {
        Part::Nodes(nodes)
    }
}

struct Projector<'a> {
    src: &'a Doc,
    out: Doc,
}

// ---- XPath helpers (source side) ---------------------------------------------

/// NormalizeSpace: XPath normalize-space over the four XML whitespace characters.
fn normalize_space(value: &str) -> String {
    let mut builder = String::with_capacity(value.len());
    let mut pending_space = false;
    for character in value.chars() {
        if matches!(character, ' ' | '\t' | '\r' | '\n') {
            pending_space = !builder.is_empty();
            continue;
        }
        if pending_space {
            builder.push(' ');
        }
        pending_space = false;
        builder.push(character);
    }
    builder
}

/// LowerAscii.
fn lower_ascii(value: &str) -> String {
    value.chars().map(|c| if c.is_ascii_uppercase() { c.to_ascii_lowercase() } else { c }).collect()
}

/// IsNotANumber: string(number(text())) = 'NaN' as .NET's XSLT number()
/// decides it (optional sign, decimal digits, optional exponent, finite).
fn is_not_a_number(value: Option<&str>) -> bool {
    let Some(value) = value else { return true };
    let trimmed = value.trim_matches(|c| matches!(c, ' ' | '\t' | '\r' | '\n'));
    if !xpath_number_matches(trimmed) {
        return true;
    }
    match trimmed.parse::<f64>() {
        Ok(number) => number.is_infinite(),
        Err(_) => true,
    }
}

/// `^[+-]?([0-9]+(\.[0-9]*)?|\.[0-9]+)([eE][+-]?[0-9]+)?$` (the input is
/// already trimmed, so .NET's `$` before a final newline cannot apply).
fn xpath_number_matches(value: &str) -> bool {
    let bytes = value.as_bytes();
    let mut index = 0;
    if index < bytes.len() && (bytes[index] == b'+' || bytes[index] == b'-') {
        index += 1;
    }
    let integer_start = index;
    while index < bytes.len() && bytes[index].is_ascii_digit() {
        index += 1;
    }
    let integer_digits = index - integer_start;
    if integer_digits > 0 {
        if index < bytes.len() && bytes[index] == b'.' {
            index += 1;
            while index < bytes.len() && bytes[index].is_ascii_digit() {
                index += 1;
            }
        }
    } else {
        if index >= bytes.len() || bytes[index] != b'.' {
            return false;
        }
        index += 1;
        let fraction_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        if index == fraction_start {
            return false;
        }
    }
    if index < bytes.len() && (bytes[index] == b'e' || bytes[index] == b'E') {
        index += 1;
        if index < bytes.len() && (bytes[index] == b'+' || bytes[index] == b'-') {
            index += 1;
        }
        let exponent_start = index;
        while index < bytes.len() && bytes[index].is_ascii_digit() {
            index += 1;
        }
        if index == exponent_start {
            return false;
        }
    }
    index == bytes.len()
}

/// FirstScalar: the first UTF-16 scalar (a surrogate pair counts as one).
fn first_scalar(value: &str) -> String {
    value.chars().next().map(String::from).unwrap_or_default()
}

/// String.Substring(1) for a color value; a split surrogate pair would make
/// the XmlWriter throw.
fn substring1(value: &str) -> Result<String, String> {
    crate::office::omml_write::proj::netstr::substring_from(value, 1)
}

impl<'a> Projector<'a> {
    // Is(element, name)
    fn is(&self, element: Option<NodeId>, name: &str) -> bool {
        element.is_some_and(|element| self.src.is(element, MML, name))
    }

    // IsAny(element, names)
    fn is_any(&self, element: Option<NodeId>, names: &[&str]) -> bool {
        element.is_some_and(|element| {
            self.src.name(element).is_some_and(|name| name.ns == MML && names.contains(&name.local.as_str()))
        })
    }

    fn is_token(&self, element: Option<NodeId>) -> bool {
        self.is_any(element, &["mi", "mn", "mo", "ms", "mtext"])
    }

    fn has_glyph(&self, element: NodeId) -> bool {
        self.src.element(element, MML, "mglyph").is_some()
    }

    // Value(element, name): @x, else @mml:x, else "".
    fn value(&self, element: Option<NodeId>, name: &str) -> String {
        let Some(element) = element else { return String::new() };
        self.src
            .attr(element, name)
            .or_else(|| self.src.attr_ns(element, MML, name))
            .unwrap_or("")
            .to_string()
    }

    fn plain(&self, element: Option<NodeId>, name: &str) -> Option<&'a str> {
        let src: &'a Doc = self.src;
        element.and_then(|element| src.attr(element, name))
    }

    fn qualified(&self, element: Option<NodeId>, name: &str) -> Option<&'a str> {
        let src: &'a Doc = self.src;
        element.and_then(|element| src.attr_ns(element, MML, name))
    }

    // AttributeIs: @x='v' or @mml:x='v'
    fn attribute_is(&self, element: NodeId, name: &str, value: &str) -> bool {
        self.plain(Some(element), name) == Some(value) || self.qualified(Some(element), name) == Some(value)
    }

    // BothBlank
    fn both_blank(&self, element: NodeId, name: &str) -> bool {
        blank(self.plain(Some(element), name)) && blank(self.qualified(Some(element), name))
    }

    // Child(element, position): the position-th child element (1-based).
    fn child(&self, element: Option<NodeId>, position: usize) -> Option<NodeId> {
        let element = element?;
        self.src.elements(element).into_iter().nth(position.saturating_sub(1))
    }

    fn preceding_sibling(&self, element: NodeId) -> Option<NodeId> {
        self.src.previous_element(element)
    }

    fn string_value(&self, element: Option<NodeId>) -> String {
        element.map(|element| self.src.value(element)).unwrap_or_default()
    }

    fn parent(&self, element: NodeId) -> Option<NodeId> {
        self.src.parent(element)
    }

    // FirstText: the first text node of an element as XPath sees it.
    fn first_text(&self, element: NodeId) -> Option<String> {
        let mut builder = String::new();
        let mut found = false;
        for &node in self.src.child_nodes(element) {
            if let Some(text) = self.src.text(node) {
                builder.push_str(text);
                found = true;
            } else if found {
                break;
            }
        }
        if found {
            Some(builder)
        } else {
            None
        }
    }

    // ---- output construction ---------------------------------------------------

    fn add_part(&mut self, element: NodeId, part: Part) {
        match part {
            Part::Node(node) => self.out.append(element, node),
            Part::Nodes(nodes) => {
                for node in nodes {
                    self.out.append(element, node);
                }
            }
            Part::Attr(ns, local, value) => self.out.set_attr_ns(element, ns, local, Some(&value)),
            Part::Text(text) => self.out.add_string(element, &text),
            Part::Nothing => {}
        }
    }

    fn add_all(&mut self, element: NodeId, nodes: Vec<NodeId>) {
        for node in nodes {
            self.out.append(element, node);
        }
    }

    // Om(name, content...)
    fn om(&mut self, name: &str, parts: Vec<Part>) -> NodeId {
        let element = self.out.new_element(Name::new(M, name));
        for part in parts {
            self.add_part(element, part);
        }
        element
    }

    // Val(value)
    fn val(value: impl Into<String>) -> Part {
        Part::Attr(M, "val", value.into())
    }

    fn w_element(&mut self, name: &str, parts: Vec<Part>) -> NodeId {
        let element = self.out.new_element(Name::new(W, name));
        for part in parts {
            self.add_part(element, part);
        }
        element
    }

    // ---- Template dispatch -------------------------------------------------------

    fn apply(&mut self, element: NodeId) -> Result<Vec<NodeId>, String> {
        let name = self.src.name(element).unwrap();
        if name.ns != MML {
            return self.apply_children(element);
        }
        if self.is_native_accent(element) {
            return Ok(vec![self.native_accent(element)?]);
        }
        let local = name.local.clone();
        match local.as_str() {
            "mspace" => {
                if self.src.attr(element, SPACE_TEXT_ATTRIBUTE).is_some() {
                    Ok(self.space(element))
                } else {
                    self.apply_children(element)
                }
            }
            "mrow" | "mstyle" => self.row(element),
            "mi" | "mn" | "mo" | "ms" | "mtext" => {
                if self.has_glyph(element) {
                    self.glyph_token(element)
                } else {
                    self.token(element)
                }
            }
            "mfrac" => Ok(vec![self.fraction(element)?]),
            "menclose" | "msqrt" => self.enclosure(element),
            "mroot" => Ok(vec![self.root(element)?]),
            "munder" => Ok(vec![self.under(element)?]),
            "mover" => Ok(vec![self.over(element)?]),
            "munderover" => Ok(vec![self.under_over(element)?]),
            "mfenced" => Ok(vec![self.fenced(element)?]),
            "msub" => Ok(vec![self.script(element, "msub")?]),
            "msup" => Ok(vec![self.script(element, "msup")?]),
            "msubsup" => Ok(vec![self.script(element, "msubsup")?]),
            "mmultiscripts" => self.multi_scripts(element),
            "mtable" => Ok(vec![self.table(element)?]),
            "mglyph" => Ok(vec![self.standalone_glyph(element)]),
            "mpadded" => self.padded(element),
            "mphantom" => Ok(vec![self.phantom(element)?]),
            _ => self.apply_children(element),
        }
    }

    fn apply_children(&mut self, element: NodeId) -> Result<Vec<NodeId>, String> {
        let mut nodes = Vec::new();
        for child in self.src.elements(element) {
            nodes.extend(self.apply(child)?);
        }
        Ok(nodes)
    }

    fn apply_optional(&mut self, element: Option<NodeId>) -> Result<Vec<NodeId>, String> {
        match element {
            None => Ok(Vec::new()),
            Some(element) => self.apply(element),
        }
    }

    fn apply_all(&mut self, elements: Vec<NodeId>) -> Result<Vec<NodeId>, String> {
        let mut nodes = Vec::new();
        for element in elements {
            nodes.extend(self.apply(element)?);
        }
        Ok(nodes)
    }

    // ---- Rows, functions, linear fractions -----------------------------------------

    fn row(&mut self, row: NodeId) -> Result<Vec<NodeId>, String> {
        if self.is_nary_argument(row) {
            return Ok(Vec::new());
        }
        if self.is_linear_fraction(Some(row)) {
            return Ok(vec![self.linear_fraction(row, row)?]);
        }
        if self.is_function(Some(row)) {
            return Ok(vec![self.function(row)?]);
        }
        self.apply_children(row)
    }

    fn is_nary_argument(&self, element: NodeId) -> bool {
        let previous = self.preceding_sibling(element);
        self.is_any(previous, &["munder", "mover", "munderover", "msub", "msup", "msubsup"])
            && self.is_nary(self.child(previous, 1))
    }

    fn is_linear_fraction(&self, element: Option<NodeId>) -> bool {
        self.is(element, "mrow")
            && self.src.elements(element.unwrap()).len() == 3
            && self.is(self.child(element, 2), "mo")
            && normalize_space(&self.string_value(self.child(element, 2))) == "/"
    }

    fn is_function(&self, element: Option<NodeId>) -> bool {
        element.is_some()
            && self.src.elements(element.unwrap()).len() == 3
            && self.is(element, "mrow")
            && self.is(self.child(element, 2), "mo")
            && normalize_space(&self.string_value(self.child(element, 2))) == "\u{2061}"
    }

    fn linear_fraction(&mut self, fraction: NodeId, context: NodeId) -> Result<NodeId, String> {
        let type_element = self.om("type", vec![Self::val("lin")]);
        let control = self.control_color(context)?;
        let properties = self.om("fPr", vec![type_element.into(), control.into()]);
        let num_props = self.argument_properties(context);
        let num_content = self.apply_optional(self.child(Some(fraction), 1))?;
        let num = self.om("num", vec![num_props.into(), num_content.into()]);
        let den_props = self.argument_properties(context);
        let den_content = self.apply_optional(self.child(Some(fraction), 3))?;
        let den = self.om("den", vec![den_props.into(), den_content.into()]);
        Ok(self.om("f", vec![properties.into(), num.into(), den.into()]))
    }

    fn function(&mut self, function: NodeId) -> Result<NodeId, String> {
        let name_content = self.apply_optional(self.child(Some(function), 1))?;
        let name = self.om("fName", vec![name_content.into()]);
        let argument_content = self.apply_optional(self.child(Some(function), 3))?;
        let argument = self.om("e", vec![argument_content.into()]);
        Ok(self.om("func", vec![name.into(), argument.into()]))
    }

    // ArgumentProperties: m:argPr for an explicit mstyle script level 0, 1 or 2.
    fn argument_properties(&mut self, context: NodeId) -> Option<NodeId> {
        let mut styles = Vec::new();
        let mut current = Some(context);
        while let Some(element) = current {
            if self.src.is(element, MML, "mstyle") {
                styles.push(element);
            }
            current = self.src.parent(element);
        }
        let level = |value: Option<&str>| matches!(value, Some("0" | "1" | "2"));
        if !styles.iter().any(|&style| level(self.plain(Some(style), "scriptlevel")))
            && !styles.iter().any(|&style| level(self.qualified(Some(style), "scriptlevel")))
        {
            return None;
        }
        let nearest = styles.iter().copied().find(|&style| self.src.attr(style, "scriptlevel").is_some());
        let value = self
            .plain(nearest, "scriptlevel")
            .or_else(|| self.qualified(nearest, "scriptlevel"))
            .unwrap_or("")
            .to_string();
        let level_element = self.om("scrLvl", vec![Self::val(value)]);
        Some(self.om("argPr", vec![level_element.into()]))
    }

    // ControlColor: VisualTeX's text color for a fraction or radical control character.
    fn control_color(&mut self, context: NodeId) -> Result<Option<NodeId>, String> {
        let color = self.plain(Some(context), "mathcolor").unwrap_or("").to_string();
        if color.is_empty() {
            return Ok(None);
        }
        let word_color = self.word_color(&color)?;
        Ok(Some(self.om("ctrlPr", vec![word_color.into()])))
    }

    // WordColor
    fn word_color(&mut self, math_color: &str) -> Result<NodeId, String> {
        let value = substring1(math_color)?;
        let color = self.w_element("color", vec![Part::Attr(W, "val", value)]);
        Ok(self.w_element("rPr", vec![color.into()]))
    }

    // ---- Tokens and runs ------------------------------------------------------------

    fn is_non_glyph_token(&self, element: Option<NodeId>) -> bool {
        self.is_token(element) && !self.has_glyph(element.unwrap())
    }

    fn token(&mut self, token: NodeId) -> Result<Vec<NodeId>, String> {
        let start_of_run = self.src.previous_element(token).is_none()
            || !self.is_non_glyph_token(self.preceding_sibling(token));
        let parent = self.parent(token);
        let collect = !self.is_linear_fraction(parent)
            && !self.is_function(parent)
            && self.is_any(parent, &["mrow", "mstyle", "msqrt", "menclose", "math", "mphantom", "mtd", "maction"]);
        if collect {
            return if start_of_run { self.runs_from(token) } else { Ok(Vec::new()) };
        }

        // A token outside a run is always whitespace-normalized; literal text
        // only skips Office's operator rewrites.
        let text = normalize_space(&self.string_value(Some(token)));
        let variant = self.value(Some(token), "mathvariant");
        let style = self.value(Some(token), "fontstyle");
        let weight = self.value(Some(token), "fontweight");
        let normal_text = self.is(Some(token), "mtext");
        let properties = self.run_properties(&variant, &style, &weight, Some(token), normal_text)?;
        let literal = self.is_any(Some(token), &["mtext", "ms"]);
        let t = self.om("t", vec![Part::Text(output_text(&text, literal))]);
        Ok(vec![self.om("r", vec![properties.into(), t.into()])])
    }

    // RunsFrom: Office's CreateRunWithSameProp.
    fn runs_from(&mut self, first: NodeId) -> Result<Vec<NodeId>, String> {
        let mut runs = Vec::new();
        let mut current = first;
        loop {
            let siblings = self.src.elements_after_self(current);
            let limit = siblings.iter().position(|&sibling| self.ends_run(current, sibling));
            let members: Vec<NodeId> = match limit {
                None => siblings.iter().copied().filter(|&sibling| self.is_token(Some(sibling))).collect(),
                Some(limit) => siblings[..limit].to_vec(),
            };

            let literal = self.is_any(Some(current), &["mtext", "ms"]);
            let mut text = self.token_run_text(current, literal);
            for &member in &members {
                if self.is(Some(member), "ms") {
                    text.push_str(&self.quoted_literal(member));
                } else if self.is_any(Some(member), &["mtext", "ms"]) {
                    text.push_str(&self.string_value(Some(member)));
                } else {
                    text.push_str(&normalize_space(&self.string_value(Some(member))));
                }
            }
            let variant = self.value(Some(current), "mathvariant");
            let style = self.value(Some(current), "fontstyle");
            let weight = self.value(Some(current), "fontweight");
            let normal_text = self.is(Some(current), "mtext");
            let properties = self.run_properties(&variant, &style, &weight, Some(current), normal_text)?;
            let t = self.om("t", vec![Part::Text(output_text(&text, literal))]);
            runs.push(self.om("r", vec![properties.into(), t.into()]));

            let Some(limit) = limit else { break };
            let next = siblings[limit];
            if !self.is_token(Some(next)) || self.src.element(next, MML, "mglyph").is_some() {
                break;
            }
            current = next;
        }
        Ok(runs)
    }

    fn token_run_text(&self, token: NodeId, literal: bool) -> String {
        if self.is(Some(token), "ms") {
            self.quoted_literal(token)
        } else if literal {
            self.string_value(Some(token))
        } else {
            normalize_space(&self.string_value(Some(token)))
        }
    }

    fn quoted_literal(&self, ms: NodeId) -> String {
        let mut text = self.left_quote(ms);
        text.push_str(&normalize_space(&self.string_value(Some(ms))));
        text.push_str(&self.right_quote(ms));
        text
    }

    fn left_quote(&self, ms: NodeId) -> String {
        if self.both_blank(ms, "lquote") {
            "\"".to_string()
        } else {
            self.value(Some(ms), "lquote")
        }
    }

    fn right_quote(&self, ms: NodeId) -> String {
        if self.both_blank(ms, "rquote") {
            "\"".to_string()
        } else {
            self.value(Some(ms), "rquote")
        }
    }

    // EndsRun: true when sibling must not join the run that starts at first.
    fn ends_run(&self, first: NodeId, sibling: NodeId) -> bool {
        if self.plain(Some(sibling), "mathcolor").unwrap_or("") != self.plain(Some(first), "mathcolor").unwrap_or("") {
            return true;
        }
        if self.plain(Some(sibling), "mathvariant").unwrap_or("") != self.plain(Some(first), "mathvariant").unwrap_or("")
        {
            return true;
        }
        if self.plain(Some(sibling), RUN_ROLE_ATTRIBUTE).unwrap_or("")
            != self.plain(Some(first), RUN_ROLE_ATTRIBUTE).unwrap_or("")
        {
            return true;
        }
        if self.is(Some(sibling), "mo") && normalize_space(&self.string_value(Some(sibling))) == "\u{2061}" {
            return true;
        }
        if !self.is_token(Some(sibling)) {
            return true;
        }
        if self.has_glyph(sibling) {
            return true;
        }
        let first_is_text = self.is(Some(first), "mtext");
        if first_is_text != self.is(Some(sibling), "mtext") {
            return true;
        }

        let font = self.font_of(
            &self.value(Some(first), "mathvariant"),
            &self.value(Some(first), "fontstyle"),
            &self.value(Some(first), "fontweight"),
            FontNode::Element(first),
        );
        !(self.font_matches(&font, sibling) && self.family_matches(&self.value(Some(first), "font-family"), sibling))
    }

    fn font_matches(&self, font: &str, s: NodeId) -> bool {
        let no_variant = self.both_blank(s, "mathvariant");
        let not_bold = self.plain(Some(s), "fontweight") != Some("bold") && self.qualified(Some(s), "fontweight") != Some("bold");
        let bold_weight = self.attribute_is(s, "fontweight", "bold");
        let italic_style = self.attribute_is(s, "fontstyle", "italic");
        let length = len16(&normalize_space(&self.string_value(Some(s))));
        let multi_letter = self.is(Some(s), "mi") && length > 1;
        let single_letter = self.is(Some(s), "mi") && length <= 1;
        let numeral = self.is(Some(s), "mn") && !is_not_a_number(self.first_text(s).as_deref());
        let non_numeral = self.is(Some(s), "mn") && is_not_a_number(self.first_text(s).as_deref());
        let no_weight = self.both_blank(s, "fontweight");

        if let Some(plain_variant) = self.plain(Some(s), "mathvariant") {
            if plain_variant == font {
                return true;
            }
        }
        if let Some(qualified_variant) = self.qualified(Some(s), "mathvariant") {
            if qualified_variant == font {
                return true;
            }
        }
        if font == "normal"
            && (self.attribute_is(s, "mathvariant", "normal")
                || no_variant && ((self.attribute_is(s, "fontstyle", "normal") && not_bold) || multi_letter || non_numeral))
        {
            return true;
        }
        if font == "italic"
            && (self.attribute_is(s, "mathvariant", "italic")
                || no_variant && ((italic_style && not_bold) || numeral || self.is(Some(s), "mo") || single_letter))
        {
            return true;
        }
        if font == "bold"
            && (self.attribute_is(s, "mathvariant", "bold")
                || no_variant && bold_weight && (self.attribute_is(s, "fontstyle", "normal") || single_letter))
        {
            return true;
        }
        if (font == "bi" || font == "bold-italic")
            && (self.attribute_is(s, "mathvariant", "bold-italic")
                || no_variant
                    && ((bold_weight && italic_style)
                        || (bold_weight && (self.is(Some(s), "mn") || self.is(Some(s), "mo") || single_letter))))
        {
            return true;
        }
        if font.is_empty()
            && ((no_variant && self.both_blank(s, "fontstyle") && no_weight)
                || self.attribute_is(s, "mathvariant", "italic")
                || no_variant
                    && ((self.attribute_is(s, "fontweight", "normal") && italic_style)
                        || (no_weight && italic_style)
                        || (no_weight && self.both_blank(s, "fontstyle"))))
        {
            return true;
        }
        if font == "normal"
            && ((self.is(Some(s), "mi")
                && blank(self.plain(Some(s), "mathvariant"))
                && self.both_blank(s, "fontstyle")
                && no_weight
                && length > 1)
                || (self.is_any(Some(s), &["ms", "mtext"])
                    && blank(self.plain(Some(s), "mathvariant"))
                    && blank(self.plain(Some(s), "fontstyle"))
                    && blank(self.qualified(Some(s), "fontweight"))))
        {
            return true;
        }
        false
    }

    fn family_matches(&self, family: &str, s: NodeId) -> bool {
        self.plain(Some(s), "font-family").is_some_and(|plain| plain == family)
            || self.qualified(Some(s), "font-family").is_some_and(|qualified| qualified == family)
            || (family.is_empty() && self.both_blank(s, "font-family"))
    }

    // FontOf: Office's GetFontCur.
    fn font_of(&self, variant: &str, style: &str, weight: &str, node: FontNode) -> String {
        if !variant.is_empty() {
            return variant.to_string();
        }
        let element = match node {
            FontNode::None => return "italic".to_string(),
            FontNode::Element(element) => Some(element),
            FontNode::Text => None,
        };
        if (self.is(element, "mi") && len16(&normalize_space(&self.string_value(element))) <= 1)
            || (self.is(element, "mn") && !is_not_a_number(self.first_text(element.unwrap()).as_deref()))
            || self.is(element, "mo")
        {
            if style == "normal" && weight == "bold" {
                return "bold".to_string();
            }
            if style == "normal" {
                return "normal".to_string();
            }
            if weight == "bold" {
                return "bi".to_string();
            }
            return "italic".to_string();
        }
        if style == "italic" && weight == "bold" {
            return "bi".to_string();
        }
        if style == "italic" {
            return "italic".to_string();
        }
        if weight == "bold" {
            return "bold".to_string();
        }
        "normal".to_string()
    }

    // RunProperties: Office's CreateRunProp with VisualTeX's run color.
    fn run_properties(
        &mut self,
        variant: &str,
        style: &str,
        weight: &str,
        node: Option<NodeId>,
        normal_text: bool,
    ) -> Result<Vec<NodeId>, String> {
        let font_node = match node {
            None => FontNode::Text,
            Some(element) => FontNode::Element(element),
        };
        self.run_properties_for(variant, style, weight, font_node, normal_text)
    }

    fn run_properties_for(
        &mut self,
        variant: &str,
        style: &str,
        weight: &str,
        node: FontNode,
        normal_text: bool,
    ) -> Result<Vec<NodeId>, String> {
        let mut properties = Vec::new();
        let font = self.font_of(variant, style, weight, node);
        if normal_text || !font.is_empty() {
            let nor = if normal_text { Some(self.om("nor", vec![])) } else { None };
            let script_and_style = self.script_and_style(&font, normal_text);
            properties.push(self.om("rPr", vec![nor.into(), script_and_style.into()]));
        }
        let color = match node {
            FontNode::Element(element) => self.src.attr(element, "mathcolor").unwrap_or("").to_string(),
            _ => String::new(),
        };
        if !color.is_empty() {
            properties.push(self.word_color(&color)?);
        }
        Ok(properties)
    }

    // ScriptAndStyle: Office's CreateMathScrStyProp with VisualTeX's monospace and italic.
    fn script_and_style(&mut self, font: &str, normal_text: bool) -> Vec<NodeId> {
        let pairs: &[(&str, &str)] = match font {
            "normal" if !normal_text => &[("sty", "p")],
            "bold" => &[("sty", "b")],
            "monospace" => &[("scr", "monospace"), ("sty", "p")],
            "italic" => &[("sty", "i")],
            "script" => &[("scr", "script")],
            "bold-script" => &[("scr", "script"), ("sty", "b")],
            "double-struck" => &[("scr", "double-struck"), ("sty", "p")],
            "fraktur" => &[("scr", "fraktur"), ("sty", "p")],
            "bold-fraktur" => &[("scr", "fraktur"), ("sty", "b")],
            "sans-serif" => &[("scr", "sans-serif"), ("sty", "p")],
            "bold-sans-serif" => &[("scr", "sans-serif"), ("sty", "b")],
            "sans-serif-italic" => &[("scr", "sans-serif")],
            "sans-serif-bold-italic" => &[("scr", "sans-serif"), ("sty", "bi")],
            "bi" | "bold-italic" => &[("sty", "bi")],
            _ => &[],
        };
        pairs.iter().map(|&(name, value)| self.om(name, vec![Self::val(value)])).collect()
    }

    fn glyph_token(&mut self, token: NodeId) -> Result<Vec<NodeId>, String> {
        let variant = self.value(Some(token), "mathvariant");
        let style = self.value(Some(token), "fontstyle");
        let weight = self.value(Some(token), "fontweight");
        let normal_text = self.is(Some(token), "mtext");
        let literal = self.is_any(Some(token), &["mtext", "ms"]);
        let mut nodes = Vec::new();
        if self.is(Some(token), "ms") {
            let properties = self.run_properties(&variant, &style, &weight, Some(token), normal_text)?;
            let quote = self.left_quote(token);
            let t = self.om("t", vec![Part::Text(output_text(&quote, literal))]);
            nodes.push(self.om("r", vec![properties.into(), t.into()]));
        }
        for &item in self.src.child_nodes(token) {
            let text;
            let force_normal;
            let font_node;
            if self.src.is(item, MML, "mglyph") {
                text = normalize_space(&self.value(Some(item), "alt"));
                force_normal = true;
                font_node = FontNode::Element(item);
            } else if let Some(text_value) = self.src.text(item) {
                text = normalize_space(text_value);
                force_normal = false;
                font_node = FontNode::Text;
            } else {
                continue;
            }
            if text.is_empty() {
                continue;
            }
            let properties = self.run_properties_for(&variant, &style, &weight, font_node, force_normal || normal_text)?;
            let t = self.om("t", vec![Part::Text(output_text(&text, false))]);
            nodes.push(self.om("r", vec![properties.into(), t.into()]));
        }
        if self.is(Some(token), "ms") {
            let properties = self.run_properties(&variant, &style, &weight, Some(token), normal_text)?;
            let quote = self.right_quote(token);
            let t = self.om("t", vec![Part::Text(output_text(&quote, literal))]);
            nodes.push(self.om("r", vec![properties.into(), t.into()]));
        }
        Ok(nodes)
    }

    fn standalone_glyph(&mut self, glyph: NodeId) -> NodeId {
        let nor = self.om("nor", vec![]);
        let properties = self.om("rPr", vec![nor.into()]);
        let text = output_text(&normalize_space(&self.value(Some(glyph), "alt")), false);
        let t = self.om("t", vec![Part::Text(text)]);
        self.om("r", vec![properties.into(), t.into()])
    }

    fn space(&mut self, space: NodeId) -> Vec<NodeId> {
        let text = self.src.attr(space, SPACE_TEXT_ATTRIBUTE).unwrap_or("").to_string();
        if text.is_empty() {
            return Vec::new();
        }
        let t = self.om("t", vec![Part::Attr(XML_NS, "space", "preserve".to_string()), Part::Text(text)]);
        vec![self.om("r", vec![t.into()])]
    }

    // ---- Fractions, radicals, enclosures ----------------------------------------------

    fn fraction(&mut self, fraction: NodeId) -> Result<NodeId, String> {
        let thickness = lower_ascii(&self.value(Some(fraction), "linethickness"));
        let bar = thickness.is_empty()
            || matches!(thickness.as_str(), "thin" | "medium" | "thick")
            || thickness.chars().any(|c| ('1'..='9').contains(&c));
        let fraction_type = if !bar {
            "noBar"
        } else if self.attribute_is(fraction, "bevelled", "true") {
            "skw"
        } else {
            "bar"
        };
        let type_element = self.om("type", vec![Self::val(fraction_type)]);
        let control = self.control_color(fraction)?;
        let properties = self.om("fPr", vec![type_element.into(), control.into()]);
        let num_props = self.argument_properties(fraction);
        let num_content = self.apply_optional(self.child(Some(fraction), 1))?;
        let num = self.om("num", vec![num_props.into(), num_content.into()]);
        let den_props = self.argument_properties(fraction);
        let den_content = self.apply_optional(self.child(Some(fraction), 2))?;
        let den = self.om("den", vec![den_props.into(), den_content.into()]);
        Ok(self.om("f", vec![properties.into(), num.into(), den.into()]))
    }

    fn enclosure(&mut self, enclosure: NodeId) -> Result<Vec<NodeId>, String> {
        let notation = lower_ascii(&self.value(Some(enclosure), "notation"));
        if notation == "radical" || notation.is_empty() || self.is(Some(enclosure), "msqrt") {
            let deg_hide = self.om("degHide", vec![Self::val("on")]);
            let control = self.control_color(enclosure)?;
            let properties = self.om("radPr", vec![deg_hide.into(), control.into()]);
            let deg_props = self.argument_properties(enclosure);
            let degree = self.om("deg", vec![deg_props.into()]);
            let e_props = self.argument_properties(enclosure);
            let content = self.apply_children(enclosure)?;
            let e = self.om("e", vec![e_props.into(), content.into()]);
            return Ok(vec![self.om("rad", vec![properties.into(), degree.into(), e.into()])]);
        }
        if notation == "actuarial" || notation == "longdiv" {
            return Ok(Vec::new());
        }

        let has = |part: &str| notation.contains(part);
        let boxed = has("box") || has("circle") || has("roundedbox");
        let top = has("top");
        let bottom = has("bottom");
        let left = has("left");
        let right = has("right");
        let strike_h = has("horizontalstrike");
        let strike_v = has("verticalstrike");
        let strike_bltr = has("updiagonalstrike");
        let strike_tlbr = has("downdiagonalstrike");
        let mut properties: Option<NodeId> = None;
        if strike_h || strike_v || strike_bltr || strike_tlbr || (!boxed && !(top && bottom && left && right)) {
            let element = self.om("borderBoxPr", vec![]);
            if !boxed {
                if !top {
                    let hide = self.om("hideTop", vec![Self::val("on")]);
                    self.out.append(element, hide);
                }
                if !bottom {
                    let hide = self.om("hideBot", vec![Self::val("on")]);
                    self.out.append(element, hide);
                }
                if !left {
                    let hide = self.om("hideLeft", vec![Self::val("on")]);
                    self.out.append(element, hide);
                }
                if !right {
                    let hide = self.om("hideRight", vec![Self::val("on")]);
                    self.out.append(element, hide);
                }
            }
            if strike_h {
                let strike = self.om("strikeH", vec![Self::val("on")]);
                self.out.append(element, strike);
            }
            if strike_v {
                let strike = self.om("strikeV", vec![Self::val("on")]);
                self.out.append(element, strike);
            }
            if strike_bltr {
                let strike = self.om("strikeBLTR", vec![Self::val("on")]);
                self.out.append(element, strike);
            }
            if strike_tlbr {
                let strike = self.om("strikeTLBR", vec![Self::val("on")]);
                self.out.append(element, strike);
            }
            properties = Some(element);
        }
        let e_props = self.argument_properties(enclosure);
        let content = self.apply_children(enclosure)?;
        let e = self.om("e", vec![e_props.into(), content.into()]);
        Ok(vec![self.om("borderBox", vec![properties.into(), e.into()])])
    }

    fn root(&mut self, root: NodeId) -> Result<NodeId, String> {
        let deg_hide = self.om("degHide", vec![Self::val("off")]);
        let control = self.control_color(root)?;
        let properties = self.om("radPr", vec![deg_hide.into(), control.into()]);
        let deg_props = self.argument_properties(root);
        let deg_content = self.apply_optional(self.child(Some(root), 2))?;
        let degree = self.om("deg", vec![deg_props.into(), deg_content.into()]);
        let e_props = self.argument_properties(root);
        let e_content = self.apply_optional(self.child(Some(root), 1))?;
        let e = self.om("e", vec![e_props.into(), e_content.into()]);
        Ok(self.om("rad", vec![properties.into(), degree.into(), e.into()]))
    }

    // ---- Under, over, scripts and n-ary operators ------------------------------------

    fn accent_of(&self, element: Option<NodeId>, under: bool) -> String {
        if element.is_none() {
            return String::new();
        }
        let name = if under { "accentunder" } else { "accent" };
        lower_ascii(self.plain(element, name).or_else(|| self.qualified(element, name)).unwrap_or(""))
    }

    fn is_nary(&self, operand: Option<NodeId>) -> bool {
        let Some(operand) = operand else { return false };
        let text = normalize_space(&self.string_value(Some(operand)));
        if !NARY_OPERATORS.contains(&text.as_str()) {
            return false;
        }
        let parent = self.parent(operand);
        let accent =
            if self.is(parent, "munder") { self.accent_of(parent, true) } else { self.accent_of(parent, false) };
        if accent == "true" {
            return false;
        }
        let subtree = self.src.descendants_and_self(operand);
        self.is(subtree.last().copied(), "mo")
            && subtree.iter().all(|&element| self.is_any(Some(element), &["mo", "mstyle", "mrow"]))
    }

    fn nary(&mut self, owner: NodeId, kind: &str, sub: Option<NodeId>, sup: Option<NodeId>) -> Result<NodeId, String> {
        let first = self.child(Some(owner), 1);
        let symbol = normalize_space(&self.string_value(first));
        let grow_value = lower_ascii(&self.value(first, "stretchy"));
        let grow = if grow_value == "true" {
            "1"
        } else if grow_value == "false" {
            "0"
        } else if GROWING_NARY_OPERATORS.contains(&symbol.as_str()) {
            "1"
        } else {
            "0"
        };
        let under_over = matches!(kind, "munder" | "mover" | "munderover");
        let chr = self.om("chr", vec![Self::val(symbol)]);
        let lim_loc = self.om("limLoc", vec![Self::val(if under_over { "undOvr" } else { "subSup" })]);
        let grow_element = self.om("grow", vec![Self::val(grow)]);
        let sub_hide = self.om("subHide", vec![Self::val(if matches!(kind, "mover" | "msup") { "on" } else { "off" })]);
        let sup_hide = self.om("supHide", vec![Self::val(if matches!(kind, "munder" | "msub") { "on" } else { "off" })]);
        let properties = self.om(
            "naryPr",
            vec![chr.into(), lim_loc.into(), grow_element.into(), sub_hide.into(), sup_hide.into()],
        );
        let sub_props = self.argument_properties(owner);
        let sub_content = self.apply_optional(sub)?;
        let sub_element = self.om("sub", vec![sub_props.into(), sub_content.into()]);
        let sup_props = self.argument_properties(owner);
        let sup_content = self.apply_optional(sup)?;
        let sup_element = self.om("sup", vec![sup_props.into(), sup_content.into()]);
        let e_props = self.argument_properties(owner);
        let argument = self.src.elements_after_self(owner).first().copied();
        let e_content = self.nary_argument(argument, owner)?;
        let e = self.om("e", vec![e_props.into(), e_content.into()]);
        Ok(self.om("nary", vec![properties.into(), sub_element.into(), sup_element.into(), e.into()]))
    }

    // NaryArgument: Office's NaryHandleMrowMstyle.
    fn nary_argument(&mut self, argument: Option<NodeId>, owner: NodeId) -> Result<Vec<NodeId>, String> {
        if self.is(argument, "mrow") {
            return if self.is_linear_fraction(argument) {
                Ok(vec![self.linear_fraction(argument.unwrap(), owner)?])
            } else {
                self.apply_children(argument.unwrap())
            };
        }
        if self.is(argument, "mstyle") {
            return self.apply_children(argument.unwrap());
        }
        Ok(Vec::new())
    }

    fn is_bar(&self, element: NodeId, under: bool) -> bool {
        if self.accent_of(Some(element), under) == "true" {
            return false;
        }
        let mark = self.child(Some(element), 2);
        if !self.is(mark, "mo") {
            return false;
        }
        let text = self.string_value(mark);
        if under {
            text == "\u{0332}" || text == "_"
        } else {
            text == "\u{0305}" || text == "\u{00AF}"
        }
    }

    // SecondEqualsSomeOperator: child::*[2] = mml:mo compares string values with every mo child.
    fn second_equals_some_operator(&self, element: NodeId) -> bool {
        let Some(second) = self.child(Some(element), 2) else { return false };
        let value = self.string_value(Some(second));
        self.src.elements_named(element, MML, "mo").into_iter().any(|mo| self.string_value(Some(mo)) == value)
    }

    fn is_accent(&self, over: NodeId) -> bool {
        let accent = self.accent_of(Some(over), false);
        let second = self.child(Some(over), 2);
        let mut mark_accent = String::new();
        if self.second_equals_some_operator(over) {
            if let Some(plain) = self.src.attr(second.unwrap(), "accent") {
                mark_accent = lower_ascii(plain);
            } else if let Some(qualified) = self.src.attr_ns(second.unwrap(), MML, "accent") {
                mark_accent = lower_ascii(qualified);
            }
        }
        let accented = mark_accent == "true" || (mark_accent.is_empty() && accent == "true");
        accented && self.second_equals_some_operator(over) && len16(&self.string_value(second)) <= 1
    }

    fn is_group_character(&self, element: NodeId, under: bool) -> bool {
        if self.accent_of(Some(element), under) != "false" {
            return false;
        }
        let children = self.src.elements(element);
        if children.len() != 2 {
            return false;
        }
        if !((self.is(Some(children[0]), "mrow") && self.is(Some(children[1]), "mo"))
            || (self.is(Some(children[0]), "mo") && self.is(Some(children[1]), "mrow")))
        {
            return false;
        }
        len16(&self.string_value(self.src.element(element, MML, "mo"))) <= 1
    }

    fn group_character(&mut self, element: NodeId, position: &str, justification: &str) -> Result<NodeId, String> {
        let chr_value = self.string_value(self.src.element(element, MML, "mo"));
        let chr = self.om("chr", vec![Self::val(chr_value)]);
        let pos = self.om("pos", vec![Self::val(position)]);
        let vert_jc = self.om("vertJc", vec![Self::val(justification)]);
        let properties = self.om("groupChrPr", vec![chr.into(), pos.into(), vert_jc.into()]);
        let rows = self.src.elements_named(element, MML, "mrow");
        let content = self.apply_all(rows)?;
        let e = self.om("e", vec![content.into()]);
        Ok(self.om("groupChr", vec![properties.into(), e.into()]))
    }

    fn under(&mut self, under: NodeId) -> Result<NodeId, String> {
        if self.is_nary(self.child(Some(under), 1)) {
            let sub = self.child(Some(under), 2);
            return self.nary(under, "munder", sub, None);
        }
        if self.is_bar(under, true) {
            let pos = self.om("pos", vec![Self::val("bot")]);
            let properties = self.om("barPr", vec![pos.into()]);
            let e_props = self.argument_properties(under);
            let content = self.apply_optional(self.child(Some(under), 1))?;
            let e = self.om("e", vec![e_props.into(), content.into()]);
            return Ok(self.om("bar", vec![properties.into(), e.into()]));
        }
        if self.is_group_character(under, true) {
            let position = if self.is(self.child(Some(under), 1), "mrow") { "bot" } else { "top" };
            return self.group_character(under, position, "top");
        }
        let e_props = self.argument_properties(under);
        let e_content = self.apply_optional(self.child(Some(under), 1))?;
        let e = self.om("e", vec![e_props.into(), e_content.into()]);
        let lim_props = self.argument_properties(under);
        let lim_content = self.apply_optional(self.child(Some(under), 2))?;
        let lim = self.om("lim", vec![lim_props.into(), lim_content.into()]);
        Ok(self.om("limLow", vec![e.into(), lim.into()]))
    }

    fn over(&mut self, over: NodeId) -> Result<NodeId, String> {
        if self.is_nary(self.child(Some(over), 1)) {
            let sup = self.child(Some(over), 2);
            return self.nary(over, "mover", None, sup);
        }
        if self.is_bar(over, false) {
            let pos = self.om("pos", vec![Self::val("top")]);
            let properties = self.om("barPr", vec![pos.into()]);
            let e_props = self.argument_properties(over);
            let content = self.apply_optional(self.child(Some(over), 1))?;
            let e = self.om("e", vec![e_props.into(), content.into()]);
            return Ok(self.om("bar", vec![properties.into(), e.into()]));
        }
        if self.is_accent(over) {
            let chr_value = to_upper_combining(&self.string_value(self.child(Some(over), 2)));
            let chr = self.om("chr", vec![Self::val(chr_value)]);
            let properties = self.om("accPr", vec![chr.into()]);
            let e_props = self.argument_properties(over);
            let content = self.apply_optional(self.child(Some(over), 1))?;
            let e = self.om("e", vec![e_props.into(), content.into()]);
            return Ok(self.om("acc", vec![properties.into(), e.into()]));
        }
        if self.is_group_character(over, false) {
            let position = if self.is(self.child(Some(over), 1), "mrow") { "top" } else { "bot" };
            return self.group_character(over, position, "bot");
        }
        let e_props = self.argument_properties(over);
        let e_content = self.apply_optional(self.child(Some(over), 1))?;
        let e = self.om("e", vec![e_props.into(), e_content.into()]);
        let lim_props = self.argument_properties(over);
        let lim_content = self.apply_optional(self.child(Some(over), 2))?;
        let lim = self.om("lim", vec![lim_props.into(), lim_content.into()]);
        Ok(self.om("limUpp", vec![e.into(), lim.into()]))
    }

    // IsNativeAccent: a one-character styled mark whose control formatting
    // carries the mark's own bold, italic and color.
    fn is_native_accent(&self, element: NodeId) -> bool {
        if !self.is(Some(element), "mover") || self.plain(Some(element), "accent") != Some("true") {
            return false;
        }
        let children = self.src.elements(element);
        children.len() == 2
            && len16(&self.string_value(Some(children[1]))) == 1
            && self.is(Some(children[1]), "mo")
            && matches!(self.plain(Some(children[1]), "mathvariant"), Some("normal" | "bold" | "italic" | "bold-italic"))
    }

    fn native_accent(&mut self, over: NodeId) -> Result<NodeId, String> {
        let mark = self.child(Some(over), 2).unwrap();
        let variant = self.plain(Some(mark), "mathvariant");
        let color = self.plain(Some(mark), "mathcolor").unwrap_or("").to_string();
        let chr_value = self.string_value(Some(mark));
        let chr = self.om("chr", vec![Self::val(chr_value)]);
        let bold = matches!(variant, Some("bold" | "bold-italic"));
        let italic = matches!(variant, Some("italic" | "bold-italic"));
        let b = self.w_element("b", vec![Part::Attr(W, "val", if bold { "1" } else { "0" }.to_string())]);
        let i = self.w_element("i", vec![Part::Attr(W, "val", if italic { "1" } else { "0" }.to_string())]);
        let color_element = if !color.is_empty() {
            let value = substring1(&color)?;
            Some(self.w_element("color", vec![Part::Attr(W, "val", value)]))
        } else {
            None
        };
        let run_properties = self.w_element("rPr", vec![b.into(), i.into(), color_element.into()]);
        let control = self.om("ctrlPr", vec![run_properties.into()]);
        let properties = self.om("accPr", vec![chr.into(), control.into()]);
        let content = self.apply_optional(self.child(Some(over), 1))?;
        let e = self.om("e", vec![content.into()]);
        Ok(self.om("acc", vec![properties.into(), e.into()]))
    }

    fn under_over(&mut self, element: NodeId) -> Result<NodeId, String> {
        if self.is_nary(self.child(Some(element), 1)) {
            let sub = self.child(Some(element), 2);
            let sup = self.child(Some(element), 3);
            return self.nary(element, "munderover", sub, sup);
        }
        let outer_e_props = self.argument_properties(element);
        let inner_e_props = self.argument_properties(element);
        let inner_e_content = self.apply_optional(self.child(Some(element), 1))?;
        let inner_e = self.om("e", vec![inner_e_props.into(), inner_e_content.into()]);
        let inner_lim_props = self.argument_properties(element);
        let inner_lim_content = self.apply_optional(self.child(Some(element), 2))?;
        let inner_lim = self.om("lim", vec![inner_lim_props.into(), inner_lim_content.into()]);
        let lim_low = self.om("limLow", vec![inner_e.into(), inner_lim.into()]);
        let outer_e = self.om("e", vec![outer_e_props.into(), lim_low.into()]);
        let outer_lim_props = self.argument_properties(element);
        let outer_lim_content = self.apply_optional(self.child(Some(element), 3))?;
        let outer_lim = self.om("lim", vec![outer_lim_props.into(), outer_lim_content.into()]);
        Ok(self.om("limUpp", vec![outer_e.into(), outer_lim.into()]))
    }

    fn script(&mut self, script: NodeId, kind: &str) -> Result<NodeId, String> {
        let base = self.child(Some(script), 1);
        if self.is_nary(base) {
            return match kind {
                "msub" => {
                    let sub = self.child(Some(script), 2);
                    self.nary(script, kind, sub, None)
                }
                "msup" => {
                    let sup = self.child(Some(script), 2);
                    self.nary(script, kind, None, sup)
                }
                _ => {
                    let sub = self.child(Some(script), 2);
                    let sup = self.child(Some(script), 3);
                    self.nary(script, kind, sub, sup)
                }
            };
        }
        let slots: &[(&str, usize)] = match kind {
            "msub" => &[("e", 1), ("sub", 2)],
            "msup" => &[("e", 1), ("sup", 2)],
            _ => &[("e", 1), ("sub", 2), ("sup", 3)],
        };
        let mut parts = Vec::new();
        for &(name, position) in slots {
            let props = self.argument_properties(script);
            let content = self.apply_optional(self.child(Some(script), position))?;
            parts.push(Part::Node(self.om(name, vec![props.into(), content.into()])));
        }
        let owner = match kind {
            "msub" => "sSub",
            "msup" => "sSup",
            _ => "sSubSup",
        };
        Ok(self.om(owner, parts))
    }

    fn multi_scripts(&mut self, multi: NodeId) -> Result<Vec<NodeId>, String> {
        let children = self.src.elements(multi);
        let prescripts_index = children.iter().position(|&child| self.is(Some(child), "mprescripts"));
        let prescripts: Vec<NodeId> = match prescripts_index {
            None => Vec::new(),
            Some(index) => children[index + 1..].to_vec(),
        };
        let prescript_count = prescripts.iter().filter(|&&child| !self.is(Some(child), "none")).count() as i64;
        let before_prescripts = |index: usize| prescripts_index.is_none_or(|p| index < p);
        let superscripts = (0..children.len())
            .filter(|&index| before_prescripts(index) && (index + 1) % 2 == 1 && !self.is(Some(children[index]), "none"))
            .count() as i64
            - 1;
        let subscripts = (0..children.len())
            .filter(|&index| before_prescripts(index) && (index + 1) % 2 == 0 && !self.is(Some(children[index]), "none"))
            .count() as i64;
        let script_count = superscripts + subscripts;
        let scripts = (0..children.len()).filter(|&index| before_prescripts(index)).count() as i64 - 1;

        if prescript_count <= 0 && script_count <= 0 {
            return self.apply_optional(self.child(Some(multi), 1));
        }
        if prescript_count <= 0 {
            let rest: Vec<NodeId> = children.iter().skip(1).copied().collect();
            if superscripts > 0 && subscripts > 0 {
                let base = self.base_slot(multi)?;
                let split = self.split(multi, &rest)?;
                let mut parts = vec![Part::Node(base)];
                parts.extend(split.into_iter().map(Part::Node));
                return Ok(vec![self.om("sSubSup", parts)]);
            }
            if subscripts > 0 {
                let base = self.base_slot(multi)?;
                let content = self.apply_all(rest)?;
                let sub = self.om("sub", vec![content.into()]);
                return Ok(vec![self.om("sSub", vec![base.into(), sub.into()])]);
            }
            let base = self.base_slot(multi)?;
            let content = self.apply_all(rest)?;
            let sup = self.om("sup", vec![content.into()]);
            return Ok(vec![self.om("sSup", vec![base.into(), sup.into()])]);
        }
        if script_count <= 0 {
            let base = self.base_slot(multi)?;
            let split = self.split(multi, &prescripts)?;
            let mut parts = vec![Part::Node(base)];
            parts.extend(split.into_iter().map(Part::Node));
            return Ok(vec![self.om("sPre", parts)]);
        }

        let take = scripts.max(0) as usize;
        let postscripts: Vec<NodeId> = children.iter().skip(1).take(take).copied().collect();
        let inner = if superscripts > 0 && subscripts > 0 {
            let base = self.base_slot(multi)?;
            let split = self.split(multi, &postscripts)?;
            let mut parts = vec![Part::Node(base)];
            parts.extend(split.into_iter().map(Part::Node));
            self.om("sSubSup", parts)
        } else if subscripts > 0 {
            let base = self.base_slot(multi)?;
            let content = self.apply_all(postscripts)?;
            let sub = self.om("sub", vec![content.into()]);
            self.om("sSub", vec![base.into(), sub.into()])
        } else {
            let base = self.base_slot(multi)?;
            let content = self.apply_all(postscripts)?;
            let sup = self.om("sup", vec![content.into()]);
            self.om("sSup", vec![base.into(), sup.into()])
        };
        let e = self.om("e", vec![inner.into()]);
        let split = self.split(multi, &prescripts)?;
        let mut parts = vec![Part::Node(e)];
        parts.extend(split.into_iter().map(Part::Node));
        Ok(vec![self.om("sPre", parts)])
    }

    // BaseSlot()
    fn base_slot(&mut self, multi: NodeId) -> Result<NodeId, String> {
        let props = self.argument_properties(multi);
        let content = self.apply_optional(self.child(Some(multi), 1))?;
        Ok(self.om("e", vec![props.into(), content.into()]))
    }

    // Split(nodes): odd positions to m:sub, even positions to m:sup.
    fn split(&mut self, multi: NodeId, nodes: &[NodeId]) -> Result<Vec<NodeId>, String> {
        let odd: Vec<NodeId> = nodes.iter().enumerate().filter(|(index, _)| (index + 1) % 2 == 1).map(|(_, &n)| n).collect();
        let even: Vec<NodeId> = nodes.iter().enumerate().filter(|(index, _)| (index + 1) % 2 == 0).map(|(_, &n)| n).collect();
        let sub_props = self.argument_properties(multi);
        let sub_content = self.apply_all(odd)?;
        let sub = self.om("sub", vec![sub_props.into(), sub_content.into()]);
        let sup_props = self.argument_properties(multi);
        let sup_content = self.apply_all(even)?;
        let sup = self.om("sup", vec![sup_props.into(), sup_content.into()]);
        Ok(vec![sub, sup])
    }

    // ---- Fences --------------------------------------------------------------------

    fn fenced(&mut self, fence: NodeId) -> Result<NodeId, String> {
        let valid = |this: &Self, name: &str| {
            this.src.attr(fence, name).is_some() || this.src.attr_ns(fence, MML, name).is_some()
        };
        let open_valid = valid(self, "open");
        let open = self.value(Some(fence), "open");
        let separators_valid = valid(self, "separators");
        let separator = first_scalar(&self.value(Some(fence), "separators"));
        let close_valid = valid(self, "close");
        let close = self.value(Some(fence), "close");

        let delimiter = self.om("d", vec![]);
        if (open_valid && open != "(") || (close_valid && close != ")") || separator != "|" {
            let properties = self.om("dPr", vec![]);
            if open_valid && open != "(" {
                let element = self.om("begChr", vec![Self::val(open.clone())]);
                self.out.append(properties, element);
            }
            if separator != "|" {
                let value = if separators_valid { separator.clone() } else { ",".to_string() };
                let element = self.om("sepChr", vec![Self::val(value)]);
                self.out.append(properties, element);
            }
            if close_valid && close != ")" {
                let element = self.om("endChr", vec![Self::val(close.clone())]);
                self.out.append(properties, element);
            }
            self.out.append(delimiter, properties);
        }
        for child in self.src.elements(fence) {
            let props = self.argument_properties(child);
            let content = self.apply(child)?;
            let e = self.om("e", vec![props.into(), content.into()]);
            self.out.append(delimiter, e);
        }
        Ok(delimiter)
    }

    // ---- Tables and equation arrays ------------------------------------------------

    fn is_equation_array(&self, table: NodeId) -> bool {
        let no_lines = |name: &str| {
            let plain = self.plain(Some(table), name);
            let qualified = self.qualified(Some(table), name);
            (blank(plain) || plain == Some("none")) && (blank(qualified) || qualified == Some("none"))
        };
        no_lines("frame")
            && no_lines("columnlines")
            && no_lines("rowlines")
            && self
                .src
                .elements_named(table, MML, "mtr")
                .into_iter()
                .all(|row| self.src.elements_named(row, MML, "mtd").len() == 1)
            && self.src.elements_named(table, MML, "mlabeledtr").is_empty()
    }

    fn table(&mut self, table: NodeId) -> Result<NodeId, String> {
        if self.is_equation_array(table) {
            let mut rows = Vec::new();
            for row in self.src.elements_named(table, MML, "mtr") {
                let mut items = Vec::new();
                for cell in self.src.elements_named(row, MML, "mtd") {
                    items.extend(self.src.elements(cell));
                }
                let content = self.equation_array_row(items)?;
                rows.push(self.om("e", vec![content.into()]));
            }
            return Ok(self.om("eqArr", vec![rows.into()]));
        }

        let mut columns = 0usize;
        for child in self.src.elements(table) {
            let count = if self.is(Some(child), "mlabeledtr") {
                self.src.elements(child).len().saturating_sub(1)
            } else if self.is(Some(child), "mtr") {
                self.src.elements(child).len()
            } else {
                1
            };
            if count > columns {
                columns = count;
            }
        }
        let base_jc = self.om("baseJc", vec![Self::val("center")]);
        let plc_hide = self.om("plcHide", vec![Self::val("on")]);
        let count = self.om("count", vec![Self::val(columns.to_string())]);
        let mc_jc = self.om("mcJc", vec![Self::val("center")]);
        let mc_pr = self.om("mcPr", vec![count.into(), mc_jc.into()]);
        let mc = self.om("mc", vec![mc_pr.into()]);
        let mcs = self.om("mcs", vec![mc.into()]);
        let properties = self.om("mPr", vec![base_jc.into(), plc_hide.into(), mcs.into()]);
        let matrix = self.om("m", vec![properties.into()]);
        for child in self.src.elements(table) {
            let row = self.om("mr", vec![]);
            let cells: Vec<NodeId> = if self.is(Some(child), "mtr") {
                self.src.elements(child)
            } else if self.is(Some(child), "mlabeledtr") {
                self.src.elements(child).into_iter().skip(1).collect()
            } else {
                vec![child]
            };
            let mut used = 0usize;
            for cell in cells {
                let content = self.apply(cell)?;
                let e = self.om("e", vec![content.into()]);
                self.out.append(row, e);
                used += 1;
            }
            for _ in used..columns {
                let e = self.om("e", vec![]);
                self.out.append(row, e);
            }
            self.out.append(matrix, row);
        }
        Ok(matrix)
    }

    fn has_alignment_descendant(&self, element: NodeId) -> bool {
        self.src
            .descendants(element)
            .into_iter()
            .any(|descendant| self.is_any(Some(descendant), &["maligngroup", "malignmark"]))
    }

    fn first_alignment_already_found(&self, element: Option<NodeId>) -> bool {
        let mut element = element;
        while let Some(current) = element {
            if self.src.elements_before_self(current).into_iter().any(|sibling| {
                self.is_any(Some(sibling), &["maligngroup", "malignmark"]) || self.has_alignment_descendant(sibling)
            }) {
                return true;
            }
            if self.is(self.parent(current), "mtd") {
                return false;
            }
            element = self.parent(current);
        }
        false
    }

    // EquationArrayRow: Office's ProcessEqArrayRow; alignment points become '&'.
    fn equation_array_row(&mut self, items: Vec<NodeId>) -> Result<Vec<NodeId>, String> {
        let mut nodes = Vec::new();
        for item in items {
            let special =
                self.is(Some(item), "mrow") && (self.is_linear_fraction(Some(item)) || self.is_function(Some(item)));
            if self.is_any(Some(item), &["maligngroup", "malignmark"]) {
                if self.is(Some(item), "malignmark") || self.first_alignment_already_found(Some(item)) {
                    let t = self.om("t", vec![Part::Text("&".to_string())]);
                    nodes.push(self.om("r", vec![t.into()]));
                }
            } else if !self.is_nary_argument(item)
                && ((self.is(Some(item), "mrow") && !special) || self.is(Some(item), "mstyle"))
            {
                let children = self.src.elements(item);
                nodes.extend(self.equation_array_row(children)?);
            } else if self.has_alignment_descendant(item) {
                let count = self
                    .src
                    .descendants(item)
                    .into_iter()
                    .filter(|&descendant| self.is_any(Some(descendant), &["maligngroup", "malignmark"]))
                    .count();
                let literal = self.is_any(Some(item), &["mtext", "ms"]);
                let t = self.om("t", vec![Part::Text(output_text(&"&".repeat(count), literal))]);
                nodes.push(self.om("r", vec![t.into()]));
                nodes.extend(self.apply(item)?);
            } else {
                nodes.extend(self.apply(item)?);
            }
        }
        Ok(nodes)
    }

    // ---- Phantoms --------------------------------------------------------------------

    fn phantom_properties(&mut self, show: i32, width: i32, height: i32, depth: i32) -> Option<NodeId> {
        if show != 0 && width != 0 && height != 0 && depth != 0 {
            return None;
        }
        let show_element = if show == 0 { Some(self.om("show", vec![Self::val("off")])) } else { None };
        let width_element = if width == 0 { Some(self.om("zeroWid", vec![Self::val("on")])) } else { None };
        let height_element = if height == 0 { Some(self.om("zeroAsc", vec![Self::val("on")])) } else { None };
        let depth_element = if depth == 0 { Some(self.om("zeroDesc", vec![Self::val("on")])) } else { None };
        Some(self.om(
            "phantPr",
            vec![show_element.into(), width_element.into(), height_element.into(), depth_element.into()],
        ))
    }

    fn padded(&mut self, padded: NodeId) -> Result<Vec<NodeId>, String> {
        if self.is(self.parent(padded), "mphantom")
            && self.src.elements_before_self(padded).is_empty()
            && self.src.elements_after_self(padded).is_empty()
        {
            return self.apply_children(padded);
        }
        let width = full(&self.value(Some(padded), "width"));
        let height = full(&self.value(Some(padded), "height"));
        let depth = full(&self.value(Some(padded), "depth"));
        let properties = self.phantom_properties(1, width, height, depth);
        let content = self.apply_children(padded)?;
        let e = self.om("e", vec![content.into()]);
        Ok(vec![self.om("phant", vec![properties.into(), e.into()])])
    }

    fn phantom(&mut self, phantom: NodeId) -> Result<NodeId, String> {
        let children = self.src.elements(phantom);
        let properties = if children.len() == 1 && self.is(Some(children[0]), "mpadded") {
            let padded = children[0];
            let width = full(&lower_ascii(&self.value(Some(padded), "width")));
            let height = full(&lower_ascii(&self.value(Some(padded), "height")));
            let depth = full(&lower_ascii(&self.value(Some(padded), "depth")));
            self.phantom_properties(0, width, height, depth)
        } else {
            self.phantom_properties(0, 1, 1, 1)
        };
        let content = self.apply_children(phantom)?;
        let e = self.om("e", vec![content.into()]);
        Ok(self.om("phant", vec![properties.into(), e.into()]))
    }
}

/// What FontOf's `XObject? node` argument is.
#[derive(Clone, Copy)]
enum FontNode {
    #[allow(dead_code)]
    None,
    Element(NodeId),
    Text,
}

// (not(@x) or @x='') for one attribute
fn blank(value: Option<&str>) -> bool {
    value.is_none_or(str::is_empty)
}

/// OutputText: Office's OutputText; literal text (mtext, ms) is written as authored.
fn output_text(input: &str, literal: bool) -> String {
    if literal {
        return input.to_string();
    }
    let mut builder = String::with_capacity(input.len());
    for character in input.chars() {
        if character == '\u{2062}' || character == '\u{200B}' {
            continue;
        }
        builder.push(if character == '\u{00A0}' { ' ' } else { character });
    }
    builder.replace('\u{2A75}', "==")
}

/// ToUpperCombining: Office's ToUpperCombining.
fn to_upper_combining(character: &str) -> String {
    match character {
        "\u{02D8}" => "\u{0306}",
        "\u{00B8}" => "\u{0312}",
        "`" => "\u{0300}",
        "-" | "\u{2212}" => "\u{0305}",
        "." | "\u{02D9}" => "\u{0307}",
        "\u{02DD}" => "\u{030B}",
        "\u{00B4}" => "\u{0301}",
        "~" | "\u{02DC}" => "\u{0303}",
        "\u{00A8}" => "\u{0308}",
        "\u{02C7}" => "\u{030C}",
        "^" => "\u{0302}",
        "\u{00AF}" => "\u{0305}",
        "\u{2192}" | "\u{27F6}" => "\u{20D7}",
        "\u{2190}" => "\u{20D6}",
        other => other,
    }
    .to_string()
}

/// Full(value)
fn full(value: &str) -> i32 {
    if value.chars().any(|c| ('1'..='9').contains(&c)) {
        return 1;
    }
    if value.contains('0') {
        return 0;
    }
    1
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn xpath_numbers() {
        for good in ["1", "+1", "-1.5", ".5", "1.", "1e5", "1E-5", "  2  "] {
            assert!(!is_not_a_number(Some(good)), "{good}");
        }
        for bad in ["", "e5", "1e", ".", "1.2.3", "x", "1e999", "0x1", "١"] {
            assert!(is_not_a_number(Some(bad)), "{bad}");
        }
        assert!(is_not_a_number(None));
    }

    #[test]
    fn normalize_space_collapses() {
        assert_eq!(normalize_space("  a \t b\n"), "a b");
        assert_eq!(normalize_space(""), "");
        assert_eq!(normalize_space("\u{A0}a"), "\u{A0}a");
    }
}
