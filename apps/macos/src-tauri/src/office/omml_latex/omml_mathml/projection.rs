//! Port of OmmlMathMlProjection.cs: VisualTeX's own OMML -> presentation
//! MathML projection (the replacement for Office's OMML2MML.XSL).

use super::dom::{Data, Doc, Name, NodeId, MATHML_NS, OMML_NS, WORD_NS};
use super::xml;
use std::sync::LazyLock;

const M: &str = OMML_NS;
const W: &str = WORD_NS;
const NO_BREAK_SPACE: &str = "\u{A0}";

// Every character OMML recognizes as an operator (Office's sOperators).
static OPERATORS: LazyLock<Vec<bool>> = LazyLock::new(|| {
    let singles = concat!(
        "\u{00A8}\u{0021}\u{0022}\u{0023}\u{0026}\u{0028}\u{0029}\u{002B}\u{002C}\u{002D}\u{002E}\u{002F}\u{003A}",
        "\u{003B}\u{003C}\u{003D}\u{003E}\u{003F}\u{0040}\u{005B}\u{005C}\u{005D}\u{005E}\u{005F}\u{0060}\u{007B}",
        "\u{007C}\u{007D}\u{007E}\u{00A1}\u{00A6}\u{00AC}\u{00AF}\u{00B0}\u{00B1}\u{00B2}\u{00B3}\u{00B4}\u{00B7}\u{00B9}\u{00BF}",
        "\u{00D7}\u{00F7}\u{02C7}\u{02D8}\u{02D9}\u{02DC}\u{02DD}",
        "\u{033F}",
        "\u{2000}\u{2001}\u{2002}\u{2003}\u{2004}\u{2005}\u{2006}\u{2009}\u{200A}\u{2010}\u{2012}\u{2013}",
        "\u{2014}\u{2016}\u{2020}\u{2021}\u{2022}\u{2024}\u{2025}\u{2026}\u{2032}\u{2033}\u{2034}\u{203C}",
        "\u{2040}\u{2044}\u{204E}\u{204F}\u{2050}\u{2057}\u{2061}\u{2062}\u{2063}\u{2070}\u{2074}\u{2075}",
        "\u{2076}\u{2077}\u{2078}\u{2079}\u{207A}\u{207B}\u{207C}\u{207D}\u{207E}",
        "\u{2140}\u{2146}",
        "\u{21B6}\u{21B7}",
        "\u{2305}\u{2306}\u{2308}\u{2309}\u{230A}\u{230B}\u{231C}\u{231D}\u{231E}\u{231F}\u{2322}\u{2323}\u{2329}\u{232A}",
        "\u{233D}\u{233F}\u{23B0}\u{23B1}\u{23DC}\u{23DD}\u{23DE}\u{23DF}\u{23E0}\u{2502}\u{251C}\u{2524}",
        "\u{252C}\u{2534}\u{2581}\u{2588}\u{2592}\u{25A0}\u{25A1}\u{25AD}\u{25B2}\u{25B3}\u{25B4}\u{25B5}",
        "\u{25B6}\u{25B7}\u{25B8}\u{25B9}\u{25BC}\u{25BD}\u{25BE}\u{25BF}\u{25C0}\u{25C1}\u{25C2}\u{25C3}",
        "\u{25C4}\u{25C5}\u{25CA}\u{25CB}\u{25E6}\u{25EB}\u{25EC}\u{25F8}\u{25F9}\u{25FA}\u{25FB}\u{25FC}",
        "\u{25FD}\u{25FE}\u{25FF}\u{2605}\u{2606}\u{2772}\u{2773}",
        "\u{29B6}\u{29B7}\u{29B8}\u{29B9}\u{29C0}\u{29C1}\u{29C4}\u{29C5}\u{29C6}\u{29C7}\u{29C8}",
        "\u{29DF}",
        "\u{29EB}",
        "\u{2B04}\u{2B06}\u{2B07}\u{2B0C}\u{2B0D}\u{3014}\u{3015}\u{3016}\u{3017}\u{3018}\u{3019}\u{FF01}\u{FF06}",
        "\u{FF08}\u{FF09}\u{FF0B}\u{FF0C}\u{FF0D}\u{FF0E}\u{FF0F}\u{FF1A}\u{FF1B}\u{FF1C}\u{FF1D}\u{FF1E}",
        "\u{FF1F}\u{FF20}\u{FF3B}\u{FF3C}\u{FF3D}\u{FF3E}\u{FF3F}\u{FF5B}\u{FF5C}\u{FF5D}",
    );
    let ranges: &[(u32, u32)] = &[
        (0x0300, 0x0338),
        (0x2080, 0x208E),
        (0x20D0, 0x20E1),
        (0x20E4, 0x20EA),
        (0x2190, 0x21B3),
        (0x21BA, 0x21E9),
        (0x21F3, 0x2204),
        (0x2206, 0x220D),
        (0x220F, 0x221D),
        (0x2223, 0x223E),
        (0x2240, 0x22A3),
        (0x22A5, 0x22BD),
        (0x22C0, 0x22FF),
        (0x27D1, 0x27EB),
        (0x27F0, 0x2980),
        (0x2982, 0x299A),
        (0x29CE, 0x29DB),
        (0x29E1, 0x29E6),
        (0x29F4, 0x2AE0),
        (0x2AE2, 0x2AF0),
        (0x2AF2, 0x2AFF),
    ];
    let mut table = vec![false; 0x10000];
    for c in singles.chars() {
        table[c as usize] = true;
    }
    for &(first, last) in ranges {
        for code in first..=last {
            table[code as usize] = true;
        }
    }
    table
});

fn is_operator(c: char) -> bool {
    (c as u32) < 0x10000 && OPERATORS[c as usize]
}

/// OmmlMathMlProjection.Project: reads `omml` with ReaderSettings() and
/// projects it. The result is a MathML math element (the document root).
pub fn project(omml: &str) -> Result<(Doc, NodeId), String> {
    let source = xml::parse(omml, xml::PROJECTION_READER)?;
    let root = source.root.ok_or_else(|| "XML input has no root element".to_string())?;
    Ok(project_tree(&source, Some(root)))
}

/// Projects an already-read source tree (whitespace as ReaderSettings() reads it).
pub fn project_tree(source: &Doc, root: Option<NodeId>) -> (Doc, NodeId) {
    let mut projector = Projector { s: source, o: Doc::new() };
    let math = projector.o.new_element(Name::new(MATHML_NS, "math"));
    if let Some(root) = root {
        for node in projector.apply(root) {
            projector.o.append(math, node);
        }
    }
    let mut output = projector.o;
    output.root = Some(math);
    settle(&mut output, math);
    (output, math)
}

// The projection was historically serialized and parsed again with the
// reader settings: whitespace-only text is dropped and line breaks normalized.
fn settle(doc: &mut Doc, root: NodeId) {
    xml::normalize_text_line_ends(doc, root);
    for node in doc.descendant_nodes(root) {
        if let Data::Text(text) = &doc.nodes[node].data {
            if text.chars().all(xml::is_xml_whitespace) {
                doc.detach(node);
            }
        }
    }
}

type Attr = (&'static str, String);

// XPath 1.0 comparisons of a node-set with a string are existential.
fn eq(set: &[String], value: &str) -> bool {
    set.iter().any(|item| item == value)
}
fn ne(set: &[String], value: &str) -> bool {
    set.iter().any(|item| item != value)
}
fn str_(set: &[String]) -> &str {
    set.first().map(String::as_str).unwrap_or("")
}

fn lower_ascii(value: &str) -> String {
    value.chars().map(|c| if c.is_ascii_uppercase() { c.to_ascii_lowercase() } else { c }).collect()
}

// Office's TFromStrVal: on/1/true, off/0/false, otherwise neither.
fn tristate(value: &str) -> i32 {
    match value {
        "on" | "1" | "true" => 1,
        "off" | "0" | "false" => 0,
        _ => -1,
    }
}

fn force_false(lower: &str) -> i32 {
    if tristate(lower) == 0 { 0 } else { 1 }
}

fn force_true(lower: &str) -> i32 {
    if tristate(lower) == 1 { 1 } else { 0 }
}

fn math_color(colors: &[String]) -> Option<Attr> {
    if !str_(colors).is_empty() && ne(colors, "auto") {
        Some(("mathcolor", format!("#{}", str_(colors))))
    } else {
        None
    }
}

fn normalize_space(value: &str) -> String {
    value.split([' ', '\t', '\r', '\n']).filter(|part| !part.is_empty()).collect::<Vec<_>>().join(" ")
}

// Office's ToNonCombining; the first matching row of its table wins.
fn to_non_combining(character: &str) -> String {
    match character {
        "\u{0306}" | "\u{032E}" => "\u{02D8}",
        "\u{0312}" | "\u{0327}" => "\u{00B8}",
        "\u{0300}" | "\u{0316}" => "`",
        "\u{0305}" | "\u{0332}" => "-",
        "\u{0323}" => ".",
        "\u{0307}" => "\u{02D9}",
        "\u{030B}" | "\u{02DD}" => "\u{02DD}",
        "\u{0301}" | "\u{0317}" => "\u{00B4}",
        "\u{0303}" | "\u{0330}" => "~",
        "\u{0308}" | "\u{0324}" => "\u{00A8}",
        "\u{030C}" | "\u{032C}" => "\u{02C7}",
        "\u{0302}" | "\u{032D}" => "^",
        "\u{20D7}" | "\u{20EF}" => "\u{2192}",
        "\u{20D6}" | "\u{20EE}" => "\u{2190}",
        other => other,
    }
    .to_string()
}

fn first_scalar(value: &str) -> String {
    value.chars().next().map(String::from).unwrap_or_default()
}

struct Projector<'a> {
    s: &'a Doc,
    o: Doc,
}

impl<'a> Projector<'a> {
    // ---- source helpers ----------------------------------------------------

    fn first(&self, element: NodeId, name: &str) -> Option<NodeId> {
        self.s.element(element, M, name)
    }

    fn last(&self, element: Option<NodeId>, name: &str) -> Option<NodeId> {
        element.and_then(|element| self.s.elements_named(element, M, name).last().copied())
    }

    // An XPath attribute node-set: the m:val of every named child, in document order.
    fn values_of(&self, owners: &[NodeId], name: &str) -> Vec<String> {
        let s = self.s;
        owners
            .iter()
            .flat_map(|&owner| s.elements_named(owner, M, name))
            .filter_map(|child| s.attr_ns(child, M, "val").map(str::to_string))
            .collect()
    }

    fn values(&self, owner: Option<NodeId>, name: &str) -> Vec<String> {
        match owner {
            Some(owner) => self.values_of(&[owner], name),
            None => Vec::new(),
        }
    }

    // An on/off property: absent is off, present is on unless explicitly false.
    fn switch(&self, properties: Option<NodeId>, name: &str) -> i32 {
        match properties {
            Some(properties) if self.s.element(properties, M, name).is_some() => {
                force_false(&lower_ascii(str_(&self.values(Some(properties), name))))
            }
            _ => 0,
        }
    }

    fn normal(&self, run: NodeId) -> i32 {
        self.switch(self.last(Some(run), "rPr"), "nor")
    }

    fn colors(&self, run_properties: &[Option<NodeId>]) -> Vec<String> {
        let s = self.s;
        run_properties
            .iter()
            .flatten()
            .flat_map(|&properties| s.elements_named(properties, W, "color"))
            .filter_map(|color| s.attr_ns(color, W, "val").map(str::to_string))
            .collect()
    }

    fn run_colors(&self, run: Option<NodeId>) -> Vec<String> {
        match run {
            None => Vec::new(),
            Some(run) => self.colors(&[self.s.elements_named(run, W, "rPr").last().copied()]),
        }
    }

    // The nearest m:r at or above the context node.
    fn context_run_colors(&self, context: NodeId) -> Vec<String> {
        let s = self.s;
        let mut chain = vec![context];
        chain.extend(s.ancestors(context));
        self.run_colors(chain.into_iter().find(|&node| s.is(node, M, "r")))
    }

    fn control_colors(&self, owner: NodeId) -> Vec<String> {
        let s = self.s;
        let controls: Vec<Option<NodeId>> = s
            .elements(owner)
            .into_iter()
            .filter(|&element| s.is(element, M, "fPr") || s.is(element, M, "radPr"))
            .flat_map(|properties| s.elements_named(properties, M, "ctrlPr"))
            .map(|control| s.elements_named(control, W, "rPr").last().copied())
            .collect();
        self.colors(&controls)
    }

    // ---- output helpers ----------------------------------------------------

    fn el(&mut self, name: &str, attrs: Vec<Option<Attr>>, children: Vec<NodeId>) -> NodeId {
        let id = self.o.new_element(Name::new(MATHML_NS, name));
        for (key, value) in attrs.into_iter().flatten() {
            self.o.set_attr_ns(id, "", key, Some(&value));
        }
        for child in children {
            self.o.append(id, child);
        }
        id
    }

    fn text(&mut self, value: &str) -> Vec<NodeId> {
        if value.is_empty() {
            Vec::new()
        } else {
            vec![self.o.new_text(value)]
        }
    }

    fn row(&mut self, element: Option<NodeId>) -> NodeId {
        let content = self.apply_optional(element);
        self.el("mrow", vec![], content)
    }

    // ---- projection --------------------------------------------------------

    fn apply(&mut self, element: NodeId) -> Vec<NodeId> {
        let s = self.s;
        let name = s.name(element).expect("apply on element");
        if name.ns != M {
            return self.apply_children(element);
        }
        match name.local.as_str() {
            "borderBox" => vec![self.border_box(element)],
            "acc" => vec![self.accent(element)],
            "sPre" => vec![self.prescripts(element)],
            "m" => vec![self.matrix(element)],
            "phant" => vec![self.phantom(element)],
            "rad" => vec![self.radical(element)],
            "nary" => vec![self.nary(element)],
            "limLow" => {
                let base = self.row(self.first(element, "e"));
                let limit = self.row(self.first(element, "lim"));
                vec![self.el("munder", vec![], vec![base, limit])]
            }
            "limUpp" => {
                let base = self.row(self.first(element, "e"));
                let limit = self.row(self.first(element, "lim"));
                vec![self.el("mover", vec![], vec![base, limit])]
            }
            "sSub" => {
                let base = self.row(self.first(element, "e"));
                let sub = self.row(self.first(element, "sub"));
                vec![self.el("msub", vec![], vec![base, sub])]
            }
            "sSup" => {
                let base = self.row(self.first(element, "e"));
                let sup = self.row(self.first(element, "sup"));
                vec![self.el("msup", vec![], vec![base, sup])]
            }
            "sSubSup" => {
                let base = self.row(self.first(element, "e"));
                let sub = self.row(self.first(element, "sub"));
                let sup = self.row(self.first(element, "sup"));
                vec![self.el("msubsup", vec![], vec![base, sub, sup])]
            }
            "groupChr" => vec![self.group_character(element)],
            "func" => vec![self.function(element)],
            "f" => vec![self.fraction(element)],
            "e" | "den" | "num" | "lim" | "sup" | "sub" => self.argument(element),
            "bar" => vec![self.bar(element)],
            "d" => vec![self.delimiter(element)],
            "r" => self.run(element),
            "eqArr" => vec![self.equation_array(element)],
            _ => self.apply_children(element),
        }
    }

    fn apply_children(&mut self, element: NodeId) -> Vec<NodeId> {
        let mut nodes = Vec::new();
        for child in self.s.elements(element) {
            nodes.extend(self.apply(child));
        }
        nodes
    }

    fn apply_optional(&mut self, element: Option<NodeId>) -> Vec<NodeId> {
        match element {
            Some(element) => self.apply(element),
            None => Vec::new(),
        }
    }

    fn border_box(&mut self, bx: NodeId) -> NodeId {
        let s = self.s;
        let properties = self.last(Some(bx), "borderBoxPr");
        let flag = |name: &str| {
            let value = self
                .last(properties, name)
                .and_then(|flag| s.attr_ns(flag, M, "val"))
                .unwrap_or("");
            force_true(&lower_ascii(value))
        };
        let hide_top = flag("hideTop");
        let hide_bottom = flag("hideBot");
        let hide_left = flag("hideLeft");
        let hide_right = flag("hideRight");
        let strike_h = flag("strikeH");
        let strike_v = flag("strikeV");
        let strike_bltr = flag("strikeBLTR");
        let strike_tlbr = flag("strikeTLBR");
        if hide_top == 1
            && hide_bottom == 1
            && hide_left == 1
            && hide_right == 1
            && strike_h == 0
            && strike_v == 0
            && strike_bltr == 0
            && strike_tlbr == 0
        {
            let content = self.apply_optional(self.first(bx, "e"));
            return self.el("mrow", vec![], content);
        }

        let mut notation;
        if hide_top == 0 && hide_bottom == 0 && hide_left == 0 && hide_right == 0 {
            notation = "box".to_string();
        } else {
            notation = "left right top bottom".to_string();
            if hide_top == 1 {
                notation = notation.replace("top", "");
            }
            if hide_bottom == 1 {
                notation = notation.replace("bottom", "");
            }
            if hide_left == 1 {
                notation = notation.replace("left", "");
            }
            if hide_right == 1 {
                notation = notation.replace("right", "");
            }
        }
        if strike_h == 1 {
            notation.push_str(" horizontalstrike");
        }
        if strike_v == 1 {
            notation.push_str(" verticalstrike");
        }
        if strike_bltr == 1 {
            notation.push_str(" updiagonalstrike");
        }
        if strike_tlbr == 1 {
            notation.push_str(" downdiagonalstrike");
        }
        let content = self.apply_optional(self.first(bx, "e"));
        self.el("menclose", vec![Some(("notation", normalize_space(&notation)))], content)
    }

    fn accent(&mut self, accent: NodeId) -> NodeId {
        let s = self.s;
        let has_chr = self.last(Some(accent), "accPr").and_then(|p| s.element(p, M, "chr")).is_some();
        let mark = if !has_chr {
            "\u{0302}".to_string()
        } else {
            let owners = s.elements_named(accent, M, "accPr");
            first_scalar(str_(&self.values_of(&owners, "chr")))
        };
        let base_slot = self.first(accent, "e");
        let base_row = self.row(base_slot);
        let result = self.el("mover", vec![Some(("accent", "true".to_string()))], vec![base_row]);
        if mark.is_empty() {
            let mo = self.el("mo", vec![], vec![]);
            self.o.append(result, mo);
            return result;
        }

        // The mark takes the style of the base's runs, as Office's stylesheet did.
        let base_run_properties: Vec<NodeId> = base_slot
            .map(|slot| s.elements(slot))
            .unwrap_or_default()
            .into_iter()
            .filter_map(|child| self.last(Some(child), "rPr"))
            .collect();
        let normal = if base_run_properties.iter().any(|&p| s.element(p, M, "nor").is_some()) {
            force_false(&lower_ascii(str_(&self.values_of(&base_run_properties, "nor"))))
        } else {
            0
        };
        let script = self.values_of(&base_run_properties, "scr");
        let style = self.values_of(&base_run_properties, "sty");
        let nodes = self.parse_text(&to_non_combining(&mark), &script, &style, normal, accent);
        for node in nodes {
            self.o.append(result, node);
        }
        result
    }

    fn prescripts(&mut self, script: NodeId) -> NodeId {
        let s = self.s;
        let base = self.row(self.first(script, "e"));
        let marker = self.el("mprescripts", vec![], vec![]);
        let slot = |this: &mut Self, slot: Option<NodeId>| match slot {
            Some(slot) if !s.elements(slot).is_empty() => {
                let content = this.apply(slot);
                this.el("mrow", vec![], content)
            }
            _ => this.el("none", vec![], vec![]),
        };
        let (sub_slot, sup_slot) = (self.first(script, "sub"), self.first(script, "sup"));
        let sub = slot(self, sub_slot);
        let sup = slot(self, sup_slot);
        self.el("mmultiscripts", vec![], vec![base, marker, sub, sup])
    }

    fn matrix(&mut self, matrix: NodeId) -> NodeId {
        let s = self.s;
        let column_properties: Vec<NodeId> = self
            .last(Some(matrix), "mPr")
            .map(|properties| s.elements_named(properties, M, "mcs"))
            .unwrap_or_default()
            .into_iter()
            .flat_map(|columns| s.elements_named(columns, M, "mc"))
            .filter_map(|column| self.last(Some(column), "mcPr"))
            .collect();
        let alignment = lower_ascii(str_(&self.values_of(&column_properties, "mcJc")));
        let mut attrs = vec![];
        if alignment == "left" || alignment == "right" {
            attrs.push(Some(("columnalign", alignment)));
        }
        let mut rows = Vec::new();
        for row in s.elements_named(matrix, M, "mr") {
            let mut cells = Vec::new();
            for cell in s.elements_named(row, M, "e") {
                let content = self.apply(cell);
                cells.push(self.el("mtd", vec![], content));
            }
            rows.push(self.el("mtr", vec![], cells));
        }
        self.el("mtable", attrs, rows)
    }

    fn phantom(&mut self, phantom: NodeId) -> NodeId {
        let s = self.s;
        let properties = self.last(Some(phantom), "phantPr");
        let flag = |name: &str| match self.last(properties, name) {
            None => 0,
            Some(flag) => force_false(&lower_ascii(s.attr_ns(flag, M, "val").unwrap_or(""))),
        };
        let zero_width = flag("zeroWid");
        let zero_ascent = flag("zeroAsc");
        let zero_descent = flag("zeroDesc");
        let show = force_false(&lower_ascii(
            self.last(properties, "show").and_then(|show| s.attr_ns(show, M, "val")).unwrap_or(""),
        ));

        let content = |this: &mut Self| {
            let mut nodes = Vec::new();
            for e in s.elements_named(phantom, M, "e") {
                nodes.extend(this.apply(e));
            }
            this.el("mrow", vec![], nodes)
        };
        let padded = |this: &mut Self| {
            let inner = content(this);
            this.el(
                "mpadded",
                vec![
                    (zero_width == 1).then(|| ("width", "0in".to_string())),
                    (zero_ascent == 1).then(|| ("height", "0in".to_string())),
                    (zero_descent == 1).then(|| ("depth", "0in".to_string())),
                ],
                vec![inner],
            )
        };
        if show == 1 {
            return padded(self);
        }
        if zero_width == 0 && zero_ascent == 0 && zero_descent == 0 {
            let inner = content(self);
            return self.el("mphantom", vec![], vec![inner]);
        }
        let inner = padded(self);
        self.el("mphantom", vec![], vec![inner])
    }

    fn radical(&mut self, radical: NodeId) -> NodeId {
        let hide_degree = self.switch(self.last(Some(radical), "radPr"), "degHide");
        let color = math_color(&self.control_colors(radical));
        if hide_degree == 1 {
            let content = self.apply_optional(self.first(radical, "e"));
            return self.el("msqrt", vec![color], content);
        }
        let base = self.row(self.first(radical, "e"));
        let degree = self.row(self.first(radical, "deg"));
        self.el("mroot", vec![color], vec![base, degree])
    }

    fn nary(&mut self, nary: NodeId) -> NodeId {
        let s = self.s;
        let properties = self.last(Some(nary), "naryPr");
        let lower = |name: &str, absent: &str| match properties {
            Some(p) if s.element(p, M, name).is_some() => lower_ascii(str_(&self.values(Some(p), name))),
            _ => absent.to_string(),
        };
        let hide_sub = force_false(&lower("subHide", "off"));
        let hide_sup = force_false(&lower("supHide", "off"));
        let grow = force_false(&lower("grow", "off"));
        let limits_as_scripts = match properties {
            Some(p) if s.element(p, M, "limLoc").is_some() => {
                lower_ascii(str_(&self.values(Some(p), "limLoc"))) == "subsup"
            }
            _ => true,
        };

        let characters = self.values(properties, "chr");
        let symbol = if characters.is_empty() || eq(&characters, "") {
            "\u{222B}".to_string()
        } else {
            str_(&characters).to_string()
        };
        let operator = |this: &mut Self| {
            let text = this.text(&symbol);
            this.el("mo", vec![Some(("stretchy", if grow == 1 { "true" } else { "false" }.to_string()))], text)
        };

        let head = if hide_sup == 1 && hide_sub == 1 {
            operator(self)
        } else if hide_sub == 1 {
            let op = operator(self);
            let sup = self.row(self.first(nary, "sup"));
            self.el(if limits_as_scripts { "msup" } else { "mover" }, vec![], vec![op, sup])
        } else if hide_sup == 1 {
            let op = operator(self);
            let sub = self.row(self.first(nary, "sub"));
            self.el(if limits_as_scripts { "msub" } else { "munder" }, vec![], vec![op, sub])
        } else {
            let op = operator(self);
            let sub = self.row(self.first(nary, "sub"));
            let sup = self.row(self.first(nary, "sup"));
            self.el(if limits_as_scripts { "msubsup" } else { "munderover" }, vec![], vec![op, sub, sup])
        };
        let body = self.row(self.first(nary, "e"));
        self.el("mrow", vec![], vec![head, body])
    }

    fn group_character(&mut self, group: NodeId) -> NodeId {
        let s = self.s;
        let properties = self.last(Some(group), "groupChrPr");
        let has_character = properties.and_then(|p| s.element(p, M, "chr")).is_some();
        let character_values = self.values(properties, "chr");
        let character_value = str_(&character_values);
        let character = if has_character && character_value.is_empty() {
            String::new()
        } else if !character_value.is_empty() {
            first_scalar(character_value)
        } else {
            "\u{23DF}".to_string()
        };
        let position = lower_ascii(str_(&self.values(properties, "pos")));
        let justification = lower_ascii(str_(&self.values(properties, "vertJc")));
        let mark = |this: &mut Self| {
            let text = this.text(&character);
            this.el("mo", vec![], text)
        };
        let content = self.row(self.first(group, "e"));
        if position == "top" {
            return if justification == "bot" {
                let m = mark(self);
                self.el("mover", vec![Some(("accent", "false".to_string()))], vec![content, m])
            } else {
                let m = mark(self);
                self.el("munder", vec![Some(("accentunder", "false".to_string()))], vec![m, content])
            };
        }
        if justification == "bot" {
            let m = mark(self);
            self.el("mover", vec![Some(("accent", "false".to_string()))], vec![m, content])
        } else {
            let m = mark(self);
            self.el("munder", vec![Some(("accentunder", "false".to_string()))], vec![content, m])
        }
    }

    fn function(&mut self, function: NodeId) -> NodeId {
        let s = self.s;
        let mut names = Vec::new();
        for name in s.elements_named(function, M, "fName") {
            for child in s.elements(name) {
                names.extend(self.apply(child));
            }
        }
        let name_row = self.el("mrow", vec![], names);
        let apply_text = self.text("\u{2061}");
        let apply = self.el("mo", vec![], apply_text);
        let mut arguments = Vec::new();
        for e in s.elements_named(function, M, "e") {
            arguments.extend(self.apply(e));
        }
        let argument_row = self.el("mrow", vec![], arguments);
        self.el("mrow", vec![], vec![name_row, apply, argument_row])
    }

    fn fraction(&mut self, fraction: NodeId) -> NodeId {
        let fraction_type = lower_ascii(str_(&self.values(self.last(Some(fraction), "fPr"), "type")));
        if fraction_type == "lin" {
            let numerator = self.row(self.first(fraction, "num"));
            let slash_text = self.text("/");
            let slash = self.el("mo", vec![], slash_text);
            let denominator = self.row(self.first(fraction, "den"));
            return self.el("mrow", vec![], vec![numerator, slash, denominator]);
        }
        let color = math_color(&self.control_colors(fraction));
        let numerator = self.row(self.first(fraction, "num"));
        let denominator = self.row(self.first(fraction, "den"));
        self.el(
            "mfrac",
            vec![
                color,
                (fraction_type == "skw").then(|| ("bevelled", "true".to_string())),
                (fraction_type == "nobar").then(|| ("linethickness", "0pt".to_string())),
            ],
            vec![numerator, denominator],
        )
    }

    fn argument(&mut self, argument: NodeId) -> Vec<NodeId> {
        let level = self.values(self.last(Some(argument), "argPr"), "scrLvl");
        if level.is_empty() {
            return self.apply_children(argument);
        }
        let content = self.apply_children(argument);
        vec![self.el("mstyle", vec![Some(("scriptlevel", str_(&level).to_string()))], content)]
    }

    fn bar(&mut self, bar: NodeId) -> NodeId {
        let owners = self.s.elements_named(bar, M, "barPr");
        let position = lower_ascii(str_(&self.values_of(&owners, "pos")));
        let base = self.row(self.first(bar, "e"));
        if position == "top" {
            let text = self.text("\u{00AF}");
            let mark = self.el("mo", vec![], text);
            self.el("mover", vec![Some(("accent", "false".to_string()))], vec![base, mark])
        } else {
            let text = self.text("_");
            let mark = self.el("mo", vec![], text);
            self.el("munder", vec![Some(("underaccent", "false".to_string()))], vec![base, mark])
        }
    }

    fn delimiter(&mut self, delimiter: NodeId) -> NodeId {
        let s = self.s;
        let properties = self.first(delimiter, "dPr");
        let mut attrs = Vec::new();
        let open = self.values(properties, "begChr");
        if !open.is_empty() && !eq(&open, "(") {
            attrs.push(Some(("open", str_(&open).to_string())));
        }
        let close = self.values(properties, "endChr");
        if !close.is_empty() && !eq(&close, ")") {
            attrs.push(Some(("close", str_(&close).to_string())));
        }
        let separator = self.values(properties, "sepChr");
        if !eq(&separator, ",") {
            attrs.push(Some((
                "separators",
                if separator.is_empty() { "|".to_string() } else { str_(&separator).to_string() },
            )));
        }
        let mut items = Vec::new();
        for item in s.elements_named(delimiter, M, "e") {
            let content = self.apply(item);
            items.push(self.el("mrow", vec![], content));
        }
        self.el("mfenced", attrs, items)
    }

    fn run(&mut self, run: NodeId) -> Vec<NodeId> {
        let s = self.s;
        let texts: Vec<NodeId> = s.descendants(run).into_iter().filter(|&d| s.is(d, M, "t")).collect();
        if self.normal(run) == 1 {
            let text = texts.first().map(|&t| s.value(t)).unwrap_or_default();
            let color = math_color(&self.run_colors(Some(run)));
            let content = self.text(&text.replace(' ', NO_BREAK_SPACE));
            return vec![self.el("mtext", vec![color], content)];
        }
        let mut nodes = Vec::new();
        for text in texts {
            let properties = self.last(s.parent(text), "rPr");
            let script = self.values(properties, "scr");
            let style = self.values(properties, "sty");
            nodes.extend(self.parse_text(&s.value(text), &script, &style, 0, text));
        }
        nodes
    }

    fn equation_array(&mut self, array: NodeId) -> NodeId {
        let s = self.s;
        let mut rows = Vec::new();
        for equation in s.elements_named(array, M, "e") {
            let level = self.values(self.last(Some(equation), "argPr"), "scrLvl");
            let mut row = vec![self.el("maligngroup", vec![], vec![])];
            row.extend(self.equation_array_row(array, equation));
            let cell = if ne(&level, "0") || level.is_empty() || eq(&level, "") {
                self.el("mrow", vec![], row)
            } else {
                self.el("mstyle", vec![Some(("scriptlevel", str_(&level).to_string()))], row)
            };
            let td = self.el("mtd", vec![], vec![cell]);
            rows.push(self.el("mtr", vec![], vec![td]));
        }
        self.el("mtable", vec![], rows)
    }

    // Office's CreateEqArrRow: '&' alternates between an alignment mark and an
    // alignment group, counted across the runs of one equation.
    fn equation_array_row(&mut self, array: NodeId, equation: NodeId) -> Vec<NodeId> {
        let s = self.s;
        let mut nodes = Vec::new();
        let mut align = 1i32;
        let array_properties = self.last(Some(array), "rPr");
        let script = self.values(array_properties, "scr");
        let style = self.values(array_properties, "sty");
        for child in s.elements(equation) {
            let all_text: String = s.elements_named(child, M, "t").into_iter().map(|t| s.value(t)).collect();
            if s.is(child, M, "r") {
                let normal = self.normal(child);
                let colors = self.run_colors(Some(child));
                nodes.extend(self.parse_equation_array_text(&all_text, &script, &style, normal, align, &colors));
            } else {
                nodes.extend(self.apply(child));
            }
            let ampersands = all_text.chars().filter(|&c| c == '&').count() as i32;
            align = (align + ampersands % 2) % 2;
        }
        nodes
    }

    fn parse_equation_array_text(
        &mut self,
        source: &str,
        script: &[String],
        style: &[String],
        normal: i32,
        mut align: i32,
        colors: &[String],
    ) -> Vec<NodeId> {
        let mut nodes = Vec::new();
        let scalars: Vec<char> = source.chars().collect();
        let pattern = token_pattern(&scalars);
        let none: Vec<String> = Vec::new();
        let mut index = 0;
        while index < scalars.len() {
            let scalar = scalars[index];
            if scalar == '&' {
                nodes.push(self.el(if align == 0 { "maligngroup" } else { "malignmark" }, vec![], vec![]));
                align = if align == 1 { 0 } else { 1 };
                index += 1;
                continue;
            }
            let color = math_color(colors);
            if pattern[index] == ZERO {
                let length = number_length(&pattern, index);
                let number: String = scalars[index..index + length].iter().collect();
                let content = self.text(&number);
                let node = if normal == 0 {
                    let mut attrs = vec![color];
                    attrs.extend(token_attributes(script, style, normal, None, "mn", &none).into_iter().map(Some));
                    self.el("mn", attrs, content)
                } else {
                    self.el("mtext", vec![color], content)
                };
                nodes.push(node);
                index += length;
            } else if pattern[index] == DASH {
                let content = self.text(&scalar.to_string());
                let node = if normal == 0 {
                    let mut attrs = vec![color, Some(("stretchy", "false".to_string()))];
                    attrs.extend(token_attributes(script, style, normal, None, "mo", &none).into_iter().map(Some));
                    self.el("mo", attrs, content)
                } else {
                    self.el("mtext", vec![color], content)
                };
                nodes.push(node);
                index += 1;
            } else {
                let text = scalar.to_string().replace(' ', NO_BREAK_SPACE);
                let content = self.text(&text);
                let node = if normal == 0 {
                    let mut attrs = vec![color];
                    attrs.extend(token_attributes(script, style, normal, Some(1), "mi", &none).into_iter().map(Some));
                    self.el("mi", attrs, content)
                } else {
                    self.el("mtext", vec![color], content)
                };
                nodes.push(node);
                index += 1;
            }
        }
        nodes
    }

    // Office's ParseMt over Unicode scalars: a run splits into identifiers,
    // numbers and single-character operators. Inside a function name the
    // identifier extends to the first number or operator.
    fn parse_text(&mut self, source: &str, script: &[String], style: &[String], normal: i32, context: NodeId) -> Vec<NodeId> {
        let s = self.s;
        let mut nodes = Vec::new();
        let scalars: Vec<char> = source.chars().collect();
        let pattern = token_pattern(&scalars);
        let in_function_name = s.ancestors(context).iter().any(|&a| s.is(a, M, "fName"));
        let literal_run_operator = s.is(context, M, "t") && s.parent(context).is_some_and(|p| s.is(p, M, "r"));
        let colors = self.context_run_colors(context);
        let mut index = 0;
        while index < scalars.len() {
            if pattern[index] == DASH {
                let mut attrs = vec![literal_run_operator.then(|| ("stretchy", "false".to_string()))];
                attrs.extend(token_attributes(script, style, normal, None, "mo", &colors).into_iter().map(Some));
                let content = self.text(&scalars[index].to_string());
                nodes.push(self.el("mo", attrs, content));
                index += 1;
            } else if pattern[index] == ZERO {
                let length = number_length(&pattern, index);
                let attrs = token_attributes(script, style, normal, None, "mn", &colors).into_iter().map(Some).collect();
                let number: String = scalars[index..index + length].iter().collect();
                let content = self.text(&number);
                nodes.push(self.el("mn", attrs, content));
                index += length;
            } else {
                let length = if in_function_name { function_name_length(&pattern, index) } else { 1 };
                let attrs = token_attributes(script, style, normal, Some(length as i64), "mi", &colors)
                    .into_iter()
                    .map(Some)
                    .collect();
                let identifier: String = scalars[index..index + length].iter().collect();
                let content = self.text(&identifier.replace(' ', NO_BREAK_SPACE));
                nodes.push(self.el("mi", attrs, content));
                index += length;
            }
        }
        nodes
    }
}

const ZERO: u16 = b'0' as u16;
const DASH: u16 = b'-' as u16;

fn number_length(pattern: &[u16], start: usize) -> usize {
    let mut end = start;
    while end < pattern.len() && pattern[end] == ZERO {
        end += 1;
    }
    end - start
}

// Office's nCharToPrint inside m:fName, for the text that remains.
fn function_name_length(pattern: &[u16], start: usize) -> usize {
    let length = pattern.len() - start;
    let mut first_operator = length;
    let mut first_number = length;
    for index in start..pattern.len() {
        if pattern[index] == DASH {
            first_operator = index - start + 1;
            break;
        }
    }
    for index in start..pattern.len() {
        if pattern[index] == ZERO {
            first_number = index - start + 1;
            break;
        }
    }
    let last = pattern[pattern.len() - 1];
    if first_operator == first_number && first_operator == length && last != ZERO && last != DASH {
        return length;
    }
    if first_operator < first_number { first_operator - 1 } else { first_number - 1 }
}

// One pattern character per scalar: '0' for number text, '-' for an
// operator, otherwise the scalar's first code unit. A '.' or ',' between
// two digits belongs to the number.
fn token_pattern(scalars: &[char]) -> Vec<u16> {
    let mut pattern: Vec<u16> = scalars
        .iter()
        .map(|&c| {
            if c.is_ascii_digit() {
                ZERO
            } else {
                let mut buffer = [0u16; 2];
                c.encode_utf16(&mut buffer)[0]
            }
        })
        .collect();
    replace_separators(&mut pattern, b'.' as u16);
    replace_separators(&mut pattern, b',' as u16);
    for index in 0..pattern.len() {
        if scalars[index].len_utf16() == 1 && pattern[index] != ZERO && is_operator(scalars[index]) {
            pattern[index] = DASH;
        }
    }
    pattern
}

fn replace_separators(pattern: &mut [u16], separator: u16) {
    let mut index = 0;
    while index + 2 < pattern.len() {
        if pattern[index] == ZERO && pattern[index + 1] == separator && pattern[index + 2] == ZERO {
            pattern[index + 1] = ZERO;
            index += 3;
        } else {
            index += 1;
        }
    }
}

// Office's CreateTokenAttributes with VisualTeX's adaptations.
fn token_attributes(
    script: &[String],
    style: &[String],
    normal: i32,
    characters: Option<i64>,
    token_type: &str,
    colors: &[String],
) -> Vec<Attr> {
    let mut attributes = Vec::new();
    if normal == 1 {
        attributes.push(("mathvariant", "normal".to_string()));
    } else {
        let variant = variant(script, style, token_type);
        let identifier = token_type == "mi";
        if token_type == "mo" || token_type == "mn" {
            if !variant.is_empty() {
                attributes.push(("mathvariant", variant.to_string()));
            }
        } else if identifier && characters == Some(1) && (variant.is_empty() || variant == "italic") {
            // A single-character identifier is italic by default.
        } else if identifier && characters.is_some_and(|c| c > 1) && (variant.is_empty() || variant == "italic") {
            attributes.push(("mathvariant", "italic".to_string()));
        } else if variant != "italic" && !variant.is_empty() {
            attributes.push(("mathvariant", variant.to_string()));
        } else {
            if !(identifier && characters == Some(1)) && !(eq(style, "p") || eq(style, "b")) {
                attributes.push(("fontstyle", "italic".to_string()));
            }
            if eq(style, "b") || eq(style, "bi") {
                attributes.push(("fontweight", "bold".to_string()));
            }
        }
    }
    if let Some(color) = math_color(colors) {
        attributes.push(color);
    }
    attributes
}

fn variant(script: &[String], style: &[String], token_type: &str) -> &'static str {
    let sc = |value: &str| eq(script, value);
    let st = |value: &str| eq(style, value);
    if token_type == "mn" && (str_(script).is_empty() || sc("roman")) && (st("p") || st("")) {
        return "";
    }
    if sc("monospace") {
        return "monospace";
    }
    if sc("sans-serif") && st("i") {
        return "sans-serif-italic";
    }
    if sc("sans-serif") && st("b") {
        return "bold-sans-serif";
    }
    if sc("sans-serif") && st("bi") {
        return "sans-serif-bold-italic";
    }
    if sc("sans-serif") {
        return "sans-serif";
    }
    if sc("fraktur") && (st("b") || st("bi")) {
        return "bold-fraktur";
    }
    if sc("fraktur") {
        return "fraktur";
    }
    if sc("double-struck") {
        return "double-struck";
    }
    if sc("script") && (st("b") || st("bi")) {
        return "bold-script";
    }
    if sc("script") {
        return "script";
    }
    let roman = sc("roman") || script.is_empty() || sc("");
    if roman && st("b") {
        return "bold";
    }
    if roman && st("i") {
        return "italic";
    }
    if roman && st("p") {
        return "normal";
    }
    if roman && st("bi") {
        return "bold-italic";
    }
    ""
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn number_separators_join_digits() {
        let scalars: Vec<char> = "1.2.3".chars().collect();
        let pattern = token_pattern(&scalars);
        assert_eq!(String::from_utf16(&pattern).unwrap(), "000-0");
    }

    #[test]
    fn function_name_identifier_extends_to_operator() {
        let scalars: Vec<char> = "sin2".chars().collect();
        let pattern = token_pattern(&scalars);
        assert_eq!(function_name_length(&pattern, 0), 3);
        let scalars: Vec<char> = "log".chars().collect();
        assert_eq!(function_name_length(&token_pattern(&scalars), 0), 3);
    }

    #[test]
    fn variant_table() {
        let s = |v: &[&str]| v.iter().map(|x| x.to_string()).collect::<Vec<_>>();
        assert_eq!(variant(&s(&["fraktur"]), &s(&["b"]), "mi"), "bold-fraktur");
        assert_eq!(variant(&s(&[]), &s(&["p"]), "mn"), "");
        assert_eq!(variant(&s(&[]), &s(&["p"]), "mi"), "normal");
    }
}
