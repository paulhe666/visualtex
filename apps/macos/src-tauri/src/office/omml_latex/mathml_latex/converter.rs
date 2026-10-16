//! Port of VisualTeX.WordVsto.MathMlToLatexConverter (MathML -> LaTeX).

use std::cell::Cell;

use unicode_segmentation::UnicodeSegmentation;

use super::contracts::{
    accent_is_stretchy, effective_color, explicit_color, is_single_mathml_character, resolve_declared_variant,
    resolve_effective_variant, resolve_effective_variant_auto, space_latex_source, wrap_color,
};
use super::dom::{Doc, NodeId};
use super::formula_xml;
use super::netstr::{
    contains_ignore_case, eq_ignore_case, first_is_math_symbol, int_try_parse, is_letter_low, is_letter_or_digit_low,
    is_lower_greek, is_null_or_white_space, is_upper_greek, len16, to_lower_invariant, to_upper_invariant, units,
};
use super::tables::{
    letterlike_identifier, named_operator, ALTERNATE_TOKEN_MAP, DELIMITER_MAP, GEOMETRY_TEXT_SYMBOLS,
    TOKEN_DEFAULT_MATH_VARIANTS, TOKEN_MAP,
};

pub fn convert(mathml: &str) -> Result<String, String> {
    if is_null_or_white_space(Some(mathml)) {
        return Err("Word OMML conversion produced empty MathML.".to_string());
    }
    let doc = formula_xml::parse(mathml, "MathML")?;
    let root = doc.root.ok_or_else(|| "Word MathML has no root element.".to_string())?;
    let converter = Converter { doc: &doc, alignment_mark_depth: Cell::new(0) };
    Ok(normalize(&converter.element(root)))
}

struct Converter<'a> {
    doc: &'a Doc,
    alignment_mark_depth: Cell<i32>,
}

fn trim_owned(value: &str) -> String {
    value.trim().to_string()
}

impl Converter<'_> {
    // ---- small XLinq helpers ----------------------------------------------

    fn local(&self, element: NodeId) -> &str {
        self.doc.local(element)
    }

    fn attr(&self, element: NodeId, name: &str) -> Option<&str> {
        self.doc.attr(element, name)
    }

    fn value(&self, element: NodeId) -> String {
        self.doc.value(element)
    }

    fn trimmed_value(&self, element: NodeId) -> String {
        trim_owned(&self.doc.value(element))
    }

    fn elements(&self, element: NodeId) -> Vec<NodeId> {
        self.doc.elements(element)
    }

    // ---- dispatch -----------------------------------------------------------

    fn element(&self, element: NodeId) -> String {
        let content = self.element_content(element);
        wrap_color(content, explicit_color(self.doc, element).as_deref())
    }

    fn element_content(&self, element: NodeId) -> String {
        match self.local(element) {
            "math" => self.children(element),
            "mrow" | "mstyle" | "mpadded" => self.classed_container(element),
            "semantics" | "maction" => {
                self.elements(element).first().map(|&first| self.element(first)).unwrap_or_default()
            }
            "annotation" | "annotation-xml" => String::new(),
            "mi" | "mn" | "mo" => self.token_element(element),
            "mtext" => self.text_element(element),
            "mspace" => self.space(element),
            "mfrac" => self.fraction(element),
            "msqrt" => format!("\\sqrt{{{}}}", self.children(element)),
            "mroot" => self.root(element),
            "msup" => self.script(element, false, true),
            "msub" => self.script(element, true, false),
            "msubsup" => self.sub_sup(element),
            "mover" => self.over(element),
            "munder" => self.under(element),
            "munderover" => self.under_over(element),
            "mfenced" => self.fenced(element),
            "mtable" => self.table(element),
            "mtr" | "mlabeledtr" => self.join_cells(element),
            "mtd" => self.children(element),
            "menclose" => self.enclose(element),
            "mphantom" => format!("\\phantom{{{}}}", self.children(element)),
            "mmultiscripts" => self.multi_scripts(element),
            "malignmark" | "maligngroup" => {
                if self.alignment_mark_depth.get() > 0 { "&".to_string() } else { String::new() }
            }
            "none" => String::new(),
            _ => self.children(element),
        }
    }

    fn join_cells(&self, row: NodeId) -> String {
        self.elements(row).into_iter().map(|cell| self.element(cell)).collect::<Vec<_>>().join(" & ")
    }

    fn space(&self, element: NodeId) -> String {
        space_latex_source(self.attr(element, "width")).unwrap_or_else(|| "\\,".to_string())
    }

    fn explicit_atom_class_command(&self, element: NodeId) -> Option<&'static str> {
        let class = to_upper_invariant(self.attr(element, "data-mjx-texclass")?.trim());
        match class.as_str() {
            "REL" => Some("\\mathrel"),
            "BIN" => Some("\\mathbin"),
            "OP" => Some("\\mathop"),
            _ => None,
        }
    }

    fn classed_container(&self, element: NodeId) -> String {
        let content = self.children(element);
        match self.explicit_atom_class_command(element) {
            Some(command) if !is_null_or_white_space(Some(&content)) => {
                format!("{command}{{{}}}", content.trim())
            }
            _ => content,
        }
    }

    fn children(&self, element: NodeId) -> String {
        let children = self.elements(element);
        let mut builder = String::new();
        let mut index = 0;
        while index < children.len() {
            if index + 2 < children.len()
                && self.is_simple_operator(children[index], "(")
                && self.is_simple_operator(children[index + 2], ")")
            {
                if let Some(binomial) = self.loose_binomial_body(children[index + 1]) {
                    builder.push_str(&binomial);
                    index += 3;
                    continue;
                }
            }
            if index + 1 < children.len() && self.is_function_application(children[index + 1]) {
                if let Some(applied) = self.applied_named_function(children[index]) {
                    builder.push_str(&applied);
                    index += 2;
                    continue;
                }
            }
            builder.push_str(&self.element(children[index]));
            index += 1;
        }
        builder
    }

    fn is_function_application(&self, element: NodeId) -> bool {
        let element = self.unwrap_transparent(element);
        self.local(element) == "mo"
            && self.trimmed_value(element) == "\u{2061}"
            && explicit_color(self.doc, element).is_none()
    }

    fn applied_named_function(&self, head: NodeId) -> Option<String> {
        let unwrapped = self.unwrap_transparent(head);
        if self.doc.descendants_and_self(head).into_iter().any(|item| explicit_color(self.doc, item).is_some()) {
            return None;
        }
        if let Some(command) = self.named_function_identifier(unwrapped) {
            return Some(command);
        }
        let kind = self.local(unwrapped).to_string();
        if kind == "munder" || kind == "munderover" {
            if let Some(&limit_base) = self.elements(unwrapped).first() {
                if let Some(command) = self.named_function_identifier(self.unwrap_transparent(limit_base)) {
                    if is_limit_like_operator_latex(&command) {
                        return Some(self.element(unwrapped) + " ");
                    }
                }
            }
        }
        if kind == "mrow" && self.explicit_atom_class_command(unwrapped).is_none() {
            let items = self.elements(unwrapped);
            if items.len() >= 2 {
                if let Some(tail) = self.applied_named_function(items[items.len() - 1]) {
                    let head: String = items[..items.len() - 1].iter().map(|&item| self.element(item)).collect();
                    return Some(head + &tail);
                }
            }
            return None;
        }
        if !matches!(kind.as_str(), "msup" | "msub" | "msubsup") {
            return None;
        }
        let parts = self.elements(unwrapped);
        if parts.len() < if kind == "msubsup" { 3 } else { 2 } {
            return None;
        }
        let command = self.named_function_identifier(self.unwrap_transparent(parts[0]))?;
        let scripts = match kind.as_str() {
            "msub" => format!("_{{{}}}", self.script_argument(parts[1])),
            "msup" => format!("^{{{}}}", self.script_argument(parts[1])),
            _ => format!("_{{{}}}^{{{}}}", self.script_argument(parts[1]), self.script_argument(parts[2])),
        };
        Some(format!("{}{} ", command.trim_end(), scripts))
    }

    fn is_operator_role(&self, element: NodeId) -> bool {
        let atom_class = self.attr(element, "data-mjx-texclass").map(|value| to_upper_invariant(value.trim()));
        atom_class.as_deref() == Some("OP")
            || (atom_class.is_none() && self.attr(element, "data-mjx-auto-op") != Some("false"))
    }

    fn named_function_identifier(&self, element: NodeId) -> Option<String> {
        if self.local(element) != "mi" {
            return None;
        }
        let token = self.trimmed_value(element);
        if len16(&token) < 2 || !token.chars().all(|c| c.is_ascii_lowercase()) {
            return None;
        }
        let declared = resolve_declared_variant(self.doc, element);
        let effective = resolve_effective_variant(self.doc, element, declared.as_deref());
        if !matches!(effective.as_deref(), Some("normal" | "upright")) {
            return None;
        }
        if !self.is_operator_role(element) {
            return None;
        }
        named_operator(&token).map(str::to_string)
    }

    fn loose_binomial_body(&self, body: NodeId) -> Option<String> {
        let body = self.unwrap_transparent(body);
        if self.local(body) == "mfrac" {
            let thickness = self.attr(body, "linethickness").unwrap_or("").trim().to_string();
            let fraction = self.elements(body);
            if !thickness.starts_with('0') || fraction.len() < 2 {
                return None;
            }
            return Some(format!("\\binom{{{}}}{{{}}}", self.element(fraction[0]), self.element(fraction[1])));
        }
        if self.local(body) != "mtable" || !eq_ignore_case(self.attr(body, "data-mtef-pile"), "true") {
            return None;
        }
        self.pile_binomial(body)
    }

    fn pile_binomial(&self, table: NodeId) -> Option<String> {
        let rows: Vec<NodeId> =
            self.elements(table).into_iter().filter(|&row| matches!(self.local(row), "mtr" | "mlabeledtr")).collect();
        if rows.len() != 2 {
            return None;
        }
        let cells: Vec<Vec<NodeId>> = rows
            .iter()
            .map(|&row| self.elements(row).into_iter().filter(|&cell| self.local(cell) == "mtd").collect())
            .collect();
        if cells.iter().any(|row| row.len() != 1) {
            return None;
        }
        Some(format!("\\binom{{{}}}{{{}}}", self.element(cells[0][0]), self.element(cells[1][0])))
    }

    fn is_simple_operator(&self, element: NodeId, value: &str) -> bool {
        let element = self.unwrap_transparent(element);
        self.local(element) == "mo" && self.trimmed_value(element) == value
    }

    fn fraction(&self, element: NodeId) -> String {
        let children = self.elements(element);
        if children.len() < 2 {
            return self.children(element);
        }
        let thickness = self.attr(element, "linethickness").unwrap_or("").trim();
        if thickness.starts_with('0') {
            return format!(
                "\\genfrac{{}}{{}}{{0pt}}{{}}{{{}}}{{{}}}",
                self.element(children[0]),
                self.element(children[1])
            );
        }
        let mut command = "\\frac";
        if let Some(style) = self.doc.parent(element) {
            if self.doc.is_element(style) && self.local(style) == "mstyle" && self.elements(style).len() == 1 {
                match self.attr(style, "displaystyle").map(str::trim) {
                    Some("false") => command = "\\tfrac",
                    Some("true") => command = "\\dfrac",
                    _ => {}
                }
            }
        }
        format!("{command}{{{}}}{{{}}}", self.element(children[0]), self.element(children[1]))
    }

    fn root(&self, element: NodeId) -> String {
        let children = self.elements(element);
        if children.len() < 2 {
            return format!("\\sqrt{{{}}}", self.children(element));
        }
        format!("\\sqrt[{}]{{{}}}", self.element(children[1]), self.element(children[0]))
    }

    fn script(&self, element: NodeId, subscript: bool, superscript: bool) -> String {
        let children = self.elements(element);
        if children.len() < 2 {
            return self.children(element);
        }
        let mut result = group_base(&self.element(children[0]));
        if superscript && self.local(children[1]) == "mo" {
            let primes = self.trimmed_value(children[1]);
            if !primes.is_empty() && primes.chars().all(|c| matches!(c, '\u{2032}' | '\u{2033}' | '\u{2034}')) {
                let count: u32 = primes.chars().map(|c| c as u32 - 0x2031).sum();
                return result + &"'".repeat(count as usize);
            }
        }
        if subscript {
            result += &format!("_{{{}}}", self.script_argument(children[1]));
        }
        if superscript {
            result += &format!("^{{{}}}", self.script_argument(children[1]));
        }
        result
    }

    fn sub_sup(&self, element: NodeId) -> String {
        let children = self.elements(element);
        if children.len() < 3 {
            return self.children(element);
        }
        format!(
            "{}_{{{}}}^{{{}}}",
            group_base(&self.element(children[0])),
            self.script_argument(children[1]),
            self.script_argument(children[2])
        )
    }

    fn script_argument(&self, element: NodeId) -> String {
        self.element(element).trim_end().to_string()
    }

    fn over(&self, element: NodeId) -> String {
        let children = self.elements(element);
        if children.len() < 2 {
            return self.children(element);
        }
        if let Some(brace) = self.annotated_horizontal_brace(children[0], children[1], true) {
            return brace;
        }
        let mut body = self.element(children[0]);
        let over = self.trimmed_value(children[1]);
        let stretchy = accent_is_stretchy(self.doc, children[1]);
        let accent_style = if matches!(
            over.as_str(),
            "\u{00AF}" | "\u{203E}" | "\u{2015}" | "\u{2192}" | "\u{2190}" | "\u{2194}" | "^" | "\u{02C6}" | "~"
                | "\u{02DC}" | "." | "\u{02D9}" | "\u{00A8}"
        ) {
            resolve_effective_variant_auto(self.doc, children[1])
        } else {
            None
        };
        let style_command = match accent_style.as_deref() {
            Some("bold") => Some("\\mathbf"),
            Some("italic") => Some("\\mathit"),
            Some("bold-italic") => Some("\\mathbfit"),
            _ => None,
        };
        if style_command.is_some() {
            body = format!("\\mathnormal{{{body}}}");
        }
        let mark_color = effective_color(self.doc, children[1]);
        let parent_color = effective_color(self.doc, element);
        let independent_color = mark_color.is_some() && mark_color != parent_color;
        if independent_color {
            let base_color = effective_color(self.doc, children[0]).unwrap_or_else(|| "#000000".to_string());
            body = wrap_color(body, Some(&base_color));
        }
        let fixed_accent = self.is_fixed_accent_over_one_character(element, children[0], children[1]);
        let mut result = match over.as_str() {
            "\u{00AF}" | "\u{203E}" | "\u{2015}" => {
                format!("{}{body}}}", if fixed_accent { "\\bar{" } else { "\\overline{" })
            }
            "\u{2192}" => format!("{}{body}}}", if stretchy && !fixed_accent { "\\overrightarrow{" } else { "\\vec{" }),
            "\u{2190}" => format!("\\overleftarrow{{{body}}}"),
            "\u{2194}" => format!("\\overleftrightarrow{{{body}}}"),
            "^" | "\u{02C6}" => format!("{}{body}}}", if stretchy && !fixed_accent { "\\widehat{" } else { "\\hat{" }),
            "~" | "\u{02DC}" => {
                format!("{}{body}}}", if stretchy && !fixed_accent { "\\widetilde{" } else { "\\tilde{" })
            }
            "." | "\u{02D9}" => format!("\\dot{{{body}}}"),
            "\u{00A8}" => format!("\\ddot{{{body}}}"),
            "\u{20DB}" => format!("\\dddot{{{body}}}"),
            "\u{20DC}" => format!("\\ddddot{{{body}}}"),
            "\u{02C7}" => format!("\\check{{{body}}}"),
            "\u{00B4}" => format!("\\acute{{{body}}}"),
            "`" => format!("\\grave{{{body}}}"),
            "\u{02D8}" => format!("\\breve{{{body}}}"),
            "\u{02DA}" => format!("\\mathring{{{body}}}"),
            "\u{23DE}" | "\u{FE37}" => format!("\\overbrace{{{body}}}"),
            _ => format!("\\overset{{{}}}{{{body}}}", self.element(children[1])),
        };
        if let Some(command) = style_command {
            result = format!("{command}{{{result}}}");
        }
        wrap_color(result, if independent_color { mark_color.as_deref() } else { None })
    }

    fn is_fixed_accent_over_one_character(&self, over: NodeId, base: NodeId, mark: NodeId) -> bool {
        if !eq_ignore_case(self.attr(over, "accent").map(str::trim), "true") {
            return false;
        }
        if self.attr(mark, "stretchy").is_some() {
            return false;
        }
        let token = self.unwrap_transparent(base);
        matches!(self.local(token), "mi" | "mn" | "mo") && self.trimmed_value(token).graphemes(true).count() == 1
    }

    fn under(&self, element: NodeId) -> String {
        let children = self.elements(element);
        if children.len() < 2 {
            return self.children(element);
        }
        if let Some(brace) = self.annotated_horizontal_brace(children[0], children[1], false) {
            return brace;
        }
        let limit_base = self.unwrap_transparent(children[0]);
        if explicit_color(self.doc, limit_base).is_none() {
            let command = self.named_function_identifier(limit_base).or_else(|| {
                if self.local(limit_base) == "mo"
                    && eq_ignore_case(self.attr(limit_base, "movablelimits").map(str::trim), "false")
                {
                    named_operator(&self.trimmed_value(limit_base)).map(str::to_string)
                } else {
                    None
                }
            });
            if let Some(command) = command {
                if is_limit_like_operator_latex(&command) {
                    return format!("{}\\limits_{{{}}}", command.trim_end(), self.script_argument(children[1]));
                }
            }
        }
        let body = self.element(children[0]);
        let under = self.trimmed_value(children[1]);
        if is_limit_like_operator_latex(&body) {
            return format!("{}_{{{}}}", group_base(&body), self.script_argument(children[1]));
        }
        match under.as_str() {
            "_" | "\u{00AF}" | "\u{203E}" | "\u{2015}" => format!("\\underline{{{body}}}"),
            "\u{23DF}" | "\u{FE38}" => format!("\\underbrace{{{body}}}"),
            _ => format!("\\underset{{{}}}{{{body}}}", self.element(children[1])),
        }
    }

    fn annotated_horizontal_brace(&self, base: NodeId, annotation: NodeId, top: bool) -> Option<String> {
        let base = self.unwrap_transparent(base);
        if self.local(base) != if top { "mover" } else { "munder" } {
            return None;
        }
        let inner = self.elements(base);
        if inner.len() < 2 {
            return None;
        }
        let expected = if top { "\u{23DE}" } else { "\u{23DF}" };
        let presentation = if top { "\u{FE37}" } else { "\u{FE38}" };
        let marker = self.trimmed_value(inner[1]);
        let replacement_only = !marker.is_empty() && marker.chars().all(|c| c == '\u{FFFD}');
        let stretchy = eq_ignore_case(self.attr(inner[1], "stretchy"), "true");
        if marker != expected && marker != presentation && !(replacement_only && stretchy) {
            return None;
        }
        let body = self.element(inner[0]);
        let note = self.element(annotation).trim_end().to_string();
        Some(if top {
            format!("\\overbrace{{{body}}}^{{{note}}}")
        } else {
            format!("\\underbrace{{{body}}}_{{{note}}}")
        })
    }

    fn unwrap_transparent(&self, element: NodeId) -> NodeId {
        let mut current = element;
        while matches!(self.local(current), "mrow" | "mstyle" | "mpadded") {
            if self.explicit_atom_class_command(current).is_some() {
                break;
            }
            let children = self.elements(current);
            if children.len() != 1 {
                break;
            }
            current = children[0];
        }
        current
    }

    fn under_over(&self, element: NodeId) -> String {
        let children = self.elements(element);
        if children.len() < 3 {
            return self.children(element);
        }
        let base = if self.local(children[0]) == "mo" {
            match self.trimmed_value(children[0]).as_str() {
                "\u{222A}" => "\\bigcup ".to_string(),
                "\u{2229}" => "\\bigcap ".to_string(),
                _ => self.element(children[0]),
            }
        } else {
            self.element(children[0])
        };
        format!(
            "{}_{{{}}}^{{{}}}",
            group_base(&base),
            self.script_argument(children[1]),
            self.script_argument(children[2])
        )
    }

    fn fenced(&self, element: NodeId) -> String {
        let open = self.attr(element, "open").unwrap_or("(").to_string();
        let close = self.attr(element, "close").unwrap_or(")").to_string();
        if let Some(binomial) = self.fenced_no_bar_fraction_binomial(element, &open, &close) {
            return binomial;
        }
        if let Some(binomial) = self.mathtype_binomial_pile(element, &open, &close) {
            return binomial;
        }
        let separators = self.attr(element, "separators").unwrap_or(",");
        let parts: Vec<String> = self.elements(element).into_iter().map(|child| self.element(child)).collect();
        // C# takes separators[0], one UTF-16 unit; an astral first separator
        // would be a lone surrogate there and is kept whole here.
        let separator = separators.chars().next().map(String::from).unwrap_or_default();
        delimiter(&open, true) + &parts.join(&separator) + &delimiter(&close, false)
    }

    fn fenced_no_bar_fraction_binomial(&self, fenced: NodeId, open: &str, close: &str) -> Option<String> {
        if open != "(" || close != ")" {
            return None;
        }
        let children = self.elements(fenced);
        if children.len() != 1 {
            return None;
        }
        let child = self.unwrap_transparent(children[0]);
        if self.local(child) != "mfrac" {
            return None;
        }
        let thickness = self.attr(child, "linethickness").unwrap_or("").trim();
        if !thickness.starts_with('0') {
            return None;
        }
        let fraction = self.elements(child);
        if fraction.len() < 2 {
            return None;
        }
        Some(format!("\\binom{{{}}}{{{}}}", self.element(fraction[0]), self.element(fraction[1])))
    }

    fn mathtype_binomial_pile(&self, fenced: NodeId, open: &str, close: &str) -> Option<String> {
        if open != "(" || close != ")" {
            return None;
        }
        let children = self.elements(fenced);
        if children.len() != 1 {
            return None;
        }
        let mut table = children[0];
        if self.local(table) == "mrow" {
            let nested = self.elements(table);
            if nested.len() == 1 && self.local(nested[0]) == "mtable" {
                table = nested[0];
            }
        }
        if self.local(table) != "mtable" || !eq_ignore_case(self.attr(table, "data-mtef-pile"), "true") {
            return None;
        }
        self.pile_binomial(table)
    }

    fn table_rows(&self, table: NodeId) -> Vec<NodeId> {
        self.elements(table).into_iter().filter(|&row| matches!(self.local(row), "mtr" | "mlabeledtr")).collect()
    }

    fn is_aligned_column_pattern(&self, table: NodeId) -> bool {
        let columns: Vec<&str> =
            self.attr(table, "data-vt-columnalign").unwrap_or("").split(' ').filter(|c| !c.is_empty()).collect();
        if columns.len() < 2 {
            return false;
        }
        columns.iter().enumerate().all(|(index, &column)| column == if index % 2 == 0 { "right" } else { "left" })
    }

    fn table(&self, element: NodeId) -> String {
        if self.is_aligned_column_pattern(element) {
            let rows: Vec<String> =
                self.table_rows(element).into_iter().map(|row| self.join_cells(row).trim().to_string()).collect();
            return format!("\\begin{{aligned}}{}\\end{{aligned}}", rows.join(" \\\\ "));
        }
        if self.doc.descendants(element).into_iter().any(|item| self.local(item) == "malignmark") {
            self.alignment_mark_depth.set(self.alignment_mark_depth.get() + 1);
            let rows: Vec<String> = self
                .table_rows(element)
                .into_iter()
                .map(|row| self.join_cells(row).trim().to_string())
                .map(|row| match row.strip_prefix('&') {
                    Some(rest) => rest.trim_start().to_string(),
                    None => row,
                })
                .collect();
            self.alignment_mark_depth.set(self.alignment_mark_depth.get() - 1);
            return format!("\\begin{{aligned}}{}\\end{{aligned}}", rows.join(" \\\\ "));
        }
        let rows: Vec<String> = self.table_rows(element).into_iter().map(|row| self.join_cells(row)).collect();
        if eq_ignore_case(self.attr(element, "data-mtef-tabs"), "true") {
            return format!("\\begin{{aligned}}{}\\end{{aligned}}", rows.join(" \\\\ "));
        }
        format!("\\begin{{matrix}}{}\\end{{matrix}}", rows.join(" \\\\ "))
    }

    fn enclose(&self, element: NodeId) -> String {
        let notation = to_lower_invariant(self.attr(element, "notation").unwrap_or(""));
        let body = self.children(element);
        if notation.contains("box") {
            return format!("\\boxed{{{body}}}");
        }
        if notation.contains("radical") {
            return format!("\\sqrt{{{body}}}");
        }
        let up = notation.contains("updiagonalstrike");
        let down = notation.contains("downdiagonalstrike");
        if up && down {
            return format!("\\xcancel{{{body}}}");
        }
        if up {
            return format!("\\cancel{{{body}}}");
        }
        if down {
            return format!("\\bcancel{{{body}}}");
        }
        body
    }

    fn multi_scripts(&self, element: NodeId) -> String {
        let children = self.elements(element);
        if children.is_empty() {
            return String::new();
        }
        let mut builder = group_base(&self.element(children[0]));
        let mut index = 1;
        let mut post_pair = 0;
        while index < children.len() && self.local(children[index]) != "mprescripts" {
            let sub = self.script_argument(children[index]);
            let sup = if index + 1 < children.len() && self.local(children[index + 1]) != "mprescripts" {
                self.script_argument(children[index + 1])
            } else {
                String::new()
            };
            let pair = post_pair;
            post_pair += 1;
            if pair > 0 && (!sub.is_empty() || !sup.is_empty()) {
                builder.push_str("{}");
            }
            if !sub.is_empty() {
                builder.push_str(&format!("_{{{sub}}}"));
            }
            if !sup.is_empty() {
                builder.push_str(&format!("^{{{sup}}}"));
            }
            index += 2;
        }
        if index < children.len() && self.local(children[index]) == "mprescripts" {
            let mut prescripts: Vec<String> = Vec::new();
            index += 1;
            while index < children.len() {
                let sub = self.script_argument(children[index]);
                let sup =
                    if index + 1 < children.len() { self.script_argument(children[index + 1]) } else { String::new() };
                if !(sub.is_empty() && sup.is_empty()) {
                    let mut column = "{}".to_string();
                    if !sub.is_empty() {
                        column.push_str(&format!("_{{{sub}}}"));
                    }
                    if !sup.is_empty() {
                        column.push_str(&format!("^{{{sup}}}"));
                    }
                    prescripts.push(column);
                }
                index += 2;
            }
            prescripts.reverse();
            builder = prescripts.concat() + &builder;
        }
        builder
    }

    fn text_element(&self, element: NodeId) -> String {
        let raw = self.value(element);
        let value = raw.trim();
        let code_units = units(value);
        if code_units.len() == 1 && code_units[0] <= 0x024F && is_letter_low(code_units[0]) {
            return format!("\\mathrm{{{}}}", escape_math_identifier(value));
        }
        if code_units.len() == 1 && GEOMETRY_TEXT_SYMBOLS.contains(&code_units[0]) {
            return convert_token(value);
        }
        format!("\\text{{{}}}", escape_text(&raw))
    }

    fn token_element(&self, element: NodeId) -> String {
        let token = self.trimmed_value(element);
        let name = self.local(element).to_string();
        let declared_variant = resolve_declared_variant(self.doc, element);
        let mut variant = declared_variant.clone().unwrap_or_default();
        let effective_variant = resolve_effective_variant(self.doc, element, declared_variant.as_deref());
        if declared_variant.is_none() && effective_variant.as_deref() == Some("italic") && first_is_math_symbol(&token)
        {
            variant = effective_variant.clone().unwrap();
        }
        let native_character_style = int_try_parse(self.attr(element, "data-mtef-character-style")).unwrap_or(0);
        if name == "mi"
            && len16(&token) > 1
            && token.chars().all(|c| c.is_ascii_alphabetic())
            && matches!(effective_variant.as_deref(), Some("normal" | "upright"))
        {
            let open = if self.is_operator_role(element) { "\\operatorname{" } else { "\\mathrm{" };
            return format!("{open}{}}}", escape_math_identifier(&token));
        }
        if name == "mo" {
            if let Some(named) = named_operator(&token) {
                if contains_ignore_case(&variant, "bold-italic") {
                    return format!("\\mathbfit{{{}}}", named.trim());
                }
                if contains_ignore_case(&variant, "bold") {
                    return format!("\\boldsymbol{{{}}}", named.trim());
                }
                if contains_ignore_case(&variant, "italic") {
                    return format!("\\mathit{{{}}}", named.trim());
                }
                return named.to_string();
            }
        }
        let explicitly_upright = contains_ignore_case(&variant, "normal") || contains_ignore_case(&variant, "upright");
        if name == "mi" {
            if let Some(letterlike) = letterlike_identifier(&token) {
                return apply_native_character_style(letterlike.to_string(), native_character_style, &variant, false);
            }
        }
        if variant == "italic" && matches!(native_character_style, 0 | 2) {
            if let Some(source) = self.intrinsic_italic_identifier_source(element, &token) {
                return source;
            }
        }
        if (name == "mi" && is_latin_identifier(&token)) || name == "mn" {
            let escaped = escape_math_identifier(&token);
            let has = |needle: &str| contains_ignore_case(&variant, needle);
            let converted = if has("double-struck") {
                Some(format!("\\mathbb{{{escaped}}}"))
            } else if has("bold-fraktur") {
                Some(format!("\\boldsymbol{{\\mathfrak{{{escaped}}}}}"))
            } else if has("fraktur") {
                Some(format!("\\mathfrak{{{escaped}}}"))
            } else if has("bold-script") {
                Some(format!("\\boldsymbol{{\\mathcal{{{escaped}}}}}"))
            } else if has("script") {
                Some(format!("\\mathcal{{{escaped}}}"))
            } else if has("monospace") {
                Some(format!("\\mathtt{{{escaped}}}"))
            } else if has("bold-sans-serif") {
                Some(format!("\\boldsymbol{{\\mathsf{{{escaped}}}}}"))
            } else if has("sans-serif") {
                Some(format!("\\mathsf{{{escaped}}}"))
            } else if has("bold-italic") {
                Some(format!("\\mathbfit{{{escaped}}}"))
            } else if has("bold") {
                Some(format!("\\mathbf{{{escaped}}}"))
            } else if name == "mn" && eq_ignore_case(Some(&variant), "italic") {
                Some(format!("\\mathit{{{escaped}}}"))
            } else if explicitly_upright {
                Some(format!("\\mathrm{{{escaped}}}"))
            } else {
                None
            };
            if let Some(converted) = converted {
                return apply_native_character_style(converted, native_character_style, &variant, false);
            }
        }
        if name == "mi" && is_greek_identifier(&token) {
            return convert_greek_identifier(
                &token,
                native_character_style,
                declared_variant.as_deref(),
                effective_variant.as_deref(),
            );
        }
        let alternate = matches!(name.as_str(), "mi" | "mo")
            && self.attr(element, "data-mjx-alternate").map(str::trim) == Some("1");
        let plain = match ALTERNATE_TOKEN_MAP.get(token.as_str()) {
            Some(command) if alternate => command.to_string(),
            _ => convert_token(&token),
        };
        if matches!(name.as_str(), "mi" | "mo") {
            if contains_ignore_case(&variant, "bold-italic") {
                return format!("\\mathbfit{{{}}}", plain.trim());
            }
            if contains_ignore_case(&variant, "bold") {
                return format!("\\mathbf{{{}}}", plain.trim());
            }
            if contains_ignore_case(&variant, "italic") {
                return format!("\\mathit{{{}}}", plain.trim());
            }
            // A prime has no upright form; VisualTeX for Mac writes f' as an
            // upright apostrophe, which is f^{\prime}, not f^{\mathrm{\prime}}.
            if !token.is_empty()
                && token
                    .chars()
                    .all(|character| matches!(character, '\'' | '\u{2032}' | '\u{2033}' | '\u{2034}' | '\u{2057}'))
            {
                return plain;
            }
            if (explicitly_upright && name == "mi")
                || ((explicitly_upright || variant.is_empty())
                    && TOKEN_DEFAULT_MATH_VARIANTS.get(token.as_str()) == Some(&"italic"))
            {
                return format!("\\mathrm{{{}}}", plain.trim());
            }
        }
        plain
    }

    fn intrinsic_italic_identifier_source(&self, element: NodeId, token: &str) -> Option<String> {
        if self.local(element) != "mi"
            || !is_single_mathml_character(token)
            || self.attr(element, "data-mjx-variant") == Some("-tex-mathit")
        {
            return None;
        }
        let source = if is_latin_identifier(token) {
            escape_math_identifier(token)
        } else if is_greek_identifier(token) && units(token).into_iter().all(is_lower_greek) {
            convert_token(token)
        } else {
            String::new()
        };
        if source.is_empty() { None } else { Some(source) }
    }
}

fn convert_greek_identifier(
    token: &str,
    native_character_style: i32,
    declared_variant: Option<&str>,
    effective_variant: Option<&str>,
) -> String {
    let all_lower = units(token).into_iter().all(is_lower_greek);
    let all_upper = units(token).into_iter().all(is_upper_greek);
    let mut variant = declared_variant.or(effective_variant).unwrap_or("").to_string();
    if declared_variant.is_none() && native_character_style == 0 && variant == "italic" && all_lower {
        return convert_token(token);
    }
    variant = match native_character_style {
        1 => "bold".to_string(),
        2 => "italic".to_string(),
        3 => "bold-italic".to_string(),
        _ => variant,
    };
    match variant.as_str() {
        "normal" | "upright" => format!("\\mathrm{{{token}}}"),
        "italic" => format!("\\mathit{{{token}}}"),
        "bold" => match TOKEN_MAP.get(token) {
            Some(command) if all_upper => format!("\\boldsymbol{{{}}}", command.trim()),
            _ => format!("\\mathbf{{{token}}}"),
        },
        "bold-italic" => {
            if all_lower {
                format!("\\boldsymbol{{{}}}", convert_token(token).trim())
            } else {
                format!("\\boldsymbol{{\\mathit{{{token}}}}}")
            }
        }
        _ => apply_native_character_style(convert_token(token), native_character_style, &variant, true),
    }
}

fn apply_native_character_style(latex: String, character_style: i32, math_variant: &str, greek: bool) -> String {
    if !(1..=3).contains(&character_style) {
        return latex;
    }
    let has_bold = contains_ignore_case(math_variant, "bold");
    let has_italic = contains_ignore_case(math_variant, "italic");
    match character_style {
        1 if has_bold => latex,
        1 if greek => format!("\\boldsymbol{{{}}}", latex.trim()),
        1 => format!("\\mathbf{{{latex}}}"),
        2 if has_italic => latex,
        2 => format!("\\mathit{{{latex}}}"),
        3 if has_bold && has_italic => latex,
        3 if greek => format!("\\boldsymbol{{{}}}", latex.trim()),
        3 => format!("\\mathbfit{{{latex}}}"),
        _ => latex,
    }
}

fn is_limit_like_operator_latex(value: &str) -> bool {
    matches!(value.trim(), "\\lim" | "\\max" | "\\min" | "\\sup" | "\\inf" | "\\limsup" | "\\liminf")
}

fn is_latin_identifier(token: &str) -> bool {
    let code_units = units(token);
    !code_units.is_empty()
        && code_units.iter().any(|&unit| is_letter_low(unit))
        && code_units
            .iter()
            .all(|&unit| (unit <= 0x024F && is_letter_or_digit_low(unit)) || unit == b'.' as u16 || unit == b',' as u16)
}

fn is_greek_identifier(token: &str) -> bool {
    let code_units = units(token);
    !code_units.is_empty()
        && code_units.iter().all(|&unit| (0x0370..=0x03FF).contains(&unit) || (0x1F00..=0x1FFF).contains(&unit))
}

fn escape_math_identifier(value: &str) -> String {
    value
        .replace('\\', "\\backslash ")
        .replace('{', "\\{")
        .replace('}', "\\}")
        .replace('#', "\\#")
        .replace('%', "\\%")
        .replace('&', "\\&")
        .replace('_', "\\_")
}

pub(super) fn convert_token(value: &str) -> String {
    let token = value.trim();
    if let Some(mapped) = TOKEN_MAP.get(token) {
        return mapped.to_string();
    }
    if token.is_empty() {
        return String::new();
    }
    escape_math_identifier(token)
}

fn delimiter(value: &str, left: bool) -> String {
    let escaped = DELIMITER_MAP.get(value).copied().unwrap_or(value);
    format!("{}{escaped} ", if left { "\\left" } else { "\\right" })
}

fn group_base(value: &str) -> String {
    let trimmed = value.trim();
    if len16(trimmed) == 1 || trimmed.starts_with('\\') {
        return trimmed.to_string();
    }
    format!("{{{trimmed}}}")
}

fn escape_text(value: &str) -> String {
    value
        .replace('\\', "\\textbackslash{}")
        .replace('{', "\\{")
        .replace('}', "\\}")
        .replace('#', "\\#")
        .replace('%', "\\%")
        .replace('&', "\\&")
        .replace('_', "\\_")
}

fn normalize(value: &str) -> String {
    let mut result = value.trim().to_string();
    while result.contains("  ") {
        result = result.replace("  ", " ");
    }
    result
}
