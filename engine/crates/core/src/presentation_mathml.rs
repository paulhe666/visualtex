//! Presentation MathML export for Office compatibility. This encoder consumes
//! the retained core tree directly; it never reparses through MathJax/MathLive
//! and never mutates the editing document or Office Session state.
use std::sync::Arc;
use crate::{
    EngineError, EngineResult, ErrorCode,
    accents::{self, Mark},
    fractions::FractionForm,
    grids::{GridColumnAlignment, GridKind},
    macros::BuiltinMacroKind,
    model::{AtomClass, Node, NodeKind, ScriptKind},
    operators::{self, OperatorMark},
    parser,
    source::SourceText,
};

const MAX_OUTPUT_BYTES: usize = 8 * 1024 * 1024;
const MAX_DEPTH: usize = 96;

#[derive(Debug, Clone, serde::Serialize)]
#[serde(rename_all = "camelCase")]
pub struct PresentationMathMl {
    pub math_ml: String,
    /// Aligned tables are identified only inside this transient core result.
    /// The host strips the private attribute before handing MathML to Office.
    pub aligned_grid_ids: Vec<u32>,
}

pub fn encode(document: &parser::ParseResult, display_mode: bool) -> EngineResult<PresentationMathMl> {
    if !document.complete {
        return Err(EngineError::new(ErrorCode::IncompleteSource, "Presentation MathML requires a complete formula"));
    }
    if !document.supported {
        return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Presentation MathML cannot export unsupported source"));
    }
    if !document.exportable {
        return Err(EngineError::new(ErrorCode::IncompleteSource, "Fill all formula slots before Presentation MathML export"));
    }
    let mut encoder = Encoder { display_mode, aligned_grid_ids: Vec::new(), output_bytes: 0 };
    let body = encoder.node(&document.tree, 0)?;
    let display = if display_mode { "block" } else { "inline" };
    let math_ml = format!("<math xmlns=\"http://www.w3.org/1998/Math/MathML\" display=\"{display}\">{body}</math>");
    if math_ml.len() > MAX_OUTPUT_BYTES {
        return Err(EngineError::new(ErrorCode::ResourceLimit, "Presentation MathML output budget exceeded"));
    }
    Ok(PresentationMathMl { math_ml, aligned_grid_ids: encoder.aligned_grid_ids })
}

struct Encoder {
    display_mode: bool,
    aligned_grid_ids: Vec<u32>,
    output_bytes: usize,
}

impl Encoder {
    fn budget(&mut self, value: String) -> EngineResult<String> {
        self.output_bytes = self.output_bytes.checked_add(value.len())
            .ok_or_else(|| EngineError::new(ErrorCode::ResourceLimit, "Presentation MathML output budget overflow"))?;
        if self.output_bytes > MAX_OUTPUT_BYTES {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Presentation MathML output budget exceeded"));
        }
        Ok(value)
    }

    fn node(&mut self, node: &Node, depth: usize) -> EngineResult<String> {
        if depth > MAX_DEPTH {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Presentation MathML nesting budget exceeded"));
        }
        let value = match &node.kind {
            NodeKind::Row { entries, .. } => self.row(entries.iter().map(|entry| entry.node.as_ref()), depth + 1)?,
            NodeKind::Grid { environment, columns, array_spec, rows, .. } =>
                self.grid(node.id.0, *environment, columns.len(), array_spec.as_ref(), rows, depth + 1)?,
            NodeKind::GridRow { cells, .. } => {
                let mut body = String::new();
                for cell in cells { body.push_str("<mtd>"); body.push_str(&self.node(cell, depth + 1)?); body.push_str("</mtd>"); }
                format!("<mtr>{body}</mtr>")
            }
            NodeKind::Symbol { value, class, .. } => token(value, *class),
            NodeKind::Text { text, .. } => format!("<mtext>{}</mtext>", escape(text)),
            NodeKind::Fraction { command, numerator, denominator, .. } => {
                let spec = crate::fractions::spec(command).ok_or_else(|| EngineError::new(ErrorCode::UnsupportedSyntax, "Unknown fraction form"))?;
                let num = self.node(numerator, depth + 1)?;
                let den = self.node(denominator, depth + 1)?;
                match spec.form {
                    FractionForm::Barred => format!("<mfrac>{num}{den}</mfrac>"),
                    FractionForm::Binomial => format!("<mrow><mo fence=\"true\" stretchy=\"true\">(</mo><mfrac linethickness=\"0\">{num}{den}</mfrac><mo fence=\"true\" stretchy=\"true\">)</mo></mrow>"),
                    FractionForm::Bevelled => format!("<mfrac bevelled=\"true\">{num}{den}</mfrac>"),
                }
            }
            NodeKind::Radical { degree, radicand, .. } => {
                let body = self.node(radicand, depth + 1)?;
                if let Some(degree) = degree {
                    format!("<mroot>{body}{}</mroot>", self.node(degree, depth + 1)?)
                } else { format!("<msqrt>{body}</msqrt>") }
            }
            NodeKind::Fence { left, body, right, .. } => format!(
                "<mrow>{}<mrow>{}</mrow>{}</mrow>", delimiter(left, true), self.node(body, depth + 1)?, delimiter(right, true)),
            NodeKind::MiddleDelimiter { delimiter: source, .. } => delimiter(source, false),
            NodeKind::Accent { command, body, .. } => self.accent(command, body, depth + 1)?,
            NodeKind::Annotation { above, annotation, base, .. } => {
                let base = self.node(base, depth + 1)?;
                let annotation = self.node(annotation, depth + 1)?;
                if *above { format!("<mover>{base}{annotation}</mover>") } else { format!("<munder>{base}{annotation}</munder>") }
            }
            NodeKind::LargeOperator { command, parts } => {
                let spec = operators::spec(command).ok_or_else(|| EngineError::new(ErrorCode::UnsupportedSyntax, "Unknown large operator"))?;
                let base = match spec.mark {
                    OperatorMark::Glyph(ch) => format!("<mo largeop=\"true\">{}</mo>", escape_char(ch)),
                    OperatorMark::Text(text) => format!("<mo>{}</mo>", escape(text)),
                };
                self.operator_tail(base, parts, spec.display_limits, depth + 1)?
            }
            NodeKind::GeneralOperator { body, parts, .. } => {
                let base = format!("<mrow>{}</mrow>", self.node(body, depth + 1)?);
                self.operator_tail(base, parts, true, depth + 1)?
            }
            NodeKind::CustomSymbol { asset, parts, .. } => {
                let fallback = asset.definition().get("ommlFallback").and_then(serde_json::Value::as_str)
                    .map(str::trim).filter(|value| !value.is_empty())
                    .ok_or_else(|| EngineError::new(ErrorCode::UnsupportedSyntax, "Custom symbol has no Office semantic fallback"))?;
                let fallback = self.fallback_source(fallback, depth + 1)?;
                if asset.role() == AtomClass::Operator || !parts.is_empty() {
                    let stacked_auto = match asset.limits() {
                        crate::custom_symbols::SymbolLimits::Auto => self.display_mode,
                        crate::custom_symbols::SymbolLimits::Limits => true,
                        crate::custom_symbols::SymbolLimits::Nolimits => false,
                    };
                    self.operator_tail(fallback, parts, stacked_auto, depth + 1)?
                } else { fallback }
            }
            NodeKind::MacroCall { macro_kind, arguments, .. } => self.macro_call(*macro_kind, arguments, depth + 1)?,
            NodeKind::Style { variant, body, .. } => format!(
                "<mstyle mathvariant=\"{}\">{}</mstyle>", math_variant(*variant), self.node(body, depth + 1)?),
            NodeKind::Space { width_mu, .. } => format!("<mspace width=\"{}em\"/>", decimal(*width_mu as f64 / 18.0)),
            NodeKind::Phantom { axis, body, .. } => {
                let body = self.node(body, depth + 1)?;
                match axis {
                    crate::spacing::PhantomAxis::Full => format!("<mphantom>{body}</mphantom>"),
                    crate::spacing::PhantomAxis::Horizontal => format!("<mpadded height=\"0\" depth=\"0\"><mphantom>{body}</mphantom></mpadded>"),
                    crate::spacing::PhantomAxis::Vertical => format!("<mpadded width=\"0\"><mphantom>{body}</mphantom></mpadded>"),
                }
            }
            NodeKind::Boxed { body, .. } => format!("<menclose notation=\"box\">{}</menclose>", self.node(body, depth + 1)?),
            NodeKind::Color { spec, body, .. } => format!("<mstyle mathcolor=\"{}\">{}</mstyle>", spec.color.hex(), self.node(body, depth + 1)?),
            NodeKind::ColorChange { .. } => String::new(), // Row handles scoped color changes.
            NodeKind::Scripts { base, slots } => {
                let base = self.node(base, depth + 1)?;
                self.scripts(base, slots, false, depth + 1)?
            },
            NodeKind::EmptySlot { .. } => return Err(EngineError::new(ErrorCode::IncompleteSource, "Unfilled slot cannot be exported to Presentation MathML")),
            NodeKind::UnparsedSource { .. } => return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Opaque source cannot be exported to Presentation MathML")),
        };
        self.budget(value)
    }

    fn row<'a>(&mut self, nodes: impl Iterator<Item=&'a Node>, depth: usize) -> EngineResult<String> {
        let mut children = Vec::new();
        let mut current_color: Option<crate::paint::RgbColor> = None;
        for node in nodes {
            if let NodeKind::ColorChange { spec } = &node.kind { current_color = Some(spec.color); continue; }
            let mut child = self.node(node, depth + 1)?;
            if let Some(color) = current_color { child = format!("<mstyle mathcolor=\"{}\">{child}</mstyle>", color.hex()); }
            children.push(child);
        }
        Ok(match children.len() {
            0 => "<mrow></mrow>".into(),
            1 => children.pop().unwrap(),
            _ => format!("<mrow>{}</mrow>", children.concat()),
        })
    }

    fn scripts(&mut self, base: String, slots: &[crate::model::ScriptSlot], stacked: bool, depth: usize) -> EngineResult<String> {
        let sub = slots.iter().find(|slot| slot.script == ScriptKind::Subscript);
        let sup = slots.iter().find(|slot| slot.script == ScriptKind::Superscript);
        let sub = match sub { Some(slot) => Some(self.node(&slot.argument, depth + 1)?), None => None };
        let sup = match sup { Some(slot) => Some(self.node(&slot.argument, depth + 1)?), None => None };
        Ok(match (stacked, sub, sup) {
            (false, Some(sub), Some(sup)) => format!("<msubsup>{base}{sub}{sup}</msubsup>"),
            (false, Some(sub), None) => format!("<msub>{base}{sub}</msub>"),
            (false, None, Some(sup)) => format!("<msup>{base}{sup}</msup>"),
            (true, Some(sub), Some(sup)) => format!("<munderover>{base}{sub}{sup}</munderover>"),
            (true, Some(sub), None) => format!("<munder>{base}{sub}</munder>"),
            (true, None, Some(sup)) => format!("<mover>{base}{sup}</mover>"),
            (_, None, None) => base,
        })
    }

    fn operator_tail(&mut self, base: String, parts: &[operators::OperatorPart], default_display_limits: bool, depth: usize) -> EngineResult<String> {
        let mode = operators::effective_mode(parts);
        let stacked = match mode {
            operators::LimitMode::Auto => default_display_limits && self.display_mode,
            operators::LimitMode::Limits => true,
            operators::LimitMode::NoLimits => false,
            operators::LimitMode::DisplayLimits => self.display_mode,
        };
        let slots: Vec<_> = operators::slots(parts).cloned().collect();
        self.scripts(base, &slots, stacked, depth + 1)
    }

    fn accent(&mut self, command: &str, body: &Arc<Node>, depth: usize) -> EngineResult<String> {
        let spec = accents::spec(command).ok_or_else(|| EngineError::new(ErrorCode::UnsupportedSyntax, "Unknown accent"))?;
        let base = self.node(body, depth + 1)?;
        let mark = match spec.mark {
            Mark::Glyph { codepoint, .. } => escape_char(codepoint),
            Mark::Rule => if spec.below { "_".into() } else { "¯".into() },
            Mark::Segment => "―".into(),
        };
        let accent = format!("<mo stretchy=\"{}\">{mark}</mo>", if matches!(spec.mark, Mark::Glyph { stretch: true, .. } | Mark::Rule | Mark::Segment) { "true" } else { "false" });
        Ok(if spec.below { format!("<munder accentunder=\"true\">{base}{accent}</munder>") }
            else { format!("<mover accent=\"true\">{base}{accent}</mover>") })
    }

    fn grid(&mut self, grid_id: u32, environment: GridKind, columns: usize,
        array_spec: Option<&crate::grids::ArrayColumnSpec>, rows: &[Arc<Node>], depth: usize) -> EngineResult<String> {
        let mut body = String::new();
        for row in rows { body.push_str(&self.node(row, depth + 1)?); }
        let alignments = match environment {
            GridKind::Aligned | GridKind::Align | GridKind::AlignStar => (0..columns).map(|index| if index % 2 == 0 { "right" } else { "left" }).collect::<Vec<_>>(),
            GridKind::Cases => vec!["left"; columns],
            GridKind::Array => array_spec.ok_or_else(|| EngineError::new(ErrorCode::InvalidInput, "Array MathML requires column metadata"))?.alignments.iter().map(|alignment| match alignment { GridColumnAlignment::Left=>"left", GridColumnAlignment::Center=>"center", GridColumnAlignment::Right=>"right" }).collect(),
            _ => vec!["center"; columns],
        };
        let columnalign = alignments.join(" ");
        let private = if environment.is_aligned() {
            self.aligned_grid_ids.push(grid_id);
            format!(" data-visualtex-core-grid-id=\"{grid_id}\" data-mtef-tabs=\"true\"")
        } else { String::new() };
        let spacing = if environment.is_aligned() && columns > 1 {
            let values = (0..columns-1).map(|index| if index % 2 == 0 { "0em" } else { "2em" }).collect::<Vec<_>>().join(" ");
            format!(" columnspacing=\"{values}\"")
        } else { String::new() };
        let table = format!("<mtable columnalign=\"{columnalign}\"{spacing}{private}>{body}</mtable>");
        Ok(match environment {
            GridKind::PMatrix => wrap_table("(", table, ")"),
            GridKind::BMatrix => wrap_table("[", table, "]"),
            GridKind::BraceMatrix => wrap_table("{", table, "}"),
            GridKind::VMatrix => wrap_table("|", table, "|"),
            GridKind::DoubleVMatrix => wrap_table("‖", table, "‖"),
            GridKind::Cases => format!("<mrow><mo fence=\"true\" stretchy=\"true\">{{</mo>{table}</mrow>"),
            _ => table,
        })
    }

    fn macro_call(&mut self, kind: BuiltinMacroKind, args: &[Arc<Node>], depth: usize) -> EngineResult<String> {
        let arg = |this: &mut Self, index: usize| this.node(args.get(index).ok_or_else(|| EngineError::new(ErrorCode::InvalidInput, "Malformed built-in macro"))?, depth + 1);
        Ok(match kind {
            BuiltinMacroKind::Absolute => wrap_table("|", arg(self,0)?, "|"),
            BuiltinMacroKind::Norm => wrap_table("‖", arg(self,0)?, "‖"),
            BuiltinMacroKind::Bra => format!("<mrow><mo>⟨</mo>{}<mo>|</mo></mrow>", arg(self,0)?),
            BuiltinMacroKind::Ket => format!("<mrow><mo>|</mo>{}<mo>⟩</mo></mrow>", arg(self,0)?),
            BuiltinMacroKind::Expectation => format!("<mrow><mo>⟨</mo>{}<mo>⟩</mo></mrow>", arg(self,0)?),
            BuiltinMacroKind::Differential => format!("<mrow><mi mathvariant=\"normal\">d</mi>{}</mrow>", arg(self,0)?),
            BuiltinMacroKind::Derivative => derivative("d", arg(self,0)?, arg(self,1)?),
            BuiltinMacroKind::PartialDerivative => derivative("∂", arg(self,0)?, arg(self,1)?),
            BuiltinMacroKind::Commutator => delimited("[", &[arg(self,0)?, arg(self,1)?], "]"),
            BuiltinMacroKind::Anticommutator | BuiltinMacroKind::PoissonBracket => delimited("{", &[arg(self,0)?, arg(self,1)?], "}"),
            BuiltinMacroKind::BraKet => format!("<mrow><mo>⟨</mo>{}<mo>|</mo>{}<mo>⟩</mo></mrow>", arg(self,0)?,arg(self,1)?),
            BuiltinMacroKind::MatrixElement => format!("<mrow><mo>⟨</mo>{}<mo>|</mo>{}<mo>|</mo>{}<mo>⟩</mo></mrow>", arg(self,0)?,arg(self,1)?,arg(self,2)?),
            BuiltinMacroKind::KetBra => format!("<mrow><mo>|</mo>{}<mo>⟩</mo><mo>⟨</mo>{}<mo>|</mo></mrow>", arg(self,0)?,arg(self,1)?),
            BuiltinMacroKind::BoldVector => format!("<mstyle mathvariant=\"bold\">{}</mstyle>", arg(self,0)?),
            BuiltinMacroKind::ArrowVector => format!("<mover accent=\"true\"><mstyle mathvariant=\"bold\">{}</mstyle><mo>→</mo></mover>", arg(self,0)?),
            BuiltinMacroKind::UnitVector => format!("<mover accent=\"true\"><mstyle mathvariant=\"bold\">{}</mstyle><mo>ˆ</mo></mover>", arg(self,0)?),
        })
    }

    fn fallback_source(&mut self, source: &str, depth: usize) -> EngineResult<String> {
        if depth > MAX_DEPTH { return Err(EngineError::new(ErrorCode::ResourceLimit, "Custom symbol fallback nesting budget exceeded")); }
        let source = SourceText::new(1, source.to_owned())?;
        let parsed = parser::parse(source, crate::session::NodeId(1), &mut crate::model::IdAllocator::new(2))?;
        if !parsed.complete || !parsed.supported || !parsed.exportable {
            return Err(EngineError::new(ErrorCode::UnsupportedSyntax, "Custom symbol Office fallback is not exportable by the core"));
        }
        self.node(&parsed.tree, depth + 1)
    }
}

fn token(value: &str, class: AtomClass) -> String {
    let escaped = escape(value);
    match class {
        AtomClass::Number => format!("<mn>{escaped}</mn>"),
        AtomClass::Binary | AtomClass::Relation | AtomClass::Open | AtomClass::Close | AtomClass::Punctuation | AtomClass::Operator => format!("<mo>{escaped}</mo>"),
        _ => format!("<mi>{escaped}</mi>"),
    }
}
fn delimiter(source: &str, fence: bool) -> String {
    let value = crate::delimiters::glyph(source).flatten().map(escape_char).unwrap_or_default();
    if value.is_empty() { String::new() } else { format!("<mo fence=\"{}\" stretchy=\"true\">{value}</mo>", if fence {"true"} else {"false"}) }
}
fn wrap_table(left: &str, body: String, right: &str) -> String { format!("<mrow><mo fence=\"true\" stretchy=\"true\">{}</mo>{body}<mo fence=\"true\" stretchy=\"true\">{}</mo></mrow>", escape(left), escape(right)) }
fn derivative(mark: &str, numerator: String, denominator: String) -> String { format!("<mfrac><mrow><mi mathvariant=\"normal\">{}</mi>{numerator}</mrow><mrow><mi mathvariant=\"normal\">{}</mi>{denominator}</mrow></mfrac>", escape(mark), escape(mark)) }
fn delimited(left: &str, args: &[String], right: &str) -> String { let body=args.join("<mo>,</mo>");format!("<mrow><mo>{}</mo>{body}<mo>{}</mo></mrow>",escape(left),escape(right)) }
fn decimal(value: f64) -> String { let mut value=format!("{value:.6}");while value.contains('.')&&value.ends_with('0'){value.pop();}if value.ends_with('.'){value.pop();}value }
fn escape(value: &str) -> String { let mut out=String::with_capacity(value.len());for ch in value.chars(){match ch{'&'=>out.push_str("&amp;"),'<'=>out.push_str("&lt;"),'>'=>out.push_str("&gt;"),'\"'=>out.push_str("&quot;"),'\''=>out.push_str("&apos;"),_=>out.push(ch)}}out }
fn escape_char(ch: char) -> String { escape(&ch.to_string()) }
fn math_variant(value: crate::styles::MathVariant) -> &'static str { use crate::styles::MathVariant::*;match value{Default=>"normal",Upright=>"normal",Bold=>"bold",Italic=>"italic",BoldItalic=>"bold-italic",Sans=>"sans-serif",SansBold=>"bold-sans-serif",SansItalic=>"sans-serif-italic",SansBoldItalic=>"sans-serif-bold-italic",Monospace=>"monospace",DoubleStruck=>"double-struck",Script=>"script",BoldScript=>"bold-script",Fraktur=>"fraktur",BoldFraktur=>"bold-fraktur"} }
