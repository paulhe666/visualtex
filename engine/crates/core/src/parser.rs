mod environments;
use std::sync::Arc;
use serde::Serialize;
use crate::error::{EngineError, EngineResult, ErrorCode};
use crate::model::{self, Grouping, IdAllocator, Node, NodeKind, RowEntry, ScriptKind, ScriptSlot, TreeFacts, MAX_DEPTH, MAX_NODES};
use crate::session::NodeId;
use crate::source::{SourceSpan, SourceText, Token, TokenKind, tokenize};
use crate::symbols::{literal_class, named_symbol};

const MAX_DIAGNOSTICS: usize = 128;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum Severity { Error, Unsupported }

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum DiagnosticCode {
    MissingArgument, UnclosedGroup, UnexpectedClosingBrace, MissingBase,
    DuplicateScript, UnknownCommand, UnsupportedTextCommand, UnfinishedCommand,
    UnsupportedToken, UnbracedStructure, UnsupportedColor, UnclosedColorModel, UnclosedOptionalArgument,
    MissingDelimiter, UnsupportedDelimiter, UnclosedFence, UnexpectedRightDelimiter, UnexpectedMiddleDelimiter,
    MissingEnvironmentName, MismatchedEnvironment, UnclosedEnvironment,
    UnsupportedEnvironment, UnsupportedEnvironmentOption, UnexpectedEnvironmentEnd, UnsupportedCustomArtwork,
}

#[derive(Debug, Clone, Serialize)]
pub struct Diagnostic {
    pub code: DiagnosticCode,
    pub severity: Severity,
    pub span: SourceSpan,
    pub message: String,
}

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParseResult {
    #[serde(skip)]
    pub source: Arc<SourceText>,
    #[serde(skip_serializing_if="crate::custom_symbols::empty_catalog")]
    pub symbol_assets: Arc<crate::custom_symbols::SymbolCatalog>,
    pub tree: Arc<Node>,
    pub tokens: Vec<Token>,
    pub diagnostics: Vec<Diagnostic>,
    pub facts: TreeFacts,
    pub complete: bool,
    pub supported: bool,
    pub exportable: bool,
}

/// Parse into an owned mathematical tree while keeping an exact lexical source
/// range for every unchanged node. No TeX expansion or third-party renderer runs.
pub fn parse(source: Arc<SourceText>, root_id: NodeId, ids: &mut IdAllocator) -> EngineResult<ParseResult> {
    let tokens = tokenize(&source)?;
    parse_tokens(source, tokens, root_id, ids)
}

pub fn parse_with_symbols(source: Arc<SourceText>, root_id: NodeId, ids: &mut IdAllocator, symbols: Arc<crate::custom_symbols::SymbolCatalog>) -> EngineResult<ParseResult> {
    let tokens=tokenize(&source)?;parse_tokens_with_symbols(source,tokens,root_id,ids,symbols)
}
pub(crate) fn parse_tokens(source: Arc<SourceText>, tokens: Vec<Token>, root_id: NodeId, ids: &mut IdAllocator) -> EngineResult<ParseResult> {
    parse_tokens_with_symbols(source,tokens,root_id,ids,Arc::default())
}
pub(crate) fn parse_tokens_with_symbols(source: Arc<SourceText>, tokens: Vec<Token>, root_id: NodeId, ids: &mut IdAllocator, symbols: Arc<crate::custom_symbols::SymbolCatalog>) -> EngineResult<ParseResult> {
    let end_position = source.len();
    let mut parser = Parser { source: Arc::clone(&source), tokens, at: 0, ids, end_position,
        diagnostics: Vec::new(), nodes: 0, allow_middle: false, symbols:Arc::clone(&symbols) };
    let tree = parser.row(Grouping::Bare, Some(root_id), 0)?;
    assemble_with_symbols(source, tree, parser.tokens, parser.diagnostics, symbols)
}

pub(crate) fn assemble_with_symbols(source: Arc<SourceText>, tree: Arc<Node>, tokens: Vec<Token>, mut diagnostics: Vec<Diagnostic>, symbol_assets: Arc<crate::custom_symbols::SymbolCatalog>) -> EngineResult<ParseResult> {
    if diagnostics.len() > MAX_DIAGNOSTICS {
        return Err(EngineError::new(ErrorCode::ResourceLimit, "Diagnostic budget exhausted"));
    }
    diagnostics.sort_by_key(|d| (d.span.byte_start, d.span.byte_end));
    let facts = model::validate_tree(&tree)?;
    let complete = !diagnostics.iter().any(|d| d.severity == Severity::Error);
    let resources_supported = if facts.custom_symbols>0 { crate::custom_symbols::validate_bindings(&tree,&symbol_assets)? } else { true };
    let supported = resources_supported && !diagnostics.iter().any(|d| d.severity == Severity::Unsupported);
    let exportable = complete && supported && tree.has_formula_content() && !tree.has_unfilled_slots();
    Ok(ParseResult { source, symbol_assets, tree, tokens, diagnostics, facts, complete, supported, exportable })
}

/// Parse a proven, token-aligned closed group using global source coordinates.
/// None means the edit escaped that group; the caller must parse the formula.
pub(crate) fn parse_braced_fragment(source: Arc<SourceText>, all_tokens: &[Token], start: usize, end: usize, symbols: Arc<crate::custom_symbols::SymbolCatalog>) -> EngineResult<Option<ParseResult>> {
    let first = all_tokens.partition_point(|t| t.span.byte_end as usize <= start);
    let after = all_tokens.partition_point(|t| (t.span.byte_start as usize) < end);
    let tokens = all_tokens[first..after].to_vec();
    if !tokens.first().is_some_and(|t| t.kind == TokenKind::OpenBrace && t.span.byte_start as usize == start)
        || !tokens.last().is_some_and(|t| t.kind == TokenKind::CloseBrace && t.span.byte_end as usize == end)
    {
        return Ok(None);
    }
    let mut ids = IdAllocator::new(2);
    let mut parser = Parser { source: Arc::clone(&source), tokens, at: 0, ids: &mut ids,
        end_position: end, diagnostics: Vec::new(), nodes: 0, allow_middle: false, symbols:Arc::clone(&symbols) };
    let tree = parser.row(Grouping::Braced, Some(NodeId(1)), 0)?;
    if parser.at != parser.tokens.len() || parser.diagnostics.iter().any(|d| d.severity == Severity::Error) {
        return Ok(None);
    }
    Ok(Some(assemble_with_symbols(source, tree, parser.tokens, parser.diagnostics, symbols)?))
}

struct Parser<'a> {
    source: Arc<SourceText>,
    tokens: Vec<Token>,
    at: usize,
    ids: &'a mut IdAllocator,
    diagnostics: Vec<Diagnostic>,
    nodes: usize,
    end_position: usize,
    allow_middle: bool,
    symbols:Arc<crate::custom_symbols::SymbolCatalog>,
}

impl Parser<'_> {
    fn position(&self) -> usize {
        self.tokens.get(self.at).map_or(self.end_position, |t| t.span.byte_start as usize)
    }

    fn kind(&self) -> Option<TokenKind> { self.tokens.get(self.at).map(|t| t.kind) }

    fn text(&self) -> &str {
        self.tokens.get(self.at).map_or("", |t| t.text(&self.source))
    }

    fn depth(&self, depth: usize) -> EngineResult<()> {
        if depth > MAX_DEPTH {
            Err(EngineError::new(ErrorCode::ResourceLimit, "Parser nesting budget exhausted"))
        } else { Ok(()) }
    }

    fn diagnostic(&mut self, code: DiagnosticCode, severity: Severity, start: usize, end: usize, message: &str) -> EngineResult<()> {
        if self.diagnostics.len() >= MAX_DIAGNOSTICS {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Diagnostic budget exhausted"));
        }
        self.diagnostics.push(Diagnostic { code, severity, span: self.source.span(start, end)?, message: message.into() });
        Ok(())
    }

    fn make(&mut self, kind: NodeKind, start: usize, end: usize, id: Option<NodeId>) -> EngineResult<Arc<Node>> {
        if self.nodes >= MAX_NODES {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Parser node budget exhausted"));
        }
        self.nodes += 1;
        Ok(Arc::new(Node { id: match id { Some(id) => id, None => self.ids.allocate()? },
            origin: Some(self.source.slice(start, end)?), kind }))
    }

    fn trivia(&mut self) -> String {
        let start = self.position();
        while self.tokens.get(self.at).is_some_and(Token::is_trivia) { self.at += 1; }
        self.source.text()[start..self.position()].to_owned()
    }

    fn row(&mut self, grouping: Grouping, id: Option<NodeId>, depth: usize) -> EngineResult<Arc<Node>> {
        self.depth(depth)?;
        let start = self.position();
        if grouping == Grouping::Braced { self.at += 1; }
        let mut entries = Vec::new();
        let trailing;
        loop {
            let leading = self.trivia();
            if self.kind().is_none() {
                trailing = leading;
                if grouping == Grouping::Braced {
                    self.diagnostic(DiagnosticCode::UnclosedGroup, Severity::Error, start, self.position(), "Group is missing its closing brace")?;
                }
                break;
            }
            if grouping == Grouping::Braced && self.kind() == Some(TokenKind::CloseBrace) {
                trailing = leading;
                self.at += 1;
                break;
            }
            let before = self.at;
            let atom = self.atom(depth + 1, true)?;
            let node = self.attach_scripts(atom, depth + 1)?;
            if self.at <= before {
                return Err(EngineError::new(ErrorCode::InvalidInput, "Parser failed to advance"));
            }
            entries.push(RowEntry { leading, node });
        }
        self.make(NodeKind::Row { grouping, entries, trailing }, start, self.position(), id)
    }

    fn atom(&mut self, depth: usize, allow_structure: bool) -> EngineResult<Arc<Node>> {
        self.depth(depth)?;
        let start = self.position();
        let Some(token) = self.tokens.get(self.at).cloned() else {
            self.diagnostic(DiagnosticCode::MissingArgument, Severity::Error, start, start, "Expected a math atom")?;
            return self.make(NodeKind::EmptySlot { label: String::new() }, start, start, None);
        };
        let raw = token.text(&self.source).to_owned();
        match token.kind {
            TokenKind::OpenBrace => self.row(Grouping::Braced, None, depth),
            TokenKind::CloseBrace => {
                self.at += 1;
                self.diagnostic(DiagnosticCode::UnexpectedClosingBrace, Severity::Error, start, self.position(), "Unexpected closing brace")?;
                self.make(NodeKind::UnparsedSource { raw }, start, self.position(), None)
            }
            TokenKind::Superscript | TokenKind::Subscript => {
                self.diagnostic(DiagnosticCode::MissingBase, Severity::Error, start, start, "Script has no base")?;
                self.make(NodeKind::EmptySlot { label: String::new() }, start, start, None)
            }
            TokenKind::ControlWord | TokenKind::ControlSymbol => {
                self.at += 1;
                if crate::operators::spec(&raw).is_some() {
                    return self.make(NodeKind::LargeOperator{command:raw,parts:Vec::new()},start,self.position(),None);
                }
                if let Some(width_mu)=crate::spacing::space_mu(&raw) {
                    return self.make(NodeKind::Space{command:raw,width_mu},start,self.position(),None);
                }
                if let Some((value, class)) = named_symbol(&raw) {
                    return self.make(NodeKind::Symbol { value: value.into(), class, command: Some(raw) }, start, self.position(), None);
                }
                if let Some(asset)=self.symbols.get(&raw) {
                    if !asset.supported(){self.diagnostic(DiagnosticCode::UnsupportedCustomArtwork,Severity::Unsupported,start,self.position(),"Custom symbol definition is retained but its artwork needs an unsupported capability")?;}
                    return self.make(NodeKind::CustomSymbol{command:raw,asset_sha256:asset.sha256().into(),asset,parts:Vec::new()},start,self.position(),None);
                }
                let structural = matches!(raw.as_str(), "\\sqrt" | "\\left" | "\\middle" | "\\text" | "\\operatorname" | "\\mathop" | "\\begin" | "\\placeholder")
                    || crate::fractions::spec(&raw).is_some()
                    || crate::accents::spec(&raw).is_some() || crate::accents::annotation_above(&raw).is_some()
                    || crate::styles::command_variant(&raw).is_some() || crate::spacing::phantom_axis(&raw).is_some()
                    || crate::macros::spec(&raw).is_some() || matches!(raw.as_str(),"\\boxed"|"\\textcolor"|"\\color");
                if structural && !allow_structure {
                    self.diagnostic(DiagnosticCode::UnbracedStructure, Severity::Error, start, self.position(), "A structural command used as an argument requires braces")?;
                    return self.make(NodeKind::UnparsedSource { raw }, start, self.position(), None);
                }
                match raw.as_str() {
                    _ if crate::macros::spec(&raw).is_some() => {
                        let spec=crate::macros::spec(&raw).unwrap();
                        let mut before_arguments=Vec::with_capacity(spec.arity);
                        let mut arguments=Vec::with_capacity(spec.arity);
                        for _ in 0..spec.arity {
                            let (before,argument)=self.argument(depth+1)?;
                            before_arguments.push(before);arguments.push(argument);
                        }
                        self.make(NodeKind::MacroCall {command:raw,macro_kind:spec.kind,before_arguments,arguments},start,self.position(),None)
                    }
                    _ if crate::fractions::spec(&raw).is_some() => {
                        let (before_numerator, numerator) = self.argument(depth + 1)?;
                        let (before_denominator, denominator) = self.argument(depth + 1)?;
                        self.make(NodeKind::Fraction { command: raw, before_numerator, numerator, before_denominator, denominator }, start, self.position(), None)
                    }
                    _ if crate::accents::spec(&raw).is_some() => {
                        let (before_body, body) = self.argument(depth + 1)?;
                        self.make(NodeKind::Accent { command: raw, before_body, body }, start, self.position(), None)
                    }
                    _ if crate::accents::annotation_above(&raw).is_some() => {
                        let above = crate::accents::annotation_above(&raw).unwrap();
                        let (before_annotation, annotation) = self.argument(depth + 1)?;
                        let (before_base, base) = self.argument(depth + 1)?;
                        self.make(NodeKind::Annotation { command: raw, above, before_annotation, annotation, before_base, base }, start, self.position(), None)
                    }
                    _ if crate::styles::command_variant(&raw).is_some() => {
                        let variant = crate::styles::command_variant(&raw).unwrap();
                        let (before_body, body) = self.argument(depth + 1)?;
                        self.make(NodeKind::Style { command: raw, before_body, variant, body }, start, self.position(), None)
                    }
                    _ if crate::spacing::phantom_axis(&raw).is_some() => {
                        let axis=crate::spacing::phantom_axis(&raw).unwrap();
                        let (before_body,body)=self.argument(depth+1)?;
                        self.make(NodeKind::Phantom {command:raw,before_body,axis,body},start,self.position(),None)
                    }
                    "\\textcolor" | "\\color" => self.color(start,raw=="\\color",depth+1),
                    "\\boxed" => {
                        let (before_body, body)=self.argument(depth+1)?;
                        self.make(NodeKind::Boxed{before_body,body},start,self.position(),None)
                    }
                    "\\mathop" => {
                        let (before_body, body) = self.argument(depth + 1)?;
                        self.make(NodeKind::GeneralOperator { command: raw, before_body, body, parts: Vec::new() }, start, self.position(), None)
                    }
                    "\\begin" => self.environment(start, depth + 1),
                    "\\end" => self.unexpected_environment_end(start),
                    "\\left" => self.fence(start, depth + 1),
                    "\\middle" => {
                        let (before_delimiter, delimiter) = self.delimiter()?;
                        if !self.allow_middle {
                            self.diagnostic(DiagnosticCode::UnexpectedMiddleDelimiter, Severity::Unsupported, start, self.position(), "Middle delimiter has no enclosing left/right fence")?;
                            let raw = self.source.text()[start..self.position()].to_owned();
                            self.make(NodeKind::UnparsedSource { raw }, start, self.position(), None)
                        } else {
                            self.make(NodeKind::MiddleDelimiter { before_delimiter, delimiter }, start, self.position(), None)
                        }
                    }
                    "\\right" => {
                        let _ = self.delimiter()?;
                        self.diagnostic(DiagnosticCode::UnexpectedRightDelimiter, Severity::Error, start, self.position(), "Closing delimiter has no matching left delimiter")?;
                        let raw = self.source.text()[start..self.position()].to_owned();
                        self.make(NodeKind::UnparsedSource { raw }, start, self.position(), None)
                    }
                    "\\sqrt" => {
                        let (before_degree, degree) = self.optional_degree(depth + 1)?;
                        let (before_radicand, radicand) = self.argument(depth + 1)?;
                        self.make(NodeKind::Radical { command: raw, before_degree, degree, before_radicand, radicand }, start, self.position(), None)
                    }
                    "\\text" | "\\operatorname" | "\\placeholder" => {
                        let before_body = self.trivia();
                        let (body_start, body_end) = self.text_body()?;
                        let text = self.decode_text(body_start, body_end)?;
                        if raw == "\\placeholder" {
                            self.make(NodeKind::EmptySlot { label: text }, start, self.position(), None)
                        } else {
                            self.make(NodeKind::Text { command: raw, before_body, text }, start, self.position(), None)
                        }
                    }
                    "\\" => {
                        self.diagnostic(DiagnosticCode::UnfinishedCommand, Severity::Error, start, self.position(), "Command is not yet complete")?;
                        self.make(NodeKind::UnparsedSource { raw }, start, self.position(), None)
                    }
                    _ => {
                        self.capture_opaque_arguments()?;
                        self.diagnostic(DiagnosticCode::UnknownCommand, Severity::Unsupported, start, self.position(), "Command is not supported by the current engine capability")?;
                        let raw = self.source.text()[start..self.position()].to_owned();
                        self.make(NodeKind::UnparsedSource { raw }, start, self.position(), None)
                    }
                }
            }
            TokenKind::Character => {
                self.at += 1;
                if crate::operators::spec(&raw).is_some() {
                    return self.make(NodeKind::LargeOperator{command:raw,parts:Vec::new()},start,self.position(),None);
                }
                if matches!(raw.as_str(), "$" | "&" | "#" | "~" | "'") || raw.chars().any(char::is_control) {
                    self.diagnostic(DiagnosticCode::UnsupportedToken, Severity::Unsupported, start, self.position(), "Token requires a language feature outside the current capability")?;
                    self.make(NodeKind::UnparsedSource { raw }, start, self.position(), None)
                } else {
                    self.make(NodeKind::Symbol { class: literal_class(&raw), value: raw, command: None }, start, self.position(), None)
                }
            }
            TokenKind::Space | TokenKind::Comment => {
                Err(EngineError::new(ErrorCode::InvalidInput, "Unexpected trivia at atom boundary"))
            }
        }
    }

    fn color(&mut self,start:usize,declaration:bool,depth:usize)->EngineResult<Arc<Node>> {
        self.depth(depth)?;
        let before=self.trivia();
        let (before_model,model,before_value)=if self.kind()==Some(TokenKind::Character)&&self.text()=="[" {
            self.at+=1;let model_start=self.position();
            while self.kind().is_some() && !(self.kind()==Some(TokenKind::Character)&&self.text()=="]")
                && !matches!(self.kind(),Some(TokenKind::OpenBrace|TokenKind::CloseBrace)) {
                self.at+=1;
                if self.position()-model_start>crate::paint::MAX_COLOR_SOURCE {
                    return Err(EngineError::new(ErrorCode::ResourceLimit,"Color model budget exceeded"));
                }
            }
            let model=self.source.text()[model_start..self.position()].to_owned();
            if self.kind()==Some(TokenKind::Character)&&self.text()=="]" {self.at+=1;}else {
                self.diagnostic(DiagnosticCode::UnclosedColorModel,Severity::Error,model_start,self.position(),"Color model has no closing bracket")?;
            }
            (before,Some(model),self.trivia())
        }else{(String::new(),None,before)};
        let (value_start,value_end)=self.text_body()?;
        let value=self.source.text()[value_start..value_end].to_owned();
        let body=if declaration{None}else{Some(self.argument(depth+1)?)};
        match crate::paint::ColorSpec::authored(before_model,model,before_value,value) {
            Ok(spec)=>{
                let kind=match body {Some((before_body,body))=>NodeKind::Color{spec,before_body,body},None=>NodeKind::ColorChange{spec}};
                self.make(kind,start,self.position(),None)
            }
            Err(error) if error.code==ErrorCode::ResourceLimit=>Err(error),
            Err(_)=>{
                self.diagnostic(DiagnosticCode::UnsupportedColor,Severity::Unsupported,start,self.position(),"Color model or value is outside the supported bounded set")?;
                let raw=self.source.text()[start..self.position()].to_owned();
                self.make(NodeKind::UnparsedSource{raw},start,self.position(),None)
            }
        }
    }

    fn delimiter(&mut self) -> EngineResult<(String, String)> {
        let before = self.trivia();
        let start = self.position();
        if matches!(self.kind(), None | Some(TokenKind::OpenBrace | TokenKind::CloseBrace | TokenKind::Superscript | TokenKind::Subscript)) {
            self.diagnostic(DiagnosticCode::MissingDelimiter, Severity::Error, start, start, "Delimiter is missing or requires escaping")?;
            return Ok((before, ".".into()));
        }
        let raw = self.text().to_owned();self.at += 1;
        if crate::delimiters::glyph(&raw).is_none() {
            self.diagnostic(DiagnosticCode::UnsupportedDelimiter, Severity::Unsupported, start, self.position(), "Delimiter is outside the supported delimiter set")?;
        }
        Ok((before, raw))
    }

    fn fence(&mut self, start: usize, depth: usize) -> EngineResult<Arc<Node>> {
        self.depth(depth)?;
        let (before_left, left) = self.delimiter()?;
        let first = self.at;let mut end_at = first;let mut braces = 0usize;let mut nested_fences = 0usize;let mut closed = false;
        while let Some(token) = self.tokens.get(end_at) {
            match token.kind {
                TokenKind::OpenBrace => braces += 1,
                TokenKind::CloseBrace if braces == 0 => break,
                TokenKind::CloseBrace => braces -= 1,
                TokenKind::ControlWord if braces == 0 && token.text(&self.source) == "\\left" => nested_fences += 1,
                TokenKind::ControlWord if braces == 0 && token.text(&self.source) == "\\right" => {
                    if nested_fences == 0 { closed = true;break; }nested_fences -= 1;
                }
                _ => {}
            }
            if braces + nested_fences > MAX_DEPTH { return Err(EngineError::new(ErrorCode::ResourceLimit, "Fence nesting budget exhausted")); }
            end_at += 1;
        }
        let end = self.tokens.get(end_at).map_or(self.end_position, |t| t.span.byte_start as usize);
        let mut nested = Parser { source: Arc::clone(&self.source), tokens: self.tokens[first..end_at].to_vec(),
            at: 0, ids: &mut *self.ids, end_position: end, diagnostics: Vec::new(), nodes: self.nodes, allow_middle: true, symbols:Arc::clone(&self.symbols) };
        let body = nested.row(Grouping::Bare, None, depth + 1)?;
        self.nodes = nested.nodes;self.diagnostics.extend(nested.diagnostics);self.at = end_at + usize::from(closed);
        let (before_right, right) = if closed { self.delimiter()? } else {
            self.diagnostic(DiagnosticCode::UnclosedFence, Severity::Error, start, end, "Left delimiter is missing a matching right delimiter")?;
            (String::new(), ".".into())
        };
        self.make(NodeKind::Fence { before_left, left, body, before_right, right }, start, self.position(), None)
    }

    fn optional_degree(&mut self, depth: usize) -> EngineResult<(String, Option<Arc<Node>>)> {
        self.depth(depth)?;
        let saved = self.at;
        let before = self.trivia();
        if self.kind() != Some(TokenKind::Character) || self.text() != "[" {
            self.at = saved;
            return Ok((String::new(), None));
        }
        self.at += 1;
        let start = self.position();
        let first = self.at;
        let mut end_at = first;
        let mut braces = 0usize;
        let mut closed = false;
        // Delimited TeX arguments are scanned before their command bodies.
        // Do not allow a nested command to consume an unprotected outer ']'.
        while let Some(token) = self.tokens.get(end_at) {
            match token.kind {
                TokenKind::OpenBrace => braces += 1,
                TokenKind::CloseBrace if braces == 0 => break,
                TokenKind::CloseBrace => braces -= 1,
                TokenKind::Character if braces == 0 && token.text(&self.source) == "]" => { closed = true; break; }
                _ => {}
            }
            if braces > MAX_DEPTH { return Err(EngineError::new(ErrorCode::ResourceLimit, "Root index nesting budget exhausted")); }
            end_at += 1;
        }
        let end = self.tokens.get(end_at).map_or(self.end_position, |t| t.span.byte_start as usize);
        let mut nested = Parser { source: Arc::clone(&self.source), tokens: self.tokens[first..end_at].to_vec(),
            at: 0, ids: &mut *self.ids, end_position: end, diagnostics: Vec::new(), nodes: self.nodes, allow_middle: self.allow_middle, symbols:Arc::clone(&self.symbols) };
        let row = nested.row(Grouping::Bare, None, depth + 1)?;
        self.nodes = nested.nodes;
        self.diagnostics.extend(nested.diagnostics);
        self.at = end_at + usize::from(closed);
        if !closed {
            self.diagnostic(DiagnosticCode::UnclosedOptionalArgument, Severity::Error, start, end, "Root index is missing its closing bracket")?;
        }
        // A protected index [{...}] uses that explicit group as its editable
        // row. Do not grow redundant nested rows on each save/reopen cycle.
        let row = match &row.kind {
            NodeKind::Row { entries, trailing, .. } if entries.len() == 1 && entries[0].leading.is_empty() && trailing.is_empty()
                && matches!(entries[0].node.kind, NodeKind::Row { grouping: Grouping::Braced, .. }) => Arc::clone(&entries[0].node),
            _ => row,
        };
        Ok((before, Some(row)))
    }

    fn argument(&mut self, depth: usize) -> EngineResult<(String, Arc<Node>)> {
        self.depth(depth)?;
        let before = self.trivia();
        let start = self.position();
        if self.kind() == Some(TokenKind::OpenBrace) {
            return Ok((before, self.row(Grouping::Braced, None, depth)?));
        }
        if matches!(self.kind(), None | Some(TokenKind::CloseBrace | TokenKind::Superscript | TokenKind::Subscript)) {
            self.diagnostic(DiagnosticCode::MissingArgument, Severity::Error, start, start, "Required argument is missing")?;
            let row = self.make(NodeKind::Row { grouping: Grouping::Bare, entries: Vec::new(), trailing: String::new() }, start, start, None)?;
            return Ok((before, row));
        }
        let node = self.atom(depth + 1, false)?;
        let row = self.make(NodeKind::Row { grouping: Grouping::Bare,
            entries: vec![RowEntry { leading: String::new(), node }], trailing: String::new() }, start, self.position(), None)?;
        Ok((before, row))
    }

    fn attach_scripts(&mut self, base: Arc<Node>, depth: usize) -> EngineResult<Arc<Node>> {
        // A color declaration changes paint state; it is not a script base.
        if matches!(base.kind,NodeKind::ColorChange{..}) {return Ok(base);}
        if crate::operators::parts(&base).is_some() {return self.attach_operator_parts(base,depth);}
        let start = base.origin.as_ref().map_or(self.position(), |o| o.span.byte_start as usize);
        let mut slots: Vec<ScriptSlot> = Vec::new();
        loop {
            let saved = self.at;
            let before_marker = self.trivia();
            let script = match self.kind() {
                Some(TokenKind::Superscript) => ScriptKind::Superscript,
                Some(TokenKind::Subscript) => ScriptKind::Subscript,
                _ => { self.at = saved; break; }
            };
            let marker = self.position();
            self.at += 1;
            if slots.iter().any(|slot| slot.script == script) {
                self.diagnostic(DiagnosticCode::DuplicateScript, Severity::Error, marker, self.position(), "A base cannot have the same script twice")?;
            }
            let (before_argument, argument) = self.argument(depth + 1)?;
            slots.push(ScriptSlot { script, before_marker, before_argument, argument });
        }
        if slots.is_empty() { Ok(base) }
        else {
            // The base is an editable container too. A symbol's TeX spelling
            // remains unbraced; this row adds positions, not source characters.
            let base = if base.is_row() { base } else {
                let end = base.origin.as_ref().map_or(start, |o| o.span.byte_end as usize);
                self.make(NodeKind::Row { grouping: Grouping::Bare,
                    entries: vec![RowEntry { leading: String::new(), node: base }], trailing: String::new() }, start, end, None)?
            };
            self.make(NodeKind::Scripts { base, slots }, start, self.position(), None)
        }
    }

    fn attach_operator_parts(&mut self, base:Arc<Node>, depth:usize)->EngineResult<Arc<Node>> {
        use crate::operators::{self,OperatorPart,LimitMode};
        let start=base.origin.as_ref().map_or(self.position(),|o|o.span.byte_start as usize);
        let mut parts=Vec::new();
        loop {
            let saved=self.at;let before=self.trivia();let marker=self.position();
            if matches!(self.kind(),Some(TokenKind::ControlWord)) && LimitMode::from_command(self.text()).is_some() {
                let command=self.text().to_owned();self.at+=1;
                parts.push(OperatorPart::Limits{before,command});
            } else if matches!(self.kind(),Some(TokenKind::Superscript|TokenKind::Subscript)) {
                let script=if self.kind()==Some(TokenKind::Superscript){ScriptKind::Superscript}else{ScriptKind::Subscript};
                self.at+=1;
                if operators::slots(&parts).any(|slot|slot.script==script) {
                    self.diagnostic(DiagnosticCode::DuplicateScript,Severity::Error,marker,self.position(),"An operator cannot have the same bound twice")?;
                }
                let (before_argument,argument)=self.argument(depth+1)?;
                parts.push(OperatorPart::Script{slot:ScriptSlot{script,before_marker:before,before_argument,argument}});
            } else {self.at=saved;break;}
            if parts.len()>operators::MAX_OPERATOR_PARTS {return Err(EngineError::new(ErrorCode::ResourceLimit,"Operator tail budget exhausted"));}
        }
        if parts.is_empty(){Ok(base)}else{
            let kind=match &base.kind {
                NodeKind::LargeOperator{command,..}=>NodeKind::LargeOperator{command:command.clone(),parts},
                NodeKind::CustomSymbol{command,asset_sha256,asset,..}=>NodeKind::CustomSymbol{command:command.clone(),asset_sha256:asset_sha256.clone(),asset:Arc::clone(asset),parts},
                NodeKind::GeneralOperator{command,before_body,body,..}=>NodeKind::GeneralOperator{
                    command:command.clone(),before_body:before_body.clone(),body:Arc::clone(body),parts},
                _=>return Err(EngineError::new(ErrorCode::InvalidInput,"Operator tail attached to a non-operator")),
            };
            self.make(kind,start,self.position(),Some(base.id))
        }
    }

    fn text_body(&mut self) -> EngineResult<(usize, usize)> {
        let start = self.position();
        if self.kind() != Some(TokenKind::OpenBrace) {
            if matches!(self.kind(), None | Some(TokenKind::CloseBrace | TokenKind::Superscript | TokenKind::Subscript)) {
                self.diagnostic(DiagnosticCode::MissingArgument, Severity::Error, start, start, "Text argument is missing")?;
                return Ok((start, start));
            }
            self.at += 1;
            return Ok((start, self.position()));
        }
        self.at += 1;
        let body_start = self.position();
        let mut nesting = 1;
        while let Some(kind) = self.kind() {
            match kind {
                TokenKind::OpenBrace => {
                    nesting += 1;
                    if nesting > MAX_DEPTH { return Err(EngineError::new(ErrorCode::ResourceLimit, "Text group nesting budget exhausted")); }
                }
                TokenKind::CloseBrace => {
                    nesting -= 1;
                    if nesting == 0 {
                        let end = self.position();
                        self.at += 1;
                        return Ok((body_start, end));
                    }
                }
                _ => {}
            }
            self.at += 1;
        }
        self.diagnostic(DiagnosticCode::UnclosedGroup, Severity::Error, start, self.position(), "Text group is missing its closing brace")?;
        Ok((body_start, self.position()))
    }

    fn decode_text(&mut self, start: usize, end: usize) -> EngineResult<String> {
        let body = self.source.text()[start..end].to_owned();
        let mut text = String::new();
        let mut at = 0;
        while at < body.len() {
            let begin = at;
            let ch = body[at..].chars().next().unwrap();
            at += ch.len_utf8();
            match ch {
                '{' | '}' => {}
                '%' => {
                    while at < body.len() && body.as_bytes()[at] != b'\n' { at += 1; }
                    if at < body.len() { at += 1; }
                }
                '\\' => {
                    let control_start = at;
                    while at < body.len() && body.as_bytes()[at].is_ascii_alphabetic() { at += 1; }
                    if at > control_start {
                        let command = &body[control_start..at];
                        let known = match command {
                            "textbackslash" => Some('\\'), "textasciicircum" => Some('^'), "textasciitilde" => Some('~'), _ => None,
                        };
                        if let Some(ch) = known {
                            text.push(ch);
                            if body[at..].starts_with("{}") { at += 2; }
                        } else {
                            self.diagnostic(DiagnosticCode::UnsupportedTextCommand, Severity::Unsupported, start + begin, start + at, "Text command is not supported")?;
                            text.push_str(&body[begin..at]);
                        }
                    } else if at < body.len() {
                        let escaped = body[at..].chars().next().unwrap();
                        at += escaped.len_utf8();
                        if matches!(escaped, '{' | '}' | '$' | '&' | '#' | '_' | '%' | ' ') {
                            text.push(escaped);
                        } else {
                            self.diagnostic(DiagnosticCode::UnsupportedTextCommand, Severity::Unsupported, start + begin, start + at, "Text control symbol is not supported")?;
                            text.push_str(&body[begin..at]);
                        }
                    } else {
                        self.diagnostic(DiagnosticCode::UnfinishedCommand, Severity::Error, start + begin, start + at, "Text command is incomplete")?;
                        text.push('\\');
                    }
                }
                '$' | '&' | '#' | '_' | '^' | '~' => {
                    self.diagnostic(DiagnosticCode::UnsupportedToken, Severity::Unsupported, start + begin, start + at, "Text token requires an explicit escape or unsupported feature")?;
                    text.push(ch);
                }
                _ => text.push(ch),
            }
        }
        Ok(text)
    }

    fn capture_opaque_arguments(&mut self) -> EngineResult<()> {
        if self.text() == "*" { self.at += 1; }
        loop {
            let saved = self.at;
            self.trivia();
            if self.kind() == Some(TokenKind::OpenBrace) {
                self.text_body()?;
            } else if self.text() == "[" {
                let start = self.position();
                self.at += 1;
                let mut brackets = 1;
                let mut braces = 0;
                while self.kind().is_some() && brackets > 0 {
                    match self.kind() {
                        Some(TokenKind::OpenBrace) => braces += 1,
                        Some(TokenKind::CloseBrace) if braces > 0 => braces -= 1,
                        _ if braces == 0 && self.text() == "[" => brackets += 1,
                        _ if braces == 0 && self.text() == "]" => brackets -= 1,
                        _ => {}
                    }
                    if braces + brackets > MAX_DEPTH {
                        return Err(EngineError::new(ErrorCode::ResourceLimit, "Opaque argument nesting budget exhausted"));
                    }
                    self.at += 1;
                }
                if brackets != 0 {
                    self.diagnostic(DiagnosticCode::UnclosedGroup, Severity::Error, start, self.position(), "Opaque optional argument is incomplete")?;
                }
            } else {
                self.at = saved;
                break;
            }
        }
        Ok(())
    }
}
