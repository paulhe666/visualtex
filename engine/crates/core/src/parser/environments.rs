//! Token-scoped environment parsing. Physical newlines remain source trivia;
//! only top-level authored row separators split contained mathematical rows.
use super::*;
use crate::grids::{ArrayColumnSpec,GridColumnId, GridKind, MAX_GRID_ROWS, MAX_GRID_COLUMNS};

struct Header { name: String, before: String, after: usize }
impl Parser<'_> {
    fn array_column_spec(&mut self)->EngineResult<Option<ArrayColumnSpec>> {
        let saved=self.at;let before_spec=self.trivia();
        if self.kind()!=Some(TokenKind::OpenBrace){self.at=saved;return Ok(None);}
        self.at+=1;let start=self.position();let mut at=self.at;let mut nested=0usize;
        while let Some(token)=self.tokens.get(at){match token.kind{
            TokenKind::OpenBrace=>nested+=1,
            TokenKind::CloseBrace if nested==0=>{let end=token.span.byte_start as usize;let source=self.source.text()[start..end].to_owned();self.at=at+1;return ArrayColumnSpec::parse(before_spec,source).map(Some);},
            TokenKind::CloseBrace=>nested-=1,
            _=>{},
        }if nested>MAX_DEPTH{return Err(EngineError::new(ErrorCode::ResourceLimit,"Array column specification nesting budget exhausted"));}at+=1;}
        self.at=saved;Ok(None)
    }

    fn environment_header(&self, after_command: usize) -> EngineResult<Option<Header>> {
        let mut at=after_command;
        while self.tokens.get(at).is_some_and(Token::is_trivia){at+=1;}
        if self.tokens.get(at).is_none_or(|t|t.kind!=TokenKind::OpenBrace){return Ok(None);}
        let opening=at;at+=1;let first=at;
        while let Some(token)=self.tokens.get(at){
            if token.kind==TokenKind::CloseBrace {
                let name=self.tokens[first..at].iter().map(|t|t.text(&self.source)).collect::<String>();
                let start=self.tokens.get(after_command).map_or(self.end_position,|t|t.span.byte_start as usize);
                let end=self.tokens[opening].span.byte_start as usize;
                return Ok(Some(Header{name,before:self.source.text()[start..end].to_owned(),after:at+1}));
            }
            if token.kind==TokenKind::OpenBrace || at-first>128 {return Ok(None);}
            at+=1;
        }
        Ok(None)
    }

    pub(super) fn environment(&mut self,start:usize,depth:usize)->EngineResult<Arc<Node>> {
        self.depth(depth)?;
        let Some(header)=self.environment_header(self.at)? else {
            // Consume only the header's bounded lexical prefix, then preserve
            // the enclosing parser's remaining source/brace recovery behavior.
            self.diagnostic(DiagnosticCode::MissingEnvironmentName,Severity::Error,start,self.position(),"Environment name requires a closed brace group")?;
            return self.make(NodeKind::UnparsedSource{raw:self.source.text()[start..self.position()].to_owned()},start,self.position(),None);
        };
        self.at=header.after;
        let mut array_spec=None;let mut unsupported_array_spec=false;
        if header.name=="array" {
            match self.array_column_spec(){
                Ok(Some(spec))=>array_spec=Some(spec),
                Ok(None)=>unsupported_array_spec=true,
                Err(error) if error.code==ErrorCode::ResourceLimit=>return Err(error),
                Err(_)=>unsupported_array_spec=true,
            }
        }
        let first=self.at;
        let mut at=first;let mut braces=0usize;let mut stack:Vec<String>=Vec::new();
        let mut breaks=Vec::new();let mut cell_breaks=Vec::new();let mut closing=None;let mut unsupported_modifier=unsupported_array_spec;
        while let Some(token)=self.tokens.get(at){
            let raw=token.text(&self.source);
            match token.kind {
                TokenKind::OpenBrace=>braces+=1,
                TokenKind::CloseBrace if braces==0=>break,
                TokenKind::CloseBrace=>braces-=1,
                TokenKind::ControlWord if braces==0 && (raw=="\\begin"||raw=="\\end")=>{
                    if let Some(nested)=self.environment_header(at+1)? {
                        if raw=="\\begin" {stack.push(nested.name);at=nested.after;continue;}
                        if let Some(expected)=stack.pop(){
                            if expected!=nested.name {
                                self.diagnostic(DiagnosticCode::MismatchedEnvironment,Severity::Error,
                                    token.span.byte_start as usize,self.tokens[nested.after-1].span.byte_end as usize,
                                    "Nested environment end does not match its begin")?;
                            }
                            at=nested.after;continue;
                        }
                        if nested.name!=header.name {
                            self.diagnostic(DiagnosticCode::MismatchedEnvironment,Severity::Error,
                                token.span.byte_start as usize,self.tokens[nested.after-1].span.byte_end as usize,
                                "Environment end does not match its begin")?;
                        }
                        closing=Some(nested);break;
                    }
                }
                TokenKind::ControlSymbol if braces==0 && stack.is_empty() && raw=="\\\\"=>{
                    breaks.push(at);
                    let mut next=at+1;while self.tokens.get(next).is_some_and(Token::is_trivia){next+=1;}
                    if self.tokens.get(next).is_some_and(|t|matches!(t.text(&self.source),"["|"*")){unsupported_modifier=true;}
                }
                TokenKind::Character if braces==0 && stack.is_empty() && raw=="&"=>cell_breaks.push(at),
                _=>{},
            }
            if braces+stack.len()+depth>MAX_DEPTH {
                return Err(EngineError::new(ErrorCode::ResourceLimit,"Environment nesting budget exhausted"));
            }
            at+=1;
        }
        let body_end=at;
        let end_position=self.tokens.get(body_end).map_or(self.end_position,|t|t.span.byte_start as usize);
        let before_end_name=closing.as_ref().map_or(String::new(),|h|h.before.clone());
        let after=closing.as_ref().map_or(body_end,|h|h.after);
        if closing.is_none(){
            self.diagnostic(DiagnosticCode::UnclosedEnvironment,Severity::Error,start,end_position,"Environment has no matching end")?;
        }
        self.at=after;
        let mut first_nontrivia=first;while self.tokens.get(first_nontrivia).is_some_and(Token::is_trivia){first_nontrivia+=1;}
        // Alignment/row-gap options are not silently treated as formula glyphs.
        let has_option=self.tokens.get(first_nontrivia).is_some_and(|t|t.text(&self.source)=="[");
        let Some(environment)=GridKind::from_name(&header.name) else {
            self.diagnostic(DiagnosticCode::UnsupportedEnvironment,Severity::Unsupported,start,self.position(),"Environment is not supported by the current grid capability")?;
            return self.make(NodeKind::UnparsedSource{raw:self.source.text()[start..self.position()].to_owned()},start,self.position(),None);
        };
        if (has_option && matches!(environment,GridKind::Gathered|GridKind::Aligned))||unsupported_modifier {
            self.diagnostic(DiagnosticCode::UnsupportedEnvironmentOption,Severity::Unsupported,start,self.position(),"Environment placement, row-gap or array column options are outside the supported bounded set")?;
            return self.make(NodeKind::UnparsedSource{raw:self.source.text()[start..self.position()].to_owned()},start,self.position(),None);
        }
        if breaks.len()+1>MAX_GRID_ROWS {return Err(EngineError::new(ErrorCode::ResourceLimit,"Contained grid row budget exhausted"));}
        // Count authored columns before allocating any grid nodes. A ragged
        // matrix remains lossless opaque source until implicit-cell syntax has
        // an explicit projection policy; never invent separators in saved input.
        let mut row_start=first;let mut column_count=None;
        for row_end in breaks.iter().copied().chain(std::iter::once(body_end)) {
            let count=1+cell_breaks.iter().filter(|&&i|i>=row_start&&i<row_end).count();
            if count>MAX_GRID_COLUMNS {return Err(EngineError::new(ErrorCode::ResourceLimit,"Grid column budget exhausted"));}
            if !environment.accepts_columns(count) || column_count.is_some_and(|n|n!=count)
                || (environment==GridKind::Array && array_spec.as_ref().is_none_or(|spec|spec.alignments.len()!=count)) {
                self.diagnostic(DiagnosticCode::UnsupportedEnvironment,Severity::Unsupported,start,self.position(),
                    "This environment requires a rectangular authored cell grid; preserve the source without inventing cells")?;
                return self.make(NodeKind::UnparsedSource{raw:self.source.text()[start..self.position()].to_owned()},start,self.position(),None);
            }
            column_count=Some(count);row_start=row_end+1;
        }
        let columns=(0..column_count.unwrap_or(1)).map(|_|GridColumnId::allocate(self.ids)).collect::<EngineResult<Vec<_>>>()?;
        let mut rows=Vec::new();let mut row_start=first;
        for boundary in breaks.into_iter().map(Some).chain(std::iter::once(None)) {
            let row_end=boundary.unwrap_or(body_end);
            let start_position=self.tokens.get(row_start).map_or(end_position,|t|t.span.byte_start as usize);
            let row_end_position=self.tokens.get(row_end).map_or(end_position,|t|t.span.byte_start as usize);
            let boundaries:Vec<_>=cell_breaks.iter().copied().filter(|&i|i>=row_start&&i<row_end).collect();
            let mut cells=Vec::new();let mut cell_start=row_start;
            for cell_end_index in boundaries.iter().copied().chain(std::iter::once(row_end)) {
                let cell_end=self.tokens.get(cell_end_index).map_or(end_position,|t|t.span.byte_start as usize);
                let mut nested=Parser{source:Arc::clone(&self.source),tokens:self.tokens[cell_start..cell_end_index].to_vec(),at:0,
                    ids:&mut *self.ids,diagnostics:Vec::new(),nodes:self.nodes,end_position:cell_end,allow_middle:self.allow_middle,symbols:Arc::clone(&self.symbols)};
                let cell=nested.row(Grouping::Bare,None,depth+2)?;
                self.nodes=nested.nodes;self.diagnostics.extend(nested.diagnostics);
                cells.push(cell);cell_start=cell_end_index+1;
            }
            let row_break=boundary.map_or(String::new(),|i|self.tokens[i].text(&self.source).to_owned());
            let row_source_end=boundary.map_or(row_end_position,|i|self.tokens[i].span.byte_end as usize);
            let separators=vec!["&".to_owned();boundaries.len()];
            let row=self.make(NodeKind::GridRow{cells,separators,row_break},start_position,row_source_end,None)?;
            rows.push(row);row_start=row_end+usize::from(boundary.is_some());
        }
        self.make(NodeKind::Grid{environment,columns,array_spec,before_begin_name:header.before,before_end_name,rows},start,self.position(),None)
    }

    pub(super) fn unexpected_environment_end(&mut self,start:usize)->EngineResult<Arc<Node>> {
        if let Some(header)=self.environment_header(self.at)?{self.at=header.after;}
        self.diagnostic(DiagnosticCode::UnexpectedEnvironmentEnd,Severity::Error,start,self.position(),"Environment end has no matching begin")?;
        self.make(NodeKind::UnparsedSource{raw:self.source.text()[start..self.position()].to_owned()},start,self.position(),None)
    }
}
