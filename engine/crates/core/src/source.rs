use std::sync::Arc;
use serde::Serialize;
use unicode_segmentation::UnicodeSegmentation;
use crate::error::{EngineError, EngineResult, ErrorCode};

pub const MAX_SOURCE_BYTES: usize = 262_144;
pub const MAX_TOKENS: usize = 32_768;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct SourceSpan {
    pub source_id: u32,
    pub byte_start: u32,
    pub byte_end: u32,
    pub utf16_start: u32,
    pub utf16_end: u32,
}

/// A single shared source buffer with explicit scalar and grapheme boundaries.
/// Offsets in the math tree are never inferred from string length.
#[derive(Debug)]
pub struct SourceText {
    id: u32,
    text: Arc<str>,
    scalar_boundaries: Vec<(u32, u32)>,
    grapheme_boundaries: Vec<u32>,
}

impl SourceText {
    pub fn new(id: u32, text: impl Into<Arc<str>>) -> EngineResult<Arc<Self>> {
        let text = text.into();
        if id == 0 || text.len() > MAX_SOURCE_BYTES {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Invalid source identity or source byte budget exceeded"));
        }
        let mut utf16 = 0u32;
        let mut scalar_boundaries = Vec::new();
        for (byte, ch) in text.char_indices() {
            scalar_boundaries.push((byte as u32, utf16));
            utf16 += ch.len_utf16() as u32;
        }
        scalar_boundaries.push((text.len() as u32, utf16));
        let mut grapheme_boundaries: Vec<u32> = text.grapheme_indices(true)
            .map(|(byte, _)| byte as u32).collect();
        grapheme_boundaries.push(text.len() as u32);
        Ok(Arc::new(Self { id, text, scalar_boundaries, grapheme_boundaries }))
    }

    pub fn id(&self) -> u32 { self.id }
    pub fn text(&self) -> &str { &self.text }
    pub fn len(&self) -> usize { self.text.len() }
    pub fn is_empty(&self) -> bool { self.text.is_empty() }

    pub fn byte_to_utf16(&self, byte: usize) -> EngineResult<u32> {
        let byte = u32::try_from(byte).map_err(|_| invalid_boundary())?;
        self.scalar_boundaries.binary_search_by_key(&byte, |p| p.0)
            .map(|index| self.scalar_boundaries[index].1).map_err(|_| invalid_boundary())
    }

    pub fn utf16_to_byte(&self, utf16: u32) -> EngineResult<usize> {
        self.scalar_boundaries.binary_search_by_key(&utf16, |p| p.1)
            .map(|index| self.scalar_boundaries[index].0 as usize).map_err(|_| invalid_boundary())
    }

    pub fn grapheme_to_byte(&self, grapheme: usize) -> EngineResult<usize> {
        self.grapheme_boundaries.get(grapheme).map(|b| *b as usize).ok_or_else(invalid_boundary)
    }

    pub fn byte_to_grapheme(&self, byte: usize) -> EngineResult<usize> {
        let byte = u32::try_from(byte).map_err(|_| invalid_boundary())?;
        self.grapheme_boundaries.binary_search(&byte).map_err(|_| invalid_boundary())
    }

    pub fn span(&self, start: usize, end: usize) -> EngineResult<SourceSpan> {
        if start > end { return Err(invalid_boundary()); }
        Ok(SourceSpan { source_id: self.id, byte_start: start as u32, byte_end: end as u32,
            utf16_start: self.byte_to_utf16(start)?, utf16_end: self.byte_to_utf16(end)? })
    }

    pub fn slice(self: &Arc<Self>, start: usize, end: usize) -> EngineResult<SourceSlice> {
        let span = self.span(start, end)?;
        Ok(SourceSlice { source: Arc::clone(self), span })
    }
}

fn invalid_boundary() -> EngineError {
    EngineError::new(ErrorCode::InvalidPosition, "Position is outside the source or splits a Unicode boundary")
}

#[derive(Debug, Clone, Serialize)]
pub struct SourceSlice {
    #[serde(skip)]
    source: Arc<SourceText>,
    #[serde(flatten)]
    pub span: SourceSpan,
}

impl SourceSlice {
    pub fn text(&self) -> &str {
        &self.source.text()[self.span.byte_start as usize..self.span.byte_end as usize]
    }
    pub fn source(&self) -> &Arc<SourceText> { &self.source }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum TokenKind { ControlWord, ControlSymbol, Character, OpenBrace, CloseBrace, Superscript, Subscript, Space, Comment }

#[derive(Debug, Clone, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Token {
    pub kind: TokenKind,
    pub span: SourceSpan,
}

impl Token {
    pub fn text<'a>(&self, source: &'a SourceText) -> &'a str {
        &source.text()[self.span.byte_start as usize..self.span.byte_end as usize]
    }
    pub fn is_trivia(&self) -> bool { matches!(self.kind, TokenKind::Space | TokenKind::Comment) }
}

pub fn tokenize(source: &SourceText) -> EngineResult<Vec<Token>> {
    let mut tokens = Vec::new();
    let mut byte = 0;
    let input = source.text();
    while byte < input.len() {
        if tokens.len() >= MAX_TOKENS {
            return Err(EngineError::new(ErrorCode::ResourceLimit, "Token budget exhausted"));
        }
        let start = byte;
        let ch = input[byte..].chars().next().expect("byte is before end");
        let kind = match ch {
            '\\' => {
                byte += 1;
                let word_start = byte;
                while byte < input.len() && input.as_bytes()[byte].is_ascii_alphabetic() { byte += 1; }
                if byte > word_start { TokenKind::ControlWord }
                else {
                    if byte < input.len() { byte += input[byte..].chars().next().unwrap().len_utf8(); }
                    TokenKind::ControlSymbol
                }
            }
            '%' => {
                byte += 1;
                while byte < input.len() && input.as_bytes()[byte] != b'\n' { byte += 1; }
                if byte < input.len() { byte += 1; }
                TokenKind::Comment
            }
            '{' => { byte += 1; TokenKind::OpenBrace }
            '}' => { byte += 1; TokenKind::CloseBrace }
            '^' => { byte += 1; TokenKind::Superscript }
            '_' => { byte += 1; TokenKind::Subscript }
            c if c.is_whitespace() => {
                byte += c.len_utf8();
                while byte < input.len() {
                    let c = input[byte..].chars().next().unwrap();
                    if !c.is_whitespace() { break; }
                    byte += c.len_utf8();
                }
                TokenKind::Space
            }
            _ => {
                let grapheme = input[byte..].graphemes(true).next().unwrap();
                byte += grapheme.len();
                TokenKind::Character
            }
        };
        tokens.push(Token { kind, span: source.span(start, byte)? });
    }
    Ok(tokens)
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn distinguishes_utf8_utf16_and_grapheme_boundaries() {
        let source = SourceText::new(1, "A😀e\u{301}中").unwrap();
        assert_eq!(source.byte_to_utf16(5).unwrap(), 3);
        assert_eq!(source.utf16_to_byte(3).unwrap(), 5);
        assert!(source.utf16_to_byte(2).is_err());
        assert!(source.byte_to_utf16(3).is_err());
        assert_eq!(source.grapheme_to_byte(3).unwrap(), 8);
        assert_eq!(source.byte_to_grapheme(8).unwrap(), 3);
        assert!(source.byte_to_grapheme(6).is_err());
        assert_eq!(source.grapheme_to_byte(4).unwrap(), source.len());
    }
    #[test]
    fn lexer_covers_every_source_byte_including_comments() {
        let source = SourceText::new(9, " \\frac %注释\r\n{x_1}{\\text{速度}} + \\% e\u{301}").unwrap();
        let tokens = tokenize(&source).unwrap();
        let joined = tokens.iter().map(|t| t.text(&source)).collect::<String>();
        assert_eq!(joined, source.text());
        for adjacent in tokens.windows(2) { assert_eq!(adjacent[0].span.byte_end, adjacent[1].span.byte_start); }
        assert!(tokens.iter().any(|t| t.kind == TokenKind::Comment));
        assert!(tokens.iter().any(|t| t.kind == TokenKind::ControlSymbol && t.text(&source) == "\\%"));
        assert!(tokens.iter().any(|t| t.text(&source) == "e\u{301}"));
    }
    #[test]
    fn empty_source_has_valid_zero_boundaries() {
        let source = SourceText::new(1, "").unwrap();
        assert_eq!(source.byte_to_utf16(0).unwrap(), 0);
        assert_eq!(source.grapheme_to_byte(0).unwrap(), 0);
        assert!(tokenize(&source).unwrap().is_empty());
        assert!(source.span(1, 0).is_err());
    }
}
