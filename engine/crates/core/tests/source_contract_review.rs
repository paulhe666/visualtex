//! Independent P03 source/token contract regressions against the public core.
//! No old formula engine or copied implementation participates in these tests.
use visualtex_engine_core::source::{SourceText, TokenKind, tokenize, MAX_SOURCE_BYTES, MAX_TOKENS};
use visualtex_engine_core::ErrorCode;

fn assert_lossless(text: &str) {
    let source = SourceText::new(17, text).expect("fixture fits source budget");
    let tokens = tokenize(&source).expect("fixture fits token budget");
    let mut byte = 0usize;
    let mut reconstructed = String::new();
    for token in &tokens {
        let span = token.span;
        assert_eq!(span.source_id, 17);
        assert_eq!(span.byte_start as usize, byte, "tokens must partition source without gaps: {text:?}");
        assert!(span.byte_end > span.byte_start);
        assert!(text.is_char_boundary(span.byte_start as usize));
        assert!(text.is_char_boundary(span.byte_end as usize));
        assert_eq!(span.utf16_start as usize, text[..span.byte_start as usize].encode_utf16().count());
        assert_eq!(span.utf16_end as usize, text[..span.byte_end as usize].encode_utf16().count());
        let slice = source.slice(span.byte_start as usize, span.byte_end as usize).unwrap();
        assert_eq!(slice.text(), token.text(&source));
        reconstructed.push_str(token.text(&source));
        byte = span.byte_end as usize;
    }
    assert_eq!(byte, text.len());
    assert_eq!(reconstructed, text);
}

#[test]
fn control_words_symbols_and_escaped_comment_markers_remain_distinct() {
    let source = SourceText::new(1, "\\frac\\%\\{\\}\\_\\^\\\\\\α").unwrap();
    let tokens = tokenize(&source).unwrap();
    assert_eq!(tokens[0].kind, TokenKind::ControlWord);
    assert_eq!(tokens[0].text(&source), "\\frac");
    for token in &tokens[1..] { assert_eq!(token.kind, TokenKind::ControlSymbol); }
    assert_eq!(tokens.last().unwrap().text(&source), "\\α");
    assert_lossless(source.text());
}

#[test]
fn scalar_offsets_round_trip_and_reject_split_utf8_or_surrogates() {
    for text in ["", "abc", "速度", "a😀b", "𝑥_𝟙", "e\u{301}", "👩‍🔬🇨🇳\r\n中"] {
        let source = SourceText::new(3, text).unwrap();
        for byte in 0..=text.len() {
            match source.byte_to_utf16(byte) {
                Ok(unit) => {
                    assert!(text.is_char_boundary(byte));
                    assert_eq!(unit as usize, text[..byte].encode_utf16().count());
                    assert_eq!(source.utf16_to_byte(unit).unwrap(), byte);
                }
                Err(error) => {
                    assert!(!text.is_char_boundary(byte));
                    assert_eq!(error.code, ErrorCode::InvalidPosition);
                }
            }
        }
        for unit in 0..=text.encode_utf16().count() as u32 {
            if let Ok(byte) = source.utf16_to_byte(unit) {
                assert_eq!(source.byte_to_utf16(byte).unwrap(), unit);
            }
        }
        assert!(source.byte_to_utf16(text.len() + 1).is_err());
        assert!(source.utf16_to_byte(text.encode_utf16().count() as u32 + 1).is_err());
        assert!(source.byte_to_utf16(usize::MAX).is_err());
    }
    let source = SourceText::new(1, "a😀b").unwrap();
    assert!(source.utf16_to_byte(2).is_err(), "middle of UTF-16 surrogate pair");
}

#[test]
fn grapheme_boundaries_are_not_code_point_boundaries() {
    // Six user-perceived characters: combining acute, astronaut sequence,
    // regional-indicator flag pair, Hangul syllable, CRLF and Chinese character.
    let parts = ["e\u{301}", "👩🏽‍🚀", "🇨🇳", "각", "\r\n", "中"];
    let text = parts.concat();
    let source = SourceText::new(1, text.clone()).unwrap();
    let mut boundary = 0;
    for (grapheme, part) in parts.iter().enumerate() {
        assert_eq!(source.grapheme_to_byte(grapheme).unwrap(), boundary);
        assert_eq!(source.byte_to_grapheme(boundary).unwrap(), grapheme);
        for byte in boundary + 1..boundary + part.len() {
            assert!(source.byte_to_grapheme(byte).is_err());
        }
        boundary += part.len();
    }
    assert_eq!(source.grapheme_to_byte(parts.len()).unwrap(), text.len());
    assert!(source.grapheme_to_byte(parts.len() + 1).is_err());
}

#[test]
fn source_slices_retain_original_buffer_identity_after_owner_drop() {
    let source = SourceText::new(71, "α+\\frac{速度}{😀}").unwrap();
    let slice = source.slice(2, source.len()).unwrap();
    drop(source);
    assert_eq!(slice.span.source_id, 71);
    assert_eq!(slice.text(), "+\\frac{速度}{😀}");
    assert_eq!(slice.source().id(), 71);
}

#[test]
fn ranges_reject_reversal_out_of_bounds_and_split_multibyte_text() {
    let source = SourceText::new(1, "中😀a").unwrap();
    for (start, end) in [(4, 3), (0, 20), (1, 3), (3, 4), (usize::MAX, usize::MAX)] {
        assert_eq!(source.span(start, end).unwrap_err().code, ErrorCode::InvalidPosition);
    }
    assert_eq!(source.slice(3, 7).unwrap().text(), "😀");
    assert_eq!(source.slice(7, 7).unwrap().text(), "");
}

#[test]
fn incomplete_unknown_and_comment_input_is_never_discarded_by_lexer() {
    for text in [
        "\\", "\\fr", "\\frac{a}", "\\frac{a}{", "a_{", "{a", "a}",
        "\\futureCommand{a}[b]", "\\input{secret}", "\\write18{command}",
        "%全部注释", "a% } \\frac\r\nb", "a% } \rb", "\\%not-comment", "\0\u{202e}中",
        "  \\dfrac %保留注释\r\n {x_1}{y^2}  ", "\\text{速度 e\u{301} 👩‍🔬}",
    ] { assert_lossless(text); }
}

#[test]
fn source_limit_is_measured_in_bytes_not_utf16_units() {
    assert!(SourceText::new(1, "a".repeat(MAX_SOURCE_BYTES)).is_ok());
    assert_eq!(SourceText::new(1, "a".repeat(MAX_SOURCE_BYTES + 1)).unwrap_err().code, ErrorCode::ResourceLimit);
    let legal = "中".repeat(MAX_SOURCE_BYTES / 3);
    assert!(SourceText::new(1, legal.clone()).is_ok());
    assert_eq!(SourceText::new(1, format!("{legal}中")).unwrap_err().code, ErrorCode::ResourceLimit);
}

#[test]
fn token_budget_accepts_boundary_then_returns_error_without_repairing_source() {
    let exact = SourceText::new(1, "x".repeat(MAX_TOKENS)).unwrap();
    assert_eq!(tokenize(&exact).unwrap().len(), MAX_TOKENS);
    let over = SourceText::new(2, "x".repeat(MAX_TOKENS + 1)).unwrap();
    assert_eq!(tokenize(&over).unwrap_err().code, ErrorCode::ResourceLimit);
    assert_eq!(over.text().len(), MAX_TOKENS + 1);
    let one = SourceText::new(3, format!("\\{}", "a".repeat(MAX_TOKENS + 1))).unwrap();
    assert_eq!(tokenize(&one).unwrap().len(), 1);
}

#[test]
fn deterministic_mixed_unicode_fuzz_partitions_every_byte() {
    const SEED: u64 = 0x56495355414c5445;
    let atoms = ["x", "1", "中", "😀", "👩‍🔬", "e\u{301}", "\r\n", "\t", " ", "{", "}", "\\", "_", "^", "%", "\\frac", "\\%", "\0", "\u{202e}"];
    let mut state = SEED;
    for case in 0..256 {
        let mut text = String::new();
        for _ in 0..(case % 73 + 1) {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            text.push_str(atoms[(state as usize) % atoms.len()]);
        }
        assert_lossless(&text);
    }
}
