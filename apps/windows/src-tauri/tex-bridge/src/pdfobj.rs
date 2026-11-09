//! PDF objects and content-stream tokens as they appear in dvipdfmx specials (`pdf:put @x << … >>`,
//! `pdf:code q 1 0 0 1 3 4 cm …`). `@name` references stay unresolved until looked up.

#[derive(Clone, Debug, PartialEq)]
pub enum Obj {
    Null,
    Bool(bool),
    Num(f64),
    Str(Vec<u8>),
    Name(String),
    Array(Vec<Obj>),
    Dict(Vec<(String, Obj)>),
    /// `@name`: a dvipdfmx named object
    Ref(String),
    /// `n g R`: an indirect reference (unresolvable here)
    Indirect,
}

impl Obj {
    pub fn num(&self) -> Option<f64> {
        if let Obj::Num(v) = self {
            Some(*v)
        } else {
            None
        }
    }
    pub fn name(&self) -> Option<&str> {
        if let Obj::Name(v) = self {
            Some(v)
        } else {
            None
        }
    }
    pub fn array(&self) -> Option<&Vec<Obj>> {
        if let Obj::Array(v) = self {
            Some(v)
        } else {
            None
        }
    }
    pub fn get(&self, key: &str) -> Option<&Obj> {
        if let Obj::Dict(entries) = self {
            entries.iter().rev().find(|(k, _)| k == key).map(|(_, v)| v)
        } else {
            None
        }
    }
    pub fn nums(&self) -> Vec<f64> {
        self.array().map(|a| a.iter().filter_map(|o| o.num()).collect()).unwrap_or_default()
    }
    /// Merge a dictionary's entries into this one (dvipdfmx `pdf:put` on a dictionary).
    pub fn merge(&mut self, other: &Obj) {
        match (self, other) {
            (Obj::Dict(a), Obj::Dict(b)) => {
                for (k, v) in b {
                    if let Some(slot) = a.iter_mut().find(|(key, _)| key == k) {
                        // nested dictionaries merge too (/ExtGState << … >> put twice)
                        if matches!((&slot.1, v), (Obj::Dict(_), Obj::Dict(_))) {
                            slot.1.merge(v);
                        } else {
                            slot.1 = v.clone();
                        }
                    } else {
                        a.push((k.clone(), v.clone()));
                    }
                }
            }
            (Obj::Array(a), other) => match other {
                Obj::Array(b) => a.extend(b.iter().cloned()),
                v => a.push(v.clone()),
            },
            _ => {}
        }
    }
}

/// A content-stream token: an operand or an operator.
#[derive(Clone, Debug, PartialEq)]
pub enum Tok {
    Operand(Obj),
    Op(String),
}

pub struct Lexer<'a> {
    s: &'a [u8],
    pub pos: usize,
}

fn is_white(c: u8) -> bool {
    matches!(c, b' ' | b'\t' | b'\r' | b'\n' | 0 | 12)
}
fn is_delim(c: u8) -> bool {
    matches!(c, b'(' | b')' | b'<' | b'>' | b'[' | b']' | b'{' | b'}' | b'/' | b'%')
}

impl<'a> Lexer<'a> {
    pub fn new(s: &'a [u8]) -> Self {
        Lexer { s, pos: 0 }
    }

    pub fn rest(&self) -> &'a [u8] {
        &self.s[self.pos.min(self.s.len())..]
    }

    pub fn skip_white(&mut self) {
        while self.pos < self.s.len() {
            let c = self.s[self.pos];
            if is_white(c) {
                self.pos += 1;
            } else if c == b'%' {
                while self.pos < self.s.len() && self.s[self.pos] != b'\n' && self.s[self.pos] != b'\r' {
                    self.pos += 1;
                }
            } else {
                break;
            }
        }
    }

    pub fn at_end(&mut self) -> bool {
        self.skip_white();
        self.pos >= self.s.len()
    }

    fn word(&mut self) -> &'a [u8] {
        let start = self.pos;
        while self.pos < self.s.len() && !is_white(self.s[self.pos]) && !is_delim(self.s[self.pos]) {
            self.pos += 1;
        }
        &self.s[start..self.pos]
    }

    /// The next bare word (an identifier such as `width`), without consuming anything else.
    pub fn ident(&mut self) -> Option<String> {
        self.skip_white();
        let save = self.pos;
        let w = self.word();
        if w.is_empty() || !(w[0].is_ascii_alphabetic() || w[0] == b'@' || w[0] == b'_') {
            self.pos = save;
            return None;
        }
        Some(String::from_utf8_lossy(w).to_string())
    }

    pub fn number(&mut self) -> Option<f64> {
        self.skip_white();
        let save = self.pos;
        let start = self.pos;
        while self.pos < self.s.len() && matches!(self.s[self.pos], b'0'..=b'9' | b'.' | b'-' | b'+') {
            self.pos += 1;
        }
        let w = &self.s[start..self.pos];
        match std::str::from_utf8(w).ok().and_then(parse_number) {
            Some(v) => Some(v),
            None => {
                self.pos = save;
                None
            }
        }
    }

    /// A length with an optional unit (dvipdfmx `width 3cm`), in big points; `mag` for `true` units.
    pub fn length(&mut self) -> Option<f64> {
        let v = self.number()?;
        self.skip_white();
        let save = self.pos;
        let w = String::from_utf8_lossy(self.word()).to_string();
        let mut unit = w.as_str();
        if let Some(rest) = unit.strip_prefix("true") {
            unit = rest;
            if unit.is_empty() {
                self.skip_white();
                let save2 = self.pos;
                let w2 = String::from_utf8_lossy(self.word()).to_string();
                return match unit_bp(&w2) {
                    Some(u) => Some(v * u),
                    None => {
                        self.pos = save2;
                        Some(v)
                    }
                };
            }
        }
        match unit_bp(unit) {
            Some(u) => Some(v * u),
            None => {
                self.pos = save;
                Some(v)
            }
        }
    }

    /// A PDF object (or None at the end / at an operator).
    pub fn object(&mut self) -> Option<Obj> {
        self.skip_white();
        let c = *self.s.get(self.pos)?;
        match c {
            b'/' => {
                self.pos += 1;
                let w = self.word();
                Some(Obj::Name(decode_name(w)))
            }
            b'(' => Some(Obj::Str(self.literal_string())),
            b'<' => {
                if self.s.get(self.pos + 1) == Some(&b'<') {
                    self.pos += 2;
                    let mut entries = vec![];
                    loop {
                        self.skip_white();
                        if self.s.get(self.pos) == Some(&b'>') && self.s.get(self.pos + 1) == Some(&b'>') {
                            self.pos += 2;
                            break;
                        }
                        if self.pos >= self.s.len() {
                            break;
                        }
                        let key = match self.object() {
                            Some(Obj::Name(k)) => k,
                            Some(_) => continue,
                            None => {
                                self.pos += 1;
                                continue;
                            }
                        };
                        let value = self.object().unwrap_or(Obj::Null);
                        entries.push((key, value));
                    }
                    Some(Obj::Dict(entries))
                } else {
                    self.pos += 1;
                    let start = self.pos;
                    while self.pos < self.s.len() && self.s[self.pos] != b'>' {
                        self.pos += 1;
                    }
                    let hex: Vec<u8> = self.s[start..self.pos].iter().copied().filter(|c| c.is_ascii_hexdigit()).collect();
                    self.pos += 1;
                    let mut bytes = vec![];
                    for pair in hex.chunks(2) {
                        let s = std::str::from_utf8(pair).unwrap_or("0");
                        let v = u8::from_str_radix(if s.len() == 1 { "0" } else { s }, 16).unwrap_or(0);
                        bytes.push(if s.len() == 1 { u8::from_str_radix(s, 16).unwrap_or(0) << 4 } else { v });
                    }
                    Some(Obj::Str(bytes))
                }
            }
            b'[' => {
                self.pos += 1;
                let mut items = vec![];
                loop {
                    self.skip_white();
                    match self.s.get(self.pos) {
                        Some(b']') => {
                            self.pos += 1;
                            break;
                        }
                        None => break,
                        _ => {}
                    }
                    match self.object() {
                        Some(o) => items.push(o),
                        None => {
                            // an operator-like word inside an array: skip it
                            let w = self.word();
                            if w.is_empty() {
                                self.pos += 1;
                            }
                        }
                    }
                }
                Some(Obj::Array(items))
            }
            b'{' | b'}' | b')' | b']' | b'>' => None,
            b'@' => {
                self.pos += 1;
                let w = self.word();
                Some(Obj::Ref(String::from_utf8_lossy(w).to_string()))
            }
            _ => {
                let save = self.pos;
                let w = self.word();
                if w.is_empty() {
                    return None;
                }
                if let Some(v) = std::str::from_utf8(w).ok().and_then(parse_number) {
                    // `n g R`
                    let after = self.pos;
                    let mut look = Lexer { s: self.s, pos: after };
                    if let (Some(_), Some(r)) = (look.number(), {
                        look.skip_white();
                        Some(look.word())
                    }) {
                        if r == b"R" {
                            self.pos = look.pos;
                            return Some(Obj::Indirect);
                        }
                    }
                    return Some(Obj::Num(v));
                }
                match w {
                    b"true" => Some(Obj::Bool(true)),
                    b"false" => Some(Obj::Bool(false)),
                    b"null" => Some(Obj::Null),
                    _ => {
                        self.pos = save;
                        None
                    }
                }
            }
        }
    }

    fn literal_string(&mut self) -> Vec<u8> {
        self.pos += 1;
        let mut depth = 1;
        let mut out = vec![];
        while self.pos < self.s.len() {
            let c = self.s[self.pos];
            self.pos += 1;
            match c {
                b'\\' => {
                    let Some(&n) = self.s.get(self.pos) else { break };
                    self.pos += 1;
                    match n {
                        b'n' => out.push(b'\n'),
                        b'r' => out.push(b'\r'),
                        b't' => out.push(b'\t'),
                        b'b' => out.push(8),
                        b'f' => out.push(12),
                        b'0'..=b'7' => {
                            let mut v = (n - b'0') as u32;
                            for _ in 0..2 {
                                match self.s.get(self.pos) {
                                    Some(&d @ b'0'..=b'7') => {
                                        v = v * 8 + (d - b'0') as u32;
                                        self.pos += 1;
                                    }
                                    _ => break,
                                }
                            }
                            out.push(v as u8);
                        }
                        b'\n' => {}
                        b'\r' => {
                            if self.s.get(self.pos) == Some(&b'\n') {
                                self.pos += 1;
                            }
                        }
                        other => out.push(other),
                    }
                }
                b'(' => {
                    depth += 1;
                    out.push(c);
                }
                b')' => {
                    depth -= 1;
                    if depth == 0 {
                        break;
                    }
                    out.push(c);
                }
                _ => out.push(c),
            }
        }
        out
    }

    /// The next content-stream token.
    pub fn token(&mut self) -> Option<Tok> {
        self.skip_white();
        if self.pos >= self.s.len() {
            return None;
        }
        if let Some(o) = self.object() {
            return Some(Tok::Operand(o));
        }
        let c = self.s[self.pos];
        if is_delim(c) {
            self.pos += 1;
            return Some(Tok::Op(String::from(c as char)));
        }
        let w = self.word();
        Some(Tok::Op(String::from_utf8_lossy(w).to_string()))
    }
}

fn decode_name(w: &[u8]) -> String {
    let mut out = vec![];
    let mut i = 0;
    while i < w.len() {
        if w[i] == b'#' && i + 2 < w.len() + 0 && i + 2 <= w.len() - 1 {
            if let Ok(v) = u8::from_str_radix(std::str::from_utf8(&w[i + 1..i + 3]).unwrap_or("zz"), 16) {
                out.push(v);
                i += 3;
                continue;
            }
        }
        out.push(w[i]);
        i += 1;
    }
    String::from_utf8_lossy(&out).to_string()
}

pub fn parse_number(s: &str) -> Option<f64> {
    if s.is_empty() || !s.bytes().all(|c| c.is_ascii_digit() || c == b'.' || c == b'-' || c == b'+') {
        return None;
    }
    if !s.bytes().any(|c| c.is_ascii_digit()) {
        return None;
    }
    s.parse::<f64>().ok().or_else(|| {
        // "--1" or "1.2.3": what a PDF reader would make of it
        let cleaned: String = s.trim_start_matches(['+', '-']).to_string();
        cleaned.parse::<f64>().ok().map(|v| if s.starts_with('-') { -v } else { v })
    })
}

/// A unit's size in big points (TeX units).
pub fn unit_bp(unit: &str) -> Option<f64> {
    Some(match unit {
        "pt" => 72.0 / 72.27,
        "in" => 72.0,
        "cm" => 72.0 / 2.54,
        "mm" => 72.0 / 25.4,
        "bp" => 1.0,
        "pc" => 12.0 * 72.0 / 72.27,
        "dd" => 1238.0 / 1157.0 * 72.0 / 72.27,
        "cc" => 12.0 * 1238.0 / 1157.0 * 72.0 / 72.27,
        "sp" => 72.0 / (72.27 * 65536.0),
        _ => return None,
    })
}

/// All tokens of a content stream.
pub fn tokens(s: &[u8]) -> Vec<Tok> {
    let mut lx = Lexer::new(s);
    let mut out = vec![];
    while let Some(t) = lx.token() {
        out.push(t);
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn content_tokens() {
        let t = tokens(b"q 1 0 0 1 3.5 -4 cm /pgf@ca0.5 gs [2 1] 0 d 0.0 0.0 m 28.3 0 l S Q");
        assert_eq!(t[0], Tok::Op("q".into()));
        assert_eq!(t[7], Tok::Op("cm".into()));
        assert_eq!(t[8], Tok::Operand(Obj::Name("pgf@ca0.5".into())));
        assert_eq!(t[10], Tok::Operand(Obj::Array(vec![Obj::Num(2.0), Obj::Num(1.0)])));
        assert_eq!(t.last(), Some(&Tok::Op("Q".into())));
    }

    #[test]
    fn dictionaries_and_refs() {
        let mut lx = Lexer::new(b"@pgfextgs << /pgf@CA0.8 << /CA 0.8 >> /X [1 2 (a\\)b)] /R 12 0 R /G @mask >>");
        assert_eq!(lx.object(), Some(Obj::Ref("pgfextgs".into())));
        let d = lx.object().unwrap();
        assert_eq!(d.get("pgf@CA0.8").and_then(|o| o.get("CA")).and_then(|o| o.num()), Some(0.8));
        assert_eq!(d.get("X").unwrap().array().unwrap()[2], Obj::Str(b"a)b".to_vec()));
        assert_eq!(d.get("R"), Some(&Obj::Indirect));
        assert_eq!(d.get("G"), Some(&Obj::Ref("mask".into())));
        let mut a = Obj::Dict(vec![("A".into(), Obj::Dict(vec![("x".into(), Obj::Num(1.0))]))]);
        a.merge(&Obj::Dict(vec![("A".into(), Obj::Dict(vec![("y".into(), Obj::Num(2.0))]))]));
        assert_eq!(a.get("A").unwrap().get("y").unwrap().num(), Some(2.0));
        assert_eq!(a.get("A").unwrap().get("x").unwrap().num(), Some(1.0));
    }

    #[test]
    fn lengths() {
        let mut lx = Lexer::new(b"width 72.27pt height 1in depth 0 bbox");
        assert_eq!(lx.ident().as_deref(), Some("width"));
        assert!((lx.length().unwrap() - 72.0).abs() < 1e-9);
        assert_eq!(lx.ident().as_deref(), Some("height"));
        assert!((lx.length().unwrap() - 72.0).abs() < 1e-9);
        assert_eq!(lx.ident().as_deref(), Some("depth"));
        assert_eq!(lx.length(), Some(0.0));
        assert_eq!(lx.ident().as_deref(), Some("bbox"));
    }
}
