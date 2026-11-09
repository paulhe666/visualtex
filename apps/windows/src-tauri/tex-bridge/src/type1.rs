//! Type 1 fonts (.pfb, .pfa) as outlines, for drawing a live paragraph's glyphs (live.rs): the font's own
//! encoding, its charstrings (eexec and charstring encryption) and the Type 1 charstring operators, with flex,
//! hint replacement and accented characters (seac). Hints are dropped: the outlines are drawn at any size.

use std::collections::HashMap;

pub struct Type1 {
    /// glyph names by code (the font's built-in encoding)
    pub encoding: Vec<Option<String>>,
    charstrings: HashMap<String, Vec<u8>>,
    subrs: Vec<Vec<u8>>,
    /// the FontMatrix (font units to text space; 0.001 in most fonts)
    pub matrix: [f64; 6],
}

fn decrypt(data: &[u8], key: u16, skip: usize) -> Vec<u8> {
    let mut r = key;
    let mut out = Vec::with_capacity(data.len());
    for &c in data {
        out.push(c ^ (r >> 8) as u8);
        r = (c as u16).wrapping_add(r).wrapping_mul(52845).wrapping_add(22719);
    }
    out.drain(..skip.min(out.len()));
    out
}

fn hex_bytes(text: &[u8]) -> Vec<u8> {
    let mut out = Vec::with_capacity(text.len() / 2);
    let mut high: Option<u8> = None;
    for &c in text {
        let v = match c {
            b'0'..=b'9' => c - b'0',
            b'a'..=b'f' => c - b'a' + 10,
            b'A'..=b'F' => c - b'A' + 10,
            _ => continue,
        };
        match high.take() {
            Some(h) => out.push(h << 4 | v),
            None => high = Some(v),
        }
    }
    out
}

fn find(hay: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    if needle.is_empty() || from >= hay.len() {
        return None;
    }
    hay[from..].windows(needle.len()).position(|w| w == needle).map(|i| i + from)
}

/// The cleartext part and the decrypted private part of a .pfb or .pfa file.
fn parts(file: &[u8]) -> Option<(Vec<u8>, Vec<u8>)> {
    if file.first() == Some(&0x80) {
        let (mut clear, mut binary) = (vec![], vec![]);
        let mut i = 0;
        while i + 6 <= file.len() && file[i] == 0x80 {
            let kind = file[i + 1];
            if kind == 3 {
                break;
            }
            let len = u32::from_le_bytes([file[i + 2], file[i + 3], file[i + 4], file[i + 5]]) as usize;
            let seg = file.get(i + 6..(i + 6 + len).min(file.len()))?;
            if kind == 1 {
                clear.extend_from_slice(seg);
            } else {
                binary.extend_from_slice(seg);
            }
            i += 6 + len;
        }
        // a binary segment written as hex (rare)
        let encrypted = if binary.len() >= 4 && binary[..4].iter().all(|c| c.is_ascii_hexdigit()) { hex_bytes(&binary) } else { binary };
        return Some((clear, decrypt(&encrypted, 55665, 4)));
    }
    let at = find(file, b"eexec", 0)?;
    let clear = file[..at].to_vec();
    let mut rest = &file[at + 5..];
    while let Some((&c, tail)) = rest.split_first() {
        if c.is_ascii_whitespace() {
            rest = tail;
        } else {
            break;
        }
    }
    let end = find(rest, b"cleartomark", 0).unwrap_or(rest.len());
    let encrypted = if rest.len() >= 4 && rest[..4].iter().all(|c| c.is_ascii_hexdigit()) { hex_bytes(&rest[..end]) } else { rest.to_vec() };
    Some((clear, decrypt(&encrypted, 55665, 4)))
}

struct Tokens<'a> {
    data: &'a [u8],
    at: usize,
}

impl<'a> Tokens<'a> {
    fn skip_space(&mut self) {
        while self.at < self.data.len() && (self.data[self.at].is_ascii_whitespace() || self.data[self.at] == 0) {
            self.at += 1;
        }
    }
    fn token(&mut self) -> Option<&'a [u8]> {
        self.skip_space();
        let start = self.at;
        if start >= self.data.len() {
            return None;
        }
        // a name's slash starts its token
        if self.data[self.at] == b'/' {
            self.at += 1;
        }
        while self.at < self.data.len() && !self.data[self.at].is_ascii_whitespace() && !b"/[]{}()<>".contains(&self.data[self.at]) {
            self.at += 1;
        }
        if self.at == start {
            self.at += 1;
        }
        Some(&self.data[start..self.at])
    }
    fn int(&mut self) -> Option<i64> {
        std::str::from_utf8(self.token()?).ok()?.parse().ok()
    }
    /// `len RD <len bytes>`: the bytes (after the RD token and one space)
    fn binary(&mut self, len: usize) -> Option<&'a [u8]> {
        self.token()?;
        self.at += 1;
        let bytes = self.data.get(self.at..self.at + len)?;
        self.at += len;
        Some(bytes)
    }
}

impl Type1 {
    pub fn parse(file: &[u8]) -> Option<Type1> {
        let (clear, private) = parts(file)?;
        let text = String::from_utf8_lossy(&clear);
        let mut matrix = [0.001, 0.0, 0.0, 0.001, 0.0, 0.0];
        if let Some(i) = text.find("/FontMatrix") {
            let rest = &text[i + 11..];
            if let (Some(a), Some(b)) = (rest.find('['), rest.find(']')) {
                let nums: Vec<f64> = rest[a + 1..b].split_whitespace().filter_map(|n| n.parse().ok()).collect();
                if nums.len() == 6 {
                    matrix.copy_from_slice(&nums);
                }
            }
        }
        let mut encoding: Vec<Option<String>> = vec![None; 256];
        if let Some(i) = text.find("/Encoding") {
            let rest = &text[i + 9..];
            if rest.trim_start().starts_with("StandardEncoding") {
                for (code, name) in STANDARD.iter().enumerate() {
                    if !name.is_empty() {
                        encoding[code + 32] = Some(name.to_string());
                    }
                }
            } else {
                let end = rest.find("readonly def").or_else(|| rest.find(" def")).unwrap_or(rest.len());
                let words: Vec<&str> = rest[..end].split_whitespace().collect();
                for w in words.windows(4) {
                    if w[0] == "dup" && w[3] == "put" && w[2].starts_with('/') {
                        if let Ok(code) = w[1].parse::<usize>() {
                            if code < 256 {
                                encoding[code] = Some(w[2][1..].to_string());
                            }
                        }
                    }
                }
            }
        }
        let len_iv = find(&private, b"/lenIV", 0)
            .and_then(|i| {
                let mut t = Tokens { data: &private, at: i + 6 };
                t.int()
            })
            .unwrap_or(4);
        let plain = |cs: &[u8]| -> Vec<u8> {
            if len_iv < 0 {
                cs.to_vec()
            } else {
                decrypt(cs, 4330, len_iv as usize)
            }
        };
        let mut subrs = vec![];
        let chars_at = find(&private, b"/CharStrings", 0)?;
        if let Some(i) = find(&private, b"/Subrs", 0).filter(|&i| i < chars_at) {
            let mut t = Tokens { data: &private, at: i + 6 };
            let count = t.int().unwrap_or(0).clamp(0, 65535) as usize;
            subrs = vec![vec![]; count];
            let mut seen = 0;
            while seen < count && t.at < chars_at {
                let Some(tok) = t.token() else { break };
                if tok != b"dup" {
                    continue;
                }
                let (Some(index), Some(len)) = (t.int(), t.int()) else { break };
                let Some(bytes) = t.binary(len.max(0) as usize) else { break };
                if (index as usize) < count {
                    subrs[index as usize] = plain(bytes);
                }
                seen += 1;
            }
        }
        let mut charstrings = HashMap::new();
        let mut t = Tokens { data: &private, at: chars_at + 12 };
        let count = t.int().unwrap_or(0).max(0) as usize;
        while charstrings.len() < count {
            let Some(tok) = t.token() else { break };
            if tok == b"end" && charstrings.len() > 0 {
                break;
            }
            if tok.len() < 2 || tok[0] != b'/' {
                continue;
            }
            let name = String::from_utf8_lossy(&tok[1..]).to_string();
            let Some(len) = t.int() else { continue };
            let Some(bytes) = t.binary(len.max(0) as usize) else { break };
            charstrings.insert(name, plain(bytes));
        }
        Some(Type1 { encoding, charstrings, subrs, matrix })
    }

    /// The outline of glyph `name` as SVG path data in font units (y up), with `enc` the glyph names by code
    /// when an encoding file says them (else the font's own).
    pub fn path_by_code(&self, code: u32, enc: Option<&[Option<String>]>) -> Option<String> {
        let names = enc.unwrap_or(&self.encoding);
        let name = names.get(code as usize)?.as_ref()?;
        self.path(name)
    }

    pub fn path(&self, name: &str) -> Option<String> {
        let cs = self.charstrings.get(name)?;
        let mut run = Run::new(self);
        run.exec(cs, 0);
        Some(run.out)
    }
}

/// An encoding file (.enc: `/Name [ /glyph ... ] def`): the 256 glyph names.
pub fn parse_encoding(text: &str) -> Vec<Option<String>> {
    let mut names: Vec<Option<String>> = vec![];
    let body: String = text.lines().map(|l| l.split('%').next().unwrap_or("")).collect::<Vec<_>>().join(" ");
    if let (Some(a), Some(b)) = (body.find('['), body.rfind(']')) {
        for w in body[a + 1..b].split(|c: char| c.is_whitespace() || c == '/').filter(|w| !w.is_empty()) {
            names.push(if w == ".notdef" { None } else { Some(w.to_string()) });
        }
    }
    names.resize(256, None);
    names
}

struct Run<'a> {
    font: &'a Type1,
    out: String,
    x: f64,
    y: f64,
    stack: Vec<f64>,
    ps: Vec<f64>,
    flex: Option<Vec<(f64, f64)>>,
    sbx: f64,
    /// a drawing command needs a moveto first (the start, after closepath)
    open: bool,
    done: bool,
    /// where an accent's origin goes (seac)
    origin: (f64, f64),
}

fn num(v: f64) -> String {
    let r = (v * 10.0).round() / 10.0;
    if r == r.trunc() {
        format!("{}", r as i64)
    } else {
        format!("{r:.1}")
    }
}

impl<'a> Run<'a> {
    fn new(font: &'a Type1) -> Run<'a> {
        Run { font, out: String::new(), x: 0.0, y: 0.0, stack: vec![], ps: vec![], flex: None, sbx: 0.0, open: false, done: false, origin: (0.0, 0.0) }
    }
    fn pt(&self, x: f64, y: f64) -> String {
        format!("{} {}", num(x + self.origin.0), num(y + self.origin.1))
    }
    fn ensure_open(&mut self) {
        if !self.open {
            let p = self.pt(self.x, self.y);
            self.out.push_str(&format!("M{p}"));
            self.open = true;
        }
    }
    fn move_to(&mut self, dx: f64, dy: f64) {
        self.x += dx;
        self.y += dy;
        if let Some(points) = self.flex.as_mut() {
            points.push((self.x, self.y));
            return;
        }
        // a moveto starts a new subpath (an open one is closed: Type 1 subpaths are closed shapes)
        if self.open {
            self.out.push('Z');
        }
        let p = self.pt(self.x, self.y);
        self.out.push_str(&format!("M{p}"));
        self.open = true;
    }
    fn line_to(&mut self, dx: f64, dy: f64) {
        self.ensure_open();
        self.x += dx;
        self.y += dy;
        let p = self.pt(self.x, self.y);
        self.out.push_str(&format!("L{p}"));
    }
    fn curve_to(&mut self, d: [f64; 6]) {
        self.ensure_open();
        let (x1, y1) = (self.x + d[0], self.y + d[1]);
        let (x2, y2) = (x1 + d[2], y1 + d[3]);
        let (x3, y3) = (x2 + d[4], y2 + d[5]);
        let (a, b, c) = (self.pt(x1, y1), self.pt(x2, y2), self.pt(x3, y3));
        self.out.push_str(&format!("C{a} {b} {c}"));
        self.x = x3;
        self.y = y3;
    }
    fn close(&mut self) {
        if self.open {
            self.out.push('Z');
            self.open = false;
        }
    }

    fn exec(&mut self, cs: &[u8], depth: usize) {
        if depth > 12 {
            return;
        }
        let mut i = 0;
        while i < cs.len() && !self.done {
            let v = cs[i];
            i += 1;
            match v {
                32..=246 => self.stack.push(v as f64 - 139.0),
                247..=250 => {
                    let w = *cs.get(i).unwrap_or(&0) as f64;
                    i += 1;
                    self.stack.push((v as f64 - 247.0) * 256.0 + w + 108.0);
                }
                251..=254 => {
                    let w = *cs.get(i).unwrap_or(&0) as f64;
                    i += 1;
                    self.stack.push(-(v as f64 - 251.0) * 256.0 - w - 108.0);
                }
                255 => {
                    let b = cs.get(i..i + 4).unwrap_or(&[0, 0, 0, 0]);
                    i += 4;
                    self.stack.push(i32::from_be_bytes([b[0], b[1], b[2], b[3]]) as f64);
                }
                12 => {
                    let op = *cs.get(i).unwrap_or(&0);
                    i += 1;
                    self.escape(op, depth);
                }
                _ => self.command(v, depth),
            }
        }
    }

    fn args<const N: usize>(&mut self) -> [f64; N] {
        let mut a = [0.0; N];
        let n = self.stack.len();
        if n >= N {
            a.copy_from_slice(&self.stack[n - N..]);
        }
        a
    }

    fn command(&mut self, op: u8, depth: usize) {
        match op {
            // hstem, vstem: hints
            1 | 3 => {}
            4 => {
                let [dy] = self.args::<1>();
                self.move_to(0.0, dy);
            }
            5 => {
                let [dx, dy] = self.args::<2>();
                self.line_to(dx, dy);
            }
            6 => {
                let [dx] = self.args::<1>();
                self.line_to(dx, 0.0);
            }
            7 => {
                let [dy] = self.args::<1>();
                self.line_to(0.0, dy);
            }
            8 => {
                let d = self.args::<6>();
                self.curve_to(d);
            }
            9 => self.close(),
            10 => {
                let index = self.stack.pop().unwrap_or(-1.0);
                if index >= 0.0 {
                    if let Some(sub) = self.font.subrs.get(index as usize) {
                        let sub = sub.clone();
                        self.exec(&sub, depth + 1);
                    }
                }
                return;
            }
            // return: the subroutine's end (its operands stay)
            11 => return,
            13 => {
                let [sbx, _wx] = self.args::<2>();
                self.sbx = sbx;
                self.x = sbx;
                self.y = 0.0;
            }
            14 => {
                self.close();
                self.done = true;
            }
            21 => {
                let [dx, dy] = self.args::<2>();
                self.move_to(dx, dy);
            }
            22 => {
                let [dx] = self.args::<1>();
                self.move_to(dx, 0.0);
            }
            30 => {
                let [dy1, dx2, dy2, dx3] = self.args::<4>();
                self.curve_to([0.0, dy1, dx2, dy2, dx3, 0.0]);
            }
            31 => {
                let [dx1, dx2, dy2, dy3] = self.args::<4>();
                self.curve_to([dx1, 0.0, dx2, dy2, 0.0, dy3]);
            }
            _ => {}
        }
        self.stack.clear();
    }

    fn escape(&mut self, op: u8, depth: usize) {
        match op {
            // dotsection, vstem3, hstem3: hints
            0 | 1 | 2 => {}
            6 => {
                let [asb, adx, ady, bchar, achar] = self.args::<5>();
                let name = |code: f64| -> Option<&'static str> {
                    let c = code as usize;
                    if (32..256).contains(&c) {
                        Some(STANDARD[c - 32]).filter(|n| !n.is_empty())
                    } else {
                        None
                    }
                };
                let (base, accent) = (name(bchar), name(achar));
                self.stack.clear();
                if let Some(cs) = base.and_then(|n| self.font.charstrings.get(n)) {
                    let cs = cs.clone();
                    self.exec(&cs, depth + 1);
                }
                self.close();
                let base_sbx = self.sbx;
                if let Some(cs) = accent.and_then(|n| self.font.charstrings.get(n)) {
                    let cs = cs.clone();
                    self.done = false;
                    self.flex = None;
                    self.origin = (adx + base_sbx - asb, ady);
                    self.exec(&cs, depth + 1);
                    self.origin = (0.0, 0.0);
                }
                self.close();
                self.done = true;
                return;
            }
            7 => {
                let [sbx, sby, _wx, _wy] = self.args::<4>();
                self.sbx = sbx;
                self.x = sbx;
                self.y = sby;
            }
            12 => {
                let b = self.stack.pop().unwrap_or(1.0);
                let a = self.stack.pop().unwrap_or(0.0);
                self.stack.push(if b != 0.0 { a / b } else { 0.0 });
                return;
            }
            16 => {
                let other = self.stack.pop().unwrap_or(-1.0) as i64;
                let n = self.stack.pop().unwrap_or(0.0).max(0.0) as usize;
                let n = n.min(self.stack.len());
                let args: Vec<f64> = self.stack.split_off(self.stack.len() - n);
                match other {
                    1 => self.flex = Some(vec![]),
                    2 => {}
                    0 => {
                        let points = self.flex.take().unwrap_or_default();
                        if points.len() >= 7 {
                            self.ensure_open();
                            for curve in [&points[1..4], &points[4..7]] {
                                let (a, b, c) = (self.pt(curve[0].0, curve[0].1), self.pt(curve[1].0, curve[1].1), self.pt(curve[2].0, curve[2].1));
                                self.out.push_str(&format!("C{a} {b} {c}"));
                            }
                            self.x = points[6].0;
                            self.y = points[6].1;
                        }
                        // "pop pop setcurrentpoint" follows: the flex's end
                        let (ex, ey) = if args.len() >= 3 { (args[1], args[2]) } else { (self.x, self.y) };
                        self.ps = vec![ey, ex];
                    }
                    // hint replacement: "pop callsubr" calls subroutine 3 (it does nothing that draws)
                    3 => self.ps = vec![3.0],
                    _ => self.ps = args,
                }
                return;
            }
            17 => {
                let v = self.ps.pop().unwrap_or(0.0);
                self.stack.push(v);
                return;
            }
            33 => {
                let [x, y] = self.args::<2>();
                self.x = x;
                self.y = y;
            }
            _ => {}
        }
        self.stack.clear();
    }
}

/// Adobe StandardEncoding, codes 32..255 (seac's base and accent are named by it).
const STANDARD: [&str; 224] = [
    "space", "exclam", "quotedbl", "numbersign", "dollar", "percent", "ampersand", "quoteright", "parenleft",
    "parenright", "asterisk", "plus", "comma", "hyphen", "period", "slash", "zero", "one", "two", "three", "four",
    "five", "six", "seven", "eight", "nine", "colon", "semicolon", "less", "equal", "greater", "question", "at", "A",
    "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X",
    "Y", "Z", "bracketleft", "backslash", "bracketright", "asciicircum", "underscore", "quoteleft", "a", "b", "c",
    "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z",
    "braceleft", "bar", "braceright", "asciitilde", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "exclamdown", "cent", "sterling",
    "fraction", "yen", "florin", "section", "currency", "quotesingle", "quotedblleft", "guillemotleft",
    "guilsinglleft", "guilsinglright", "fi", "fl", "", "endash", "dagger", "daggerdbl", "periodcentered", "",
    "paragraph", "bullet", "quotesinglbase", "quotedblbase", "quotedblright", "guillemotright", "ellipsis",
    "perthousand", "", "questiondown", "", "grave", "acute", "circumflex", "tilde", "macron", "breve", "dotaccent",
    "dieresis", "", "ring", "cedilla", "", "hungarumlaut", "ogonek", "caron", "emdash", "", "", "", "", "", "", "",
    "", "", "", "", "", "", "", "", "", "AE", "", "ordfeminine", "", "", "", "", "Lslash", "Oslash", "OE",
    "ordmasculine", "", "", "", "", "", "ae", "", "", "", "dotlessi", "", "", "lslash", "oslash", "oe", "germandbls",
    "", "", "", "",
];

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn standard_encoding_places() {
        assert_eq!(STANDARD[b'A' as usize - 32], "A");
        assert_eq!(STANDARD[0xC1 - 32], "grave");
        assert_eq!(STANDARD[0xE1 - 32], "AE");
        assert_eq!(STANDARD[0xFB - 32], "germandbls");
    }

    #[test]
    fn encodings_read_names() {
        let enc = parse_encoding("% comment\n/Test [ /.notdef /A % x\n /B ] def\n");
        assert_eq!(enc[0], None);
        assert_eq!(enc[1].as_deref(), Some("A"));
        assert_eq!(enc[2].as_deref(), Some("B"));
        assert_eq!(enc.len(), 256);
    }
}
