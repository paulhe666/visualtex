//! DVI and XDV (XeTeX) files: read into pages of raw operations, written back after the pages were rewritten.
//! Only what the page rewriting needs is decoded; every operation's bytes are kept as they are.

use std::collections::BTreeMap;

pub const ID_DVI: u8 = 2;
pub const ID_XDV: u8 = 7;

#[derive(Clone, Debug)]
pub enum FontKind {
    /// a TFM font (or a virtual font): widths from its .tfm
    Tfm { checksum: u32, scale: i32, design: i32, name: String },
    /// an XDV native font (OpenType/TrueType file): glyph runs carry their own widths
    Native { size: i32, name: String },
}

#[derive(Clone, Debug)]
pub struct FontDef {
    pub kind: FontKind,
    /// the definition's bytes, opcode included
    pub raw: Vec<u8>,
}

#[derive(Clone, Debug)]
pub struct Page {
    pub counts: [i32; 10],
    /// operations between bop and eop
    pub body: Vec<u8>,
    /// font definitions (or nops) before this page's bop
    pub before: Vec<u8>,
}

#[derive(Clone, Debug)]
pub struct Dvi {
    pub id: u8,
    pub num: u32,
    pub den: u32,
    pub mag: u32,
    pub comment: Vec<u8>,
    pub pages: Vec<Page>,
    pub fonts: BTreeMap<u32, FontDef>,
    /// postamble font definitions (raw), in order
    pub post_fonts: Vec<u8>,
    pub max_v: u32,
    pub max_h: u32,
}

/// One operation of a page body.
#[derive(Debug, Clone)]
pub enum Op<'a> {
    SetChar(u32),
    PutChar(u32),
    SetRule(i32, i32),
    PutRule(i32, i32),
    Nop,
    Push,
    Pop,
    Right(i32),
    W0,
    W(i32),
    X0,
    X(i32),
    Down(i32),
    Y0,
    Y(i32),
    Z0,
    Z(i32),
    Fnt(u32),
    Xxx(&'a [u8]),
    FntDef(u32, FontKind),
    /// XDV glyph run (253) or text-and-glyphs (254): the advance
    Glyphs(i32),
    /// XDV picture (251): no movement
    Pic,
    /// pTeX direction
    Dir(u8),
}

#[derive(Debug)]
pub struct DviError(pub String);

impl std::fmt::Display for DviError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "DVI: {}", self.0)
    }
}

fn err<T>(msg: &str) -> Result<T, DviError> {
    Err(DviError(msg.to_string()))
}

pub struct Reader<'a> {
    pub data: &'a [u8],
    pub pos: usize,
}

impl<'a> Reader<'a> {
    pub fn new(data: &'a [u8]) -> Self {
        Reader { data, pos: 0 }
    }
    fn need(&self, n: usize) -> Result<(), DviError> {
        if self.pos + n > self.data.len() {
            err("truncated file")
        } else {
            Ok(())
        }
    }
    pub fn u(&mut self, n: usize) -> Result<u32, DviError> {
        self.need(n)?;
        let mut v: u32 = 0;
        for i in 0..n {
            v = (v << 8) | self.data[self.pos + i] as u32;
        }
        self.pos += n;
        Ok(v)
    }
    pub fn s(&mut self, n: usize) -> Result<i32, DviError> {
        let v = self.u(n)?;
        Ok(match n {
            1 => v as u8 as i8 as i32,
            2 => v as u16 as i16 as i32,
            3 => ((v << 8) as i32) >> 8,
            _ => v as i32,
        })
    }
    pub fn bytes(&mut self, n: usize) -> Result<&'a [u8], DviError> {
        self.need(n)?;
        let b = &self.data[self.pos..self.pos + n];
        self.pos += n;
        Ok(b)
    }

    /// The operation at the current position (inside a page body).
    pub fn op(&mut self) -> Result<Op<'a>, DviError> {
        self.need(1)?;
        let o = self.data[self.pos];
        self.pos += 1;
        Ok(match o {
            0..=127 => Op::SetChar(o as u32),
            128..=131 => Op::SetChar(self.u((o - 127) as usize)?),
            132 => {
                let a = self.s(4)?;
                Op::SetRule(a, self.s(4)?)
            }
            133..=136 => Op::PutChar(self.u((o - 132) as usize)?),
            137 => {
                let a = self.s(4)?;
                Op::PutRule(a, self.s(4)?)
            }
            138 => Op::Nop,
            141 => Op::Push,
            142 => Op::Pop,
            143..=146 => Op::Right(self.s((o - 142) as usize)?),
            147 => Op::W0,
            148..=151 => Op::W(self.s((o - 147) as usize)?),
            152 => Op::X0,
            153..=156 => Op::X(self.s((o - 152) as usize)?),
            157..=160 => Op::Down(self.s((o - 156) as usize)?),
            161 => Op::Y0,
            162..=165 => Op::Y(self.s((o - 161) as usize)?),
            166 => Op::Z0,
            167..=170 => Op::Z(self.s((o - 166) as usize)?),
            171..=234 => Op::Fnt((o - 171) as u32),
            235..=238 => Op::Fnt(self.u((o - 234) as usize)?),
            239..=242 => {
                let k = self.u((o - 238) as usize)? as usize;
                Op::Xxx(self.bytes(k)?)
            }
            243..=246 => {
                let k = self.u((o - 242) as usize)?;
                let checksum = self.u(4)?;
                let scale = self.s(4)?;
                let design = self.s(4)?;
                let a = self.u(1)? as usize;
                let l = self.u(1)? as usize;
                let name_bytes = self.bytes(a + l)?;
                let name = String::from_utf8_lossy(&name_bytes[a..]).to_string();
                Op::FntDef(k, FontKind::Tfm { checksum, scale, design, name })
            }
            252 => {
                let k = self.u(4)?;
                let size = self.s(4)?;
                let flags = self.u(2)?;
                let len = self.u(1)? as usize;
                let name = String::from_utf8_lossy(self.bytes(len)?).to_string();
                let _index = self.u(4)?;
                for bit in [0x0200u32, 0x1000, 0x2000, 0x4000] {
                    if flags & bit != 0 {
                        self.u(4)?;
                    }
                }
                Op::FntDef(k, FontKind::Native { size, name })
            }
            253 => {
                let w = self.s(4)?;
                let n = self.u(2)? as usize;
                self.bytes(n * 10)?;
                Op::Glyphs(w)
            }
            254 => {
                let l = self.u(2)? as usize;
                self.bytes(l * 2)?;
                let w = self.s(4)?;
                let n = self.u(2)? as usize;
                self.bytes(n * 10)?;
                Op::Glyphs(w)
            }
            251 => {
                self.u(1)?;
                self.bytes(24)?;
                self.u(2)?;
                let l = self.u(2)? as usize;
                self.bytes(l)?;
                Op::Pic
            }
            255 => Op::Dir(self.u(1)? as u8),
            _ => return err(&format!("unexpected opcode {o} at {}", self.pos - 1)),
        })
    }
}

impl Dvi {
    pub fn parse(data: &[u8]) -> Result<Dvi, DviError> {
        Self::parse_impl(data, false)
    }

    /// The complete pages of a file TeX is still writing (no postamble yet): the font definitions met so far
    /// stand in for the postamble's.
    pub fn parse_partial(data: &[u8]) -> Result<Dvi, DviError> {
        Self::parse_impl(data, true)
    }

    fn parse_impl(data: &[u8], partial: bool) -> Result<Dvi, DviError> {
        let mut r = Reader::new(data);
        if r.u(1)? != 247 {
            return err("not a DVI file");
        }
        let id = r.u(1)? as u8;
        let num = r.u(4)?;
        let den = r.u(4)?;
        let mag = r.u(4)?;
        let k = r.u(1)? as usize;
        let comment = r.bytes(k)?.to_vec();
        let mut dvi = Dvi { id, num, den, mag, comment, pages: vec![], fonts: BTreeMap::new(), post_fonts: vec![], max_v: 0, max_h: 0 };
        let mut before = Vec::new();
        loop {
            let start = r.pos;
            let Some(&o) = data.get(r.pos) else {
                if partial {
                    break;
                }
                return err("no postamble");
            };
            match o {
                138 => {
                    r.pos += 1;
                }
                243..=246 | 252 => {
                    let op = r.op()?;
                    if let Op::FntDef(k, kind) = op {
                        dvi.fonts.entry(k).or_insert(FontDef { kind, raw: data[start..r.pos].to_vec() });
                    }
                    before.extend_from_slice(&data[start..r.pos]);
                }
                139 if partial => {
                    // a page is taken only when its eop is there
                    let saved = r.pos;
                    match Self::page_at(data, &mut r, &mut dvi) {
                        Ok(Some(page)) => dvi.pages.push(Page { before: std::mem::take(&mut before), ..page }),
                        _ => {
                            let _ = saved;
                            break;
                        }
                    }
                }
                139 => {
                    r.pos += 1;
                    let mut counts = [0i32; 10];
                    for c in counts.iter_mut() {
                        *c = r.s(4)?;
                    }
                    r.u(4)?;
                    let body_start = r.pos;
                    loop {
                        if *data.get(r.pos).ok_or_else(|| DviError("page without eop".into()))? == 140 {
                            break;
                        }
                        let op_start = r.pos;
                        let op = r.op()?;
                        if let Op::FntDef(k, kind) = op {
                            dvi.fonts.entry(k).or_insert(FontDef { kind, raw: data[op_start..r.pos].to_vec() });
                        }
                    }
                    let body = data[body_start..r.pos].to_vec();
                    r.pos += 1; // eop
                    dvi.pages.push(Page { counts, body, before: std::mem::take(&mut before) });
                }
                248 if partial => break,
                248 => {
                    r.pos += 1;
                    r.u(4)?; // last bop
                    r.u(12)?; // num den mag
                    dvi.max_v = r.u(4)?;
                    dvi.max_h = r.u(4)?;
                    r.u(4)?; // max stack, total pages
                    let fonts_start = r.pos;
                    loop {
                        let o = *data.get(r.pos).ok_or_else(|| DviError("truncated postamble".into()))?;
                        if o == 249 {
                            break;
                        }
                        if o == 138 {
                            r.pos += 1;
                            continue;
                        }
                        let s = r.pos;
                        if let Op::FntDef(k, kind) = r.op()? {
                            dvi.fonts.entry(k).or_insert(FontDef { kind, raw: data[s..r.pos].to_vec() });
                        }
                    }
                    dvi.post_fonts = data[fonts_start..r.pos].to_vec();
                    break;
                }
                _ if partial => break,
                _ => return err(&format!("unexpected opcode {o} between pages")),
            }
        }
        if partial && dvi.post_fonts.is_empty() {
            for def in dvi.fonts.values() {
                dvi.post_fonts.extend_from_slice(&def.raw);
            }
        }
        Ok(dvi)
    }

    /// The page starting at r.pos (a bop), or None when it is not complete.
    fn page_at(data: &[u8], r: &mut Reader, dvi: &mut Dvi) -> Result<Option<Page>, DviError> {
        r.pos += 1;
        let mut counts = [0i32; 10];
        for c in counts.iter_mut() {
            *c = r.s(4)?;
        }
        r.u(4)?;
        let body_start = r.pos;
        let mut defs = vec![];
        loop {
            match data.get(r.pos) {
                None => return Ok(None),
                Some(140) => break,
                _ => {}
            }
            let op_start = r.pos;
            let op = r.op()?;
            if let Op::FntDef(k, kind) = op {
                defs.push((k, FontDef { kind, raw: data[op_start..r.pos].to_vec() }));
            }
        }
        let body = data[body_start..r.pos].to_vec();
        r.pos += 1;
        for (k, def) in defs {
            dvi.fonts.entry(k).or_insert(def);
        }
        Ok(Some(Page { counts, body, before: vec![] }))
    }

    /// The file with `bodies` as the pages' operations (one per page).
    pub fn write_with(&self, bodies: &[Vec<u8>]) -> Vec<u8> {
        let mut out = Vec::with_capacity(bodies.iter().map(|b| b.len() + 64).sum::<usize>() + 1024);
        out.push(247);
        out.push(self.id);
        put(&mut out, self.num, 4);
        put(&mut out, self.den, 4);
        put(&mut out, self.mag, 4);
        out.push(self.comment.len() as u8);
        out.extend_from_slice(&self.comment);
        let mut prev: i64 = -1;
        let mut max_stack = 0u32;
        for (page, body) in self.pages.iter().zip(bodies) {
            out.extend_from_slice(&page.before);
            let bop = out.len() as i64;
            out.push(139);
            for c in page.counts {
                put(&mut out, c as u32, 4);
            }
            put(&mut out, prev as i32 as u32, 4);
            prev = bop;
            out.extend_from_slice(body);
            out.push(140);
            max_stack = max_stack.max(stack_depth(body, self.id));
        }
        let post = out.len();
        out.push(248);
        put(&mut out, prev as i32 as u32, 4);
        put(&mut out, self.num, 4);
        put(&mut out, self.den, 4);
        put(&mut out, self.mag, 4);
        put(&mut out, self.max_v, 4);
        put(&mut out, self.max_h, 4);
        put(&mut out, max_stack.max(1), 2);
        put(&mut out, self.pages.len() as u32, 2);
        out.extend_from_slice(&self.post_fonts);
        out.push(249);
        put(&mut out, post as u32, 4);
        out.push(self.id);
        out.extend_from_slice(&[223, 223, 223, 223]);
        while out.len() % 4 != 0 {
            out.push(223);
        }
        out
    }

    /// Size of one DVI unit in big points.
    pub fn unit_bp(&self) -> f64 {
        self.num as f64 / self.den as f64 * self.mag as f64 / 1000.0 * 72.0 / 254000.0
    }
}

pub fn put(out: &mut Vec<u8>, v: u32, n: usize) {
    for i in (0..n).rev() {
        out.push((v >> (8 * i)) as u8);
    }
}

/// A special as DVI bytes (xxx1 or xxx4).
pub fn xxx(out: &mut Vec<u8>, text: &[u8]) {
    if text.len() < 256 {
        out.push(239);
        out.push(text.len() as u8);
    } else {
        out.push(242);
        put(out, text.len() as u32, 4);
    }
    out.extend_from_slice(text);
}

pub fn op_right(out: &mut Vec<u8>, v: i32) {
    out.push(146);
    put(out, v as u32, 4);
}
pub fn op_down(out: &mut Vec<u8>, v: i32) {
    out.push(160);
    put(out, v as u32, 4);
}
pub fn op_fnt(out: &mut Vec<u8>, k: u32) {
    out.push(238);
    put(out, k, 4);
}

fn stack_depth(body: &[u8], _id: u8) -> u32 {
    let mut r = Reader::new(body);
    let (mut d, mut m) = (0i32, 0i32);
    while r.pos < body.len() {
        match r.op() {
            Ok(Op::Push) => {
                d += 1;
                m = m.max(d);
            }
            Ok(Op::Pop) => d -= 1,
            Ok(_) => {}
            Err(_) => break,
        }
    }
    m.max(0) as u32
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample() -> Vec<u8> {
        // a two-page DVI: pre, a font, chars, a special, post
        let mut d = vec![247, 2];
        put(&mut d, 25400000, 4);
        put(&mut d, 473628672, 4);
        put(&mut d, 1000, 4);
        d.push(0);
        let mut fnt = vec![243, 0];
        put(&mut fnt, 0x1234, 4);
        put(&mut fnt, 655360, 4);
        put(&mut fnt, 655360, 4);
        fnt.push(0);
        fnt.push(5);
        fnt.extend_from_slice(b"cmr10");
        let mut prev: i64 = -1;
        for p in 0..2 {
            let bop = d.len() as i64;
            d.push(139);
            for c in 0..10 {
                put(&mut d, if c == 0 { p + 1 } else { 0 }, 4);
            }
            put(&mut d, prev as i32 as u32, 4);
            prev = bop;
            if p == 0 {
                d.extend_from_slice(&fnt);
            }
            d.push(171);
            d.extend_from_slice(&[141, 72, 105, 142]);
            xxx(&mut d, b"color push rgb 1 0 0");
            d.push(140);
        }
        let post = d.len();
        d.push(248);
        put(&mut d, prev as u32, 4);
        put(&mut d, 25400000, 4);
        put(&mut d, 473628672, 4);
        put(&mut d, 1000, 4);
        put(&mut d, 100, 4);
        put(&mut d, 100, 4);
        put(&mut d, 1, 2);
        put(&mut d, 2, 2);
        d.extend_from_slice(&fnt);
        d.push(249);
        put(&mut d, post as u32, 4);
        d.push(2);
        d.extend_from_slice(&[223; 4]);
        while d.len() % 4 != 0 {
            d.push(223);
        }
        d
    }

    #[test]
    fn parse_and_write_round_trip() {
        let data = sample();
        let dvi = Dvi::parse(&data).unwrap();
        assert_eq!(dvi.pages.len(), 2);
        assert_eq!(dvi.fonts.len(), 1);
        let bodies: Vec<Vec<u8>> = dvi.pages.iter().map(|p| p.body.clone()).collect();
        let out = dvi.write_with(&bodies);
        assert_eq!(out, data);
        let again = Dvi::parse(&out).unwrap();
        assert_eq!(again.pages[1].body, dvi.pages[1].body);
    }

    #[test]
    fn partial_files() {
        let data = sample();
        // cut inside the second page: one page, fonts kept for the postamble
        let cut = data.iter().rposition(|&b| b == 140).unwrap() - 3;
        let dvi = Dvi::parse_partial(&data[..cut]).unwrap();
        assert_eq!(dvi.pages.len(), 1);
        assert!(!dvi.post_fonts.is_empty());
        let out = dvi.write_with(&[dvi.pages[0].body.clone()]);
        assert_eq!(Dvi::parse(&out).unwrap().pages.len(), 1);
        assert_eq!(Dvi::parse_partial(&data).unwrap().pages.len(), 2);
    }

    #[test]
    fn ops_decode() {
        let data = sample();
        let dvi = Dvi::parse(&data).unwrap();
        let mut r = Reader::new(&dvi.pages[0].body);
        let mut specials = 0;
        while r.pos < dvi.pages[0].body.len() {
            if let Op::Xxx(t) = r.op().unwrap() {
                assert_eq!(t, b"color push rgb 1 0 0");
                specials += 1;
            }
        }
        assert_eq!(specials, 1);
        assert!((dvi.unit_bp() - 72.0 / 72.27 / 65536.0).abs() < 1e-12);
    }
}
